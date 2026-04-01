
#include "pch.h"
#include "PackageTests.h"
#include "../FelixPackage/Z80Xml.h"
#include "../FelixPackageUi/resource.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(ProjectTests)
	{


		TEST_METHOD(RenameFileAndCheckSorted)
		{
			HRESULT hr;

			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"RenameFileAndCheckSorted\\");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));
			auto close = wil::scope_exit([&hier] { hier->Close(); });
			auto hierAsParent = hier.try_query<IParentNode>();

			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };

			// Add file named "B".
			hr = hier.try_query<IVsProject>()->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"B", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// Add file named "C".
			hr = hier.try_query<IVsProject>()->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"C", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// First one should be "B", second one should be "C"
			wil::unique_bstr name;
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild())->get_Path(&name);
			Assert::AreEqual(L"B", name.get());
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild()->Next())->get_Path(&name);
			Assert::AreEqual(L"C", name.get());

			// Rename second one to "A".
			hr = hier->SetProperty (hierAsParent->FirstChild()->Next()->GetItemId(), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"A"));
			Assert::IsTrue(SUCCEEDED(hr));

			// First one should be "A", second one should be "B"
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild())->get_Path(&name);
			Assert::AreEqual(L"A", name.get());
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild()->Next())->get_Path(&name);
			Assert::AreEqual(L"B", name.get());

			Assert::Fail(L"TODO: check notifications");
		}

		TEST_METHOD(DeleteItems_File)
		{
			HRESULT hr;
			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"DeleteItems_File");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));

			hr = hier.try_query<IVsProject>()->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file.asm", 1, (LPCOLESTR*)TemplatePath_EmptyFile.addressof(), nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant fileItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_bstr fileMk;
			hr = hier.try_query<IVsProject>()->GetMkDocument(V_VSITEMID(&fileItemId), &fileMk);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(fileMk.get()));

			hr = hier.try_query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, (VSITEMID*)&V_VSITEMID(&fileItemId), DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&fileItemId));
			Assert::IsFalse(PathFileExists(fileMk.get()));
			hier->Close();
		}

		TEST_METHOD(DeleteItems_Folder)
		{
			HRESULT hr;
			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"DeleteItems_Folder");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));

			// Add folder and check directory exists in file system.
			wil::unique_variant folderItemId;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &folderItemId);
			wil::unique_variant folderName;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_SaveName, &folderName);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BSTR, folderName.vt);
			auto directoryFullPath = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s\\%s", testPath.get(), folderName.bstrVal);
			Assert::IsTrue(PathFileExists(directoryFullPath.get()));

			// Add file in folder and check it exists in file system.
			hr = hier.try_query<IVsProject>()->AddItem (V_VSITEMID(&folderItemId), VSADDITEMOP_CLONEFILE, L"file.asm", 1, (LPCOLESTR*)TemplatePath_EmptyFile.addressof(), nullptr, nullptr);
			wil::unique_variant fileItemId;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_FirstChild, &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_bstr fileMk;
			hr = hier.try_query<IVsProject>()->GetMkDocument(V_VSITEMID(&fileItemId), &fileMk);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(fileMk.get()));

			// Delete folder.
			hr = hier.try_query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, (VSITEMID*)&V_VSITEMID(&folderItemId), DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			// Check there's no node in hierarchy
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&folderItemId));

			// Check there's no file or directory in the file system.
			Assert::IsFalse(PathFileExists(fileMk.get()));
			Assert::IsFalse(PathFileExists(directoryFullPath.get()));
			hier->Close();
		}

		TEST_METHOD(DeleteItems_FolderAndOneOfTwoMemberFiles)
		{
		}

		TEST_METHOD(ProjectDirPropertyEndsWithBackslash)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant projDir;
			hr = hier->GetProperty (VSITEMID_ROOT, VSHPROPID_ProjectDir, &projDir);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BSTR, projDir.vt);
			Assert::AreEqual(L'\\', projDir.bstrVal[SysStringLen(projDir.bstrVal) - 1]);
		}

		TEST_METHOD(Macro_CircularReference)
		{
			com_ptr<IVsSolution> sol;
			auto hr = serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));
			Assert::IsTrue(SUCCEEDED(hr));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"Macro_CircularReference");

			com_ptr<IProjectNode> proj;
			hr = sol->CreateProject (FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&proj));
			Assert::IsTrue(SUCCEEDED(hr));

			auto cfg = AddDebugProjectConfig(proj->AsHierarchy());

			auto val = wil::make_bstr_failfast(L"%OUTPUT_NAME%");
			hr = cfg->GeneralProps()->put_OutputName(val.get());
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_bstr fn;
			hr = cfg->GeneralProps()->get_OutputFilename(&fn); // this tries to resolve the macro from above
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(wcsstr(fn.get(), L"%OUTPUT_NAME%"));

			proj->AsHierarchy()->Close();
		}
	};
}
