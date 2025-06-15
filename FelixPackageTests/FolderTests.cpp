
#include "pch.h"
#include "shared/com.h"
#include "../FelixPackage/FelixPackage.h"
#include "../FelixPackage/Z80Xml.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(FolderTests)
	{
	public:
		TEST_METHOD(AddFolderCommand)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant expandable;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_FALSE, expandable.boolVal);

			bool expandableChanged = false;
			auto propChanged = [&expandableChanged](VSITEMID itemid, VSHPROPID propid, DWORD flags)
				{
					if (itemid == VSITEMID_ROOT && propid == VSHPROPID_Expandable)
						expandableChanged = true;
				};

			auto sink = MakeMockHierarchyEventSink(std::move(propChanged));
			VSCOOKIE hierEventsCookie;
			hr = hier->AdviseHierarchyEvents(sink, &hierEventsCookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([hier=hier.get(), hierEventsCookie]() { hier->UnadviseHierarchyEvents(hierEventsCookie); });

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant childDisp;
			hr = hier->GetProperty(V_VSITEMID(&child), VSHPROPID_BrowseObject, &childDisp);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IFolderNodeProperties> folderProps;
			hr = childDisp.pdispVal->QueryInterface(IID_PPV_ARGS(&folderProps));
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsTrue(expandableChanged);

			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_TRUE, expandable.boolVal);
		}

		TEST_METHOD(AddSubFolderCommand)
		{
			HRESULT hr;

			wchar_t projDir[MAX_PATH];
			swprintf_s(projDir, L"%s\\AddSubFolderCommand\0", &tempPath[0]);
			if (PathFileExists(projDir))
			{
				SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = projDir, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
				int ires = SHFileOperation(&file_op);
				Assert::AreEqual(0, ires);
			}
			BOOL bres = CreateDirectory(projDir, 0);
			Assert::IsTrue(bres);

			com_ptr<IVsUIHierarchy> hier;
			hr = MakeProjectNode (nullptr, projDir, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant folderItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant folderDisp;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_BrowseObject, &folderDisp);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant folderSaveName;
			hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_SaveName, &folderSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hlocal_string folderPath;
			hr = wil::str_concat_nothrow (folderPath, projDir, L"\\", folderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(folderPath.get()));

			wil::unique_variant expandable;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_FALSE, expandable.boolVal);

			bool expandableChanged = false;
			auto propChanged = [&expandableChanged, id=V_VSITEMID(&folderItemId)](VSITEMID itemid, VSHPROPID propid, DWORD flags)
				{
					if (itemid == id && propid == VSHPROPID_Expandable)
						expandableChanged = true;
				};
			auto sink = MakeMockHierarchyEventSink(std::move(propChanged));
			VSCOOKIE hierEventsCookie;
			hr = hier->AdviseHierarchyEvents(sink, &hierEventsCookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([hier=hier.get(), hierEventsCookie]() { hier->UnadviseHierarchyEvents(hierEventsCookie); });

			hr = hier->ExecCommand(V_VSITEMID(&folderItemId), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant subFolderItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &subFolderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant subFolderDisp;
			hr = hier->GetProperty(V_VSITEMID(&subFolderItemId), VSHPROPID_BrowseObject, &subFolderDisp);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IFolderNodeProperties> folderProps;
			hr = subFolderDisp.pdispVal->QueryInterface(IID_PPV_ARGS(&folderProps));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant subFolderSaveName;
			hier->GetProperty(V_VSITEMID(&subFolderItemId), VSHPROPID_SaveName, &subFolderSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hlocal_string subFolderPath;
			hr = wil::str_concat_nothrow (subFolderPath, folderPath, L"\\", subFolderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(subFolderPath.get()));

			Assert::IsTrue(expandableChanged);

			hr = hier->GetProperty(V_VSITEMID(&subFolderItemId), VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_TRUE, expandable.boolVal);
		}

		TEST_METHOD(AddFolderCommand_FolderExistsInHier)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(AddFolderCommand_FolderExistsOnDisk)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);

			com_ptr<IVsHierarchyDeleteHandler3> dh;
			hr = hier->QueryInterface(IID_PPV_ARGS(&dh));
			Assert::IsTrue(SUCCEEDED(hr));
			dh->DeleteItems(1, DELITEMOP_RemoveFromProject, (VSITEMID*)&child.lVal, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(AddSubFolder_FolderMissingOnDisk)
		{
			com_ptr<IVsSolution> sol;
			auto hr = serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));
			Assert::IsTrue(SUCCEEDED(hr));

			static const wchar_t ProjFileName[] = L"TestProject.flx";
			com_ptr<IVsUIHierarchy> hier;
			hr = sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), tempPath, ProjFileName, CPF_CLONEFILE, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant folderItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant folderSaveName;
			hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_SaveName, &folderSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hlocal_string folderPath;
			hr = wil::str_concat_nothrow (folderPath, tempPath, L"\\", folderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(folderPath.get()));
			BOOL bres = RemoveDirectoryW(folderPath.get());
			Assert::IsTrue(bres);

			hr = hier->ExecCommand(V_VSITEMID(&folderItemId), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(RenameFolderAndCheckSorted)
		{
			HRESULT hr;

			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"RenameFolderAndCheckSorted");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));

			// Add folder named "NewFolder1".
			wil::unique_variant one;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &one);
			Assert::IsTrue(SUCCEEDED(hr));

			// Rename it to "B".
			hr = hier->SetProperty (V_VSITEMID(&one), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"B"));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_process_heap_string pathToB;
			wil::str_concat_nothrow(pathToB, testPath, L"\\B");
			Assert::IsTrue(PathFileExists(pathToB.get()));

			// Add folder named "NewFolder1".
			wil::unique_variant other;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &other);
			Assert::IsTrue(SUCCEEDED(hr));

			// First one should be "B", second one should be "NewFolder1"
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&one), hier.try_query<IParentNode>()->FirstChild()->GetItemId());
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&other), hier.try_query<IParentNode>()->FirstChild()->Next()->GetItemId());

			// Rename the second one to "A".
			hr = hier->SetProperty (V_VSITEMID(&other), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"A"));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_process_heap_string pathToA;
			wil::str_concat_nothrow(pathToA, testPath, L"\\A");
			Assert::IsTrue(PathFileExists(pathToA.get()));

			// First one should be "A", second one should be "B"
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&other), hier.try_query<IParentNode>()->FirstChild()->GetItemId());
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&one), hier.try_query<IParentNode>()->FirstChild()->Next()->GetItemId());
		}

		TEST_METHOD(CreateFolderSameNameDifferentCase)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IFolderNode> folder1;
			hr = GetOrCreateChildFolder (hier.try_query<IParentNode>(), L"ABC", true, &folder1);
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IFolderNode> folder2;
			hr = GetOrCreateChildFolder (hier.try_query<IParentNode>(), L"abc", true, &folder2);
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsTrue(folder1 == folder2);
		}
	};
}

