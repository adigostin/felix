
#include "pch.h"
#include "PackageTests.h"
#include "../FelixPackage/Z80Xml.h"
#include "../FelixPackageUi/resource.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(ProjectTests)
	{
		TEST_METHOD(GetItemsPutItems_WithFolders)
		{
			auto testPath = wil::str_concat_failfast<wil::unique_hglobal_string>(tempPath, L"GetItemsPutItems_WithFolders");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			com_ptr<IVsUIHierarchy> hier1;
			auto hr = MakeProjectNode (nullptr, testPath.get(), nullptr, 0, IID_PPV_ARGS(&hier1));
			Assert::IsTrue(SUCCEEDED(hr));
			auto close1 = wil::scope_exit([&hier1] { hier1->Close(); });

			wil::unique_variant folder1;
			hr = hier1->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &folder1);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier1->SetProperty (V_VSITEMID(&folder1), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"folder"));
			Assert::IsTrue(SUCCEEDED(hr));

			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };
			hr = hier1.try_query<IVsProject>()->AddItem(V_VSITEMID(&folder1), VSADDITEMOP_CLONEFILE, L"test.asm", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// ------------------------------------------------

			auto stream = com_ptr(SHCreateMemStream(nullptr, 0));
			hr = SaveToXml(hier1.try_query<IProjectNodeProperties>(), L"Temp", 0, stream);
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsHierarchy> hier2;
			hr = MakeProjectNode (nullptr, testPath.get(), nullptr, 0, IID_PPV_ARGS(&hier2));
			Assert::IsTrue(SUCCEEDED(hr));
			auto close2 = wil::scope_exit([&hier2] { hier2->Close(); });

			hr = stream->Seek({ 0 }, STREAM_SEEK_SET, nullptr);
			hr = LoadFromXml(hier2.try_query<IProjectNodeProperties>(), L"Temp", stream);
			Assert::IsTrue(SUCCEEDED(hr));

			// ---------------------

			auto pip2 = hier2.try_query<IParentNode>();
			Assert::IsNotNull(pip2->FirstChild());
			auto folder2 = wil::try_com_query_nothrow<IFolderNode>(pip2->FirstChild());
			Assert::IsNotNull(pip2->FirstChild());
			wil::unique_bstr folder2Name;
			hr = folder2.try_query<IFolderNodeProperties>()->get_Name(&folder2Name);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"folder", folder2Name.get());

			auto file2 = folder2.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(file2);
			wil::unique_bstr file2Path;
			hr = wil::try_com_query_nothrow<IFileNodeProperties>(file2)->get_Path(&file2Path);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"test.asm", file2Path.get());
		}

		TEST_METHOD(AddItemSort)
		{
			HRESULT hr;

			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"AddItemSort");

			com_ptr<IProjectNode> project;
			hr = sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&project));
			Assert::IsTrue(SUCCEEDED(hr));
			auto close = wil::scope_exit([&project] { project->AsHierarchy()->Close(); });

			auto proj = wil::try_com_query_failfast<IVsProject2>(project);
			auto hier = wil::try_com_query_failfast<IVsUIHierarchy>(project);
			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };
			hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"start.asm", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"lib.asm", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant more;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &more);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->SetProperty (V_VSITEMID(&more), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"More"));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant evenMore;
			hr = hier->ExecCommand (V_VSITEMID(&more), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &evenMore);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->SetProperty (V_VSITEMID(&evenMore), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"EvenMore"));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj->AddItem(V_VSITEMID(&evenMore), VSADDITEMOP_CLONEFILE, L"file.inc", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant generatedFiles;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &generatedFiles);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->SetProperty (V_VSITEMID(&generatedFiles), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"GeneratedFiles"));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj->AddItem(V_VSITEMID(&generatedFiles), VSADDITEMOP_CLONEFILE, L"preinclude.inc", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj->AddItem(V_VSITEMID(&generatedFiles), VSADDITEMOP_CLONEFILE, L"postinclude.inc", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			auto c = hier.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(c);
			auto genFilesFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(genFilesFolder.get());
			Assert::AreEqual(L"GeneratedFiles", GetProperty_String(hier, genFilesFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = genFilesFolder.try_query<IParentNode>()->FirstChild();
			auto postinc = wil::try_com_query_nothrow<IFileNode>(c);
			Assert::IsNotNull(postinc.get());
			Assert::AreEqual(0, _wcsicmp(L"postinclude.inc", GetProperty_String(hier, postinc->GetItemId(), VSHPROPID_SaveName).get()));

			auto preinc = wil::try_com_query_nothrow<IFileNode>(postinc->Next());
			Assert::IsNotNull(preinc.get());
			Assert::AreEqual(0, _wcsicmp(L"preinclude.inc", GetProperty_String(hier, preinc->GetItemId(), VSHPROPID_SaveName).get()));

			c = genFilesFolder->Next();
			Assert::IsNotNull(c);
			auto moreFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(moreFolder.get());
			Assert::AreEqual(L"More", GetProperty_String(hier, moreFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = moreFolder.try_query<IParentNode>()->FirstChild();
			auto evenMoreFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(evenMoreFolder.get());
			Assert::AreEqual(L"EvenMore", GetProperty_String(hier, evenMoreFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = evenMoreFolder.try_query<IParentNode>()->FirstChild();
			auto fileinc = wil::try_com_query_nothrow<IFileNode>(c);
			Assert::IsNotNull(fileinc.get());
			Assert::AreEqual(L"file.inc", GetProperty_String(hier, fileinc->GetItemId(), VSHPROPID_SaveName).get());

			auto libasm = wil::try_com_query_nothrow<IFileNode>(moreFolder->Next());
			Assert::IsNotNull(libasm.get());
			Assert::AreEqual(L"lib.asm", GetProperty_String(hier, libasm->GetItemId(), VSHPROPID_SaveName).get());

			auto startasm = wil::try_com_query_nothrow<IFileNode>(libasm->Next());
			Assert::IsNotNull(startasm.get());
			Assert::AreEqual(L"start.asm", GetProperty_String(hier, startasm->GetItemId(), VSHPROPID_SaveName).get());

			Assert::IsNull(startasm->Next());
		}

		TEST_METHOD(RenameFilePresentOnFileSystem)
		{
			static const char xml[] = ""
				"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
				"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
				"  <Configurations>\r\n"
				"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\">\r\n"
				"      <AssemblerProperties GeneratePrePostIncludeFiles=\"False\" />\r\n"
				"    </Configuration>\r\n"
				"  </Configurations>\r\n"
				"  <Items>\r\n"
				"    <File Path=\"file.asm\" BuildTool=\"Assembler\" />\r\n"
				"  </Items>\r\n"
				"</Z80Project>\r\n";

			auto s = SHCreateMemStream((BYTE*)xml, sizeof(xml) - 1);
			com_ptr<IStream> stream;
			stream.attach(s);

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto close = wil::scope_exit([&hier] { hier->Close(); });

			hr = LoadFromXml(hier.try_query<IProjectNodeProperties>(), ProjectElementName, stream);
			Assert::IsTrue(SUCCEEDED(hr));

			auto file = hier.try_query<IParentNode>()->FirstChild();

			wil::unique_bstr oldFullPath;
			hr = hier.try_query<IVsProject>()->GetMkDocument(file->GetItemId(), &oldFullPath);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_hfile handle (CreateFile(oldFullPath.get(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			Assert::IsTrue(handle.is_valid());
			handle.reset();

			wchar_t newFullPath[MAX_PATH];
			PathCombine(newFullPath, tempPath, L"new.asm");
			BOOL bres = DeleteFile(newFullPath);
			Assert::IsTrue(bres || GetLastError() == ERROR_FILE_NOT_FOUND);

			hr = hier->SetProperty(file->GetItemId(), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant newSaveName;
			hr = hier->GetProperty(file->GetItemId(), VSHPROPID_SaveName, &newSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"new.asm", newSaveName.bstrVal);

			Assert::IsTrue(PathFileExists(newFullPath));
			DeleteFile(newFullPath);
		}

		TEST_METHOD(RenameFileMissingOnFileSystem)
		{
		}

		TEST_METHOD(RenameFileMissingOnFileSystem_NewNameExists)
		{
		}

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
