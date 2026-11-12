
#include "pch.h"
#include "shared/com.h"
#include "UITests.h"

namespace UITests
{
	extern void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount);

	TEST_CLASS(MiscTests)
	{
		static inline wil::unique_process_heap_string classPath;

		TEST_CLASS_INITIALIZE(MiscTestsInit)
		{
			classPath = str_concat(tempPath, L"MiscTest");
			if (PathFileExists(classPath.get()))
			{
				// Clear leftovers before running the class; Visual Studio can leave per-solution .vs metadata behind.
				ClearDirectoryContents(classPath.get());
			}
			else
				Assert::IsTrue(CreateDirectory(classPath.get(), nullptr));
		}

		TEST_CLASS_CLEANUP(MiscTestsCleanup)
		{
			if (classPath)
			{
				// Best effort: Visual Studio may still hold files in .vs open after the last test.
				TryRemoveDirectoryTree(classPath.get());
				classPath.reset();
			}
		}

		TEST_METHOD(OpenSpecificEditor)
		{
			HRESULT hr;
			TD td(classPath.get(), L"OpenSpecificEditor", TemplatePath_TwoConfigsOneFile.get());
			// At some point between VS 17.10 and 17.14, the VS implementation of IVsUIShellOpenDocument::OpenStandardEditor
			// started trying to open our .asm files not with "Source Code (Text) Editor", but with
			// "Common Language Editor Supporting TextMate Bundles" (or at least this was the new default shown in Open With...).
			// The implementation started returning E_NOTIMPL in most cases.
			// The solution was to call OpenSpecificEditor with GUID_TextEditorFactory, instead of OpenStandardEditor.
			// This test verifies this solution. The test fails with the old code that calls OpenStandardEditor,
			// and succeeds with the new code that calls OpenSpecificEditor.

			VSITEMID itemid;
			hr = td.hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IVsWindowFrame> wf;
			hr = td.vsproj->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->Show();
			Assert::IsTrue(SUCCEEDED(hr));
			hr = td.sln->Close();
			Assert::IsTrue(SUCCEEDED(hr));

			hr = td.sln->Open(wil::make_bstr_failfast(td.slnFilePath.get()).get());
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<VxDTE::Projects> projects;
			hr = td.sln->get_Projects(&projects);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::com_ptr_failfast<VxDTE::Project> reopened;
			hr = projects->Item(wil::make_variant_bstr_failfast(L"proj.flx"), &reopened);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = reopened.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = reopened.query<IVsProject2>()->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf); // This would return E_NOTIMPL before the fix.
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(NavigateToErrorInFileInSubdir)
		{
			HRESULT hr;
			TD td(classPath.get(), L"NavigateToErrorInFileInSubdir", TemplatePath_TwoConfigsOneFile.get());
			// When the user double-clicks an error in the Error List window, VS goes out of its way to locate the file
			// and find the project it contains. To repro the bug, we need to put the project file in a subdirectory
			// (not next to the .sln file), and the offending file in a subdirectory of its own. In this case, when the user
			// double-clicks the error in the Error List window, VS used to be unable to locate the file due to bugs in
			// ProjectNode::ParseCanonicalName and ProjectNode::IsDocumentInProject (and ultimately in ParseCanonicalName).
			// This test verifies the fixes in these functions.

			long buildFailCount;
			BuildSolution(td.sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);

			auto file1Path = str_concat(td.projDir, L"\\subdir\\file1.asm");
			WriteFileCreateDirs (file1Path, "\t555555");

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.vsproj->AddItem (VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(file1Path.addressof()), NULL, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = td.proj->Save();
			Assert::IsTrue(SUCCEEDED(hr));

			BuildSolution(td.sln, &buildFailCount);
			Assert::IsTrue(buildFailCount >= 1);

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"View.ErrorList").get());
			Assert::IsTrue(SUCCEEDED(hr));
			wil::com_ptr_failfast<VxDTE::ToolWindows> toolWindows;
			hr = td.dte->get_ToolWindows(&toolWindows);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(toolWindows.get());
			wil::com_ptr_failfast<VxDTE::ErrorList> errorList;
			hr = toolWindows->get_ErrorList(&errorList);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(errorList.get());
			wil::com_ptr_failfast<VxDTE::ErrorItems> errorItems;
			hr = errorList->get_ErrorItems(&errorItems);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(errorItems.get());
			long errorCount;
			hr = errorItems->get_Count(&errorCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(errorCount >= 1);

			VARIANT v; v.vt = VT_I4; v.lVal = 1;
			com_ptr<VxDTE::ErrorItem> errorItem;
			hr = errorItems->Item(v, &errorItem);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = errorItem->Navigate(); // This call succeeds, even though VS couldn't open the file.
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<VxDTE::Document> doc;
			hr = td.dte->get_ActiveDocument(&doc);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(doc.get()); // This would fail before the fix.

			BOOL found = FALSE;
			VSDOCUMENTPRIORITY prio = { };
			VSITEMID itemid = { };
			hr = td.vsproj->IsDocumentInProject(L"subdir\\file1.asm", &found, &prio, &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = td.vsproj->IsDocumentInProject(L"subdir/..\\subdir/.\\file1.asm", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = td.vsproj->IsDocumentInProject(L"SubDir\\File1.ASM", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);

			// For the canonical name, we don't bother with tricky paths such as "subdir/..\\subdir/.\\file1.asm".
			hr = td.hier->ParseCanonicalName(L"subdir\\file1.asm", &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			hr = td.hier->ParseCanonicalName(L"SubDir\\File1.ASM", &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(GenPrePostInclude_OnlyActiveCfg)
		{
			HRESULT hr;
			TD td(classPath.get(), L"GenPrePostInclude_OnlyActiveCfg", TemplatePath_TwoConfigsOneFile.get());
			// Verify that the PreInclude.asm and PostInclude.asm are regenerated when editing the active configuration,
			// and that they are _not_ regenerated when editing an inactive configuration.

			// Let's not go through the configuration manager since we haven't implemented Project::get_ConfigurationManager yet.
			//Microsoft_VisualStudio_Interop::_SolutionPtr sln = sln0.get();
			//auto configs = sln->SolutionBuild->SolutionConfigurations;
			//Assert::IsTrue(configs->Count >= 2);
			//Assert::AreEqual<void*>(configs->Item[1].GetInterfacePtr(), sln->SolutionBuild->ActiveConfiguration.GetInterfacePtr());

			com_ptr<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// Make a change in the active configuration and verify that the Pre/PostInclude files are generated.
			auto genFilesPath = str_concat(td.projDir, L"\\GeneratedFiles");
			Assert::IsTrue(PathFileExists(genFilesPath.get()));
			RemoveDirectoryTree(genFilesPath.get());

			com_ptr<IProjectConfigProperties> props;
			hr = cfgs[0]->QueryInterface(IID_PPV_ARGS(&props));
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IProjectConfigAssemblerProperties> asmProps;
			hr = props->get_AssemblerProperties(&asmProps);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = asmProps->put_BaseAddress(1234);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(genFilesPath.get()));

			// Now make a change in the inactive configuration.
			RemoveDirectoryTree(genFilesPath.get());

			hr = cfgs[1]->QueryInterface(IID_PPV_ARGS(&props));
			Assert::IsTrue(SUCCEEDED(hr));
			hr = props->get_AssemblerProperties(&asmProps);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = asmProps->put_BaseAddress(1234);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(!PathFileExists(genFilesPath.get()));
		}

		TEST_METHOD(GenPrePostInclude_ChangeBuildTool)
		{
			TD td(classPath.get(), L"GenPrePostInclude_ChangeBuildTool", TemplatePath_TwoConfigsOneFile.get());
			// Verify that, with a single .asm file in the project, the PreInclude.asm and PostInclude.asm
			// are generated when changing the BuildTool of the .asm file to Assembler, and deleted when
			// changing the BuildTool to something else.

			// The pre/post-include files should have been generated on project creation.
			auto genFilesPath = str_concat(td.projDir, L"\\GeneratedFiles");
			Assert::IsTrue(PathFileExists(genFilesPath.get()));

			VSITEMID itemid;
			auto hr = td.hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant filevar;
			td.hier->GetProperty(itemid, VSHPROPID_BrowseObject, &filevar);
			auto fileProps = wil::com_query_failfast<IFileNodeProperties>(filevar.pdispVal);
			fileProps->put_BuildTool(BuildToolKind::None);
			Assert::IsTrue(!PathFileExists(genFilesPath.get()));

			fileProps->put_BuildTool(BuildToolKind::Assembler);
			Assert::IsTrue(PathFileExists(genFilesPath.get()));
		}
		TEST_METHOD(RemoveFileClosesEditor)
		{
			HRESULT hr;
			TD td(classPath.get(), L"RemoveFileClosesEditor", TemplatePath_TwoConfigsOneFile.get());
			VSITEMID itemId;
			hr = td.hier->ParseCanonicalName(L"file.asm", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IVsWindowFrame> wf;
			hr = td.vsproj->OpenItem(itemId, LOGVIEWID_Code, nullptr, &wf);
			wf->Show();
			hr = wf->IsVisible();
			Assert::AreEqual (S_OK, hr);

			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &itemId, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->IsVisible();
			Assert::AreNotEqual (S_OK, hr);
		}

		TEST_METHOD(RemoveFolderAndFile_FolderFirstInList)
		{
			TD td(classPath.get(), L"RemoveFolderAndFile_FolderFirstInList", TemplatePath_TwoConfigsOneFile.get());
			wil::unique_variant folderItemId;
			auto hr = td.hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folderItemId);
			Assert::AreEqual (S_OK, hr);

			VSITEMID fileItemId;
			hr = td.hier->ParseCanonicalName(L"file.asm", &fileItemId);
			Assert::AreEqual (S_OK, hr);

			VSITEMID itemIds[2] = { V_UI4(&folderItemId), fileItemId };
			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(2, DELITEMOP_DeleteFromStorage, itemIds, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(RemoveFolderClosesEditors)
		{
			TD td(classPath.get(), L"RemoveFolderClosesEditors", TemplatePath_TwoConfigsOneFile.get());
			wil::unique_variant folderItemId;
			auto hr = td.hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folderItemId);
			Assert::AreEqual (S_OK, hr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.vsproj->AddItem (V_UI4(&folderItemId), oper, L"file2.asm", 1, TemplateEmptyFile, NULL, &addResult);
			Assert::AreEqual (S_OK, hr);

			VSITEMID itemId;
			hr = td.hier->ParseCanonicalName(L"NewFolder1\\file2.asm", &itemId);
			Assert::AreEqual (S_OK, hr);
			com_ptr<IVsWindowFrame> wf;
			hr = td.vsproj->OpenItem(itemId, LOGVIEWID_Code, nullptr, &wf);
			wf->Show();
			hr = wf->IsVisible();
			Assert::AreEqual (S_OK, hr);

			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &V_UI4(&folderItemId), DHO_SUPPRESS_UI);
			Assert::AreEqual (S_OK, hr);
			hr = wf->IsVisible();
			Assert::AreNotEqual (S_OK, hr);
		}

		TEST_METHOD(AddNewFile_SameNameAsFileOutsideProjectDir)
		{
			HRESULT hr;
			TD td(classPath.get(), L"AddNewFile_SameNameAsFileOutsideProjectDir", TemplatePath_TwoConfigsOneFile.get());
			const wchar_t file1RelPath[] = L"test.asm";
			auto file1Path = str_concat(td.projDir, L"\\", file1RelPath);
			WriteFileCreateDirs (file1Path, "; comment");
			const wchar_t file2RelPath[] = L"..\\test.asm";
			auto file2Path = str_concat(td.projDir, L"\\", file2RelPath);
			WriteFileCreateDirs (file2Path, "; comment");
			wchar_t file3Path[] = L"D:\\FelixTest\\test.asm";
			WriteFileCreateDirs (file3Path, "; comment");

			LPCOLESTR filesToOpen[] = { file1Path.get(), file2Path.get(), file3Path };
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.vsproj->AddItem(VSITEMID_ROOT, oper, L"", 3, filesToOpen, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Adding them again should fail.
			for (const wchar_t* fileToOpen : filesToOpen)
			{
				hr = td.vsproj->AddItem(VSITEMID_ROOT, oper, L"", 1, &fileToOpen, nullptr, &addResult);
				Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), hr);
			}

			VSITEMID itemids[3];
			for (int i = 0; i < 3; i++)
			{
				BOOL found;
				VSDOCUMENTPRIORITY prio;
				hr = td.vsproj->IsDocumentInProject(filesToOpen[i], &found, &prio, &itemids[i]);
				Assert::IsTrue(SUCCEEDED(hr) && found);
			}

			wil::unique_bstr canonicalNames[3];
			for (int i = 0; i < 3; i++)
			{
				hr = td.hier->GetCanonicalName(itemids[i], &canonicalNames[i]);
				Assert::IsTrue(SUCCEEDED(hr));
			}

			for (int i = 0; i < 3; i++)
			{
				VSITEMID itemidx;
				hr = td.hier->ParseCanonicalName(canonicalNames[i].get(), &itemidx);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual(itemids[i], itemidx);
			}
		}

		TEST_METHOD(CloneFile_SameNameAsFileOutsideProjectDir)
		{
			HRESULT hr;
			TD td(classPath.get(), L"CloneFile_SameNameAsFileOutsideProjectDir", TemplatePath_TwoConfigsOneFile.get());
			auto templatePath = str_concat(td.testDir, L"\\template.asm");
			WriteFileCreateDirs (templatePath, "; comment");

			static const wchar_t file1RelPath[] = L"..\\test.asm";
			auto file1Path = str_concat(td.projDir, L"\\", file1RelPath);
			WriteFileCreateDirs(file1Path, "; comment");
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.vsproj->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(file1Path.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.vsproj->AddItem(VSITEMID_ROOT, oper, L"test.asm", 1, const_cast<LPCOLESTR*>(templatePath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(CloneFileToFolderMissingOnDisk)
		{
			HRESULT hr;
			TD td(classPath.get(), L"CloneFileToFolderMissingOnDisk", TemplatePath_TwoConfigsOneFile.get());
			auto templatePath = str_concat(td.testDir, L"\\template.asm");
			WriteFileCreateDirs (templatePath, "; comment");

			wil::unique_variant folderItemId;
			hr = td.hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));

			// Delete the folder on disk, then try to clone a file into it. We must at least not crash.

			wil::unique_bstr folderMk;
			hr = td.vsproj->GetMkDocument(V_VSITEMID(&folderItemId), &folderMk);
			Assert::IsTrue(SUCCEEDED(hr));

			if (PathFileExists(folderMk.get()))
			{
				BOOL bres = RemoveDirectory(folderMk.get());
				Assert::IsTrue(bres);
			}

			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			VSADDRESULT addResult;
			hr = td.vsproj->AddItem(V_VSITEMID(&folderItemId), oper, L"file.asm", 1, const_cast<LPCOLESTR*>(templatePath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		static inline const char OldXmlWithGeneratedFilesFolder[] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
			"<Z80Project Guid=\"{4839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
			"<Configurations>\r\n"
			"  <Configuration ConfigName=\"SnaOutput\" PlatformName=\"ZX Spectrum 48K\" />\r\n"
			"</Configurations>\r\n"
			"<Items>\r\n"
			"  <Folder Name=\"GeneratedFiles\">\r\n"
			"    <Items>\r\n"
			"      <File Path=\"PostInclude.asm\" IsGenerated=\"1\" BuildTool=\"Assembler\" />\r\n"
			"      <File Path=\"PreInclude.asm\" IsGenerated=\"1\" BuildTool=\"Assembler\" />\r\n"
			"    </Items>\r\n"
			"  </Folder>\r\n"
			"</Items>\r\n"
			"</Z80Project>\r\n";

		TEST_METHOD(GenPrePostInclude_LoadOldXml)
		{
			HRESULT hr;
			TD td(classPath.get(), L"GenPrePostInclude_LoadOldXml", nullptr);

			auto projPath = str_concat(td.projDir, L"\\oldproj.flx");
			WriteFileCreateDirs(projPath, OldXmlWithGeneratedFilesFolder);

			wil::com_ptr_failfast<VxDTE::Project> proj;
			hr = td.sln->AddFromFile (wil::make_bstr_failfast(projPath.get()).get(), VARIANT_TRUE, &proj);
			Assert::IsTrue(SUCCEEDED(hr));

			// The folder named GeneratedFiles should have been ignored while loading from XML,
			// and nothing should have been generated since there are no other files with
			// BuildTool==Assembler in the XML.
			wil::unique_variant var;
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &var);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, var.vt);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&var));
			auto genFilesPath = str_concat(td.projDir, L"\\GeneratedFiles");
			Assert::IsFalse(PathFileExists(genFilesPath.get()));

			// And the project should not be dirty just after loading from XML.
			auto pff = proj.query<IPersistFileFormat>();
			BOOL dirty;
			hr = pff->IsDirty(&dirty); Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsFalse(dirty);

			// Now if we add an .asm file, we should have the generated files.
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"file.asm", 1, TemplateEmptyFile, NULL, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &var);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, var.vt);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&var));
			Assert::IsTrue(PathFileExists(genFilesPath.get()));

			// The project should now be dirty. We used to have a bug that triggered a save upon generation.
			hr = pff->IsDirty(&dirty); Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(dirty);
		}

		TEST_METHOD(GenPrePostInclude_GeneratedFilesNotSaved)
		{
			TD td(classPath.get(), L"GenPrePostInclude_GeneratedFilesNotSaved", TemplatePath_TwoConfigsOneFile.get());
			auto hr = td.proj->Save(nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			auto content = ReadAllText(td.projFilePath.get());
			Assert::IsNull(strstr(content.get(), "GeneratedFiles"));
		}

		TEST_METHOD(GenPrePostInclude_AddFirstAsm_RemoveLastAsm)
		{
			HRESULT hr;
			TD td(classPath.get(), L"GenPrePostInclude_AddFirstAsm_RemoveLastAsm", TemplatePath_TwoConfigsOneFile.get());

			// We have a file in the project with BuildTool==Assembler. We should have a GeneratedFiles folder.
			wil::unique_variant firstChild;
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Remove the only file
			VSITEMID fileItemId;
			hr = td.hier->ParseCanonicalName(L"file.asm", &fileItemId);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_RemoveFromProject, &fileItemId, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);

			// Project should be empty (no file and no GeneratedFiles folder)
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&firstChild));

			// Add the file back.
			auto filePath = str_concat(td.projDir, L"\\file.asm");
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.vsproj->AddItem (VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(filePath.addressof()), nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			// First child should be "GeneratedFiles"
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));
		}

		TEST_METHOD(GenPrePostInclude_EnableDisableGeneration)
		{
			HRESULT hr;
			TD td(classPath.get(), L"GenPrePostInclude_EnableDisableGeneration", TemplatePath_TwoConfigsOneFile.get());

			// Assume first config is the active config.
			wil::com_ptr_failfast<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// We have a file in the project with BuildTool==Assembler. We should have a GeneratedFiles folder.
			wil::unique_variant firstChild;
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Disable generation. The folder should be gone.
			com_ptr<IProjectConfigAssemblerProperties> asmProps;
			cfgs[0].query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Enable again generation. Folder should be back.
			hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Disable generation in the inactive config. Folder should remain.
			cfgs[1].query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));
		}

		TEST_METHOD(GenPrePostInclude_SwitchActiveConfig)
		{
			HRESULT hr;
			TD td(classPath.get(), L"GenPrePostInclude_SwitchActiveConfig", TemplatePath_TwoConfigsOneFile.get());

			wil::com_ptr_failfast<VxDTE::SolutionConfigurations> slnConfigs;
			hr = td.slnBuild->get_SolutionConfigurations(&slnConfigs);
			Assert::AreEqual(S_OK, hr);

			long configCount;
			hr = slnConfigs->get_Count(&configCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(2l, configCount);

			wil::com_ptr_failfast<VxDTE::SolutionConfiguration> slnConfig0, slnConfig1, activeSlnConfig;
			VARIANT i; i.vt = VT_INT; i.intVal = 1;
			hr = slnConfigs->Item(i, &slnConfig0);
			Assert::AreEqual(S_OK, hr);
			i.intVal = 2;
			hr = slnConfigs->Item(i, &slnConfig1);
			Assert::AreEqual(S_OK, hr);
			hr = td.slnBuild->get_ActiveConfiguration(&activeSlnConfig);
			Assert::AreEqual(S_OK, hr);
			Assert::IsTrue(activeSlnConfig == slnConfig0);

			wil::com_ptr_failfast<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// Generation is by default enabled. We should have a GeneratedFiles folder.
			wil::unique_variant firstChild;
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Disable generation in the second config. We should still have a GeneratedFiles folder.
			wil::com_ptr_failfast<IProjectConfigAssemblerProperties> config1AsmProps;
			hr = cfgs[1].query<IProjectConfigProperties>()->get_AssemblerProperties(&config1AsmProps);
			Assert::AreEqual(S_OK, hr);
			hr = config1AsmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Switch the active config to the one with generation disabled. We should no longer have a GeneratedFiles folder.
			hr = slnConfig1->Activate();
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Switch the active config back. We should have a GeneratedFiles folder again.
			hr = slnConfig0->Activate();
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));
		}

		TEST_METHOD(GenPrePostInclude_SwitchPlatform)
		{
			HRESULT hr;
			TD td(classPath.get(), L"GenPrePostInclude_UpdatePlatform", nullptr);

			auto projPath = str_concat(td.projDir, L"\\proj.flx");
			WriteFileCreateDirs (projPath,
				"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
				"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
				"  <Configurations>\r\n"
				"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />\r\n"
				"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 128K\" />\r\n"
				"  </Configurations>\r\n"
				"  <Items>\r\n"
				"    <File Path=\"file.asm\" BuildTool=\"Assembler\" />\r\n"
				"  </Items>\r\n"
				"</Z80Project>\r\n");

			wil::com_ptr_failfast<VxDTE::Project> proj;
			hr = td.sln->AddFromFile(wil::make_bstr_failfast(projPath.get()).get(), VARIANT_TRUE, &proj);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::SolutionConfigurations> slnConfigs;
			hr = td.slnBuild->get_SolutionConfigurations(&slnConfigs);
			Assert::AreEqual(S_OK, hr);
			long configCount;
			hr = slnConfigs->get_Count(&configCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(2l, configCount);

			wil::com_ptr_failfast<VxDTE::SolutionConfiguration> config48K, config128K;
			for (long index = 1; index <= configCount; ++index)
			{
				VARIANT i; i.vt = VT_INT; i.intVal = index;
				wil::com_ptr_failfast<VxDTE::SolutionConfiguration> slnConfig;
				hr = slnConfigs->Item(i, &slnConfig);
				Assert::AreEqual(S_OK, hr);
				auto slnConfig2 = wil::com_query_failfast<VxDTE::SolutionConfiguration2>(slnConfig);
				wil::unique_bstr platformName;
				hr = slnConfig2->get_PlatformName(&platformName);
				Assert::AreEqual(S_OK, hr);
				if (wcscmp(platformName.get(), L"ZX Spectrum 48K") == 0)
					config48K = slnConfig;
				else if (wcscmp(platformName.get(), L"ZX Spectrum 128K") == 0)
					config128K = slnConfig;
			}
			Assert::IsNotNull(config48K.get());
			Assert::IsNotNull(config128K.get());

			wil::unique_bstr preincludePath;
			hr = proj.query<IVsProject>()->GetMkDocument(VSITEMID_PREINCLUDE, &preincludePath);
			Assert::AreEqual(S_OK, hr);

			hr = config48K->Activate();
			Assert::AreEqual(S_OK, hr);
			auto content = ReadAllText(preincludePath.get());
			Assert::IsNotNull(strstr(content.get(), "DEVICE ZXSPECTRUM48"));

			hr = config128K->Activate();
			Assert::AreEqual(S_OK, hr);
			content = ReadAllText(preincludePath.get());
			Assert::IsNotNull(strstr(content.get(), "DEVICE ZXSPECTRUM128"));
		}

		TEST_METHOD(Macro_CircularReference)
		{
			HRESULT hr;
			TD td(classPath.get(), L"Macro_CircularReference", TemplatePath_TwoConfigsOneFile.get());

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider2>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<IProjectConfigGeneralProperties> generalProps;
			hr = cfg.query<IProjectConfigProperties>()->get_GeneralProperties(&generalProps);
			Assert::AreEqual(S_OK, hr);

			hr = generalProps->put_OutputName(wil::make_bstr_failfast(L"%OUTPUT_NAME%").get());
			Assert::AreEqual(S_OK, hr);

			wil::unique_bstr fn;
			hr = generalProps->get_OutputFilename(&fn); // this tries to resolve the macro from above
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(wcsstr(fn.get(), L"%OUTPUT_NAME%"));
		}
	};
}
