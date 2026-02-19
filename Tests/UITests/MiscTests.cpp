
#include "pch.h"
#include "shared/com.h"
#include "../TestsCommon.h"
#include "../FelixPackage/dispids.h"

#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>
namespace VxDTE
{
	#include <dte80.h>
	#include <dte90.h>
}

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

static constexpr VSITEMID VSITEMID_GENFILES = 50;

namespace UITests
{
	extern wil::com_ptr_failfast<VxDTE::_DTE> dte;
	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount);

	TEST_CLASS(MiscTests)
	{
		wil::unique_process_heap_string testPath;
		wil::unique_process_heap_string slnFilePath;
		wil::unique_process_heap_string projPath;
		wil::unique_process_heap_string projFilePath;
		wil::com_ptr_failfast<VxDTE::_Solution> sln;
		wil::com_ptr_failfast<VxDTE::Project> proj;

		TEST_METHOD_INITIALIZE(MiscTestInit)
		{
			testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"MiscTest");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			std::tie(sln, proj) = CreateSolutionAndProject(testPath.get(), L"test", L"proj");
			slnFilePath = CombinePath(testPath.get(), L"test.sln");
			projPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj");
			projFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\proj.flx");
		}

		TEST_METHOD_CLEANUP(MiscTestCleanup)
		{
			if (sln)
			{
				sln->Close();
				sln.reset();
				proj.reset();
				slnFilePath.reset();
			}

			if (testPath)
			{
				RemoveDirectoryTree(testPath.get());
				testPath.reset();
			}
		}

		TEST_METHOD(CloneProject)
		{
			HRESULT hr;

			com_ptr<IUnknown> solution;
			hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::IsTrue(SUCCEEDED(hr));
			auto sln = wil::com_query_failfast<VxDTE::_Solution>(solution);
			hr = sln->Create(wil::make_bstr_failfast(testPath.get()).get(), wil::make_bstr_failfast(L"test.sln").get());
			Assert::IsTrue(SUCCEEDED(hr));
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			com_ptr<VxDTE::Project> proj;
			hr = sln->AddFromTemplate (
				wil::make_bstr_failfast(TemplatePath_TwoConfigsOneFile.get()).get(),
				wil::make_bstr_failfast(testPath.get()).get(),
				wil::make_bstr_failfast(L"test.flx").get(), VARIANT_TRUE, &proj);
			Assert::IsTrue(SUCCEEDED(hr), wil::str_printf_failfast<wil::unique_process_heap_string>(L"0x%08x", hr).get());
		}

		TEST_METHOD(OpenSpecificEditor)
		{
			// At some point between VS 17.10 and 17.14, the VS implementation of IVsUIShellOpenDocument::OpenStandardEditor
			// started trying to open our .asm files not with "Source Code (Text) Editor", but with
			// "Common Language Editor Supporting TextMate Bundles" (or at least this was the new default shown in Open With...).
			// The implementation started returning E_NOTIMPL in most cases.
			// The solution was to call OpenSpecificEditor with GUID_TextEditorFactory, instead of OpenStandardEditor.
			// This test verifies this solution. The test fails with the old code that calls OpenStandardEditor,
			// and succeeds with the new code that calls OpenSpecificEditor.

			HRESULT hr;
			auto hier = proj.query<IVsUIHierarchy>();
			VSITEMID itemid;
			hr = hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			auto vsp2 = proj.query<IVsProject2>();
			com_ptr<IVsWindowFrame> wf;
			hr = vsp2->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->Show();
			Assert::IsTrue(SUCCEEDED(hr));
			hr = sln->Close();
			Assert::IsTrue(SUCCEEDED(hr));

			hr = sln->Open(wil::make_bstr_failfast(slnFilePath.get()).get()); 
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<VxDTE::Projects> projects;
			hr = sln->get_Projects(&projects);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = projects->Item(wil::make_variant_bstr_failfast(L"proj.flx"), &proj);
			Assert::IsTrue(SUCCEEDED(hr));

			hier = proj.query<IVsUIHierarchy>();
			hr = hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			vsp2 = proj.query<IVsProject2>();
			hr = vsp2->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf); // This would return E_NOTIMPL before the fix.
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(NavigateToErrorInFileInSubdir)
		{
			// When the user double-clicks an error in the Error List window, VS goes out of its way to locate the file
			// and find the project it contains. To repro the bug, we need to put the project file in a subdirectory
			// (not next to the .sln file), and the offending file in a subdirectory of its own. In this case, when the user
			// double-clicks the error in the Error List window, VS used to be unable to locate the file due to bugs in
			// ProjectNode::ParseCanonicalName and ProjectNode::IsDocumentInProject (and ultimately in ParseCanonicalName).
			// This test verifies the fixes in these functions.

			HRESULT hr;
			long buildFailCount;
			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);

			auto file1Path = CombinePath(projPath.get(), L"subdir\\file1.asm");
			WriteFileOnDisk (file1Path.get(), "\t555555");

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(file1Path.addressof()), NULL, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj->Save();
			Assert::IsTrue(SUCCEEDED(hr));

			BuildSolution(sln, &buildFailCount);
			Assert::IsTrue(buildFailCount >= 1);

			hr = dte->ExecuteCommand(wil::make_bstr_failfast(L"View.ErrorList").get());
			auto dte2 = wil::com_query_failfast<VxDTE::DTE2>(dte);
			wil::com_ptr_failfast<VxDTE::ToolWindows> toolWindows;
			hr = dte2->get_ToolWindows(&toolWindows);
			wil::com_ptr_failfast<VxDTE::ErrorList> errorList;
			hr = toolWindows->get_ErrorList(&errorList);
			wil::com_ptr_failfast<VxDTE::ErrorItems> errorItems;
			hr = errorList->get_ErrorItems(&errorItems);
			long errorCount;
			hr = errorItems->get_Count(&errorCount);
			Assert::IsTrue(errorCount >= 1);

			VARIANT v; v.vt = VT_I4; v.lVal = 1;
			com_ptr<VxDTE::ErrorItem> errorItem;
			hr = errorItems->Item(v, &errorItem);
			Assert::IsTrue(SUCCEEDED(hr));

			errorItem->Navigate(); // This call succeeds, even though VS couldn't open the file.

			com_ptr<VxDTE::Document> doc;
			hr = dte->get_ActiveDocument(&doc);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(doc.get()); // This would fail before the fix.

			BOOL found = FALSE;
			VSDOCUMENTPRIORITY prio = { };
			VSITEMID itemid = { };
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"subdir\\file1.asm", &found, &prio, &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"subdir/..\\subdir/.\\file1.asm", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"SubDir\\File1.ASM", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);

			// For the canonical name, we don't bother with tricky paths such as "subdir/..\\subdir/.\\file1.asm".
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"subdir\\file1.asm", &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"SubDir\\File1.ASM", &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(GenPrePostInclude_OnlyActiveCfg)
		{
			// Verify that the PreInclude.asm and PostInclude.asm are regenerated when editing the active configuration,
			// and that they are _not_ regenerated when editing an inactive configuration.

			HRESULT hr;
			// Let's not go through the configuration manager since we haven't implemented Project::get_ConfigurationManager yet.
			//Microsoft_VisualStudio_Interop::_SolutionPtr sln = sln0.get();
			//auto configs = sln->SolutionBuild->SolutionConfigurations;
			//Assert::IsTrue(configs->Count >= 2);
			//Assert::AreEqual<void*>(configs->Item[1].GetInterfacePtr(), sln->SolutionBuild->ActiveConfiguration.GetInterfacePtr());

			com_ptr<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// Make a change in the active configuration and verify that the Pre/PostInclude files are generated.
			auto genFilesPath = CombinePath(projPath.get(), L"GeneratedFiles");
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
			// Verify that, with a single .asm file in the project, the PreInclude.asm and PostInclude.asm 
			// are generated when changing the BuildTool of the .asm file to Assembler, and deleted when
			// changing the BuildTool to something else.

			// The pre/post-include files should have been generated on project creation.
			auto genFilesPath = CombinePath(projPath.get(), L"GeneratedFiles");
			Assert::IsTrue(PathFileExists(genFilesPath.get()));

			VSITEMID itemid;
			auto hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant filevar;
			proj.query<IVsHierarchy>()->GetProperty(itemid, VSHPROPID_BrowseObject, &filevar);
			auto fileProps = wil::com_query_failfast<IFileNodeProperties>(filevar.pdispVal);
			fileProps->put_BuildTool(BuildToolKind::None);
			Assert::IsTrue(!PathFileExists(genFilesPath.get()));

			fileProps->put_BuildTool(BuildToolKind::Assembler);
			Assert::IsTrue(PathFileExists(genFilesPath.get()));
		}
		TEST_METHOD(RemoveFileClosesEditor)
		{
			HRESULT hr;
			VSITEMID itemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IVsWindowFrame> wf;
			hr = proj.query<IVsProject2>()->OpenItem(itemId, LOGVIEWID_Code, nullptr, &wf);
			wf->Show();
			hr = wf->IsVisible();
			Assert::AreEqual (S_OK, hr);

			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &itemId, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->IsVisible();
			Assert::AreNotEqual (S_OK, hr);
		}

		TEST_METHOD(RemoveFolderAndFile_FolderFirstInList)
		{
			wil::unique_variant folderItemId;
			auto hr = proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folderItemId);
			Assert::AreEqual (S_OK, hr);
			
			VSITEMID fileItemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &fileItemId);
			Assert::AreEqual (S_OK, hr);

			VSITEMID itemIds[2] = { V_UI4(&folderItemId), fileItemId };
			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(2, DELITEMOP_DeleteFromStorage, itemIds, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(RemoveFolderClosesEditors)
		{
			wil::unique_variant folderItemId;
			auto hr = proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folderItemId);
			Assert::AreEqual (S_OK, hr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem (V_UI4(&folderItemId), oper, L"file2.asm", 1, const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), NULL, &addResult);
			Assert::AreEqual (S_OK, hr);

			VSITEMID itemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"NewFolder1\\file2.asm", &itemId);
			Assert::AreEqual (S_OK, hr);
			com_ptr<IVsWindowFrame> wf;
			hr = proj.query<IVsProject2>()->OpenItem(itemId, LOGVIEWID_Code, nullptr, &wf);
			wf->Show();
			hr = wf->IsVisible();
			Assert::AreEqual (S_OK, hr);

			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &V_UI4(&folderItemId), DHO_SUPPRESS_UI);
			Assert::AreEqual (S_OK, hr);
			hr = wf->IsVisible();
			Assert::AreNotEqual (S_OK, hr);
		}

		TEST_METHOD(AddNewFile_SameNameAsFileOutsideProjectDir)
		{
			HRESULT hr;
			const wchar_t file1RelPath[] = L"test.asm";
			auto file1Path = CombinePath(projPath.get(), file1RelPath);
			WriteFileOnDisk (file1Path.get(), "; comment");
			const wchar_t file2RelPath[] = L"..\\test.asm";
			auto file2Path = CombinePath (projPath.get(), file2RelPath);
			WriteFileOnDisk (file2Path.get(), "; comment");
			wchar_t file3Path[] = L"D:\\FelixTest\\test.asm";
			WriteFileOnDisk (file3Path, "; comment");

			LPCOLESTR filesToOpen[] = { file1Path.get(), file2Path.get(), file3Path };
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 3, filesToOpen, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Adding them again should fail.
			for (const wchar_t* fileToOpen : filesToOpen)
			{
				hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, &fileToOpen, nullptr, &addResult);
				Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), hr);
			}

			VSITEMID itemids[3];
			for (int i = 0; i < 3; i++)
			{
				BOOL found;
				VSDOCUMENTPRIORITY prio;
				hr = proj.query<IVsProject>()->IsDocumentInProject(filesToOpen[i], &found, &prio, &itemids[i]);
				Assert::IsTrue(SUCCEEDED(hr) && found);
			}

			wil::unique_bstr canonicalNames[3];
			for (int i = 0; i < 3; i++)
			{
				hr = proj.query<IVsHierarchy>()->GetCanonicalName(itemids[i], &canonicalNames[i]);
				Assert::IsTrue(SUCCEEDED(hr));
			}

			for (int i = 0; i < 3; i++)
			{
				VSITEMID itemidx;
				hr = proj.query<IVsHierarchy>()->ParseCanonicalName(canonicalNames[i].get(), &itemidx);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual(itemids[i], itemidx);
			}
		}

		TEST_METHOD(CloneFile_SameNameAsFileOutsideProjectDir)
		{
			HRESULT hr;
			auto templatePath = CombinePath(testPath.get(), L"template.asm");
			WriteFileOnDisk (templatePath.get(), "; comment");

			static const wchar_t file1RelPath[] = L"..\\test.asm";
			auto file1Path = CombinePath(projPath.get(), file1RelPath);
			WriteFileOnDisk(file1Path.get(), "; comment");
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(file1Path.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"test.asm", 1, const_cast<LPCOLESTR*>(templatePath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(CloneFileToFolderMissingOnDisk)
		{
			HRESULT hr;
			auto templatePath = CombinePath (testPath.get(), L"template.asm");
			WriteFileOnDisk (templatePath.get(), "; comment");

			wil::unique_variant folderItemId;
			hr = proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));

			// Delete the folder on disk, then try to clone a file into it. We must at least not crash.

			wil::unique_bstr folderMk;
			hr = proj.query<IVsProject>()->GetMkDocument(V_VSITEMID(&folderItemId), &folderMk);
			Assert::IsTrue(SUCCEEDED(hr));

			if (PathFileExists(folderMk.get()))
			{
				BOOL bres = RemoveDirectory(folderMk.get());
				Assert::IsTrue(bres);
			}

			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			VSADDRESULT addResult;
			hr = proj.query<IVsProject>()->AddItem(V_VSITEMID(&folderItemId), oper, L"file.asm", 1, const_cast<LPCOLESTR*>(templatePath.addressof()), nullptr, &addResult);
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
			auto testPath = wil::str_concat_failfast<wil::unique_hglobal_string>(tempPath, L"GenPrePostInclude_LoadOldXml");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			wil::com_ptr_failfast<IUnknown> solution;
			hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			auto sln = solution.query<VxDTE::_Solution>();
			//hr = sln->Create(wil::make_bstr_failfast(testPath.get()).get(), wil::make_bstr_failfast(L"test").get());
			//Assert::IsTrue(SUCCEEDED(hr));
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			auto projPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj.flx");
			WriteFileOnDisk(projPath.get(), OldXmlWithGeneratedFilesFolder);

			wil::com_ptr_failfast<VxDTE::Project> proj;
			hr = sln->AddFromFile (wil::make_bstr_failfast(projPath.get()).get(), VARIANT_TRUE, &proj);
			Assert::IsTrue(SUCCEEDED(hr));

			// The folder named GeneratedFiles should have been ignored while loading from XML,
			// and nothing should have been generated since there are no other files with
			// BuildTool==Assembler in the XML.
			wil::unique_variant var;
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &var);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, var.vt);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&var));
			auto genFilesPath = CombinePath(testPath.get(), L"GeneratedFiles");
			Assert::IsFalse(PathFileExists(genFilesPath.get()));

			// And the project should not be dirty just after loading from XML.
			auto pff = proj.query<IPersistFileFormat>();
			BOOL dirty;
			hr = pff->IsDirty(&dirty); Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsFalse(dirty);

			// Now if we add an .asm file, we should have the generated files.
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"file.asm", 1, const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), NULL, &addResult);
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
			auto hr = proj->Save(nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hfile file (CreateFile(projFilePath.get(), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0));
			Assert::IsTrue(file.is_valid());
			DWORD fileSize = GetFileSize(file.get(), 0);
			auto buffer = wil::make_process_heap_ansistring(nullptr, fileSize);
			BOOL bres = ReadFile(file.get(), buffer.get(), fileSize, nullptr, nullptr);
			Assert::IsTrue(bres);
			Assert::IsNull(strstr(buffer.get(), "GeneratedFiles"));
		}

		TEST_METHOD(GenPrePostInclude_AddFirstAsm_RemoveLastAsm)
		{
			HRESULT hr;

			// We have a file in the project with BuildTool==Assembler. We should have a GeneratedFiles folder.
			wil::unique_variant firstChild;
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Remove the only file
			VSITEMID fileItemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &fileItemId);
			Assert::AreEqual(S_OK, hr);
			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_RemoveFromProject, &fileItemId, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);

			// Project should be empty (no file and no GeneratedFiles folder)
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&firstChild));

			// Add the file back.
			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\file.asm");
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(filePath.addressof()), nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			// First child should be "GeneratedFiles"
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));
		}

		TEST_METHOD(GenPrePostInclude_EnableDisableGeneration)
		{
			HRESULT hr;

			// Assume first config is the active config. 
			wil::com_ptr_failfast<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// We have a file in the project with BuildTool==Assembler. We should have a GeneratedFiles folder.
			wil::unique_variant firstChild;
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Disable generation. The folder should be gone.
			com_ptr<IProjectConfigAssemblerProperties> asmProps;
			cfgs[0].query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
			Assert::AreEqual(S_OK, hr);
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Enable again generation. Folder should be back.
			hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Disable generation in the inactive config. Folder should remain.
			cfgs[1].query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
			Assert::AreEqual(S_OK, hr);
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));
		}

		TEST_METHOD(GenPrePostInclude_SwitchActiveConfig)
		{
			HRESULT hr;

			wil::com_ptr_failfast<VxDTE::SolutionBuild> slnBuild;
			hr = sln->get_SolutionBuild(&slnBuild);
			Assert::AreEqual(S_OK, hr);
			wil::com_ptr_failfast<VxDTE::SolutionConfigurations> slnConfigs;
			hr = slnBuild->get_SolutionConfigurations(&slnConfigs);
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
			hr = slnBuild->get_ActiveConfiguration(&activeSlnConfig);
			Assert::AreEqual(S_OK, hr);
			Assert::IsTrue(activeSlnConfig == slnConfig0);
			
			wil::com_ptr_failfast<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// Generation is by default enabled. We should have a GeneratedFiles folder.
			wil::unique_variant firstChild;
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Disable generation in the second config. We should still have a GeneratedFiles folder.
			wil::com_ptr_failfast<IProjectConfigAssemblerProperties> config1AsmProps;
			hr = cfgs[1].query<IProjectConfigProperties>()->get_AssemblerProperties(&config1AsmProps);
			Assert::AreEqual(S_OK, hr);
			hr = config1AsmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
			Assert::AreEqual(S_OK, hr);
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Switch the active config to the one with generation disabled. We should no longer have a GeneratedFiles folder.
			hr = slnConfig1->Activate();
			Assert::AreEqual(S_OK, hr);
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));

			// Switch the active config back. We should have a GeneratedFiles folder again.
			hr = slnConfig0->Activate();
			Assert::AreEqual(S_OK, hr);
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_GENFILES, V_VSITEMID(&firstChild));
		}
	};
}
