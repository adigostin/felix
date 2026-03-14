
#include "pch.h"
#include "shared/com.h"
#include "../TestsCommon.h"

#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>
namespace VxDTE
{
	#include <dte80.h>
	#include <dte90.h>
}

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace UITests
{
	extern wil::com_ptr_failfast<VxDTE::_DTE> dte;
	extern wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path);
	extern wil::unique_bstr GetBuildOutputWindowPaneContent();

	TEST_CLASS(BuildTests)
	{
		struct TD
		{
			wil::unique_process_heap_string testDir;
			wil::unique_process_heap_string slnFilePath;
			wil::unique_process_heap_string projDir;
			wil::unique_process_heap_string projFilePath;
			wil::com_ptr_failfast<VxDTE::_Solution> sln;
			wil::com_ptr_failfast<VxDTE::Project> proj;
			com_ptr<VxDTE::SolutionBuild> slnBuild;

			TD (const wchar_t* projectTemplatePath)
			{
				HRESULT hr;
				testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildTests");
				Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));
				
				wil::com_ptr_failfast<IUnknown> solution;
				hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
				Assert::IsTrue(SUCCEEDED(hr));
				sln = solution.query<VxDTE::_Solution>();
				hr = sln->Create(wil::make_bstr_failfast(testDir.get()).get(), wil::make_bstr_failfast(L"test").get());
				Assert::IsTrue(SUCCEEDED(hr));

				projDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\proj");
				Assert::IsTrue(CreateDirectory(projDir.get(), nullptr));

				hr = sln->AddFromTemplate (
					wil::make_bstr_failfast(projectTemplatePath).get(),
					wil::make_bstr_failfast(projDir.get()).get(),
					wil::make_bstr_failfast(L"proj.flx").get(), VARIANT_FALSE, &proj);
				Assert::IsTrue(SUCCEEDED(hr));
				hr = sln->SaveAs(wil::make_bstr_failfast(L"test").get());
				Assert::IsTrue(SUCCEEDED(hr));

				hr = sln->get_SolutionBuild(&slnBuild);
				Assert::IsTrue(SUCCEEDED(hr));

				slnFilePath = CombinePath(testDir.get(), L"test.sln");
				projFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(projDir, L"\\proj.flx");
			}

			~TD()
			{
				sln->Close();
				RemoveDirectoryTree(testDir.get());
			}
		};

		TEST_METHOD(BuildProject)
		{
			TD td (TemplatePath_TwoConfigsOneFile.get());
			auto hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			// LastBuildInfo returns the number of failed projects, despite the parameter name.
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(0l, buildFailCount);
		}

		TEST_METHOD(BuildProjectWithError)
		{
			HRESULT hr;
			TD td (TemplatePath_TwoConfigsOneFile.get());

			long buildFailCount;
			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			// LastBuildInfo returns the number of failed projects, despite the parameter name.
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(0l, buildFailCount);

			WriteFileOnDisk (CombinePath(td.projDir.get(), L"file.asm").get(), "\tabcde");
			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			// LastBuildInfo returns the number of failed projects, despite the parameter name.
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(1l, buildFailCount);

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
			Assert::IsTrue(errorCount > 1);
		}

		TEST_METHOD(BuildOutDirNoBackslash)
		{
			HRESULT hr;
			TD td (TemplatePath_TwoConfigsOneFile.get());

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);

			com_ptr<IProjectConfigGeneralProperties> generalProps;
			cfg.query<IProjectConfigProperties>()->get_GeneralProperties(&generalProps);
			generalProps->put_OutputDirectory(wil::make_bstr_failfast(L"%PROJECT_DIR%Out").get());
			generalProps->put_OutputFileType(OutputFileType::Sna);

			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(0l, buildFailCount);
			Assert::IsTrue(PathFileExists(wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\Out\\proj.sna").get()));
		}

		TEST_METHOD(BuildOutDirOutsideProjectDir)
		{
			HRESULT hr;
			TD td (TemplatePath_TwoConfigsOneFile.get());

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);

			com_ptr<IProjectConfigGeneralProperties> generalProps;
			cfg.query<IProjectConfigProperties>()->get_GeneralProperties(&generalProps);

			auto buildIt = [&generalProps, &td](const wchar_t* outputDirExpected)
				{
					// Sna
					generalProps->put_OutputFileType(OutputFileType::Sna);
					auto hr = td.slnBuild->Build(VARIANT_TRUE);
					Assert::IsTrue(SUCCEEDED(hr));
					long buildFailCount;
					hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
					Assert::IsTrue(SUCCEEDED(hr));
					Assert::AreEqual(0l, buildFailCount);
					auto outputFile = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s\\proj.sna", outputDirExpected);
					Assert::IsTrue(PathFileExists(outputFile.get()));
					Assert::IsTrue(DeleteFile(outputFile.get()));

					// Binary
					generalProps->put_OutputFileType(OutputFileType::Binary);
					hr = td.slnBuild->Build(VARIANT_TRUE);
					Assert::IsTrue(SUCCEEDED(hr));
					hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
					Assert::IsTrue(SUCCEEDED(hr));
					Assert::AreEqual(0l, buildFailCount);
					outputFile = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s\\proj.bin", outputDirExpected);
					Assert::IsTrue(PathFileExists(outputFile.get()));
					Assert::IsTrue(DeleteFile(outputFile.get()));
				};

			// Outside project dir but on same drive, full path.
			generalProps->put_OutputDirectory(wil::make_bstr_failfast(L"%PROJECT_DIR%..").get());
			buildIt(td.testDir.get());

			// Outside project dir but on same drive, relative path.
			generalProps->put_OutputDirectory(wil::make_bstr_failfast(L"..").get());
			buildIt(td.testDir.get());

			// Different drive
			auto testPathOtherDrive = MakeVolumeGuidPath(td.testDir.get());
			wchar_t outputPathOtherDrive[MAX_PATH];
			PathCombine(outputPathOtherDrive, testPathOtherDrive.get(), L"newdir");
			{
				// Quick test that the path is writeable, before asking the build system to write to it
				wil::CreateDirectoryDeep(outputPathOtherDrive);
				wchar_t test[MAX_PATH];
				PathCombine(test, outputPathOtherDrive, L"test.txt");
				Assert::IsTrue(wil::unique_hfile(CreateFile(test, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, 0, NULL)).is_valid());
				Assert::IsTrue(DeleteFile(test));
				Assert::IsTrue(RemoveDirectory(outputPathOtherDrive));
			}
			generalProps->put_OutputDirectory(wil::make_bstr_failfast(outputPathOtherDrive).get());
			buildIt(outputPathOtherDrive);
		}

		TEST_METHOD(BuildWithFilesInFolders)
		{
			TD td (TemplatePath_TwoConfigsOneFile.get());
			HRESULT hr;

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			wil::com_ptr_failfast<IProjectConfigGeneralProperties> generalProps;
			cfg.query<IProjectConfigProperties>()->get_GeneralProperties(&generalProps);
			generalProps->put_OutputFileType(OutputFileType::Binary);
			wil::com_ptr_failfast<IProjectConfigAssemblerProperties> asmProps;
			cfg.query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			asmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);

			VSITEMID itemid;
			hr = td.proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemid);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &itemid, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);

			// Check that it's empty.
			wil::unique_variant fc;
			hr = td.proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &fc);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, fc.vt);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&fc));

			// Add new file
			auto filePath = CombinePath(td.projDir.get(), L"Folder\\single.asm");
			WriteFileOnDisk(filePath.get(), "\tnop\r\n\tend");
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(filePath.addressof()), NULL, &addResult);
			hr = td.proj.query<IVsHierarchy>()->ParseCanonicalName(L"Folder\\single.asm", &itemid);
			Assert::AreEqual(S_OK, hr);

			// Try it with the assembler.
			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(0l, buildFailCount);
			auto outputFilePath = CombinePath(td.projDir.get(), L"Out\\Debug\\proj.bin");
			wil::unique_hfile file (CreateFile(outputFilePath.get(), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0));
			Assert::IsTrue(file.is_valid());
			Assert::AreEqual(1ul, GetFileSize(file.get(), NULL));
			file.reset();
			Assert::IsTrue(DeleteFile(outputFilePath.get()));

			// Now switch to custom build tool and try again.
			wil::unique_variant filevar;
			hr = td.proj.query<IVsHierarchy>()->GetProperty(itemid, VSHPROPID_BrowseObject, &filevar);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_DISPATCH, filevar.vt);
			wil::com_ptr_failfast<IFileNodeProperties> fileProps;
			hr = filevar.pdispVal->QueryInterface(&fileProps);
			Assert::AreEqual(S_OK, hr);
			fileProps->put_BuildTool(BuildToolKind::CustomBuildTool);
			wil::com_ptr_failfast<ICustomBuildToolProperties> cbtProps;
			fileProps->get_CustomBuildToolProperties(&cbtProps);
			cbtProps->put_CommandLine(wil::make_bstr_failfast(L"cmd /c echo > %OUTPUT_DIR%\\proj.bin").get());

			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(0l, buildFailCount);
			file.reset (CreateFile(outputFilePath.get(), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0));
			Assert::IsTrue(file.is_valid());
			Assert::AreNotEqual(0ul, GetFileSize(file.get(), NULL));
			file.reset();
		}

		TEST_METHOD(BuildFilesInFoldersAndSubfolders)
		{
			TD td (TemplatePath_EmptyProject.get());
			HRESULT hr;
			auto hier = td.proj.query<IVsUIHierarchy>();

			wil::unique_variant folder;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
				OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"folder").addressof(), &folder);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, folder.vt);
			//hr = proj->AsHierarchy()->SetProperty (V_VSITEMID(&folder), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"folder"));
			//Assert::IsTrue(SUCCEEDED(hr));

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&folder), oper, L"file1.asm", 1,
				const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), NULL, &addResult);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant subfolder;
			hr = hier->ExecCommand (V_VSITEMID(&folder), &CMDSETID_StandardCommandSet97, cmdidNewFolder,
				OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_nothrow(L"subfolder").addressof(), &subfolder);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, subfolder.vt);
			//hr = hier->SetProperty (V_VSITEMID(&subfolder), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"subfolder"));
			//Assert::IsTrue(SUCCEEDED(hr));

			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&subfolder), oper, L"file2.asm", 1, 
				const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), NULL, &addResult);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			wil::com_ptr_failfast<IProjectConfigAssemblerProperties> asmProps;
			cfg.query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			wil::unique_bstr cmdLine;
			asmProps->get_CommandLine(&cmdLine);

			Assert::IsNotNull(wcsstr(cmdLine.get(), L" folder\\file1.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), L" folder\\subfolder\\file2.asm"));
		}

		TEST_METHOD(BuildFilesNotInProjectDir)
		{
			TD td (TemplatePath_EmptyProject.get());
			HRESULT hr;
			auto hier = td.proj.query<IVsUIHierarchy>();

			// file 1 in project dir
			auto file1FullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\file1.asm");
			WriteFileOnDisk(file1FullPath.get(), "");
			// file 2 in project sub dir
			auto file2FullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\subdir\\file2.asm");
			WriteFileOnDisk(file2FullPath.get(), "");
			// file 3 outside project dir but on same drive
			auto file3FullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.testDir, L"\\file3.asm");
			WriteFileOnDisk(file3FullPath.get(), "");
			// file 4 on different drive
			auto temp = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\file4.asm");
			auto file4FullPath = MakeVolumeGuidPath(temp.get());
			WriteFileOnDisk(file4FullPath.get(), "");

			LPCOLESTR files[] = { file1FullPath.get(), file2FullPath.get(), file3FullPath.get(), file4FullPath.get() };

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", (ULONG)_countof(files), files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			wil::com_ptr_failfast<IProjectConfigAssemblerProperties> asmProps;
			cfg.query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			wil::unique_bstr cmdLine;
			asmProps->get_CommandLine(&cmdLine);

			Assert::IsNotNull(wcsstr(cmdLine.get(), L" file1.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), L" subdir\\file2.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), L" ..\\file3.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), file4FullPath.get()));
		}

		TEST_METHOD(SjasmCommandLine_ExitCodeZero)
		{
			HRESULT hr;
			TD td (TemplatePath_TwoConfigsOneFile.get());
			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(0l, buildFailCount);
		}

		TEST_METHOD(SjasmCommandLine_ExitCodeNonzero)
		{
			HRESULT hr;
			TD td (TemplatePath_TwoConfigsOneFile.get());
			WriteFileOnDisk (CombinePath(td.projDir.get(), L"file.asm").get(), "\tabcde");
			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(1l, buildFailCount);
		}

		TEST_METHOD(BuildOnlySynchronousSteps)
		{
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());
			HRESULT hr;

			VSITEMID itemid;
			hr = td.proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemid);
			Assert::AreEqual(S_OK, hr);
			wil::unique_variant filevar;
			hr = td.proj.query<IVsHierarchy>()->GetProperty(itemid, VSHPROPID_BrowseObject, &filevar);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_DISPATCH, filevar.vt);
			wil::com_ptr_failfast<IFileNodeProperties> fileProps;
			hr = filevar.pdispVal->QueryInterface(&fileProps);
			Assert::AreEqual(S_OK, hr);
			wil::com_ptr_failfast<ICustomBuildToolProperties> cbtProps;
			fileProps->get_CustomBuildToolProperties(&cbtProps);
			cbtProps->put_CommandLine(nullptr);
			static const wchar_t description[] = L"CBT Description";
			cbtProps->put_Description(wil::make_bstr_nothrow(description).get());

			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);
			long buildFailCount = -1;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(0l, buildFailCount);

			auto text = GetBuildOutputWindowPaneContent();
			auto p = wcsstr(text.get(), description);
			Assert::IsNotNull(p);
		}

		TEST_METHOD(CustomBuildToolFilesInFolders)
		{
			// TODO:
		}

		TEST_METHOD(BuildFailsOnEmptyProject)
		{
			TD td (TemplatePath_EmptyProject.get());
			HRESULT hr;

			hr = td.slnBuild->BuildProject(wil::make_bstr_failfast(L"Debug").get(),
				wil::make_bstr_failfast(L"proj\\proj.flx").get(), VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(1l, buildFailCount);
		}

		TEST_METHOD(PrePostBuildEvents)
		{
			TD td (TemplatePath_EmptyProject.get());
			HRESULT hr;

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			auto config = cfg.query<IProjectConfigProperties>();

			com_ptr<IProjectConfigPrePostBuildProperties> preBuildProps;
			hr = config->get_PreBuildProperties(&preBuildProps);
			Assert::AreEqual(S_OK, hr);
			hr = preBuildProps->put_CommandLine(wil::make_bstr_nothrow(L"cmd /c echo XXXX").get());
			Assert::AreEqual(S_OK, hr);

			com_ptr<IProjectConfigPrePostBuildProperties> postBuildProps;
			hr = config->get_PostBuildProperties(&postBuildProps);
			Assert::AreEqual(S_OK, hr);
			hr = postBuildProps->put_CommandLine(wil::make_bstr_nothrow(L"cmd /c echo YYYY").get());
			Assert::AreEqual(S_OK, hr);

			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);

			auto output = GetBuildOutputWindowPaneContent();
			auto preMessage = wcsstr(output.get(), L"XXXX");
			Assert::IsNotNull(preMessage);
			auto postMessage = wcsstr(output.get(), L"YYYY");
			Assert::IsNotNull(preMessage);
		}

		static void SetCustomBuildToolParams (VxDTE::Project* proj, const wchar_t* fileCanonicalName, const wchar_t* commandLine, const wchar_t* description)
		{
			auto hier = wil::com_query_failfast<IVsHierarchy>(proj);
			VSITEMID itemId;
			auto hr = hier->ParseCanonicalName(L"file.asm", &itemId); 
			Assert::AreEqual(S_OK, hr);
			wil::unique_variant file;
			hr = hier->GetProperty(itemId, VSHPROPID_BrowseObject, &file);
			Assert::AreEqual(S_OK, hr);
			auto fileProps = wil::com_query_failfast<IFileNodeProperties>(file.pdispVal);
			
			wil::com_ptr_failfast<ICustomBuildToolProperties> cbtProps;
			hr = fileProps->get_CustomBuildToolProperties(&cbtProps); 
			Assert::AreEqual(S_OK, hr);
			
			if (commandLine)
			{
				hr = cbtProps->put_CommandLine(wil::make_bstr_failfast(commandLine).get());
				Assert::AreEqual(S_OK, hr);
			}

			if (description)
			{
				hr = cbtProps->put_Description(wil::make_bstr_failfast(description).get());
				Assert::AreEqual(S_OK, hr);
			}
		}

		TEST_METHOD(CustomBuildToolOnlyWhitespaceCommands)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			SetCustomBuildToolParams(td.proj, L"file.asm", L"   \r\n   \t   ", nullptr);

			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(1l, buildFailCount);
			auto output = GetBuildOutputWindowPaneContent();
			auto x = wcsstr(output.get(), L"0x80070012"); // HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES)
			Assert::IsNotNull(x, output.get());
			/*
			wil::com_ptr_failfast<IDispatch> aodisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"TestHelper").get(), &aodisp);
			Assert::AreEqual(S_OK, hr);
			auto ao = aodisp.query<IFelixTestHelper>();
			wil::com_ptr_failfast<IUnknown> runnerUnk;
			hr = ao->CreateInstanceFromLocalRegistry(__uuidof(IBuildRunner), runnerUnk.addressof());
			Assert::AreEqual(S_OK, hr);
			wil::com_ptr_failfast<IBuildRunner> runner;
			hr = runnerUnk->QueryInterface(IID_PPV_ARGS(runner.addressof()));
			Assert::AreEqual(S_OK, hr);
			*/
		}

		TEST_METHOD(CustomBuildToolSomeWhitespaceCommands)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			static const wchar_t cmdLine[] = L"  cmd /c type file.asm  \t\r\n   \t   ";
			SetCustomBuildToolParams(td.proj, L"file.asm", cmdLine, nullptr);

			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);
			long buildFailCount;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(0l, buildFailCount);
			auto output = GetBuildOutputWindowPaneContent();
			auto x = wcsstr(output.get(), L"start:");
			Assert::IsNotNull(x, output.get());
		}

		TEST_METHOD(CustomBuildToolWaitingUserInput)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			SetCustomBuildToolParams(td.proj, L"file.asm", L"cmd /c pause", nullptr);
			hr = td.slnBuild->Build(VARIANT_FALSE);
			Assert::AreEqual(S_OK, hr);

			auto tickCount = GetTickCount();
			VxDTE::vsBuildState buildState;
			while (true)
			{
				hr = td.slnBuild->get_BuildState(&buildState);
				Assert::AreEqual(S_OK, hr);
				if (buildState == VxDTE::vsBuildStateDone)
					break;
				if (GetTickCount() - tickCount >= 1000)
					break;
				Sleep(50);
			}

			Assert::AreEqual<int>(VxDTE::vsBuildStateInProgress, buildState);

			hr = dte->ExecuteCommand(wil::make_bstr_failfast(L"Build.Cancel").get());
			Assert::AreEqual(S_OK, hr);

			while (SUCCEEDED(td.slnBuild->get_BuildState(&buildState)) && buildState == VxDTE::vsBuildStateInProgress)
				Sleep(50);

			Assert::AreEqual<int>(VxDTE::vsBuildStateDone, buildState);

			long buildFailCount = 0;
			hr = td.slnBuild->get_LastBuildInfo(&buildFailCount);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(1l, buildFailCount);
		}

		class HeavyLoad
		{
			wil::unique_event_failfast event;
			vector_nothrow<wil::unique_handle> threads;

		public:
			HeavyLoad()
			{
				SYSTEM_INFO si;
				GetSystemInfo (&si);
				threads.try_resize(si.dwNumberOfProcessors);
				event.create(wil::EventOptions::ManualReset);
				for (auto& h : threads)
					h.reset(CreateThread(nullptr, 0, ThreadProc, this, 0, nullptr));
			}

			~HeavyLoad()
			{
				event.SetEvent();
				while(!threads.empty())
				{
					WaitForSingleObject(threads.back().get(), INFINITE);
					threads.remove_back();
				}
			}

		private:
			static DWORD WINAPI ThreadProc (void* arg)
			{
				HeavyLoad* _this = (HeavyLoad*)arg;
				DWORD tickStart = GetTickCount();
				while (!_this->event.is_signaled())
					;
				return (DWORD)0;
			}
		};

		TEST_METHOD(CustomBuildToolOutputWithNoEOL)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			HeavyLoad hl;

			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\file.asm");
			WriteFileOnDisk(filePath.get(), "content_xxx");
			SetCustomBuildToolParams (td.proj, L"file.asm", L"cmd /c type file.asm", nullptr);
			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);

			auto output = GetBuildOutputWindowPaneContent();
			static const wchar_t cnt[] = L"content_xxx";
			auto x = wcsstr(output.get(), cnt);
			Assert::IsNotNull(x);
			wchar_t charAfter = x[std::size(cnt) - 1];
			Assert::AreNotEqual(L'\r', charAfter);
			Assert::AreNotEqual(L'\n', charAfter);
		}

		TEST_METHOD(CustomBuildToolOutputWithEOL)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			HeavyLoad hl;

			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\file.asm");
			WriteFileOnDisk(filePath.get(), "content_xxx\r\n");
			SetCustomBuildToolParams (td.proj, L"file.asm", L"cmd /c type file.asm", nullptr);
			hr = td.slnBuild->Build(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);

			auto output = GetBuildOutputWindowPaneContent();
			static const wchar_t cnt[] = L"content_xxx\r\n";
			auto x = wcsstr(output.get(), cnt);
			Assert::IsNotNull(x);
		}
		
		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode0)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			SetCustomBuildToolParams(td.proj, L"file.asm", L"cmd /c exit 0", nullptr);

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			auto config = cfg.query<IVsBuildableProjectCfg2>();
			hr = config->StartBuildEx(0, NULL, 0x0001'0000);
			Assert::AreEqual(S_FALSE, hr);
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode1)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			SetCustomBuildToolParams(td.proj, L"file.asm", L"cmd /c exit 1", nullptr);

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			auto config = cfg.query<IVsBuildableProjectCfg2>();
			hr = config->StartBuildEx(0, NULL, 0x0002'0000);
			Assert::AreEqual(S_FALSE, hr);
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode0_NotOnLastCmd)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			auto tempFilename = wil::str_concat_failfast<wil::unique_process_heap_string>(td.testDir, L"TST");
			Assert::IsTrue(wil::unique_hfile(CreateFile(tempFilename.get(), GENERIC_WRITE, 0, 0, CREATE_NEW, 0, 0)).is_valid());
			auto cmd = wil::str_printf_failfast<wil::unique_process_heap_string>(L"cmd /c exit 0\r\ncmd /c del \"%s\"", tempFilename);
			SetCustomBuildToolParams (td.proj, L"file.asm", cmd.get(), nullptr);

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			auto config = cfg.query<IVsBuildableProjectCfg2>();
			hr = config->StartBuildEx(0, NULL, 0x0002'0000);
			Assert::AreEqual(S_FALSE, hr);

			// Since we canceled the build right after the first command, the second command
			// (the one that deletes the temporary file), shouldn't have been executed.
			Assert::IsTrue(PathFileExists(tempFilename.get()));
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode1_NotOnLastCmd)
		{
			HRESULT hr;
			TD td (TemplatePath_OneConfigOneCustomBuildTool.get());

			auto tempFilename = wil::str_concat_failfast<wil::unique_process_heap_string>(td.testDir, L"TST");
			Assert::IsTrue(wil::unique_hfile(CreateFile(tempFilename.get(), GENERIC_WRITE, 0, 0, CREATE_NEW, 0, 0)).is_valid());
			auto cmd = wil::str_printf_failfast<wil::unique_process_heap_string>(L"cmd /c exit 1\r\ncmd /c del \"%s\"", tempFilename);
			SetCustomBuildToolParams (td.proj, L"file.asm", cmd.get(), nullptr);

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			hr = td.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			auto config = cfg.query<IVsBuildableProjectCfg2>();
			hr = config->StartBuildEx(0, NULL, 0x0002'0000);
			Assert::AreEqual(S_FALSE, hr);

			// Since we canceled the build right after the first command, the second command
			// (the one that deletes the temporary file), shouldn't have been executed.
			Assert::IsTrue(PathFileExists(tempFilename.get()));
		}
	};
}
