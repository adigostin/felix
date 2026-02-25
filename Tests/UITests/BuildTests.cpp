
#include "pch.h"
#include "shared/com.h"
#include "FelixPackage.h"
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
	};
}
