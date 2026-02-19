
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
	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount);
	extern wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path);

	TEST_CLASS(BuildTests)
	{
	public:
		TEST_METHOD(BuildProject)
		{
			HRESULT hr;

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildProject");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			long buildFailCount;
			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);
		}

		TEST_METHOD(BuildProjectWithError)
		{
			HRESULT hr;

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildProjectWithError");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			long buildFailCount;
			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);

			WriteFileOnDisk (CombinePath(testPath.get(), L"file.asm").get(), "\tabcde");

			BuildSolution(sln, &buildFailCount);
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

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildOutDirNoBackslash");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"proj");
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);

			com_ptr<IProjectConfigGeneralProperties> generalProps;
			cfg.query<IProjectConfigProperties>()->get_GeneralProperties(&generalProps);
			generalProps->put_OutputDirectory(wil::make_bstr_failfast(L"%PROJECT_DIR%Out").get());
			generalProps->put_OutputFileType(OutputFileType::Sna);

			com_ptr<VxDTE::SolutionBuild> solutionBuild;
			sln->get_SolutionBuild(&solutionBuild);
			hr = solutionBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));
			long buildFailCount;
			hr = solutionBuild->get_LastBuildInfo(&buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(0l, buildFailCount);
			Assert::IsTrue(PathFileExists(wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj\\Out\\proj.sna").get()));
		}

		TEST_METHOD(BuildOutDirOutsideProjectDir)
		{
			HRESULT hr;

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildOutDirOutsideProjectDir");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"proj");
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			com_ptr<VxDTE::SolutionBuild> solutionBuild;
			sln->get_SolutionBuild(&solutionBuild);

			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);

			com_ptr<IProjectConfigGeneralProperties> generalProps;
			cfg.query<IProjectConfigProperties>()->get_GeneralProperties(&generalProps);

			auto buildIt = [&generalProps, &solutionBuild](const wchar_t* outputDirExpected)
				{
					// Sna
					generalProps->put_OutputFileType(OutputFileType::Sna);
					auto hr = solutionBuild->Build(VARIANT_TRUE);
					Assert::IsTrue(SUCCEEDED(hr));
					long buildFailCount;
					hr = solutionBuild->get_LastBuildInfo(&buildFailCount);
					Assert::IsTrue(SUCCEEDED(hr));
					Assert::AreEqual(0l, buildFailCount);
					auto outputFile = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s\\proj.sna", outputDirExpected);
					Assert::IsTrue(PathFileExists(outputFile.get()));
					Assert::IsTrue(DeleteFile(outputFile.get()));

					// Binary
					generalProps->put_OutputFileType(OutputFileType::Binary);
					hr = solutionBuild->Build(VARIANT_TRUE);
					Assert::IsTrue(SUCCEEDED(hr));
					hr = solutionBuild->get_LastBuildInfo(&buildFailCount);
					Assert::IsTrue(SUCCEEDED(hr));
					Assert::AreEqual(0l, buildFailCount);
					outputFile = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s\\proj.bin", outputDirExpected);
					Assert::IsTrue(PathFileExists(outputFile.get()));
					Assert::IsTrue(DeleteFile(outputFile.get()));
				};

			// Outside project dir but on same drive, full path.
			generalProps->put_OutputDirectory(wil::make_bstr_failfast(L"%PROJECT_DIR%..").get());
			buildIt(testPath.get());

			// Outside project dir but on same drive, relative path.
			generalProps->put_OutputDirectory(wil::make_bstr_failfast(L"..").get());
			buildIt(testPath.get());

			// Different drive
			auto testPathOtherDrive = MakeVolumeGuidPath(testPath.get());
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
	};
}
