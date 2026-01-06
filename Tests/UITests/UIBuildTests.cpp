
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
	extern com_ptr<VxDTE::_DTE> dte;
	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount);

	TEST_CLASS(UIBuildTests)
	{
	public:
		TEST_METHOD(BuildProject)
		{
			HRESULT hr;

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildProject");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

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
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

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
			Assert::AreEqual(1l, errorCount);
		}
	};
}
