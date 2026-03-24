
#include "pch.h"
#include "UITests.h"

namespace UITests
{
	TEST_CLASS(PackageTests)
	{
		TEST_METHOD(PackageUnloads)
		{
			HRESULT hr;
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"PackageUnloads\\");
			auto dte = LaunchVS(testDir.get());
			auto closeVS = wil::scope_exit([&dte] { CloseVS(dte); });

			Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));
			wil::com_ptr_failfast<IUnknown> solution;
			hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::AreEqual(S_OK, hr);
			auto sln = solution.query<VxDTE::_Solution>();
			hr = sln->Create(wil::make_bstr_failfast(testDir.get()).get(), wil::make_bstr_failfast(L"test").get());
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Project> proj;
			hr = sln->AddFromTemplate (
				wil::make_bstr_failfast(TemplatePath_EmptyProject.get()).get(),
				wil::make_bstr_failfast(testDir.get()).get(),
				wil::make_bstr_failfast(L"proj.flx").get(), VARIANT_FALSE, &proj);
			Assert::AreEqual(S_OK, hr);

			auto path = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"FelixPackageImpl");
			Assert::IsTrue(PathFileExists(path.get()));

			closeVS.reset();

			Assert::IsFalse(PathFileExists(path.get()));
		}

		TEST_METHOD(ProjectUnloads)
		{
			HRESULT hr;
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"ProjectUnloads\\");
			auto dte = LaunchVS(testDir.get());
			auto closeVS = wil::scope_exit([&dte] { CloseVS(dte); });

			auto path = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"ProjectNode");
			Assert::IsFalse(PathFileExists(path.get()));

			{
				Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));
				wil::com_ptr_failfast<IUnknown> solution;
				hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
				Assert::AreEqual(S_OK, hr);
				auto sln = solution.query<VxDTE::_Solution>();
				hr = sln->Create(wil::make_bstr_failfast(testDir.get()).get(), wil::make_bstr_failfast(L"test").get());
				Assert::AreEqual(S_OK, hr);

				wil::com_ptr_failfast<VxDTE::Project> proj;
				hr = sln->AddFromTemplate (
					wil::make_bstr_failfast(TemplatePath_EmptyProject.get()).get(),
					wil::make_bstr_failfast(testDir.get()).get(),
					wil::make_bstr_failfast(L"proj.flx").get(), VARIANT_FALSE, &proj);
				Assert::AreEqual(S_OK, hr);
			}

			Assert::IsTrue(PathFileExists(path.get()));

			closeVS.reset();

			Assert::IsFalse(PathFileExists(path.get()));
		}
	};
}
