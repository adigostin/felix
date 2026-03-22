
#include "pch.h"
#include "UITests.h"

namespace UITests
{
	TEST_CLASS(PackageTests)
	{
		TEST_METHOD(PackageUnloads)
		{
			HRESULT hr;

			auto dte = LaunchVS();
			auto closeVS = wil::scope_exit([&dte] { CloseVS(dte); });

			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"PackageUnloads\\");
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

			auto path = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"FelixPackageImpl");
			Assert::IsTrue(PathFileExists(path.get()));

			closeVS.reset();

			Assert::IsFalse(PathFileExists(path.get()));
		}
	};
}
