
#include "pch.h"
#include "shared/com.h"
#include "UITests.h"

namespace UITests
{
	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (VxDTE::DTE2* dte, PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path);

	TEST_CLASS(ProjectTests)
	{
		static inline wil::unique_process_heap_string classPath;

		TEST_CLASS_INITIALIZE(ProjectTestsInit)
		{
			classPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"ProjectTests\\");
			Assert::IsTrue(CreateDirectory(classPath .get(), nullptr));
		}

		TEST_CLASS_CLEANUP(ProjectTestsCleanup)
		{
			RemoveDirectoryTree(classPath.get());
		}

		TEST_METHOD(AddItemNew)
		{
			HRESULT hr;
			auto dte = GetDefaultVSInstance();

			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(classPath, L"AddItemNew\\");
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
				wil::make_bstr_failfast(L"test").get(), VARIANT_TRUE, &proj);
			Assert::AreEqual(S_OK, hr);
			hr = sln->SaveAs(wil::make_bstr_failfast(L"test").get());
			Assert::AreEqual(S_OK, hr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"test1.asm", 1,  const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant firstChildItemId;
			hr = proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChildItemId);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, firstChildItemId.vt);

			wil::unique_variant parentItemId;
			hr = proj.query<IVsHierarchy>()->GetProperty(V_VSITEMID(&firstChildItemId), VSHPROPID_Parent, &parentItemId);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, firstChildItemId.vt);
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, V_VSITEMID(&parentItemId));
		}
	};
}

