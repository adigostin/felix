
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

		struct TestData
		{
			wil::unique_process_heap_string testDir;
			wil::com_ptr_failfast<VxDTE::_Solution> sln;
			wil::com_ptr_failfast<VxDTE::Project> proj;

			TestData (const wchar_t* testName, const wchar_t* projTemplatePath)
			{
				HRESULT hr;
				testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(classPath, testName, L"\\");
				Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));

				wil::com_ptr_failfast<IUnknown> solution;
				hr = GetDefaultVSInstance()->get_Solution((VxDTE::Solution**)solution.addressof());
				Assert::AreEqual(S_OK, hr);
				sln = solution.query<VxDTE::_Solution>();
				hr = sln->Create(wil::make_bstr_failfast(testDir.get()).get(), wil::make_bstr_failfast(L"test").get());
				Assert::AreEqual(S_OK, hr);

				hr = sln->AddFromTemplate (
					wil::make_bstr_failfast(projTemplatePath).get(),
					wil::make_bstr_failfast(testDir.get()).get(),
					wil::make_bstr_failfast(L"test").get(), VARIANT_TRUE, &proj);
				Assert::AreEqual(S_OK, hr);
				hr = sln->SaveAs(wil::make_bstr_failfast(L"test").get());
				Assert::AreEqual(S_OK, hr);
			}

			~TestData()
			{
				sln->Close();
				RemoveDirectoryTree(testDir.get());
			}
		};

		TEST_METHOD(AddItemNew)
		{
			HRESULT hr;
			TestData td (L"AddItemNew", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"test1.asm", 1,  const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant firstChildItemId;
			hr = td.proj.query<IVsHierarchy>()->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChildItemId);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, firstChildItemId.vt);

			wil::unique_variant parentItemId;
			hr = td.proj.query<IVsHierarchy>()->GetProperty(V_VSITEMID(&firstChildItemId), VSHPROPID_Parent, &parentItemId);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, firstChildItemId.vt);
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, V_VSITEMID(&parentItemId));
		}

		TEST_METHOD(AddItemTwoFilesInTwoFolders)
		{
			HRESULT hr;
			TestData td (L"AddItemTwoFilesInTwoFolders", TemplatePath_EmptyProject.get());
			auto hier = td.proj.query<IVsUIHierarchy>();

			wil::unique_variant tf1;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &tf1);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf1.vt);
			hr = hier->SetProperty (V_VSITEMID(&tf1), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"testfolder1"));
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant tf2;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &tf2);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf2.vt);
			hr = hier->SetProperty (V_VSITEMID(&tf2), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"testfolder2"));
			Assert::AreEqual(S_OK, hr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			auto tmpl = const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof());
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&tf1), oper, L"file1.asm", 1, tmpl, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&tf2), oper, L"file2.asm", 1, tmpl, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			// First project child should be a folder.
			auto folder1ItemId = GetProperty_VSITEMID(hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			auto folder1Disp = GetProperty_Dispatch(hier, folder1ItemId, VSHPROPID_BrowseObject);
			//auto folder1 = folder1Disp.query<IFolderNode>();
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, folder1ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"testfolder1", GetProperty_String(hier, folder1ItemId, VSHPROPID_SaveName).get());

			// Next project child should be the other folder.
			auto folder2ItemId = GetProperty_VSITEMID(hier, folder1ItemId, VSHPROPID_NextSibling);
			auto folder2Disp = GetProperty_Dispatch(hier, folder2ItemId, VSHPROPID_BrowseObject);
			//folder2 = folder2Disp.try_query<IFolderNode>();
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, folder2ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"testfolder2", GetProperty_String(hier, folder2ItemId, VSHPROPID_SaveName).get());

			// There should be no more nodes after that.
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, folder2ItemId, VSHPROPID_NextSibling));

			// First child in first folder should be our first file.
			auto file1ItemId = GetProperty_VSITEMID (hier, folder1ItemId, VSHPROPID_FirstChild);
			Assert::AreEqual<VSITEMID>(folder1ItemId, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"file1.asm", GetProperty_String(hier, file1ItemId, VSHPROPID_SaveName).get());
			// and then no more nodes
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_NextSibling));

			// First child in second folder should be our second file.
			auto file2ItemId = GetProperty_VSITEMID (hier, folder2ItemId, VSHPROPID_FirstChild);
			Assert::AreEqual<VSITEMID>(folder2ItemId, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"file2.asm", GetProperty_String(hier, file2ItemId, VSHPROPID_SaveName).get());
			// and then no more nodes
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_NextSibling));
		}
	};
}

