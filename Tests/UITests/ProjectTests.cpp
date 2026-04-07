
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
			RemoveDirectoryTree(classPath);
		}

		struct TestData
		{
			wil::unique_process_heap_string testDir;
			wil::com_ptr_failfast<VxDTE::_Solution> sln;
			wil::com_ptr_failfast<VxDTE::Project> dteproj;
			wil::com_ptr_failfast<IVsUIHierarchy> hier;
			wil::com_ptr_failfast<IVsProject> proj;

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
					wil::make_bstr_failfast(L"test").get(), VARIANT_TRUE, &dteproj);
				Assert::AreEqual(S_OK, hr);
				hr = sln->SaveAs(wil::make_bstr_failfast(L"test").get());
				Assert::AreEqual(S_OK, hr);

				hier = dteproj.query<IVsUIHierarchy>();
				proj = dteproj.query<IVsProject>();
			}

			~TestData()
			{
				sln->Close();
				RemoveDirectoryTree(testDir);
			}
		};

		TEST_METHOD(AddItemNew)
		{
			HRESULT hr;
			TestData td (L"AddItemNew", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"test1.asm", 1,  TemplateEmptyFile, nullptr, &addResult);
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

		TEST_METHOD(AddItemNew_NameAlreadyExists)
		{
			HRESULT hr;
			TestData td (L"AddItemNew_NameAlreadyExists", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"file.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			hr = td.proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"file.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), hr);
		}

		TEST_METHOD(AddItemNewToExistingFolder)
		{
			HRESULT hr;
			TestData td (L"AddItemNewToExistingFolder", TemplatePath_EmptyProject.get());

			wil::unique_variant folder;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folder);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->SetProperty (V_VSITEMID(&folder), VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"folder"));
			Assert::AreEqual(S_OK, hr);

			VSADDRESULT result;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&folder), oper, L"file.asm", 1, TemplateEmptyFile, nullptr, &result);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant folderFirstChild;
			hr = td.hier->GetProperty(V_VSITEMID(&folder), VSHPROPID_FirstChild, &folderFirstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, folderFirstChild.vt);

			wil::unique_variant fileName;
			hr = td.hier->GetProperty(V_VSITEMID(&folderFirstChild), VSHPROPID_SaveName, &fileName);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(L"file.asm", fileName.bstrVal);
		}

		TEST_METHOD(AddItemTwoFilesInTwoFolders)
		{
			HRESULT hr;
			TestData td (L"AddItemTwoFilesInTwoFolders", TemplatePath_EmptyProject.get());

			wil::unique_variant tf1;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &tf1);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf1.vt);
			hr = td.hier->SetProperty (V_VSITEMID(&tf1), VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"testfolder1"));
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant tf2;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &tf2);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf2.vt);
			hr = td.hier->SetProperty (V_VSITEMID(&tf2), VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"testfolder2"));
			Assert::AreEqual(S_OK, hr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&tf1), oper, L"file1.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&tf2), oper, L"file2.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			// First project child should be a folder.
			auto folder1ItemId = GetProperty_VSITEMID(td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			auto folder1Disp = GetProperty_Dispatch(td.hier, folder1ItemId, VSHPROPID_BrowseObject);
			//auto folder1 = folder1Disp.query<IFolderNode>();
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(td.hier, folder1ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"testfolder1", GetProperty_String(td.hier, folder1ItemId, VSHPROPID_SaveName).get());

			// Next project child should be the other folder.
			auto folder2ItemId = GetProperty_VSITEMID(td.hier, folder1ItemId, VSHPROPID_NextSibling);
			auto folder2Disp = GetProperty_Dispatch(td.hier, folder2ItemId, VSHPROPID_BrowseObject);
			//folder2 = folder2Disp.try_query<IFolderNode>();
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(td.hier, folder2ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"testfolder2", GetProperty_String(td.hier, folder2ItemId, VSHPROPID_SaveName).get());

			// There should be no more nodes after that.
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(td.hier, folder2ItemId, VSHPROPID_NextSibling));

			// First child in first folder should be our first file.
			auto file1ItemId = GetProperty_VSITEMID (td.hier, folder1ItemId, VSHPROPID_FirstChild);
			Assert::AreEqual<VSITEMID>(folder1ItemId, GetProperty_VSITEMID(td.hier, file1ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"file1.asm", GetProperty_String(td.hier, file1ItemId, VSHPROPID_SaveName).get());
			// and then no more nodes
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(td.hier, file1ItemId, VSHPROPID_NextSibling));

			// First child in second folder should be our second file.
			auto file2ItemId = GetProperty_VSITEMID (td.hier, folder2ItemId, VSHPROPID_FirstChild);
			Assert::AreEqual<VSITEMID>(folder2ItemId, GetProperty_VSITEMID(td.hier, file2ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"file2.asm", GetProperty_String(td.hier, file2ItemId, VSHPROPID_SaveName).get());
			// and then no more nodes
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(td.hier, file2ItemId, VSHPROPID_NextSibling));
		}

		TEST_METHOD(AddExistingItemWithHierarchyEventSinks)
		{
			HRESULT hr;
			TestData td (L"AddExistingItemWithHierarchyEventSinks", TemplatePath_EmptyProject.get());

			auto sink = MakeTestHierarchyEventSink();
			VSCOOKIE cookie;
			hr = td.hier->AdviseHierarchyEvents(sink, &cookie);
			Assert::AreEqual(S_OK, hr);
			auto unadvise = wil::scope_exit([cookie, &td] { td.hier->UnadviseHierarchyEvents(cookie); });

			auto fullPathSource = wil::str_concat_failfast<wil::unique_process_heap_string>(td.testDir, L"file.asm");
			wil::unique_hfile (CreateFile(fullPathSource.get(), GENERIC_WRITE, 0, 0, CREATE_NEW, 0, 0));
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fullPathSource.addressof()), nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<int>(ADDRESULT_Success & 0xFF, addResult);

			VSITEMID fileid;
			hr = td.hier->ParseCanonicalName(L"file.asm", &fileid);
			Assert::AreEqual(S_OK, hr);

			Assert::IsTrue(sink->ItemAdded(VSITEMID_ROOT, fileid));
		}

		TEST_METHOD(AddItemOpen)
		{
		}

		TEST_METHOD(AddItemNewWithSubfolder_TestHierarchyEvents)
		{
			// test that OnItemAdded and OnPropertyChanged are called
		}

		TEST_METHOD(AddItemOutsideOfProjectDir)
		{
		}

		TEST_METHOD(AddItem_DirtyAfter)
		{
		}

		TEST_METHOD(AddItemSort)
		{
			HRESULT hr;
			TestData td (L"AddItemSort", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);

			hr = td.proj->AddItem(VSITEMID_ROOT, oper, L"start.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj->AddItem(VSITEMID_ROOT, oper, L"lib.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant more;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &more);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->SetProperty (V_VSITEMID(&more), VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"More"));
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant evenMore;
			hr = td.hier->ExecCommand (V_VSITEMID(&more), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &evenMore);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->SetProperty (V_VSITEMID(&evenMore), VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"EvenMore"));
			Assert::AreEqual(S_OK, hr);

			hr = td.proj->AddItem(V_VSITEMID(&evenMore), oper, L"file.inc", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant generatedFiles;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &generatedFiles);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->SetProperty (V_VSITEMID(&generatedFiles), VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"GeneratedFiles"));
			Assert::AreEqual(S_OK, hr);

			hr = td.proj->AddItem(V_VSITEMID(&generatedFiles), oper, L"preinclude.inc", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj->AddItem(V_VSITEMID(&generatedFiles), oper, L"postinclude.inc", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			auto genFilesFolder = GetProperty_VSITEMID(td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, genFilesFolder);
			Assert::AreEqual(L"GeneratedFiles", GetProperty_String(td.hier, genFilesFolder, VSHPROPID_SaveName).get());

			auto postinc = GetProperty_VSITEMID(td.hier, genFilesFolder, VSHPROPID_FirstChild);
			Assert::AreEqual(0, _wcsicmp(L"postinclude.inc", GetProperty_String(td.hier, postinc, VSHPROPID_SaveName).get()));

			auto preinc = GetProperty_VSITEMID(td.hier, postinc, VSHPROPID_NextSibling);
			Assert::AreEqual(0, _wcsicmp(L"preinclude.inc", GetProperty_String(td.hier, preinc, VSHPROPID_SaveName).get()));

			auto moreFolder = GetProperty_VSITEMID(td.hier, genFilesFolder, VSHPROPID_NextSibling);
			Assert::AreEqual(L"More", GetProperty_String(td.hier, moreFolder, VSHPROPID_SaveName).get());

			auto evenMoreFolder = GetProperty_VSITEMID(td.hier, moreFolder, VSHPROPID_FirstChild);
			Assert::AreEqual(L"EvenMore", GetProperty_String(td.hier, evenMoreFolder, VSHPROPID_SaveName).get());

			auto fileinc = GetProperty_VSITEMID(td.hier, evenMoreFolder, VSHPROPID_FirstChild);
			Assert::AreEqual(L"file.inc", GetProperty_String(td.hier, fileinc, VSHPROPID_SaveName).get());

			auto libasm = GetProperty_VSITEMID(td.hier, moreFolder, VSHPROPID_NextSibling);
			Assert::AreEqual(L"lib.asm", GetProperty_String(td.hier, libasm, VSHPROPID_SaveName).get());

			auto startasm = GetProperty_VSITEMID(td.hier, libasm, VSHPROPID_NextSibling);
			Assert::AreEqual(L"start.asm", GetProperty_String(td.hier, startasm, VSHPROPID_SaveName).get());

			auto emp = GetProperty_VSITEMID(td.hier, startasm, VSHPROPID_NextSibling);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, emp);
		}

		TEST_METHOD(PutItemsTwoFilesOneInFolder)
		{
			/*
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file1 = MakeFileNode(L"file1.asm");
			auto file2 = MakeFileNode(L"testfolder/file2.asm");

			com_ptr<IFolderNode> folder;

			{
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, file1->GetItemId());
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, file2->GetItemId());

			SAFEARRAYBOUND bound = { .cElements = 2, .lLbound = 0 };
			auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
			LONG i = 0;
			SafeArrayPutElement (sa.get(), &i, file1.try_query<IDispatch>());
			i++;
			SafeArrayPutElement (sa.get(), &i, file2.try_query<IDispatch>());
			hr = hier.try_query<IProjectNodeProperties>()->put_Items(sa.get());
			Assert::IsTrue(SUCCEEDED(hr));

			// The project should have assigned them item ids
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file1->GetItemId());
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file2->GetItemId());

			// First project child should be a folder.
			auto folderItemId = GetProperty_VSITEMID(hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			auto folderDisp = GetProperty_Dispatch(hier, folderItemId, VSHPROPID_BrowseObject);
			folder = folderDisp.try_query<IFolderNode>();
			Assert::IsNotNull(folder.get());
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, folderItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"testfolder", GetProperty_String(hier, folderItemId, VSHPROPID_SaveName).get());

			// Next project child should be our first file.
			auto file1ItemId = GetProperty_VSITEMID(hier, folderItemId, VSHPROPID_NextSibling);
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"file1.asm", GetProperty_String(hier, file1ItemId, VSHPROPID_SaveName).get());

			// There should be no more files after that.
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_NextSibling));

			// First child in the folder should be our second file.
			auto file2ItemId = GetProperty_VSITEMID (hier, folderItemId, VSHPROPID_FirstChild);
			Assert::AreEqual<VSITEMID>(folderItemId, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_Parent));
			Assert::AreEqual(L"file2.asm", GetProperty_String(hier, file2ItemId, VSHPROPID_SaveName).get());

			// There should be no more files after that.
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_NextSibling));
			}

			Assert::AreEqual<ULONG>(0, hier.detach()->Release());
			Assert::AreEqual<ULONG>(0, folder.detach()->Release());
			Assert::AreEqual<ULONG>(0, file1.detach()->Release());
			Assert::AreEqual<ULONG>(0, file2.detach()->Release());
			*/
		}

		TEST_METHOD(PutItemsTwoDirsUnsorted)
		{
		}

		TEST_METHOD(PutItemsOneFileInOneSubfolder)
		{
		}

		TEST_METHOD(PutItemsOutsideOfProjectDir)
		{
		}

		TEST_METHOD(PutItemsFourFilesUnsorted)
		{
			HRESULT hr;
			TestData td (L"PutItemsFourFilesUnsorted", TemplatePath_EmptyProject.get());

			auto proj = td.proj.try_query<IVsProject>();
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);

			// Add to empty parent
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file3.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Insert in first pos
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file1.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Insert between the two above
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file2.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Add at the end
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file4.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// First project child should be file1.
			auto file1ItemId = GetProperty_VSITEMID(td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreEqual(L"file1.asm", GetProperty_String(td.hier, file1ItemId, VSHPROPID_SaveName).get());

			// Next should be file2.
			auto file2ItemId = GetProperty_VSITEMID(td.hier, file1ItemId, VSHPROPID_NextSibling);
			Assert::AreEqual(L"file2.asm", GetProperty_String(td.hier, file2ItemId, VSHPROPID_SaveName).get());

			// Then file3.
			auto file3ItemId = GetProperty_VSITEMID(td.hier, file2ItemId, VSHPROPID_NextSibling);
			Assert::AreEqual(L"file3.asm", GetProperty_String(td.hier, file3ItemId, VSHPROPID_SaveName).get());

			// And then file4.
			auto file4ItemId = GetProperty_VSITEMID(td.hier, file3ItemId, VSHPROPID_NextSibling);
			Assert::AreEqual(L"file4.asm", GetProperty_String(td.hier, file4ItemId, VSHPROPID_SaveName).get());
		} 

		TEST_METHOD(GetItemsPutItems_WithFolders)
		{
			HRESULT hr;
			auto testDir = str_concat(classPath, L"GetItemsPutItems_WithFolders\\");
			Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));

			wil::com_ptr_failfast<IUnknown> solution;
			hr = GetDefaultVSInstance()->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::AreEqual(S_OK, hr);
			auto sln = solution.query<VxDTE::_Solution>();
			hr = sln->Create(wil::make_bstr_failfast(testDir.get()).get(), wil::make_bstr_failfast(L"test").get());
			Assert::AreEqual(S_OK, hr);
			auto closesln = wil::scope_exit([&sln] { sln->Close(); });

			wil::com_ptr_failfast<VxDTE::Project> proj1;
			hr = sln->AddFromTemplate (
				wil::make_bstr_failfast(TemplatePath_EmptyProject.get()).get(),
				wil::make_bstr_failfast(testDir.get()).get(),
				wil::make_bstr_failfast(L"test.flx").get(), VARIANT_TRUE, &proj1);
			Assert::AreEqual(S_OK, hr);
			hr = sln->SaveAs(wil::make_bstr_failfast(L"test").get());
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant folder1;
			hr = proj1.query<IVsUIHierarchy>()->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folder1);
			Assert::AreEqual(S_OK, hr);
			hr = proj1.query<IVsUIHierarchy>()->SetProperty (V_VSITEMID(&folder1), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"folder"));
			Assert::AreEqual(S_OK, hr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = proj1.query<IVsProject>()->AddItem(V_VSITEMID(&folder1), oper, L"test.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			hr = proj1->Save(nullptr);
			hr = sln->Close(VARIANT_TRUE);
			Assert::AreEqual(S_OK, hr);

			// ------------------------------------------------

			auto projfn = str_concat(testDir, L"test.flx");
			wil::com_ptr_failfast<VxDTE::Project> proj2;
			hr = sln->AddFromFile (wil::make_bstr_failfast(projfn.get()).get(), VARIANT_TRUE, &proj2);
			Assert::AreEqual(S_OK, hr);

			auto folder2 = GetProperty_VSITEMID(proj2.query<IVsHierarchy>(), VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, folder2);
			auto folder2Name = GetProperty_String(proj2.query<IVsHierarchy>(), folder2, VSHPROPID_Name);
			Assert::AreEqual(L"folder", folder2Name.get());

			auto file2 = GetProperty_VSITEMID(proj2.query<IVsHierarchy>(), folder2, VSHPROPID_FirstChild);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file2);
			auto file2disp = GetProperty_Dispatch(proj2.query<IVsHierarchy>(), file2, VSHPROPID_BrowseObject);
			auto file2Props = file2disp.query<IFileNodeProperties>();
			wil::unique_bstr file2Path;
			hr = file2Props->get_Path(&file2Path);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(L"test.asm", file2Path.get());
		}

		TEST_METHOD(RemoveItemsFromRootNode)
		{
			HRESULT hr;
			TestData td (L"RemoveItemsFromRootNode", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"file.asm", 1,  TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);

			VSITEMID id;
			hr = td.hier->ParseCanonicalName(L"file.asm", &id);
			Assert::AreEqual(S_OK, hr);
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, id);

			auto dh = td.proj.query<IVsHierarchyDeleteHandler3>();
			hr = dh->DeleteItems (1, DELITEMOP_RemoveFromProject, &id, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant firstChild;
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&firstChild));
		}

		TEST_METHOD(RemoveItemsFromFolderNode)
		{
		}

		TEST_METHOD(RemoveNestedFoldersNoFiles)
		{
		}

		TEST_METHOD(RemoveItem_FirstOfTwo)
		{
			HRESULT hr;
			TestData td (L"RemoveItemsFromRootNode", TemplatePath_EmptyProject.get());

			auto path1 = wil::str_concat_failfast<wil::unique_process_heap_string>(td.testDir, L"file1.asm");
			WriteFileCreateDirs(path1, "");
			auto path2 = wil::str_concat_failfast<wil::unique_process_heap_string>(td.testDir, L"file2.asm");
			WriteFileCreateDirs(path2, "");
			const wchar_t* files[] = { path1.get(), path2.get() };
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", _countof(files), files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<int>(ADDRESULT_Success & 0xFF, addResult);

			wil::unique_variant file1ItemId;
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &file1ItemId);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant file2ItemId;
			hr = td.hier->GetProperty(V_VSITEMID(&file1ItemId), VSHPROPID_NextSibling, &file2ItemId);
			Assert::AreEqual(S_OK, hr);

			auto dh = td.proj.query<IVsHierarchyDeleteHandler3>();
			hr = dh->DeleteItems(1, DELITEMOP_RemoveFromProject, (VSITEMID*)&V_VSITEMID(&file1ItemId), DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant firstChildItemId;
			hr = td.hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChildItemId);
			Assert::AreEqual(S_OK, hr);

			Assert::AreEqual(V_VSITEMID(&file2ItemId), V_VSITEMID(&firstChildItemId));
		}

		TEST_METHOD(DeleteItems_File)
		{
			HRESULT hr;
			TestData td (L"DeleteItems_File", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.hier.try_query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"file.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			auto fileItemId = GetProperty_VSITEMID (td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);

			wil::unique_bstr fileMk;
			hr = td.proj->GetMkDocument(fileItemId, &fileMk);
			Assert::AreEqual(S_OK, hr);
			Assert::IsTrue(PathFileExists(fileMk.get()));

			hr = td.dteproj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &fileItemId, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);

			fileItemId = GetProperty_VSITEMID (td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, fileItemId);
			Assert::IsFalse(PathFileExists(fileMk.get()));
		}

		TEST_METHOD(DeleteItems_Folder)
		{
			HRESULT hr;
			TestData td (L"DeleteItems_Folder", TemplatePath_EmptyProject.get());

			// Add folder and check directory exists in file system.
			wil::unique_variant itemid;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &itemid);
			Assert::AreEqual(S_OK, hr);
			VSITEMID folderItemId = V_VSITEMID(&itemid);

			auto folderName = GetProperty_String (td.hier, folderItemId, VSHPROPID_SaveName);
			auto directoryFullPath = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s\\%s", td.testDir.get(), folderName.get());
			Assert::IsTrue(PathFileExists(directoryFullPath.get()));

			// Add file in folder and check it exists in file system.
			VSADDRESULT result;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj->AddItem (folderItemId, oper, L"file.asm", 1, TemplateEmptyFile, nullptr, &result);
			Assert::AreEqual(S_OK, hr);

			auto fileItemId = GetProperty_VSITEMID (td.hier, folderItemId, VSHPROPID_FirstChild);
			wil::unique_bstr fileMk;
			hr = td.proj->GetMkDocument(fileItemId, &fileMk);
			Assert::AreEqual(S_OK, hr);
			Assert::IsTrue(PathFileExists(fileMk.get()));

			// Delete folder.
			hr = td.dteproj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &folderItemId, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);

			// Check there's no node in hierarchy
			folderItemId = GetProperty_VSITEMID (td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, folderItemId);

			// Check there's no file or directory in the file system.
			Assert::IsFalse(PathFileExists(fileMk.get()));
			Assert::IsFalse(PathFileExists(directoryFullPath.get()));
		}

		TEST_METHOD(DeleteItems_FolderAndOneOfTwoMemberFiles)
		{
		}

		TEST_METHOD(RenameFilePresentOnFileSystem)
		{
			HRESULT hr;
			TestData td (L"RenameFilePresentOnFileSystem", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj->AddItem(VSITEMID_ROOT, oper, L"file.asm", 1,  TemplateEmptyFile, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<int>(ADDRESULT_Success & 0xFF, addResult);

			VSITEMID file;
			hr = td.hier->ParseCanonicalName(L"file.asm", &file);
			Assert::AreEqual(S_OK, hr);

			wil::unique_bstr oldFullPath;
			hr = td.proj->GetMkDocument(file, &oldFullPath);
			Assert::AreEqual(S_OK, hr);
			PathFindFileName(oldFullPath.get())[0] = 0;
			auto newFullPath = str_concat(oldFullPath, L"new.asm");
			Assert::IsFalse(PathFileExists(newFullPath.get()));

			hr = td.hier->SetProperty(file, VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant newSaveName;
			hr = td.hier->GetProperty(file, VSHPROPID_SaveName, &newSaveName);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(L"new.asm", newSaveName.bstrVal);

			Assert::IsTrue(PathFileExists(newFullPath.get()));
		}

		TEST_METHOD(RenameFileMissingOnFileSystem)
		{
		}

		TEST_METHOD(RenameFileMissingOnFileSystem_NewNameExists)
		{
		}

		TEST_METHOD(RenameFileAndCheckSorted)
		{
			HRESULT hr;
			TestData td (L"RenameFileAndCheckSorted", TemplatePath_EmptyProject.get());

			auto sink = MakeTestHierarchyEventSink();
			VSCOOKIE hecookie;
			hr = td.hier->AdviseHierarchyEvents(sink, &hecookie);
			Assert::AreEqual(S_OK, hr);
			auto unadvise = wil::scope_exit([hecookie, &td] { td.hier->UnadviseHierarchyEvents(hecookie); });

			// Add files named "B", "C", "D".
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj->AddItem (VSITEMID_ROOT, oper, L"B", 1, TemplateEmptyFile, nullptr, &addResult); Assert::AreEqual(S_OK, hr);
			hr = td.proj->AddItem (VSITEMID_ROOT, oper, L"C", 1, TemplateEmptyFile, nullptr, &addResult); Assert::AreEqual(S_OK, hr);
			hr = td.proj->AddItem (VSITEMID_ROOT, oper, L"D", 1, TemplateEmptyFile, nullptr, &addResult); Assert::AreEqual(S_OK, hr);

			// First one should be "B", second one should be "C", third "D"
			auto b = GetProperty_VSITEMID (td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreEqual(L"B", GetProperty_String(td.hier, b, VSHPROPID_SaveName).get());
			auto c = GetProperty_VSITEMID (td.hier, b, VSHPROPID_NextSibling);
			Assert::AreEqual(L"C", GetProperty_String(td.hier, c, VSHPROPID_SaveName).get());
			auto d = GetProperty_VSITEMID (td.hier, c, VSHPROPID_NextSibling);
			Assert::AreEqual(L"D", GetProperty_String(td.hier, d, VSHPROPID_SaveName).get());

			Assert::IsFalse(sink->ChildItemsInvalidated(VSITEMID_ROOT));

			// Rename "C" to "A".
			hr = td.hier->SetProperty (c, VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"A"));
			Assert::AreEqual(S_OK, hr);
			auto a = GetProperty_VSITEMID (td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreEqual(L"A", GetProperty_String(td.hier, a, VSHPROPID_SaveName).get());
			b = GetProperty_VSITEMID (td.hier, a, VSHPROPID_NextSibling);
			Assert::AreEqual(L"B", GetProperty_String(td.hier, b, VSHPROPID_SaveName).get());
			d = GetProperty_VSITEMID (td.hier, b, VSHPROPID_NextSibling);
			Assert::AreEqual(L"D", GetProperty_String(td.hier, d, VSHPROPID_SaveName).get());

			Assert::IsTrue(sink->ChildItemsInvalidated(VSITEMID_ROOT));

			// Rename "B" to "Z"
			hr = td.hier->SetProperty (b, VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"Z"));
			Assert::AreEqual(S_OK, hr);
			a = GetProperty_VSITEMID (td.hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
			Assert::AreEqual(L"A", GetProperty_String(td.hier, a, VSHPROPID_SaveName).get());
			d = GetProperty_VSITEMID (td.hier, a, VSHPROPID_NextSibling);
			Assert::AreEqual(L"D", GetProperty_String(td.hier, d, VSHPROPID_SaveName).get());
			auto z = GetProperty_VSITEMID (td.hier, d, VSHPROPID_NextSibling);
			Assert::AreEqual(L"Z", GetProperty_String(td.hier, z, VSHPROPID_SaveName).get());
		}

		TEST_METHOD(ProjectDirPropertyEndsWithBackslash)
		{
			HRESULT hr;
			TestData td (L"ProjectDirPropertyEndsWithBackslash", TemplatePath_EmptyProject.get());

			wil::unique_variant projDir;
			hr = td.hier->GetProperty (VSITEMID_ROOT, VSHPROPID_ProjectDir, &projDir);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BSTR, projDir.vt);
			Assert::AreEqual(L'\\', projDir.bstrVal[SysStringLen(projDir.bstrVal) - 1]);
		}
	};
}

