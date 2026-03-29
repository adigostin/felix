
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
			wil::com_ptr_failfast<IVsUIHierarchy> hier;

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

				hier = proj.query<IVsUIHierarchy>();
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

		TEST_METHOD(AddItemNewToExistingFolder)
		{
			HRESULT hr;
			TestData td (L"AddItemNewToExistingFolder", TemplatePath_EmptyProject.get());

			wil::unique_variant folder;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folder);
			Assert::AreEqual(S_OK, hr);
			hr = td.hier->SetProperty (V_VSITEMID(&folder), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"folder"));
			Assert::AreEqual(S_OK, hr);

			const wchar_t* templateName = TemplatePath_EmptyFile.get();
			VSADDRESULT result;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&folder), oper, L"file.asm", 1, &templateName, nullptr, &result);
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
			hr = td.hier->SetProperty (V_VSITEMID(&tf1), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"testfolder1"));
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant tf2;
			hr = td.hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &tf2);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf2.vt);
			hr = td.hier->SetProperty (V_VSITEMID(&tf2), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"testfolder2"));
			Assert::AreEqual(S_OK, hr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			auto tmpl = const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof());
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&tf1), oper, L"file1.asm", 1, tmpl, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj.query<IVsProject>()->AddItem (V_VSITEMID(&tf2), oper, L"file2.asm", 1, tmpl, nullptr, &addResult);
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

			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };
			auto proj = td.proj.try_query<IVsProject>();
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);

			// Add to empty parent
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file3.asm", 1, templateasm, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Insert in first pos
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file1.asm", 1, templateasm, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Insert between the two above
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file2.asm", 1, templateasm, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Add at the end
			hr = proj->AddItem(VSITEMID_ROOT, oper, L"file4.asm", 1, templateasm, nullptr, &addResult);
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

		TEST_METHOD(RemoveItemsFromRootNode)
		{
			HRESULT hr;
			TestData td (L"RemoveItemsFromRootNode", TemplatePath_EmptyProject.get());

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"file.asm", 1,  const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), nullptr, &addResult);
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
			WriteFileOnDisk(path1.get(), "");
			auto path2 = wil::str_concat_failfast<wil::unique_process_heap_string>(td.testDir, L"file2.asm");
			WriteFileOnDisk(path2.get(), "");
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
	};
}

