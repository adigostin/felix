
#include "pch.h"
#include "shared/com.h"
#include "../FelixPackage/FelixPackage.h"
#include "../FelixPackage/Z80Xml.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(ProjectTests)
	{
	public:
		TEST_METHOD(PutItemsOneFile)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file = MakeFileNode(L"test.asm");
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, file->GetItemId());

			{
				SAFEARRAYBOUND bound = { .cElements = 1, .lLbound = 0 };
				auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
				Assert::IsTrue(SUCCEEDED(hr));
				LONG i = 0;
				SafeArrayPutElement (sa.get(), &i, file.try_query<IDispatch>());
				hr = hier.try_query<IProjectNodeProperties>()->put_Items(sa.get());
				Assert::IsTrue(SUCCEEDED(hr));

				Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file->GetItemId());

				auto pip = hier.try_query<IParentNode>();
				IUnknown* firstChild = wil::try_com_query_nothrow<IUnknown>(pip->FirstChild());
				Assert::AreEqual<void*>(file.try_query<IUnknown>().get(), firstChild);

				wil::unique_variant parentID;
				hr = file->GetProperty(VSHPROPID_Parent, &parentID);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, V_VSITEMID(&parentID));
			}

			ULONG refCount = hier.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);

			refCount = file.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);
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

		TEST_METHOD(AddItemTwoFilesInTwoFolders)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IFolderNode> folder1;
			com_ptr<IFolderNode> folder2;

			{
				wil::unique_variant tf1;
				hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &tf1);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf1.vt);
				hr = hier->SetProperty (V_VSITEMID(&tf1), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"testfolder1"));
				Assert::IsTrue(SUCCEEDED(hr));

				wil::unique_variant tf2;
				hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &tf2);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf2.vt);
				hr = hier->SetProperty (V_VSITEMID(&tf2), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"testfolder2"));
				Assert::IsTrue(SUCCEEDED(hr));

				LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };
				auto proj = hier.try_query<IVsProject>();
				hr = proj->AddItem (V_VSITEMID(&tf1), VSADDITEMOP_CLONEFILE, L"file1.asm", 1, templateasm, nullptr, nullptr);
				Assert::IsTrue(SUCCEEDED(hr));
				hr = proj->AddItem (V_VSITEMID(&tf2), VSADDITEMOP_CLONEFILE, L"file2.asm", 1, templateasm, nullptr, nullptr);
				Assert::IsTrue(SUCCEEDED(hr));

				// First project child should be a folder.
				auto folder1ItemId = GetProperty_VSITEMID(hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
				auto folder1Disp = GetProperty_Dispatch(hier, folder1ItemId, VSHPROPID_BrowseObject);
				folder1 = folder1Disp.try_query<IFolderNode>();
				Assert::IsNotNull(folder1.get());
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, folder1ItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"testfolder1", GetProperty_String(hier, folder1ItemId, VSHPROPID_SaveName).get());

				// Next project child should be the other folder.
				auto folder2ItemId = GetProperty_VSITEMID(hier, folder1ItemId, VSHPROPID_NextSibling);
				auto folder2Disp = GetProperty_Dispatch(hier, folder2ItemId, VSHPROPID_BrowseObject);
				folder2 = folder2Disp.try_query<IFolderNode>();
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

			Assert::AreEqual<ULONG>(0, hier.detach()->Release());
			Assert::AreEqual<ULONG>(0, folder1.detach()->Release());
			Assert::AreEqual<ULONG>(0, folder2.detach()->Release());
		}

		TEST_METHOD(PutItemsFourFilesUnsorted)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file1 = MakeFileNode(L"file1.asm");
			auto file2 = MakeFileNode(L"file2.asm");
			auto file3 = MakeFileNode(L"file3.asm");
			auto file4 = MakeFileNode(L"file4.asm");

			{
				LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };

				auto proj = hier.try_query<IVsProject>();
				// Add to empty parent
				hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file3.asm", 1, templateasm, nullptr, nullptr);
				Assert::IsTrue(SUCCEEDED(hr));

				// Insert in first pos
				hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file1.asm", 1, templateasm, nullptr, nullptr);
				Assert::IsTrue(SUCCEEDED(hr));

				// Insert between the two above
				hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file2.asm", 1, templateasm, nullptr, nullptr);
				Assert::IsTrue(SUCCEEDED(hr));

				// Add at the end
				hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file4.asm", 1, templateasm, nullptr, nullptr);
				Assert::IsTrue(SUCCEEDED(hr));

				// First project child should be file1.
				auto file1ItemId = GetProperty_VSITEMID(hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
				Assert::AreEqual(L"file1.asm", GetProperty_String(hier, file1ItemId, VSHPROPID_SaveName).get());

				// Next should be file2.
				auto file2ItemId = GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_NextSibling);
				Assert::AreEqual(L"file2.asm", GetProperty_String(hier, file2ItemId, VSHPROPID_SaveName).get());

				// Then file3.
				auto file3ItemId = GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_NextSibling);
				Assert::AreEqual(L"file3.asm", GetProperty_String(hier, file3ItemId, VSHPROPID_SaveName).get());

				// And then file4.
				auto file4ItemId = GetProperty_VSITEMID(hier, file3ItemId, VSHPROPID_NextSibling);
				Assert::AreEqual(L"file4.asm", GetProperty_String(hier, file4ItemId, VSHPROPID_SaveName).get());
			}

			Assert::AreEqual<ULONG>(0, hier.detach()->Release());
			Assert::AreEqual<ULONG>(0, file1.detach()->Release());
			Assert::AreEqual<ULONG>(0, file2.detach()->Release());
			Assert::AreEqual<ULONG>(0, file3.detach()->Release());
			Assert::AreEqual<ULONG>(0, file4.detach()->Release());
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

		TEST_METHOD(RemoveItemsFromRootNode)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file = MakeFileNode(L"file.asm");
			SAFEARRAYBOUND bound = { .cElements = 1, .lLbound = 0 };
			auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
			LONG i = 0;
			SafeArrayPutElement (sa.get(), &i, file.try_query<IDispatch>());
			hr = hier.try_query<IProjectNodeProperties>()->put_Items(sa.get());
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsHierarchyDeleteHandler3> dh;
			hr = hier->QueryInterface(IID_PPV_ARGS(&dh));
			Assert::IsTrue(SUCCEEDED(hr));

			VSITEMID id = file->GetItemId();
			hr = dh->DeleteItems (1, DELITEMOP_RemoveFromProject, &id, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant firstChild;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChild);
			Assert::IsTrue(SUCCEEDED(hr));
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
			com_ptr<IVsSolution> sol;
			auto hr = serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol)); Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsUIHierarchy> hier;
			hr = sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), tempPath, L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier)); Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsProject> proj;
			hr = hier->QueryInterface(IID_PPV_ARGS(&proj)); Assert::IsTrue(SUCCEEDED(hr));

			auto path1 = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"file1.asm");
			auto path2 = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"file2.asm");
			const wchar_t* files[] = { path1.get(), path2.get() };
			VSADDRESULT addResult;
			hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, _countof(files), files, nullptr, &addResult); Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<DWORD>(ADDRESULT_Success, addResult);

			wil::unique_variant file1ItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &file1ItemId); Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant file2ItemId;
			hr = hier->GetProperty(V_VSITEMID(&file1ItemId), VSHPROPID_NextSibling, &file2ItemId); Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsHierarchyDeleteHandler3> dh;
			hr = hier->QueryInterface(IID_PPV_ARGS(&dh)); Assert::IsTrue(SUCCEEDED(hr));
			hr = dh->DeleteItems(1, DELITEMOP_RemoveFromProject, (VSITEMID*)&V_VSITEMID(&file1ItemId), DHO_SUPPRESS_UI); Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant firstChildItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &firstChildItemId); Assert::IsTrue(SUCCEEDED(hr));

			Assert::AreEqual(V_VSITEMID(&file2ItemId), V_VSITEMID(&firstChildItemId));
		}

		TEST_METHOD(AddFolderCommand)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant expandable;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_FALSE, expandable.boolVal);

			bool expandableChanged = false;
			auto propChanged = [&expandableChanged](VSITEMID itemid, VSHPROPID propid, DWORD flags)
				{
					if (itemid == VSITEMID_ROOT && propid == VSHPROPID_Expandable)
						expandableChanged = true;
				};

			auto sink = MakeMockHierarchyEventSink(std::move(propChanged));
			VSCOOKIE hierEventsCookie;
			hr = hier->AdviseHierarchyEvents(sink, &hierEventsCookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([hier=hier.get(), hierEventsCookie]() { hier->UnadviseHierarchyEvents(hierEventsCookie); });

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant childDisp;
			hr = hier->GetProperty(V_VSITEMID(&child), VSHPROPID_BrowseObject, &childDisp);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IFolderNodeProperties> folderProps;
			hr = childDisp.pdispVal->QueryInterface(IID_PPV_ARGS(&folderProps));
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsTrue(expandableChanged);

			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_TRUE, expandable.boolVal);
		}

		TEST_METHOD(AddSubFolderCommand)
		{
			HRESULT hr;

			wchar_t projDir[MAX_PATH];
			swprintf_s(projDir, L"%s\\AddSubFolderCommand\0", &tempPath[0]);
			if (PathFileExists(projDir))
			{
				SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = projDir, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
				int ires = SHFileOperation(&file_op);
				Assert::AreEqual(0, ires);
			}
			BOOL bres = CreateDirectory(projDir, 0);
			Assert::IsTrue(bres);

			com_ptr<IVsUIHierarchy> hier;
			hr = MakeProjectNode (nullptr, projDir, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant folderItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant folderDisp;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_BrowseObject, &folderDisp);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant folderSaveName;
			hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_SaveName, &folderSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hlocal_string folderPath;
			hr = wil::str_concat_nothrow (folderPath, projDir, L"\\", folderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(folderPath.get()));

			wil::unique_variant expandable;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_FALSE, expandable.boolVal);

			bool expandableChanged = false;
			auto propChanged = [&expandableChanged, id=V_VSITEMID(&folderItemId)](VSITEMID itemid, VSHPROPID propid, DWORD flags)
				{
					if (itemid == id && propid == VSHPROPID_Expandable)
						expandableChanged = true;
				};
			auto sink = MakeMockHierarchyEventSink(std::move(propChanged));
			VSCOOKIE hierEventsCookie;
			hr = hier->AdviseHierarchyEvents(sink, &hierEventsCookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([hier=hier.get(), hierEventsCookie]() { hier->UnadviseHierarchyEvents(hierEventsCookie); });

			hr = hier->ExecCommand(V_VSITEMID(&folderItemId), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant subFolderItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &subFolderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant subFolderDisp;
			hr = hier->GetProperty(V_VSITEMID(&subFolderItemId), VSHPROPID_BrowseObject, &subFolderDisp);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IFolderNodeProperties> folderProps;
			hr = subFolderDisp.pdispVal->QueryInterface(IID_PPV_ARGS(&folderProps));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant subFolderSaveName;
			hier->GetProperty(V_VSITEMID(&subFolderItemId), VSHPROPID_SaveName, &subFolderSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hlocal_string subFolderPath;
			hr = wil::str_concat_nothrow (subFolderPath, folderPath, L"\\", subFolderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(subFolderPath.get()));

			Assert::IsTrue(expandableChanged);

			hr = hier->GetProperty(V_VSITEMID(&subFolderItemId), VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_TRUE, expandable.boolVal);
		}

		TEST_METHOD(AddFolderCommand_FolderExistsInHier)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(AddFolderCommand_FolderExistsOnDisk)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);

			com_ptr<IVsHierarchyDeleteHandler3> dh;
			hr = hier->QueryInterface(IID_PPV_ARGS(&dh));
			Assert::IsTrue(SUCCEEDED(hr));
			dh->DeleteItems(1, DELITEMOP_RemoveFromProject, (VSITEMID*)&child.lVal, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(AddSubFolder_FolderMissingOnDisk)
		{
			com_ptr<IVsSolution> sol;
			auto hr = serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));
			Assert::IsTrue(SUCCEEDED(hr));

			static const wchar_t ProjFileName[] = L"TestProject.flx";
			com_ptr<IVsUIHierarchy> hier;
			hr = sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), tempPath, ProjFileName, CPF_CLONEFILE, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant folderItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant folderSaveName;
			hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_SaveName, &folderSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hlocal_string folderPath;
			hr = wil::str_concat_nothrow (folderPath, tempPath, L"\\", folderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(folderPath.get()));
			BOOL bres = RemoveDirectoryW(folderPath.get());
			Assert::IsTrue(bres);

			hr = hier->ExecCommand(V_VSITEMID(&folderItemId), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(AddItemNew)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsProject2> proj;
			hr = hier->QueryInterface(&proj);
			Assert::IsTrue(SUCCEEDED(hr));

			wchar_t templateFullPath[MAX_PATH];
			PathCombine(templateFullPath, tempPath, L"template.asm");
			HANDLE h = CreateFile(templateFullPath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
			Assert::AreNotEqual(INVALID_HANDLE_VALUE, h);
			CloseHandle(h);

			const wchar_t* templateName = templateFullPath;
			VSADDRESULT result;
			hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"test.asm", 1, &templateName, nullptr, &result);
			Assert::IsTrue(SUCCEEDED(hr));

			auto pip = hier.try_query<IParentNode>();
			com_ptr<IChildNode> firstChild = pip->FirstChild();
			Assert::IsNotNull(firstChild.get());
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, firstChild->GetItemId());
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, firstChild->GetItemId(), VSHPROPID_Parent));

			pip.reset();
			proj.reset();

			ULONG refCount = hier.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);

			refCount = firstChild.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);
		}

		TEST_METHOD(AddItemNewToExistingFolder)
		{
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\AddItemNewToExistingFolder");

			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, testPath.get(), nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant folder;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &folder);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->SetProperty (V_VSITEMID(&folder), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"folder"));
			Assert::IsTrue(SUCCEEDED(hr));

			{
				com_ptr<IVsProject2> proj;
				hr = hier->QueryInterface(&proj);
				Assert::IsTrue(SUCCEEDED(hr));

				wchar_t templateFullPath[MAX_PATH];
				PathCombine(templateFullPath, testPath.get(), L"template.asm");
				HANDLE h = CreateFile(templateFullPath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
				Assert::AreNotEqual(INVALID_HANDLE_VALUE, h);
				CloseHandle(h);

				const wchar_t* templateName = templateFullPath;
				VSADDRESULT result;
				hr = proj->AddItem (V_VSITEMID(&folder), VSADDITEMOP_CLONEFILE, L"file.asm", 1, &templateName, nullptr, &result);
				Assert::IsTrue(SUCCEEDED(hr));

				wil::unique_variant folderFirstChild;
				hr = hier->GetProperty(V_VSITEMID(&folder), VSHPROPID_FirstChild, &folderFirstChild);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual<VARTYPE>(VT_VSITEMID, folderFirstChild.vt);
				
				wil::unique_variant fileName;
				hr = hier->GetProperty(V_VSITEMID(&folderFirstChild), VSHPROPID_SaveName, &fileName);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual(L"file.asm", fileName.bstrVal);
			}

			ULONG refCount = hier.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);
		}

		TEST_METHOD(AddExistingItemWithHierarchyEventSinks)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			auto sink = MakeMockHierarchyEventSink(nullptr);
			VSCOOKIE cookie;
			hr = hier->AdviseHierarchyEvents(sink, &cookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([cookie, hier=hier.get()] { hier->UnadviseHierarchyEvents(cookie); });

			auto fullPathSource = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine(fullPathSource.get(), tempPath, L"file.asm");
			VSADDRESULT result;
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fullPathSource.addressof(), nullptr, &result);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<DWORD>(ADDRESULT_Success, result);
		}

		TEST_METHOD(AddItemNew_NameAlreadyExists)
		{
			com_ptr<IVsSolution> sol;
			auto hr = serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));
			Assert::IsTrue(SUCCEEDED(hr));

			static const wchar_t ProjFileName[] = L"TestProject.flx";
			com_ptr<IVsUIHierarchy> hier;
			hr = sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), tempPath, ProjFileName, CPF_CLONEFILE, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			WriteFileOnDisk(tempPath, L"template.asm");
			wil::unique_process_heap_string templateFileFullPath;
			wil::str_concat_nothrow(templateFileFullPath, tempPath, L"\\template.asm");
			auto proj = hier.try_query<IVsProject>();
			VSADDRESULT addResult;
			hr = proj->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file.asm", 1, (LPCOLESTR*)templateFileFullPath.addressof(), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file.asm", 1, (LPCOLESTR*)templateFileFullPath.addressof(), nullptr, &addResult);
			Assert::AreEqual<HRESULT>(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), hr);
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

		TEST_METHOD(DirtyAfterAddItem)
		{
		}

		TEST_METHOD(GetItemsPutItems_WithFolders)
		{
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\GetItemsPutItems_WithFolders");
			com_ptr<IVsUIHierarchy> hier1;
			auto hr = MakeProjectNode (nullptr, testPath.get(), nullptr, 0, IID_PPV_ARGS(&hier1));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant folder1;
			hr = hier1->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &folder1);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier1->SetProperty (V_VSITEMID(&folder1), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"folder"));
			Assert::IsTrue(SUCCEEDED(hr));

			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };
			hr = hier1.try_query<IVsProject>()->AddItem(V_VSITEMID(&folder1), VSADDITEMOP_CLONEFILE, L"test.asm", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// ------------------------------------------------

			auto stream = com_ptr(SHCreateMemStream(nullptr, 0));
			hr = SaveToXml(hier1.try_query<IProjectNodeProperties>(), L"Temp", 0, stream);
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsHierarchy> hier2;
			hr = MakeProjectNode (nullptr, testPath.get(), nullptr, 0, IID_PPV_ARGS(&hier2));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = stream->Seek({ 0 }, STREAM_SEEK_SET, nullptr);
			hr = LoadFromXml(hier2.try_query<IProjectNodeProperties>(), L"Temp", stream);
			Assert::IsTrue(SUCCEEDED(hr));

			// ---------------------

			auto pip2 = hier2.try_query<IParentNode>();
			Assert::IsNotNull(pip2->FirstChild());
			auto folder2 = wil::try_com_query_nothrow<IFolderNode>(pip2->FirstChild());
			Assert::IsNotNull(pip2->FirstChild());
			wil::unique_bstr folder2Name;
			hr = folder2.try_query<IFolderNodeProperties>()->get_Name(&folder2Name);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"folder", folder2Name.get());

			auto file2 = folder2.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(file2);
			wil::unique_bstr file2Path;
			hr = wil::try_com_query_nothrow<IFileNodeProperties>(file2)->get_Path(&file2Path);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"test.asm", file2Path.get());
		}

		TEST_METHOD(AddItemSort)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto proj = hier.try_query<IVsProject>();
			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };
			hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"start.asm", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"lib.asm", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant more;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &more);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->SetProperty (V_VSITEMID(&more), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"More"));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant evenMore;
			hr = hier->ExecCommand (V_VSITEMID(&more), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &evenMore);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->SetProperty (V_VSITEMID(&evenMore), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"EvenMore"));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj->AddItem(V_VSITEMID(&evenMore), VSADDITEMOP_CLONEFILE, L"file.inc", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant generatedFiles;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &generatedFiles);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->SetProperty (V_VSITEMID(&generatedFiles), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"GeneratedFiles"));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj->AddItem(V_VSITEMID(&generatedFiles), VSADDITEMOP_CLONEFILE, L"preinclude.inc", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj->AddItem(V_VSITEMID(&generatedFiles), VSADDITEMOP_CLONEFILE, L"postinclude.inc", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));


			auto c = hier.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(c);
			auto genFilesFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(genFilesFolder.get());
			Assert::AreEqual(L"GeneratedFiles", GetProperty_String(hier, genFilesFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = genFilesFolder.try_query<IParentNode>()->FirstChild();
			auto postinc = wil::try_com_query_nothrow<IFileNode>(c);
			Assert::IsNotNull(postinc.get());
			Assert::AreEqual(L"postinclude.inc", GetProperty_String(hier, postinc->GetItemId(), VSHPROPID_SaveName).get());

			auto preinc = wil::try_com_query_nothrow<IFileNode>(postinc->Next());
			Assert::IsNotNull(preinc.get());
			Assert::AreEqual(L"preinclude.inc", GetProperty_String(hier, preinc->GetItemId(), VSHPROPID_SaveName).get());

			c = genFilesFolder->Next();
			Assert::IsNotNull(c);
			auto moreFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(moreFolder.get());
			Assert::AreEqual(L"More", GetProperty_String(hier, moreFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = moreFolder.try_query<IParentNode>()->FirstChild();
			auto evenMoreFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(evenMoreFolder.get());
			Assert::AreEqual(L"EvenMore", GetProperty_String(hier, evenMoreFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = evenMoreFolder.try_query<IParentNode>()->FirstChild();
			auto fileinc = wil::try_com_query_nothrow<IFileNode>(c);
			Assert::IsNotNull(fileinc.get());
			Assert::AreEqual(L"file.inc", GetProperty_String(hier, fileinc->GetItemId(), VSHPROPID_SaveName).get());

			auto libasm = wil::try_com_query_nothrow<IFileNode>(moreFolder->Next());
			Assert::IsNotNull(libasm.get());
			Assert::AreEqual(L"lib.asm", GetProperty_String(hier, libasm->GetItemId(), VSHPROPID_SaveName).get());

			auto startasm = wil::try_com_query_nothrow<IFileNode>(libasm->Next());
			Assert::IsNotNull(startasm.get());
			Assert::AreEqual(L"start.asm", GetProperty_String(hier, startasm->GetItemId(), VSHPROPID_SaveName).get());

			Assert::IsNull(startasm->Next());
		}

		TEST_METHOD(RenameFilePresentOnFileSystem)
		{
			static const char xml[] = ""
				"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
				"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">"
				"<Configurations>"
				"<Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />"
				"</Configurations>"
				"<Items>"
				"<File Path=\"file.asm\" BuildTool=\"Assembler\" />"
				"</Items>"
				"</Z80Project>";

			auto s = SHCreateMemStream((BYTE*)xml, sizeof(xml) - 1);
			com_ptr<IStream> stream;
			stream.attach(s);

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = LoadFromXml(hier.try_query<IProjectNodeProperties>(), ProjectElementName, stream);
			Assert::IsTrue(SUCCEEDED(hr));

			auto file = hier.try_query<IParentNode>()->FirstChild();

			wil::unique_bstr oldFullPath;
			hr = hier.try_query<IVsProject>()->GetMkDocument(file->GetItemId(), &oldFullPath);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_hfile handle (CreateFile(oldFullPath.get(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			Assert::IsTrue(handle.is_valid());
			handle.reset();

			wchar_t newFullPath[MAX_PATH];
			PathCombine(newFullPath, tempPath, L"new.asm");
			BOOL bres = DeleteFile(newFullPath);
			Assert::IsTrue(bres || GetLastError() == ERROR_FILE_NOT_FOUND);

			hr = hier->SetProperty(file->GetItemId(), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant newSaveName;
			hr = hier->GetProperty(file->GetItemId(), VSHPROPID_SaveName, &newSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"new.asm", newSaveName.bstrVal);

			Assert::IsTrue(PathFileExists(newFullPath));
			DeleteFile(newFullPath);
		}

		TEST_METHOD(RenameFileMissingOnFileSystem)
		{
		}

		TEST_METHOD(RenameFileMissingOnFileSystem_NewNameExists)
		{
		}

		TEST_METHOD(RenameFolderAndCheckSorted)
		{
			HRESULT hr;

			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"RenameFolderAndCheckSorted");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));

			// Add folder named "NewFolder1".
			wil::unique_variant one;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &one);
			Assert::IsTrue(SUCCEEDED(hr));

			// Rename it to "B".
			hr = hier->SetProperty (V_VSITEMID(&one), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"B"));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_process_heap_string pathToB;
			wil::str_concat_nothrow(pathToB, testPath, L"\\B");
			Assert::IsTrue(PathFileExists(pathToB.get()));

			// Add folder named "NewFolder1".
			wil::unique_variant other;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &other);
			Assert::IsTrue(SUCCEEDED(hr));

			// First one should be "B", second one should be "NewFolder1"
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&one), hier.try_query<IParentNode>()->FirstChild()->GetItemId());
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&other), hier.try_query<IParentNode>()->FirstChild()->Next()->GetItemId());

			// Rename the second one to "A".
			hr = hier->SetProperty (V_VSITEMID(&other), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"A"));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_process_heap_string pathToA;
			wil::str_concat_nothrow(pathToA, testPath, L"\\A");
			Assert::IsTrue(PathFileExists(pathToA.get()));

			// First one should be "A", second one should be "B"
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&other), hier.try_query<IParentNode>()->FirstChild()->GetItemId());
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&one), hier.try_query<IParentNode>()->FirstChild()->Next()->GetItemId());
		}

		TEST_METHOD(RenameFileAndCheckSorted)
		{
			HRESULT hr;

			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"RenameFileAndCheckSorted\\");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));
			auto hierAsParent = hier.try_query<IParentNode>();

			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };

			// Add file named "B".
			hr = hier.try_query<IVsProject>()->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"B", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// Add file named "C".
			hr = hier.try_query<IVsProject>()->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"C", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// First one should be "B", second one should be "C"
			wil::unique_bstr name;
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild())->get_Path(&name);
			Assert::AreEqual(L"B", name.get());
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild()->Next())->get_Path(&name);
			Assert::AreEqual(L"C", name.get());

			// Rename second one to "A".
			hr = hier->SetProperty (hierAsParent->FirstChild()->Next()->GetItemId(), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"A"));
			Assert::IsTrue(SUCCEEDED(hr));

			// First one should be "A", second one should be "B"
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild())->get_Path(&name);
			Assert::AreEqual(L"A", name.get());
			wil::try_com_query_nothrow<IFileNodeProperties>(hierAsParent->FirstChild()->Next())->get_Path(&name);
			Assert::AreEqual(L"B", name.get());

			Assert::Fail(L"TODO: check notifications");
		}

		TEST_METHOD(DeleteItems_File)
		{
			HRESULT hr;
			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"DeleteItems_File");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));

			hr = hier.try_query<IVsProject>()->AddItem (VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file.asm", 1, (LPCOLESTR*)TemplatePath_EmptyFile.addressof(), nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant fileItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_bstr fileMk;
			hr = hier.try_query<IVsProject>()->GetMkDocument(V_VSITEMID(&fileItemId), &fileMk);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(fileMk.get()));

			hr = hier.try_query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, (VSITEMID*)&V_VSITEMID(&fileItemId), DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&fileItemId));
			Assert::IsFalse(PathFileExists(fileMk.get()));
		}

		TEST_METHOD(DeleteItems_Folder)
		{
			HRESULT hr;
			com_ptr<IVsSolution> sol;
			serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"DeleteItems_Folder");

			com_ptr<IVsUIHierarchy> hier;
			sol->CreateProject(FelixProjectType, TemplatePath_EmptyProject.get(), testPath.get(), L"TestProject.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));

			// Add folder and check directory exists in file system.
			wil::unique_variant folderItemId;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &folderItemId);
			wil::unique_variant folderName;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_SaveName, &folderName);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BSTR, folderName.vt);
			auto directoryFullPath = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s\\%s", testPath.get(), folderName.bstrVal);
			Assert::IsTrue(PathFileExists(directoryFullPath.get()));

			// Add file in folder and check it exists in file system.
			hr = hier.try_query<IVsProject>()->AddItem (V_VSITEMID(&folderItemId), VSADDITEMOP_CLONEFILE, L"file.asm", 1, (LPCOLESTR*)TemplatePath_EmptyFile.addressof(), nullptr, nullptr);
			wil::unique_variant fileItemId;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_FirstChild, &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_bstr fileMk;
			hr = hier.try_query<IVsProject>()->GetMkDocument(V_VSITEMID(&fileItemId), &fileMk);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(fileMk.get()));

			// Delete folder.
			hr = hier.try_query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, (VSITEMID*)&V_VSITEMID(&folderItemId), DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			// Check there's no node in hierarchy
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, V_VSITEMID(&folderItemId));

			// Check there's no file or directory in the file system.
			Assert::IsFalse(PathFileExists(fileMk.get()));
			Assert::IsFalse(PathFileExists(directoryFullPath.get()));
		}

		TEST_METHOD(DeleteItems_FolderAndOneOfTwoMemberFiles)
		{
		}

		TEST_METHOD(ProjectDirPropertyEndsWithBackslash)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant projDir;
			hr = hier->GetProperty (VSITEMID_ROOT, VSHPROPID_ProjectDir, &projDir);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BSTR, projDir.vt);
			Assert::AreEqual(L'\\', projDir.bstrVal[SysStringLen(projDir.bstrVal) - 1]);
		}
	};
}
