
#include "pch.h"
#include "shared/com.h"
#include "UITests.h"

namespace UITests
{
	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (VxDTE::DTE2* dte, PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path);

	TEST_CLASS(FolderTests)
	{
		struct TD
		{
			wil::com_ptr_failfast<VxDTE::DTE2> dte;
			wil::unique_process_heap_string testDir;
			wil::unique_process_heap_string slnFilePath;
			wil::unique_process_heap_string projDir;
			wil::unique_process_heap_string projFilePath;
			wil::com_ptr_failfast<VxDTE::_Solution> sln;
			wil::com_ptr_failfast<VxDTE::Project> proj;

			TD()
			{
				dte = GetDefaultVSInstance();
				testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"FolderTest");
				Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));
				std::tie(sln, proj) = CreateSolutionAndProject(dte, testDir.get(), L"test", L"proj");
				slnFilePath = CombinePath(testDir.get(), L"test.sln");
				projDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\proj");
				projFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(projDir, L"\\proj.flx");
			}

			~TD()
			{
				sln->Close();
				RemoveDirectoryTree(testDir.get());
			}
		};

		TEST_METHOD(CreateFolderNameExists)
		{
			TD td;
			HRESULT hr;

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"Name", 1, TemplateEmptyFile, NULL, &addResult);
			Assert::AreEqual(S_OK, hr);

			VSITEMID fileItemId;
			hr = td.proj.query<IVsHierarchy>()->ParseCanonicalName(L"Name", &fileItemId);
			Assert::AreEqual(S_OK, hr);
			wil::unique_bstr fileMk;
			hr = td.proj.query<IVsProject>()->GetMkDocument(fileItemId, &fileMk);
			Assert::AreEqual(S_OK, hr);

			VSITEMID nextVSItemID;
			hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextVSItemID);
			Assert::AreEqual(S_OK, hr);

			auto tryAddFolder = [&td, nextVSItemID]
				{
					// Try to add a folder with the same name as the existing file, with same casing and then different casing.
					// We should get ERROR_ALREADY_EXISTS, and the hierarchy should not be changed.
					wil::unique_variant folderItemId;
					auto hr = td.proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
						OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"Name").addressof(), &folderItemId);
					Assert::IsFalse(SUCCEEDED(hr));
					VSITEMID nextNow;
					hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextNow);
					Assert::AreEqual(nextVSItemID, nextNow);

					// Try to add it again, with the same name and different casing. We should get ERROR_ALREADY_EXISTS.
					hr = td.proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
						OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"NAME").addressof(), &folderItemId);
					Assert::IsFalse(SUCCEEDED(hr));
					hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextNow);
					Assert::AreEqual(nextVSItemID, nextNow);

					// Try to add it again, with the same name and different casing. We should get ERROR_ALREADY_EXISTS.
					hr = td.proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
						OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"name").addressof(), &folderItemId);
					Assert::IsFalse(SUCCEEDED(hr));
					hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextNow);
					Assert::AreEqual(nextVSItemID, nextNow);
				};

			tryAddFolder();

			// Now delete the file from disk (keep it in the hierarchy), and try again to add the folder.
			BOOL bres = DeleteFile(fileMk.get());
			Assert::IsTrue(bres);
			tryAddFolder();
			RemoveDirectory(fileMk.get()); //remove any leftover directory (we're ok with the leftover for this test)

			// Now put the file back on disk and remove it from the hierarchy, and try again to add the folder.
			WriteFileCreateDirs(fileMk, "");
			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_RemoveFromProject, &fileItemId, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);
			tryAddFolder();

			// TODO: do the same with a directory in a directory. With the outer directory present and then missing.
		}

		TEST_METHOD(AddFileInFolder_DifferentFolderCase)
		{
			TD td;
			wil::unique_variant folderItemId;
			auto hr = td.proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
				OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"Name").addressof(), &folderItemId);
			Assert::AreEqual(S_OK, hr);

			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\NAME\\file.txt");
			WriteFileCreateDirs(filePath, "");

			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000u);
			VSADDRESULT addResult;
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(filePath.addressof()), nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
		}

		TEST_METHOD(AddFolderCommand)
		{
			TD td;
			HRESULT hr;

			DisableGeneratedFiles(td.proj);

			auto hier = td.proj.query<IVsUIHierarchy>();
			
			VSITEMID itemid;
			hier->ParseCanonicalName(L"file.asm", &itemid);
			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &itemid, DHO_SUPPRESS_UI);
			
			wil::unique_variant expandable;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Expandable, &expandable);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_FALSE, expandable.boolVal);

			auto sink = MakeTestHierarchyEventSink();
			VSCOOKIE hierEventsCookie;
			hr = hier->AdviseHierarchyEvents(sink, &hierEventsCookie);
			Assert::AreEqual(S_OK, hr);
			auto unadvise = wil::scope_exit([&hier, hierEventsCookie]() { hier->UnadviseHierarchyEvents(hierEventsCookie); });

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
			Assert::AreEqual(S_OK, hr);

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);
			Assert::AreEqual(S_OK, hr);
			wil::unique_variant childDisp;
			hr = hier->GetProperty(V_VSITEMID(&child), VSHPROPID_BrowseObject, &childDisp);
			Assert::AreEqual(S_OK, hr);
			com_ptr<IFolderNodeProperties> folderProps;
			hr = childDisp.pdispVal->QueryInterface(IID_PPV_ARGS(&folderProps));
			Assert::AreEqual(S_OK, hr);

			Assert::IsTrue(sink->PropertyChanged(VSITEMID_ROOT, VSHPROPID_Expandable));

			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Expandable, &expandable);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_TRUE, expandable.boolVal);
		}

		TEST_METHOD(AddFolderCommand_FolderExistsOnDisk)
		{
			TD td;
			auto hier = td.proj.query<IVsUIHierarchy>();
			auto hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);

			com_ptr<IVsHierarchyDeleteHandler3> dh;
			hr = hier->QueryInterface(IID_PPV_ARGS(&dh));
			Assert::IsTrue(SUCCEEDED(hr));
			dh->DeleteItems(1, DELITEMOP_RemoveFromProject, (VSITEMID*)&child.lVal, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(AddSubFolder_FolderMissingOnDisk)
		{
			TD td;
			HRESULT hr;

			auto hier = td.proj.query<IVsUIHierarchy>();

			DisableGeneratedFiles(td.proj);

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant folderItemId;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant folderSaveName;
			hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_SaveName, &folderSaveName);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hlocal_string folderPath;
			hr = wil::str_concat_nothrow (folderPath, td.projDir, L"\\", folderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(folderPath.get()));
			BOOL bres = RemoveDirectoryW(folderPath.get());
			Assert::IsTrue(bres);

			hr = hier->ExecCommand(V_VSITEMID(&folderItemId), &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(AddSubFolderCommand)
		{
			TD td;
			HRESULT hr;
			auto hier = td.proj.query<IVsUIHierarchy>();
			DisableGeneratedFiles(td.proj);

			hr = hier->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
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
			hr = wil::str_concat_nothrow (folderPath, td.projDir, L"\\", folderSaveName.bstrVal);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(folderPath.get()));

			wil::unique_variant expandable;
			hr = hier->GetProperty(V_VSITEMID(&folderItemId), VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_FALSE, expandable.boolVal);

			auto sink = MakeTestHierarchyEventSink();
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

			Assert::IsTrue(sink->PropertyChanged(V_VSITEMID(&folderItemId), VSHPROPID_Expandable));

			hr = hier->GetProperty(V_VSITEMID(&subFolderItemId), VSHPROPID_Expandable, &expandable);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BOOL, expandable.vt);
			Assert::AreEqual(VARIANT_TRUE, expandable.boolVal);
		}

		TEST_METHOD(RenameFolderAndCheckSorted)
		{
			TD td;
			HRESULT hr;
			auto hier = td.proj.query<IVsUIHierarchy>();
			DisableGeneratedFiles(td.proj);

			// Add folder named "NewFolder1".
			wil::unique_variant one;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &one);
			Assert::IsTrue(SUCCEEDED(hr));

			// Rename it to "B".
			hr = hier->SetProperty (V_VSITEMID(&one), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"B"));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_process_heap_string pathToB;
			wil::str_concat_nothrow(pathToB, td.projDir, L"\\B");
			Assert::IsTrue(PathFileExists(pathToB.get()));

			// Add folder named "NewFolder1".
			wil::unique_variant other;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &other);
			Assert::IsTrue(SUCCEEDED(hr));

			// First one should be "B", second one should be "NewFolder1"
			wil::unique_variant first, second;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &first);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&one), V_VSITEMID(&first));
			hr = hier->GetProperty(V_VSITEMID(&first), VSHPROPID_NextSibling, &second);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&other), V_VSITEMID(&second));

			// Rename the second one to "A".
			hr = hier->SetProperty (V_VSITEMID(&other), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"A"));
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_process_heap_string pathToA;
			wil::str_concat_nothrow(pathToA, td.projDir, L"\\A");
			Assert::IsTrue(PathFileExists(pathToA.get()));

			// First one should be "A", second one should be "B"
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &first);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&other), V_VSITEMID(&first));
			hr = hier->GetProperty(V_VSITEMID(&first), VSHPROPID_NextSibling, &second);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<VSITEMID>(V_VSITEMID(&one), V_VSITEMID(&second));
		}
	};
}
