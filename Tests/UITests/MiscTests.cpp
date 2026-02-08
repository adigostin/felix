
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

	TEST_CLASS(MiscTests)
	{
	public:
		TEST_METHOD(LaunchVS)
		{
			com_ptr<IUnknown> solution;
			auto hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::IsTrue(SUCCEEDED(hr));
			auto sln = wil::com_query_failfast<VxDTE::_Solution>(solution);
			sln->Close();
		}

		TEST_METHOD(CloneProject)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"CloneProject");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			com_ptr<IUnknown> solution;
			hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::IsTrue(SUCCEEDED(hr));
			auto sln = wil::com_query_failfast<VxDTE::_Solution>(solution);
			hr = sln->Create(wil::make_bstr_failfast(testPath.get()).get(), wil::make_bstr_failfast(L"test.sln").get());
			Assert::IsTrue(SUCCEEDED(hr));
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			com_ptr<VxDTE::Project> proj;
			hr = sln->AddFromTemplate (
				wil::make_bstr_failfast(TemplatePath_TwoConfigsOneFile.get()).get(),
				wil::make_bstr_failfast(testPath.get()).get(),
				wil::make_bstr_failfast(L"test.flx").get(), VARIANT_TRUE, &proj);
			Assert::IsTrue(SUCCEEDED(hr), wil::str_printf_failfast<wil::unique_process_heap_string>(L"0x%08x", hr).get());
		}

		TEST_METHOD(OpenSpecificEditor)
		{
			// At some point between VS 17.10 and 17.14, the VS implementation of IVsUIShellOpenDocument::OpenStandardEditor
			// started trying to open our .asm files not with "Source Code (Text) Editor", but with
			// "Common Language Editor Supporting TextMate Bundles" (or at least this was the new default shown in Open With...).
			// The implementation started returning E_NOTIMPL in most cases.
			// The solution was to call OpenSpecificEditor with GUID_TextEditorFactory, instead of OpenStandardEditor.
			// This test verifies this solution. The test fails with the old code that calls OpenStandardEditor,
			// and succeeds with the new code that calls OpenSpecificEditor.

			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"OpenSpecificEditor");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject(testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			auto hier = proj.query<IVsUIHierarchy>();
			VSITEMID itemid;
			hr = hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			auto vsp2 = proj.query<IVsProject2>();
			com_ptr<IVsWindowFrame> wf;
			hr = vsp2->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->Show();
			Assert::IsTrue(SUCCEEDED(hr));
			hr = sln->Close();
			Assert::IsTrue(SUCCEEDED(hr));

			auto fullSlnPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\test.sln");
			hr = sln->Open(wil::make_bstr_failfast(fullSlnPath.get()).get()); 
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<VxDTE::Projects> projects;
			hr = sln->get_Projects(&projects);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = projects->Item(wil::make_variant_bstr_failfast(L"test.flx"), &proj);
			Assert::IsTrue(SUCCEEDED(hr));

			hier = proj.query<IVsUIHierarchy>();
			hr = hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			vsp2 = proj.query<IVsProject2>();
			hr = vsp2->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf); // This would return E_NOTIMPL before the fix.
			Assert::IsTrue(SUCCEEDED(hr));
		}

		struct FileChangeEvents : IVsFileChangeEvents
		{
			ULONG _refCount = 0;
			bool _changed = false;

			#pragma region IUnknown
			virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
			{
				if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<IVsFileChangeEvents>(this, riid, ppvObject))
					return S_OK;

				*ppvObject = nullptr;
				return E_NOINTERFACE;
			}
			virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
			virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
			#pragma endregion

			#pragma region IVsFileChangeEvents
			virtual HRESULT STDMETHODCALLTYPE FilesChanged (DWORD cChanges, LPCOLESTR rgpszFile[], VSFILECHANGEFLAGS rggrfChange[]) override
			{
				_changed = true;
				return S_OK;
			}

			virtual HRESULT STDMETHODCALLTYPE DirectoryChanged (LPCOLESTR pszDirectory) override
			{
				RETURN_HR(E_NOTIMPL);
			}
			#pragma endregion
		};

		TEST_METHOD(AdviseFileChangeNotCalledOnProjectSave)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"AdviseFileChangeNotCalledOnProjectSave");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			com_ptr<IDispatch> aodisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"TestHelper").get(), &aodisp);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IFelixTestHelper> ao;
			hr = aodisp->QueryInterface(IID_PPV_ARGS(&ao));
			Assert::IsTrue(SUCCEEDED(hr));

			auto fce = com_ptr(new (std::nothrow) FileChangeEvents()); FAIL_FAST_IF_NULL_ALLOC(fce);
			DWORD cookie;
			hr = ao->AdviseProjectFileChange(proj, fce, &cookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([ao=ao.get(), cookie]() {
				ao->UnadviseProjectFileChange(cookie);
				});

			hr = proj->Save(nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// give it some time to notice the file changed on disk, and to call our callback
			DWORD tickStart = GetTickCount();
			while (!fce->_changed && GetTickCount() - tickStart < 1000)
			{
				MSG msg;
				while(PeekMessage(&msg,0,0,0,PM_NOREMOVE))
				{
					if (::GetMessage(&msg, NULL, 0, 0) > 0)
						::DispatchMessage(&msg);
				}

				Sleep(10);
			}

			bool changed = fce->_changed;
			hr = CoDisconnectObject(fce, 0);
			Assert::IsTrue(SUCCEEDED(hr));
			fce = nullptr;

			Assert::IsFalse(changed); // this would fail before the fix (IgnoreFile called in ProjectNode::Save)
		}

		TEST_METHOD(NavigateToErrorInFileInSubdir)
		{
			// When the user double-clicks an error in the Error List window, VS goes out of its way to locate the file
			// and find the project it contains. To repro the bug, we need to put the project file in a subdirectory
			// (not next to the .sln file), and the offending file in a subdirectory of its own. In this case, when the user
			// double-clicks the error in the Error List window, VS used to be unable to locate the file due to bugs in
			// ProjectNode::ParseCanonicalName and ProjectNode::IsDocumentInProject (and ultimately in ParseCanonicalName).
			// This test verifies the fixes in these functions.

			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"NavigateToErrorInFileInSubdir");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"testproj");
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			long buildFailCount;
			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);

			auto file1Path = CombinePath(testPath.get(), L"testproj\\subdir\\file1.asm");
			WriteFileOnDisk (file1Path.get(), "\t555555");

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(file1Path.addressof()), NULL, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj->Save();
			Assert::IsTrue(SUCCEEDED(hr));

			BuildSolution(sln, &buildFailCount);
			Assert::IsTrue(buildFailCount >= 1);

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
			Assert::IsTrue(errorCount >= 1);

			VARIANT v; v.vt = VT_I4; v.lVal = 1;
			com_ptr<VxDTE::ErrorItem> errorItem;
			hr = errorItems->Item(v, &errorItem);
			Assert::IsTrue(SUCCEEDED(hr));

			errorItem->Navigate(); // This call succeeds, even though VS couldn't open the file.

			com_ptr<VxDTE::Document> doc;
			hr = dte->get_ActiveDocument(&doc);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(doc.get()); // This would fail before the fix.

			BOOL found = FALSE;
			VSDOCUMENTPRIORITY prio = { };
			VSITEMID itemid = { };
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"subdir\\file1.asm", &found, &prio, &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"subdir/..\\subdir/.\\file1.asm", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"SubDir\\File1.ASM", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);

			// For the canonical name, we don't bother with tricky paths such as "subdir/..\\subdir/.\\file1.asm".
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"subdir\\file1.asm", &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"SubDir\\File1.ASM", &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(GenPrePostInclude_OnlyActiveCfg)
		{
			// Verify that the PreInclude.asm and PostInclude.asm are regenerated when editing the active configuration,
			// and that they are _not_ regenerated when editing an inactive configuration.

			HRESULT hr;

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"GenPrePostInclude_OnlyActiveCfg");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln0, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"testproj");
			auto close = wil::scope_exit([sln=sln0.get()] { sln->Close(); });

			// Let's not go through the configuration manager since we haven't implemented Project::get_ConfigurationManager yet.
			//Microsoft_VisualStudio_Interop::_SolutionPtr sln = sln0.get();
			//auto configs = sln->SolutionBuild->SolutionConfigurations;
			//Assert::IsTrue(configs->Count >= 2);
			//Assert::AreEqual<void*>(configs->Item[1].GetInterfacePtr(), sln->SolutionBuild->ActiveConfiguration.GetInterfacePtr());

			com_ptr<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// Make a change in the active configuration and verify that the Pre/PostInclude files are generated.
			auto genFilesPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\testproj\\GeneratedFiles");
			Assert::IsTrue(PathFileExists(genFilesPath.get()));
			RemoveDirectoryTree(genFilesPath.get());

			com_ptr<IProjectConfigProperties> props;
			hr = cfgs[0]->QueryInterface(IID_PPV_ARGS(&props));
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IProjectConfigAssemblerProperties> asmProps;
			hr = props->get_AssemblerProperties(&asmProps);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = asmProps->put_BaseAddress(1234);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(genFilesPath.get()));

			// Now make a change in the inactive configuration.
			RemoveDirectoryTree(genFilesPath.get());

			hr = cfgs[1]->QueryInterface(IID_PPV_ARGS(&props));
			Assert::IsTrue(SUCCEEDED(hr));
			hr = props->get_AssemblerProperties(&asmProps);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = asmProps->put_BaseAddress(1234);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(!PathFileExists(genFilesPath.get()));
		}

		TEST_METHOD(RemoveFileClosesEditor)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"RemoveFileClosesEditor");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln0, proj] = CreateSolutionAndProject (testPath.get(), L"test", NULL);
			auto close = wil::scope_exit([sln=sln0.get()] { sln->Close(); });

			VSITEMID itemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IVsWindowFrame> wf;
			hr = proj.query<IVsProject2>()->OpenItem(itemId, LOGVIEWID_Code, nullptr, &wf);
			wf->Show();
			hr = wf->IsVisible();
			Assert::AreEqual (S_OK, hr);

			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &itemId, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->IsVisible();
			Assert::AreNotEqual (S_OK, hr);
		}

		TEST_METHOD(RemoveFolderAndFile_FolderFirstInList)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"RemoveFolderAndFile_FolderFirstInList");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln0, proj] = CreateSolutionAndProject (testPath.get(), L"test", NULL);
			auto close = wil::scope_exit([sln=sln0.get()] { sln->Close(); });

			auto hier = proj.query<IVsHierarchy>();
			VSITEMID itemIds[2];
			hr = hier->ParseCanonicalName(L"GeneratedFiles", &itemIds[0]);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = hier->ParseCanonicalName(L"GeneratedFiles\\PreInclude.asm", &itemIds[1]);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(2, DELITEMOP_DeleteFromStorage, itemIds, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(RemoveFolderClosesEditors)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"RemoveFolderClosesEditors");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln0, proj] = CreateSolutionAndProject (testPath.get(), L"test", NULL);
			auto close = wil::scope_exit([sln=sln0.get()] { sln->Close(); });

			VSITEMID itemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"GeneratedFiles\\PreInclude.asm", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IVsWindowFrame> wf;
			hr = proj.query<IVsProject2>()->OpenItem(itemId, LOGVIEWID_Code, nullptr, &wf);
			wf->Show();
			hr = wf->IsVisible();
			Assert::AreEqual (S_OK, hr);

			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"GeneratedFiles", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &itemId, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->IsVisible();
			Assert::AreNotEqual (S_OK, hr);
		}

		TEST_METHOD(AddNewFile_SameNameAsFileOutsideProjectDir)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_hglobal_string>(tempPath, L"AddNewFile_SameNameAsFileOutsideProjectDir");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"proj");
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			auto projPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj\\");

			const wchar_t file1RelPath[] = L"test.asm";
			auto file1Path = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, file1RelPath);
			WriteFileOnDisk (file1Path.get(), "; comment");
			const wchar_t file2RelPath[] = L"..\\test.asm";
			auto file2Path = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, file2RelPath);
			WriteFileOnDisk (file2Path.get(), "; comment");
			wchar_t file3Path[] = L"D:\\FelixTest\\test.asm";
			WriteFileOnDisk (file3Path, "; comment");

			LPCOLESTR filesToOpen[] = { file1Path.get(), file2Path.get(), file3Path };
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 3, filesToOpen, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			// Adding them again should fail.
			for (const wchar_t* fileToOpen : filesToOpen)
			{
				hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, &fileToOpen, nullptr, &addResult);
				Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), hr);
			}

			VSITEMID itemids[3];
			for (int i = 0; i < 3; i++)
			{
				BOOL found;
				VSDOCUMENTPRIORITY prio;
				hr = proj.query<IVsProject>()->IsDocumentInProject(filesToOpen[i], &found, &prio, &itemids[i]);
				Assert::IsTrue(SUCCEEDED(hr) && found);
			}

			wil::unique_bstr canonicalNames[3];
			for (int i = 0; i < 3; i++)
			{
				hr = proj.query<IVsHierarchy>()->GetCanonicalName(itemids[i], &canonicalNames[i]);
				Assert::IsTrue(SUCCEEDED(hr));
			}

			for (int i = 0; i < 3; i++)
			{
				VSITEMID itemidx;
				hr = proj.query<IVsHierarchy>()->ParseCanonicalName(canonicalNames[i].get(), &itemidx);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual(itemids[i], itemidx);
			}
		}

		TEST_METHOD(CloneFile_SameNameAsFileOutsideProjectDir)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_hglobal_string>(tempPath, L"CloneFile_SameNameAsFileOutsideProjectDir");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"proj");
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			auto templatePath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\template.asm");
			WriteFileOnDisk (templatePath.get(), "; comment");

			auto projPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj\\");

			static const wchar_t file1RelPath[] = L"..\\test.asm";
			auto file1Path = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, file1RelPath);
			WriteFileOnDisk(file1Path.get(), "; comment");
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(file1Path.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"test.asm", 1, const_cast<LPCOLESTR*>(templatePath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(CloneFileToFolderMissingOnDisk)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_hglobal_string>(tempPath, L"CloneFileToFolderMissingOnDisk");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { RemoveDirectoryTree(tp); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"proj");
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			auto templatePath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\template.asm");
			WriteFileOnDisk (templatePath.get(), "; comment");

			wil::unique_variant folderItemId;
			hr = proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &folderItemId);
			Assert::IsTrue(SUCCEEDED(hr));

			// Delete the folder on disk, then try to clone a file into it. We must at least not crash.

			wil::unique_bstr folderMk;
			hr = proj.query<IVsProject>()->GetMkDocument(V_VSITEMID(&folderItemId), &folderMk);
			Assert::IsTrue(SUCCEEDED(hr));

			if (PathFileExists(folderMk.get()))
			{
				BOOL bres = RemoveDirectory(folderMk.get());
				Assert::IsTrue(bres);
			}

			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			VSADDRESULT addResult;
			hr = proj.query<IVsProject>()->AddItem(V_VSITEMID(&folderItemId), oper, L"file.asm", 1, const_cast<LPCOLESTR*>(templatePath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

		}
	};
}
