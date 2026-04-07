
#include "pch.h"
#include "shared/com.h"
#include "../TestsCommon.h"
#include "../FelixPackage/dispids.h"
#include "UITests.h"

namespace UITests
{
	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (VxDTE::DTE2* dte, PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount);

	TEST_CLASS(NotifyTests)
	{
		wil::com_ptr_failfast<VxDTE::DTE2> dte;
		wil::unique_process_heap_string testPath;
		wil::unique_process_heap_string slnFilePath;
		wil::unique_process_heap_string projPath;
		wil::com_ptr_failfast<VxDTE::_Solution> sln;
		wil::com_ptr_failfast<VxDTE::Project> proj;

		TEST_METHOD_INITIALIZE(NotifyTestInit)
		{
			dte = GetDefaultVSInstance();
			testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"NotifyTest");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			std::tie(sln, proj) = CreateSolutionAndProject(dte, testPath.get(), L"test", L"proj");
			slnFilePath = CombinePath(testPath.get(), L"test.sln");
			projPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj");
		}

		TEST_METHOD_CLEANUP(NotifyTestCleanup)
		{
			if (sln)
			{
				sln->Close();
				sln.reset();
				proj.reset();
			}

			if (testPath)
			{
				RemoveDirectoryTree(testPath.get());
				testPath.reset();
			}
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
			com_ptr<IDispatch> aodisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"TestHelper").get(), &aodisp);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IFelixTestHelper> ao;
			hr = aodisp->QueryInterface(IID_PPV_ARGS(&ao));
			Assert::IsTrue(SUCCEEDED(hr));

			auto fce = com_ptr(new (std::nothrow) FileChangeEvents()); FAIL_FAST_IF_NULL_ALLOC(fce);
			auto disconnect = wil::scope_exit([&fce] { CoDisconnectObject(fce, 0); });

			DWORD cookie;
			hr = ao->AdviseProjectFileChange(proj, fce, &cookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([ao=ao.get(), cookie]() { ao->UnadviseProjectFileChange(cookie); });

			hr = proj->Save(nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// give it some time to notice the file changed on disk, and to call our callback
			bool changed = WaitWithMessageLoop([&fce] { return fce->_changed; }, 1000);

			Assert::IsTrue(SUCCEEDED(hr));
			fce = nullptr;

			Assert::IsFalse(changed); // this would fail before the fix (IgnoreFile called in ProjectNode::Save)
		}

		TEST_METHOD(NotifyPropertyChangedFileFolderNodes)
		{
			HRESULT hr;
			auto hier = proj.query<IVsUIHierarchy>();
			auto sink = MakeTestHierarchyEventSink();
			VSCOOKIE hierEventsCookie;
			hier->AdviseHierarchyEvents(sink, &hierEventsCookie);
			auto unadvise = wil::scope_exit([hier=hier.get(), hierEventsCookie]() { hier->UnadviseHierarchyEvents(hierEventsCookie); });

			// VSHPROPID_Name on file
			VSITEMID fileItemId = VSITEMID_NIL;
			hier->ParseCanonicalName (L"file.asm", &fileItemId);
			hier->SetProperty(fileItemId, VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"file1.asm"));
			Assert::IsTrue(sink->PropertyChanged(fileItemId, VSHPROPID_Name));

			// VSHPROPID_Name on folder
			wil::unique_variant tf1;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &tf1);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf1.vt);
			hr = hier->SetProperty (V_VSITEMID(&tf1), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"testfolder1"));
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(sink->PropertyChanged(V_VSITEMID(&tf1), VSHPROPID_Name));
		}

		TEST_METHOD(NotifyItemInsertedRemoved)
		{
			HRESULT hr;
			auto hier = proj.query<IVsUIHierarchy>();
			auto sink = MakeTestHierarchyEventSink();
			VSCOOKIE hierEventsCookie;
			hier->AdviseHierarchyEvents(sink, &hierEventsCookie);
			auto unadvise = wil::scope_exit([hier=hier.get(), hierEventsCookie]() { hier->UnadviseHierarchyEvents(hierEventsCookie); });

			// Add file
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			VSADDRESULT addResult;
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"1.asm", 1, TemplateEmptyFile, nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
			VSITEMID itemIdFileAdded;
			hr = hier->ParseCanonicalName(L"1.asm", &itemIdFileAdded);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(sink->ItemAdded(VSITEMID_ROOT, itemIdFileAdded));

			// Add folder
			wil::unique_variant tf1;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &tf1);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, tf1.vt);
			Assert::IsTrue(sink->ItemAdded(VSITEMID_ROOT, V_VSITEMID(&tf1)));

			// Remove file
			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, &itemIdFileAdded, DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(sink->ItemRemoved(itemIdFileAdded));

			// Remove folder
			hr = proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_DeleteFromStorage, (VSITEMID*)&V_VSITEMID(&tf1), DHO_SUPPRESS_UI);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(sink->ItemRemoved(V_VSITEMID(&tf1)));	
		}

		TEST_METHOD(NotifyPropertyChangingChanged_File)
		{
			// Tests how well these events are propagated, not necessarily if they are generated for every single property.

			HRESULT hr;
			auto changeSink = MakeTestPropertyChangeSink();
			auto notifySink = MakeTestPropertyNotifySink();
			auto disconnectSinks = wil::scope_exit([&changeSink, &notifySink]
				{
					CoDisconnectObject(changeSink, 0);
					CoDisconnectObject(notifySink, 0);
				});

			VSITEMID fileItemID;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &fileItemID);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant filevar;
			hr = proj.query<IVsHierarchy>()->GetProperty(fileItemID, VSHPROPID_BrowseObject, &filevar);
			Assert::IsTrue(SUCCEEDED(hr));
			auto fileProps = wil::com_query_failfast<IFileNodeProperties>(filevar.pdispVal);
			AdviseSinkToken fileChangeSinkToken;
			hr = AdviseSink<IPropertyChangeSink>(fileProps, changeSink, &fileChangeSinkToken);
			AdviseSinkToken fileNotifySinkToken;
			hr = AdviseSink<IPropertyNotifySink>(fileProps, notifySink, &fileNotifySinkToken);
			Assert::IsTrue(SUCCEEDED(hr));
			fileProps->put_BuildTool(BuildToolKind::CustomBuildTool);
			bool called = WaitWithMessageLoop([&changeSink, &notifySink, fileDisp=filevar.pdispVal]
				{
					return changeSink->Called(fileDisp, { dispidBuildToolKind })
						&& notifySink->Called({ dispidBuildToolKind });
				}, 1000);
			Assert::IsTrue(called);
			BOOL projectDirty = FALSE;
			proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
			Assert::IsTrue(projectDirty);
			proj->Save(nullptr);
			proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
			Assert::IsFalse(projectDirty);

			wil::com_ptr_failfast<ICustomBuildToolProperties> cbtProps;
			hr = fileProps->get_CustomBuildToolProperties(&cbtProps);
			Assert::IsTrue(SUCCEEDED(hr));

			AdviseSinkToken cbtChangeSinkToken;
			hr = AdviseSink<IPropertyChangeSink>(cbtProps, changeSink, &cbtChangeSinkToken);
			AdviseSinkToken cbtNotifySinkToken;
			hr = AdviseSink<IPropertyNotifySink>(cbtProps, notifySink, &cbtNotifySinkToken);
			Assert::IsTrue(SUCCEEDED(hr));
			auto cbtDisp = cbtProps.query<IDispatch>();

			cbtProps->put_CommandLine(wil::make_bstr_failfast(L"CmdLine").get());
			cbtProps->put_Description(wil::make_bstr_failfast(L"Desc").get());
			cbtProps->put_Outputs(wil::make_bstr_failfast(L"Outputs").get());
			called = WaitWithMessageLoop([&changeSink, &notifySink, &cbtDisp]
				{ 
					return changeSink->Called(cbtDisp, { dispidCommandLine, dispidDescription, dispidOutputs })
						&& notifySink->Called({ dispidCommandLine, dispidDescription, dispidOutputs });
				}, 1000);
			Assert::IsTrue(called);
			wil::unique_bstr value;
			cbtProps->get_CommandLine(&value);
			Assert::AreEqual(L"CmdLine", (wchar_t*)value.get());
			cbtProps->get_Description(&value);
			Assert::AreEqual(L"Desc", (wchar_t*)value.get());
			cbtProps->get_Outputs(&value);
			Assert::AreEqual(L"Outputs", (wchar_t*)value.get());

			proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
			Assert::IsTrue(projectDirty);
		}

		TEST_METHOD(NotifyPropertyChangingChanged_Folder)
		{
			// Tests how well these events are propagated, not necessarily if they are generated for every single property.

			HRESULT hr;
			auto changeSink = MakeTestPropertyChangeSink();
			auto notifySink = MakeTestPropertyNotifySink();
			auto disconnectSinks = wil::scope_exit([&changeSink, &notifySink]
				{
					CoDisconnectObject(changeSink, 0);
					CoDisconnectObject(notifySink, 0);
				});

			wil::unique_variant itemId;
			hr = proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, &itemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant var;
			hr = proj.query<IVsHierarchy>()->GetProperty(V_VSITEMID(&itemId), VSHPROPID_BrowseObject, &var);
			Assert::IsTrue(SUCCEEDED(hr));
			AdviseSinkToken changeSinkToken;
			hr = AdviseSink<IPropertyChangeSink>(var.pdispVal, changeSink, &changeSinkToken);
			AdviseSinkToken notifySinkToken;
			hr = AdviseSink<IPropertyNotifySink>(var.pdispVal, notifySink, &notifySinkToken);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj.query<IVsHierarchy>()->SetProperty(V_VSITEMID(&itemId), VSHPROPID_EditLabel, wil::make_variant_bstr_failfast(L"newname"));
			Assert::IsTrue(SUCCEEDED(hr));
			bool called = WaitWithMessageLoop([&changeSink, &notifySink, &var]
				{
					return changeSink->Called(var.pdispVal, { dispidFolderName })
						&& notifySink->Called({ dispidFolderName });
				}, 1000);
			Assert::IsTrue(called);

			BOOL projectDirty = FALSE;
			proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
			Assert::IsTrue(projectDirty);
		}

		TEST_METHOD(NotifyPropertyChangingChanged_Config)
		{
			HRESULT hr;
			auto changeSink = MakeTestPropertyChangeSink();
			auto notifySink = MakeTestPropertyNotifySink();
			auto disconnectSinks = wil::scope_exit([&changeSink, &notifySink]
				{
					CoDisconnectObject(changeSink, 0);
					CoDisconnectObject(notifySink, 0);
				});

			BOOL projectDirty;
			proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
			Assert::IsFalse(projectDirty);

			auto cfgProvider = proj.query<IVsCfgProvider2>();
			wil::com_ptr_failfast<IVsCfg> cfgs[2];
			VSCFGFLAGS cfgFlags[2];
			ULONG cfgActual;
			hr = cfgProvider->GetCfgs(2, cfgs[0].addressof(), &cfgActual, cfgFlags);
			Assert::IsTrue(SUCCEEDED(hr));

			{
				hr = cfgs[0].query<IProjectConfigProperties>()->put_ConfigName(wil::make_bstr_failfast(L"Debug1").get());
				Assert::IsTrue(SUCCEEDED(hr));
				proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
				Assert::IsTrue(projectDirty);
			}

			{
				proj->Save(nullptr);
				wil::com_ptr_failfast<IProjectConfigGeneralProperties> generalProps;
				hr = cfgs[0].query<IProjectConfigProperties>()->get_GeneralProperties(&generalProps);
				Assert::IsTrue(SUCCEEDED(hr));
				hr = generalProps->put_OutputName(wil::make_bstr_failfast(L"output1.bin").get());
				Assert::IsTrue(SUCCEEDED(hr));
				proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
				Assert::IsTrue(projectDirty);
			}

			{
				proj->Save(nullptr);
				wil::com_ptr_failfast<IProjectConfigAssemblerProperties> asmProps;
				hr = cfgs[0].query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
				Assert::IsTrue(SUCCEEDED(hr));
				hr = asmProps->put_BaseAddress(123);
				Assert::IsTrue(SUCCEEDED(hr));
				proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
				Assert::IsTrue(projectDirty);
			}

			{
				proj->Save(nullptr);
				wil::com_ptr_failfast<IProjectConfigDebugProperties> dbgProps;
				hr = cfgs[0].query<IProjectConfigProperties>()->get_DebuggingProperties(&dbgProps);
				Assert::IsTrue(SUCCEEDED(hr));
				hr = dbgProps->put_LaunchTarget(wil::make_bstr_failfast(L"output1.bin").get());
				Assert::IsTrue(SUCCEEDED(hr));
				proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
				Assert::IsTrue(projectDirty);
			}

			{
				proj->Save(nullptr);
				wil::com_ptr_failfast<IProjectConfigPrePostBuildProperties> preBuildProps;
				hr = cfgs[0].query<IProjectConfigProperties>()->get_PreBuildProperties(&preBuildProps);
				Assert::IsTrue(SUCCEEDED(hr));
				hr = preBuildProps->put_CommandLine(wil::make_bstr_failfast(L"cmd.exe").get());
				Assert::IsTrue(SUCCEEDED(hr));
				proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
				Assert::IsTrue(projectDirty);
			}

			{
				proj->Save(nullptr);
				wil::com_ptr_failfast<IProjectConfigPrePostBuildProperties> postBuildProps;
				hr = cfgs[0].query<IProjectConfigProperties>()->get_PostBuildProperties(&postBuildProps);
				Assert::IsTrue(SUCCEEDED(hr));
				hr = postBuildProps->put_CommandLine(wil::make_bstr_failfast(L"cmd.exe").get());
				Assert::IsTrue(SUCCEEDED(hr));
				proj.query<IPersistFileFormat>()->IsDirty(&projectDirty);
				Assert::IsTrue(projectDirty);
			}
		}
	};
}
