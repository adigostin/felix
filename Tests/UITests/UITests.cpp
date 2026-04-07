
#include "pch.h"
#include "UITests.h"
#include "shared/com.h"
#include <unordered_map>
#include <set>

namespace UITests
{
	static wil::critical_section cs;
	static wil::com_ptr_failfast<IUIAutomation> automation;
	static wil::com_ptr_failfast<VxDTE::DTE2> defaultInstance;

	#pragma region DTE Initialization
	static HRESULT GetDTE (DWORD processId, VxDTE::DTE2** ppDTE)
	{
		HRESULT hr;

		//MessageFilter.Register();

		wil::unique_process_heap_string fn;
		hr = wil::GetModuleFileNameW(nullptr, fn); RETURN_IF_FAILED_EXPECTED(hr);
		DWORD dummydw;
		DWORD verlen = GetFileVersionInfoSize (fn.get(), &dummydw); RETURN_LAST_ERROR_IF_EXPECTED(verlen == 0);
		auto verbuffer = wil::make_unique_nothrow<char[]>(verlen);
		BOOL bres = GetFileVersionInfo(fn.get(), 0, verlen, verbuffer.get()); RETURN_IF_WIN32_BOOL_FALSE_EXPECTED(bres);
		void* valbuffer;
		UINT vallen;
		bres = VerQueryValueW (verbuffer.get(), L"\\", &valbuffer, &vallen); RETURN_HR_IF_EXPECTED(E_FAIL, !bres);
		VS_FIXEDFILEINFO* fi = (VS_FIXEDFILEINFO*)valbuffer;
		DWORD vsMajorVersion = fi->dwProductVersionMS >> 16;
		auto progId = wil::str_printf_failfast<wil::unique_process_heap_string> (L"!VisualStudio.DTE.%u.0:%u", vsMajorVersion, processId);

		wil::com_ptr_failfast<IUnknown> runningObject;
		wil::com_ptr_failfast<IBindCtx> bindCtx;
		wil::com_ptr_failfast<IRunningObjectTable> rot;
		wil::com_ptr_failfast<IEnumMoniker> enumMonikers;

		hr = CreateBindCtx(0, &bindCtx); RETURN_IF_FAILED_EXPECTED(hr);
		hr = bindCtx->GetRunningObjectTable(&rot); RETURN_IF_FAILED_EXPECTED(hr);
		hr = rot->EnumRunning(&enumMonikers); RETURN_IF_FAILED_EXPECTED(hr);

		wil::com_ptr_failfast<IMoniker> moniker;
		ULONG numberFetched = 0;
		while (enumMonikers->Next (1, moniker.addressof(), &numberFetched) == S_OK)
		{
			wil::unique_cotaskmem_string name;
			hr = moniker->GetDisplayName(bindCtx, nullptr, &name);
			if (FAILED(hr))
			{
				if (hr == E_ACCESSDENIED)
				{
					// Do nothing, there is something in the ROT that we do not have access to.
				}
				else
					return hr;
			}

			if (name && !wcscmp(name.get(), progId.get()))
			{
				hr = rot->GetObject(moniker, &runningObject);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::IsNotNull(runningObject.get());
				hr = runningObject->QueryInterface(IID_PPV_ARGS(ppDTE)); RETURN_IF_FAILED_EXPECTED(hr);
				return S_OK;
			}
		}

		return E_FAIL;
	}

	static HWND FindTopLevelWindow (DWORD process_id)
	{
		struct handle_data {
			DWORD process_id;
			HWND window_handle;
		};

		handle_data data;
		data.process_id = process_id;
		data.window_handle = 0;

		auto callback = [](HWND handle, LPARAM lParam) -> BOOL
			{
				handle_data* data = (handle_data*)lParam;
				DWORD process_id = 0;
				GetWindowThreadProcessId(handle, &process_id);
				if (data->process_id != process_id || GetWindow(handle, GW_OWNER) || !IsWindowVisible(handle))
					return TRUE;
				data->window_handle = handle;
				return FALSE;   
			};
		EnumWindows(callback, (LPARAM)&data);
		return data.window_handle;
	}

	static void Attach (DWORD targetVSProcessID, VxDTE::_DTE* targetVSDTE, DWORD debuggerVsProcessId, VxDTE::Debugger* debugger)
	{
		HRESULT hr;
		auto dbg3 = wil::com_query_failfast<VxDTE::Debugger3>(debugger);
		wil::com_ptr_failfast<VxDTE::Processes> processes;
		hr = dbg3->get_LocalProcesses(&processes);
		Assert::IsTrue(SUCCEEDED(hr));
		wil::com_ptr_failfast<IUnknown> newEnum;
		hr = processes->_NewEnum(&newEnum);
		Assert::IsTrue(SUCCEEDED(hr));
		auto enumProcesses = wil::com_query_failfast<IEnumVARIANT>(newEnum);
		wil::unique_variant var;
		ULONG fetched;
		while (enumProcesses->Next(1, &var, &fetched) == S_OK)
		{
			if (auto targetProcess = wil::try_com_query_failfast<VxDTE::Process3>(var.pdispVal))
			{
				long pid;
				if (SUCCEEDED(targetProcess->get_ProcessID(&pid))
					&& (DWORD)pid == targetVSProcessID)
				{
					targetVSDTE->put_SuppressUI(VARIANT_TRUE);
					//hr = targetProcess->Attach2(wil::make_variant_bstr_failfast(L"Managed/Native"));
					hr = targetProcess->Attach2(wil::make_variant_bstr_failfast(L"Native"));
					targetVSDTE->put_SuppressUI(VARIANT_FALSE);
					Assert::IsTrue(SUCCEEDED(hr));
					return;
				}
			}
		}

		Assert::Fail();
	}

	static void FindAndAttach (DWORD targetVsProcessId, VxDTE::_DTE* targetVSDTE)
	{
		HRESULT hr;

		// We (testhost.exe) are being debugged. Find the VS instance that's debugging us
		// and tell it to debug the target VS too.
		DWORD selfId = GetCurrentProcessId();

		// Get all running processes.
		DWORD cap = 300;
		wil::unique_hlocal_ptr<DWORD[]> pids;
		DWORD processCount;
		while(true)
		{
			pids = wil::make_unique_hlocal_failfast<DWORD[]>(cap);
			DWORD neededBytes;
			BOOL bres = EnumProcesses (pids.get(), cap * sizeof(DWORD), &neededBytes);
			Assert::IsTrue(bres);
			processCount = neededBytes / sizeof(DWORD);
			if (processCount < cap)
				break;
			cap += cap;
		}

		// Retain only the "devenv.exe" processes.
		DWORD vsCount = 0;
		for (DWORD i = 0; i < cap; i++)
		{
			wil::unique_process_handle h (OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, pids.get()[i]));
			if (h.is_valid())
			{
				wil::unique_process_heap_string path;
				hr = wil::QueryFullProcessImageNameW (h.get(), 0, path);
				if (SUCCEEDED(hr))
				{
					auto p = wcsrchr(path.get(), L'\\');
					if (p && !wcscmp(p + 1, L"devenv.exe"))
						pids.get()[vsCount++] = pids.get()[i];
				}
			}
		}

		// Now find the VS process that's debugging us.
		for (DWORD i = 0; i < vsCount; i++)
		{
			DWORD vspid = pids.get()[i];
			if (vspid == targetVsProcessId)
				continue;
			wil::com_ptr_failfast<VxDTE::DTE2> dte;
			if (SUCCEEDED(GetDTE(vspid, &dte)))
			{
				wil::com_ptr_failfast<VxDTE::Debugger> debugger;
				VxDTE::dbgDebugMode mode;
				wil::com_ptr_failfast<VxDTE::Processes> processes;
				wil::com_ptr_failfast<IUnknown> newEnum;
				wil::com_ptr_failfast<IEnumVARIANT> enumVar;
				if (SUCCEEDED(dte->get_Debugger(&debugger))
					&& SUCCEEDED(debugger->get_CurrentMode(&mode))
					&& mode != VxDTE::dbgDesignMode
					&& SUCCEEDED(debugger->get_DebuggedProcesses(&processes))
					&& SUCCEEDED(processes->_NewEnum(&newEnum))
					&& SUCCEEDED(newEnum->QueryInterface(IID_PPV_ARGS(&enumVar))))
				{
					wil::unique_variant var;
					ULONG fetched;
					while (enumVar->Next(1, &var, &fetched) == S_OK)
					{
						wil::com_ptr_failfast<VxDTE::Process> process;
						if (var.vt == VT_DISPATCH && SUCCEEDED(var.pdispVal->QueryInterface(IID_PPV_ARGS(&process))))
						{
							long pid;
							if (SUCCEEDED(process->get_ProcessID(&pid))
								&& (DWORD)pid == selfId)
							{
								// This is the correct VS, so attach and return.
								Attach (targetVsProcessId, targetVSDTE, vspid, debugger);
								return;
							}
						}
					}
				}
			}
		}

		Assert::Fail();
	}

	wil::com_ptr_failfast<VxDTE::DTE2> LaunchVS (const wchar_t* envVar)
	{
		HRESULT hr;

		auto lock = cs.lock();
		auto resetev = wil::scope_exit([] { SetEnvironmentVariable(L"FelixTestUI", nullptr); });
		if (envVar)
			SetEnvironmentVariable(L"FelixTestUI", envVar);
		else
			resetev.release();

		wil::unique_process_heap_string devenvExe;
		hr = wil::GetEnvironmentVariableW (L"VSAPPIDDIR", devenvExe);
		Assert::IsTrue(SUCCEEDED(hr));
		devenvExe = wil::str_concat_failfast<wil::unique_process_heap_string>(devenvExe, L"devenv.exe");
		const wchar_t* devenvArguments = L"/rootSuffix Exp";

		auto cmdLine = wil::str_concat_failfast<wil::unique_process_heap_string>(L"\"", devenvExe, L"\" ", devenvArguments);
		STARTUPINFO si = { .cb = sizeof(si) };
		wil::unique_process_information pi;
		BOOL bres = CreateProcessW (devenvExe.get(), cmdLine.get(), nullptr, nullptr, TRUE, 0, nullptr, nullptr, &si, &pi);
		Assert::IsTrue(bres);
		DWORD exitCode;
		bres = GetExitCodeProcess (pi.hProcess, &exitCode);
		Assert::IsTrue(bres);
		Assert::AreEqual(STILL_ACTIVE, exitCode);

		wil::com_ptr_failfast<VxDTE::DTE2> dte;
		DWORD startTime = GetTickCount();
		while (GetTickCount() - startTime <= 60'000 && !dte)
		{
			if (SUCCEEDED(GetDTE(pi.dwProcessId, &dte)))
				break;
			Sleep(250);
		}

		Assert::IsNotNull(dte.get(), L"Failed to start VS");

		// The window we get from DTE is the main window (even if it's invisible), not the splash window.
		wil::com_ptr_failfast<VxDTE::Window> dteMainWindow;
		hr = dte->get_MainWindow(&dteMainWindow);
		Assert::IsTrue(SUCCEEDED(hr));
		long dteMainWindowHWnd;
		hr = dteMainWindow->get_HWnd(&dteMainWindowHWnd);
		Assert::IsTrue(SUCCEEDED(hr));

		auto handle = FindTopLevelWindow(pi.dwProcessId);
		Assert::IsNotNull(handle);
		if (handle != (HWND)(size_t)(DWORD)dteMainWindowHWnd)
		{
			// We have the splash window up.
			wil::com_ptr_failfast<IUIAutomationElement> topLevelWindow;
			hr = automation->ElementFromHandle(handle, &topLevelWindow);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::com_ptr_failfast<IUIAutomationCondition> condition;
			hr = automation->CreatePropertyCondition(UIA_NamePropertyId, wil::make_variant_bstr_failfast(L"Continue without code"), &condition);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::com_ptr_failfast<IUIAutomationElement> cwc;
			hr = topLevelWindow->FindFirst(TreeScope_Descendants, condition, &cwc);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::com_ptr_failfast<IUnknown> patternUnk;
			hr = cwc->GetCurrentPattern(UIA_InvokePatternId, &patternUnk);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::com_ptr_failfast<IUIAutomationInvokePattern> invpat;
			hr = patternUnk->QueryInterface(IID_PPV_ARGS(&invpat));
			Assert::IsTrue(SUCCEEDED(hr));
			hr = invpat->Invoke();
			Assert::IsTrue(SUCCEEDED(hr));

			// Let's wait for the main window to become the top level window
			DWORD tickStart = GetTickCount();
			while (true)
			{
				handle = FindTopLevelWindow(pi.dwProcessId);
				if (handle == (HWND)(size_t)(DWORD)dteMainWindowHWnd)
					break;
				if (GetTickCount() - tickStart >= 5000)
					Assert::Fail();
				Sleep(100);
			}
		}

		if (IsDebuggerPresent())
			FindAndAttach(pi.dwProcessId, dte);

		return dte;
	}

	void CloseVS (VxDTE::DTE2* dte, bool hard)
	{
		HRESULT hr;

		hr = dte->put_SuppressUI(VARIANT_TRUE);
		Assert::IsTrue(SUCCEEDED(hr));

		wil::com_ptr_failfast<IUnknown> solution;
		hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
		Assert::IsTrue(SUCCEEDED(hr));
		hr = solution.query<VxDTE::_Solution>()->Close();
		Assert::IsTrue(SUCCEEDED(hr));

		wil::com_ptr_failfast<VxDTE::Window> dteMainWindow;
		hr = dte->get_MainWindow(&dteMainWindow);
		Assert::IsTrue(SUCCEEDED(hr));
		long dteMainWindowHWnd;
		hr = dteMainWindow->get_HWnd(&dteMainWindowHWnd);
		Assert::IsTrue(SUCCEEDED(hr));
		DWORD processID;
		GetWindowThreadProcessId((HWND)(size_t)(DWORD)dteMainWindowHWnd, &processID);
		wil::unique_handle hProcess (OpenProcess (SYNCHRONIZE, FALSE, processID));

		bool exited = false;
		//if (!hard && SUCCEEDED(dte->Quit())) dte->Quit() seems to return immediately and run asynchronously
		if (!hard && SUCCEEDED(dte->ExecuteCommand(wil::make_bstr_failfast(L"File.Exit").get(), wil::make_bstr_failfast(L"").get())))
		{
			auto processExited = [&hProcess] { return WaitForSingleObject(hProcess.get(), 0) == WAIT_OBJECT_0; };
			exited = WaitWithMessageLoop (processExited, IsDebuggerPresent() ? INFINITE : 10000);
		}

		if (!exited)
			TerminateProcess(hProcess.get(), 1234);
	}
	#pragma endregion

	static void __stdcall WilLoggingCallback (const wil::FailureInfo& fi) noexcept
	{
		Assert::Fail(fi.pszMessage);
	}

	TEST_MODULE_INITIALIZE(UITestsInitialize)
	{
		HRESULT hr;

		wil::SetResultLoggingCallback (WilLoggingCallback);

		MakeTemplates (L"FelixTestUI");

		hr = CoCreateInstance (__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation));
		Assert::IsTrue(SUCCEEDED(hr));
	}

	TEST_MODULE_CLEANUP(UITestsCleanup)
	{
		if (defaultInstance)
		{
			CloseVS(defaultInstance);
			defaultInstance = nullptr;
		}

		automation.reset();

		// To easy debugging, delete only the contents of the test directory, not the test directory itself.
		auto buffer = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s*.*%c", tempPath, L'\0');
		SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = buffer.get(), .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
		int ires = SHFileOperation(&file_op);
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::AreEqual(0, ires);

		wil::SetResultLoggingCallback(nullptr);
	}

	wil::com_ptr_failfast<VxDTE::DTE2> GetDefaultVSInstance()
	{
		if (!defaultInstance)
			defaultInstance = LaunchVS();

		return defaultInstance;
	}

	// <param name="testDir">The directory in which to create the solution.</param>
	// <param name="solutionName">The name of the solution to create, without extension.</param>
	// <param name="projectName">NULL to create a project with the same name as the solution in the same dir, filename without extension to create project in subdir with different name.</param>
	std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (VxDTE::DTE2* dte, PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName)
	{
		HRESULT hr;
		wil::com_ptr_failfast<IUnknown> solution;
		hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
		Assert::IsTrue(SUCCEEDED(hr));
		auto sln = solution.query<VxDTE::_Solution>();
		hr = sln->Create(wil::make_bstr_failfast(testDir).get(), wil::make_bstr_failfast(solutionName).get());
		Assert::IsTrue(SUCCEEDED(hr));

		wil::unique_process_heap_string projDirBuffer, projName;
		PCWSTR projDir;
		if (projectName)
		{
			projDirBuffer = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\", projectName);
			projDir = projDirBuffer.get();
			projName = wil::str_concat_failfast<wil::unique_process_heap_string>(projectName, L".flx");
			Assert::IsTrue(CreateDirectory(projDir, nullptr));
		}
		else
		{
			projDir = testDir;
			projName = wil::str_concat_failfast<wil::unique_process_heap_string>(solutionName, L".flx");
		}

		wil::com_ptr_failfast<VxDTE::Project> proj;
		hr = sln->AddFromTemplate (
			wil::make_bstr_failfast(TemplatePath_TwoConfigsOneFile.get()).get(),
			wil::make_bstr_failfast(projDir).get(),
			wil::make_bstr_failfast(projName.get()).get(), VARIANT_FALSE, &proj);
		Assert::IsTrue(SUCCEEDED(hr));
		hr = sln->SaveAs(wil::make_bstr_failfast(solutionName).get());
		Assert::IsTrue(SUCCEEDED(hr));

		return { std::move(sln), std::move(proj) };
	}

	void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount)
	{
		HRESULT hr;

		wil::com_ptr_failfast<VxDTE::SolutionBuild> solutionBuild;
		hr = sln->get_SolutionBuild(&solutionBuild);
		Assert::IsTrue(SUCCEEDED(hr));

		hr = solutionBuild->Build(VARIANT_TRUE);
		Assert::IsTrue(SUCCEEDED(hr));

		// LastBuildInfo returns the number of failed projects, despite the parameter name.
		hr = solutionBuild->get_LastBuildInfo(buildFailCount);
		Assert::IsTrue(SUCCEEDED(hr));
	}

	wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path)
	{
		wchar_t volPathName[50];
		BOOL bres = GetVolumePathNameW(path, volPathName, _countof(volPathName));
		Assert::IsTrue(bres);
		wchar_t volumeName[50];
		bres = GetVolumeNameForVolumeMountPointW (volPathName, volumeName, _countof(volumeName));
		Assert::IsTrue(bres);
		auto pathWithoutDrive = PathSkipRootW(path);
		auto testPathOtherDrive = wil::make_process_heap_string_failfast(nullptr, MAX_PATH);
		auto pres = PathCombine(testPathOtherDrive.get(), volumeName, pathWithoutDrive);
		Assert::IsNotNull(pres);
		return testPathOtherDrive;
	}

	wil::unique_bstr GetBuildOutputWindowPaneContent(VxDTE::DTE2* dte)
	{
		HRESULT hr;
		auto dte2 = wil::com_query_failfast<VxDTE::DTE2>(dte);
		wil::com_ptr_failfast<VxDTE::ToolWindows> toolWindows;
		dte2->get_ToolWindows(&toolWindows);
		wil::com_ptr_failfast<VxDTE::OutputWindow> outputWindow;
		toolWindows->get_OutputWindow(&outputWindow);
		wil::com_ptr_failfast<VxDTE::OutputWindowPanes> panes;
		outputWindow->get_OutputWindowPanes(&panes);
		wil::com_ptr_failfast<VxDTE::OutputWindowPane> buildPane;
		hr = panes->Item (wil::make_variant_bstr_failfast(L"Build"), &buildPane);
		Assert::AreEqual(S_OK, hr);
		wil::com_ptr_failfast<VxDTE::TextDocument> text;
		hr = buildPane->get_TextDocument(&text);
		Assert::AreEqual(S_OK, hr);
		wil::com_ptr_failfast<VxDTE::TextSelection> selection;
		text->get_Selection(&selection);
		hr = selection->SelectAll();
		Assert::AreEqual(S_OK, hr);
		wil::unique_bstr output;
		hr = selection->get_Text(&output);
		Assert::AreEqual(S_OK, hr);

		// Wait until no more text is added.
		while(true)
		{
			Sleep(200);
			hr = selection->SelectAll();
			Assert::AreEqual(S_OK, hr);
			wil::unique_bstr outputNew;
			hr = selection->get_Text(&outputNew);
			Assert::AreEqual(S_OK, hr);
			if (!wcscmp(output.get(), outputNew.get()))
				break;
			output = std::move(outputNew);
		}

		return output;
	}

	void DisableGeneratedFiles (VxDTE::Project* proj)
	{
		wil::com_ptr_failfast<IVsCfg> cfg;
		ULONG actual;
		VSCFGFLAGS flags;
		wil::com_query_failfast<IVsCfgProvider>(proj)->GetCfgs(1, &cfg, &actual, &flags);
		wil::com_ptr_failfast<IProjectConfigAssemblerProperties> asmProps;
		cfg.query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
		auto hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
		Assert::AreEqual(S_OK, hr);
	}


	struct TestHierarchyEventSink : ITestHierarchyEventSink
	{
		ULONG _refCount = 0;
		std::unordered_map<VSITEMID, std::set<VSHPROPID>> _changedProps;

		struct Added { VSITEMID itemidParent; VSITEMID itemidSiblingPrev; VSITEMID itemidAdded; };
		std::vector<Added> _added;

		std::set<VSITEMID> _removed;

		std::set<VSITEMID> _childItemsInvalidated;

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(this, riid, ppvObject)
				|| TryQI<IVsHierarchyEvents>(this, riid, ppvObject)
				|| TryQI<ITestHierarchyEventSink>(this, riid, ppvObject)
				)
				return S_OK;

			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}
		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IVsHierarchyEvents
		virtual HRESULT STDMETHODCALLTYPE OnItemAdded (VSITEMID itemidParent, VSITEMID itemidSiblingPrev, VSITEMID itemidAdded) override
		{
			_added.push_back({ itemidParent, itemidSiblingPrev, itemidAdded });
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE OnItemsAppended (VSITEMID itemidParent) override
		{
			Assert::Fail();
		}

		virtual HRESULT STDMETHODCALLTYPE OnItemDeleted (VSITEMID itemid) override
		{
			_removed.insert(itemid);
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged (VSITEMID itemid, VSHPROPID propid, DWORD flags) override
		{
			_changedProps[itemid].insert(propid);
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE OnInvalidateItems (VSITEMID itemidParent) override
		{
			_childItemsInvalidated.insert(itemidParent);
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE OnInvalidateIcon (HICON hicon) override
		{
			return S_OK;
		}
		#pragma endregion

		#pragma region ITestHierarchyEventSink
		virtual bool PropertyChanged (VSITEMID itemid, VSHPROPID propid) const override
		{
			auto it = _changedProps.find(itemid);
			if (it == _changedProps.end())
				return false;
			return it->second.contains(propid);
		}

		virtual bool ItemAdded (VSITEMID itemidParent, VSITEMID itemidAdded) const override
		{
			for (auto& a : _added)
			{
				if (a.itemidParent == itemidParent && a.itemidAdded == itemidAdded)
					return true;
			}

			return false;
		}

		virtual bool ItemRemoved (VSITEMID itemid) const override
		{
			return _removed.contains(itemid);
		}

		virtual bool ChildItemsInvalidated(VSITEMID itemidParent) const override
		{
			return _childItemsInvalidated.contains(itemidParent);
		}
		#pragma endregion
	};

	wil::com_ptr_failfast<ITestHierarchyEventSink> MakeTestHierarchyEventSink()
	{
		return wil::com_ptr_failfast(new (std::nothrow) TestHierarchyEventSink());
	}

	struct TestPropertyChangeSink : ITestPropertyChangeSink
	{
		ULONG _refCount = 0;
		std::unordered_map<wil::com_ptr_failfast<IDispatch>, std::set<DISPID>, std::hash<IDispatch*>> _changing;
		std::unordered_map<wil::com_ptr_failfast<IDispatch>, std::set<DISPID>, std::hash<IDispatch*>> _changed;

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(this, riid, ppvObject)
				|| TryQI<IPropertyChangeSink>(this, riid, ppvObject)
				|| TryQI<ITestPropertyChangeSink>(this, riid, ppvObject)
				)
				return S_OK;

			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}
		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IPropertyChangeSink
		virtual HRESULT STDMETHODCALLTYPE OnPropertyChanging (IDispatch* pObject, DISPID dispID, PropertyChangeArgs args) override
		{
			_changing[pObject].insert(dispID);
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged (IDispatch* pObject, DISPID dispID, PropertyChangeArgs args) override
		{
			_changed[pObject].insert(dispID);
			return S_OK;
		}
		#pragma endregion

		#pragma region ITestPropertyChangeSink
		virtual bool Called (IDispatch* obj, std::initializer_list<DISPID> dispIDs) const override
		{
			for (DISPID dispID : dispIDs)
			{
				auto it = _changing.find(obj);
				if (it == _changing.end() || !it->second.contains(dispID))
					return false;
				it = _changed.find(obj);
				if (it == _changed.end() || !it->second.contains(dispID))
					return false;
			}

			return true;
		}

		#pragma endregion
	};

	wil::com_ptr_failfast<ITestPropertyChangeSink> MakeTestPropertyChangeSink()
	{
		return new TestPropertyChangeSink();
	}

	struct TestPropertyNotifySink : ITestPropertyNotifySink
	{
		ULONG _refCount = 0;
		std::set<DISPID> _changed;

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(static_cast<IPropertyNotifySink*>(this), riid, ppvObject)
				|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
				|| TryQI<ITestPropertyNotifySink>(this, riid, ppvObject)
				)
				return S_OK;

			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IPropertyNotifySink
		virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
		{
			_changed.insert(dispID);
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
		{
			return E_NOTIMPL;
		}
		#pragma endregion

		#pragma region ITestPropertyNotifySink
		virtual bool Called (std::initializer_list<DISPID> dispIDs) const override
		{
			for (auto dispID : dispIDs)
			{
				if (_changed.find(dispID) == _changed.end())
					return false;
			}

			return true;
		}
		#pragma endregion
	};

	wil::com_ptr_failfast<ITestPropertyNotifySink> MakeTestPropertyNotifySink()
	{
		return new (std::nothrow) TestPropertyNotifySink();
	}
}
