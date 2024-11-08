
#include "pch.h"
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/com.h"
#include "shared/string_builder.h"
#include "dispids.h"
#include "guids.h"
#include "Z80Xml.h"

// Useful doc: https://learn.microsoft.com/en-us/visualstudio/extensibility/internals/managing-configuration-options?view=vs-2022

#pragma comment (lib, "Synchronization.lib")

static const HINSTANCE hinstance = reinterpret_cast<HINSTANCE>(&__ImageBase);

static constexpr DWORD LoadAddressDefaultValue = 0x8000;
static constexpr WORD EntryPointAddressDefaultValue = 0x8000;
static constexpr LaunchType LaunchTypeDefaultValue = LaunchType::PrintUsr;

struct ProjectConfig
	: IProjectConfig
	, IVsDebuggableProjectCfg
	, IVsBuildableProjectCfg
	, IVsBuildableProjectCfg2
	, ISpecifyPropertyPages
	, IXmlParent
	//, public IVsProjectCfgDebugTargetSelection
	//, public IVsProjectCfgDebugTypeSelection
{
	ULONG _refCount = 0;
	com_ptr<IVsUIHierarchy> _hier;
	DWORD _threadId;
	wil::unique_bstr _configName;
	wil::unique_bstr _platformName;
	unordered_map_nothrow<VSCOOKIE, wil::com_ptr_nothrow<IVsBuildStatusCallback>> _buildStatusCallbacks;
	VSCOOKIE _buildStatusNextCookie = VSCOOKIE_NIL + 1;

	DWORD _loadAddress = LoadAddressDefaultValue;
	WORD _entryPointAddress = EntryPointAddressDefaultValue;
	LaunchType _launchType = LaunchTypeDefaultValue;

	bool _saveListing = false;
	wil::unique_bstr _listingFilename;

	static constexpr wchar_t wnd_class_name[] = L"ProjectConfig-{9838F078-469A-4F89-B08E-881AF33AE76D}";
	using unique_atom = wil::unique_any<ATOM, void(ATOM), [](ATOM a) { BOOL bres = UnregisterClassW((LPCWSTR)a, hinstance); LOG_IF_WIN32_BOOL_FALSE(bres); }>;
	static inline unique_atom atom;
	static constexpr UINT WM_OUTPUT_LINE = WM_APP + 0;
	static constexpr UINT WM_BUILD_COMPLETE = WM_APP + 1;

	struct pending_build_info_t
	{
		wil::slim_event_auto_reset thread_started_event;
		wil::unique_event_nothrow exit_request_event;
		wil::unique_handle stdoutReadHandle;
		wil::com_ptr_nothrow<IVsOutputWindowPane> output_window;
		wil::com_ptr_nothrow<IVsOutputWindowPaneNoPump> output_window_no_pump;
		wil::unique_process_information process_info;
		wil::unique_hwnd hwnd;
		wil::unique_handle thread_handle;
		vector_nothrow<vector_nothrow<wchar_t>> string_queue; // TODO: change this from vector to queue, or maybe a Win32 SList
		wil::srwlock string_queue_lock;
	};

	wistd::unique_ptr<pending_build_info_t> _pending_build;

public:
	HRESULT InitInstance (IVsUIHierarchy* hier)
	{
		_hier = hier;
		_threadId = GetCurrentThreadId();
		_platformName = wil::make_bstr_nothrow(L"ZX Spectrum 48K"); RETURN_IF_NULL_ALLOC(_platformName);

		if (!atom)
		{
			WNDCLASS wc = { };
			wc.lpfnWndProc = window_proc;
			wc.hInstance = hinstance;
			wc.lpszClassName = wnd_class_name;
			atom.reset (RegisterClassW(&wc)); RETURN_LAST_ERROR_IF(!atom);
		}

		return S_OK;
	}

	static LRESULT CALLBACK window_proc (HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
	{
		if (msg == WM_OUTPUT_LINE)
		{
			auto p = reinterpret_cast<ProjectConfig*>(GetWindowLongPtr (hwnd, GWLP_USERDATA)); WI_ASSERT(p);
			auto lock = p->_pending_build->string_queue_lock.lock_exclusive();
			auto str = p->_pending_build->string_queue.remove(p->_pending_build->string_queue.begin());
			lock.reset();
			if (p->_pending_build->output_window_no_pump)
				p->_pending_build->output_window_no_pump->OutputStringNoPump(str.data());
			else
				p->_pending_build->output_window->OutputStringThreadSafe(str.data());
			return 0;
		}

		if (msg == WM_BUILD_COMPLETE)
		{
			auto p = reinterpret_cast<ProjectConfig*>(GetWindowLongPtr (hwnd, GWLP_USERDATA)); WI_ASSERT(p);
			WaitForSingleObject(p->_pending_build->thread_handle.get(), IsDebuggerPresent() ? INFINITE : 5000);
			DWORD exit_code_hr = E_FAIL;
			BOOL bres = GetExitCodeThread (p->_pending_build->thread_handle.get(), &exit_code_hr); LOG_IF_WIN32_BOOL_FALSE(bres);
			p->_pending_build.reset();

			for (auto rit = p->_buildStatusCallbacks.rbegin(); rit != p->_buildStatusCallbacks.rend(); rit++)
				rit->second->BuildEnd(SUCCEEDED(exit_code_hr));

			return 0;
		}

		return DefWindowProc (hwnd, msg, wparam, lparam);
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_POINTER;

		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IVsDebuggableProjectCfg*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProjectConfig>(this, riid, ppvObject)
			|| TryQI<IXmlParent>(this, riid, ppvObject)
			|| TryQI<IVsCfg>(this, riid, ppvObject)
			|| TryQI<IVsProjectCfg>(this, riid, ppvObject)
			|| TryQI<IVsDebuggableProjectCfg>(this, riid, ppvObject)
			|| TryQI<IVsBuildableProjectCfg>(this, riid, ppvObject)
			|| TryQI<IVsBuildableProjectCfg2>(this, riid, ppvObject)
			|| TryQI<ISpecifyPropertyPages>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		// These will never be implemented.
		if (   riid == IID_IManagedObject
			|| riid == IID_IInspectable
			|| riid == IID_ICustomTypeDescriptor
			|| riid == IID_IComponent
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_INoIdea3
			|| riid == IID_INoIdea4
			|| riid == IID_IConnectionPointContainer
			|| riid == IID_IPerPropertyBrowsing
///			|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IProvideMultipleClassInfo
			|| riid == IID_IWeakReferenceSource
			|| riid == IID_ICustomCast
		)
			return E_NOINTERFACE;

		// These may be implemented at a later time.
		if (   riid == IID_IVsDeployableProjectCfg
			|| riid == IID_IVsPublishableProjectCfg
			|| riid == IID_IVsProjectCfg2
			|| riid == IID_ICategorizeProperties
			|| riid == IID_IProvidePropertyBuilder
			|| riid == IID_IVsDebuggableProjectCfg2
			|| riid == IID_IVsPerPropertyBrowsing
		)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IProjectConfig);

	#pragma region IVsCfg
	virtual HRESULT STDMETHODCALLTYPE get_DisplayName(BSTR * pbstrDisplayName) override
	{
		wstring_builder sb;
		sb << _configName.get() << L"|" << _platformName.get();
		*pbstrDisplayName = SysAllocStringLen(sb.data(), sb.size()); RETURN_IF_NULL_ALLOC(*pbstrDisplayName);
		return S_OK;
	}

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_IsDebugOnly(BOOL * pfIsDebugOnly) override { return E_NOTIMPL; }

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_IsReleaseOnly(BOOL * pfIsReleaseOnly) override { return E_NOTIMPL; }
	#pragma endregion

	#pragma region IVsProjectCfg
	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE EnumOutputs(IVsEnumOutputs ** ppIVsEnumOutputs) override { return E_NOTIMPL; }

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE OpenOutput(LPCOLESTR szOutputCanonicalName, IVsOutput ** ppIVsOutput) override { return E_NOTIMPL; }

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_ProjectCfgProvider(IVsProjectCfgProvider ** ppIVsProjectCfgProvider) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE get_BuildableProjectCfg(IVsBuildableProjectCfg ** ppIVsBuildableProjectCfg) override
	{
		*ppIVsBuildableProjectCfg = this;
		AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_CanonicalName(BSTR* pbstrCanonicalName) override
	{
		*pbstrCanonicalName = SysAllocString(_configName.get()); RETURN_IF_NULL_ALLOC(*pbstrCanonicalName);
		return S_OK;
	}

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_Platform(GUID * pguidPlatform) override { return E_NOTIMPL; }

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_IsPackaged(BOOL * pfIsPackaged) override { return E_NOTIMPL; }

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_IsSpecifyingOutputSupported(BOOL * pfIsSpecifyingOutputSupported) override { return E_NOTIMPL; }

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_TargetCodePage(UINT * puiTargetCodePage) override { return E_NOTIMPL; }

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_UpdateSequenceNumber(ULARGE_INTEGER * puliUSN) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE get_RootURL(BSTR * pbstrRootURL) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	HRESULT MakeLaunchOptionsString (BSTR* pOptionsString)
	{
		com_ptr<IFelixLaunchOptions> opts;
		auto hr = MakeLaunchOptions(&opts); RETURN_IF_FAILED(hr);

		wil::unique_variant projectDir;
		_hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr);

		RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);
		hr = opts->put_ProjectDir(projectDir.bstrVal); RETURN_IF_FAILED(hr);
		
		com_ptr<IProjectConfigDebugProperties> dbgProps;
		hr = this->get_DebuggingProperties(&dbgProps); RETURN_IF_FAILED(hr);
		hr = opts->put_DebuggingProperties(dbgProps); RETURN_IF_FAILED(hr);

		com_ptr<IStream> stream;
		hr = CreateStreamOnHGlobal (nullptr, TRUE, &stream); RETURN_IF_FAILED(hr);
		UINT UTF16CodePage = 1200;
		hr = SaveToXml (opts.get(), L"Options", stream.get(), UTF16CodePage); RETURN_IF_FAILED(hr);

		STATSTG stat;
		hr = stream->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(ERROR_FILE_TOO_LARGE, !!stat.cbSize.HighPart);

		HGLOBAL hg;
		hr = GetHGlobalFromStream(stream.get(), &hg); RETURN_IF_FAILED(hr);
		auto buffer = GlobalLock(hg); WI_ASSERT(buffer);
		*pOptionsString = SysAllocStringLen((OLECHAR*)buffer, stat.cbSize.LowPart / 2);
		GlobalUnlock (hg);
		RETURN_IF_NULL_ALLOC(*pOptionsString);
		
		return S_OK;
	}

	#pragma region IVsDebuggableProjectCfg
	virtual HRESULT STDMETHODCALLTYPE DebugLaunch(VSDBGLAUNCHFLAGS grfLaunch) override
	{
		// https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/launching-a-program?view=vs-2022

		wil::com_ptr_nothrow<IVsDebugger> debugger;
		auto hr = serviceProvider->QueryService (SID_SVsShellDebugger, &debugger); RETURN_IF_FAILED(hr);

		wil::unique_bstr output_dir;
		hr = GetOutputDirectory(&output_dir); RETURN_IF_FAILED(hr);
		wil::unique_bstr output_filename;
		hr = GetOutputFileName(&output_filename); RETURN_IF_FAILED(hr);

		DWORD PathFlags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
		wil::unique_hlocal_string exe_path;
		hr = PathAllocCombine (output_dir.get(), output_filename.get(), PathFlags, &exe_path); RETURN_IF_FAILED(hr);
		auto exePathBstr = wil::make_bstr_nothrow(exe_path.get()); RETURN_IF_NULL_ALLOC(exePathBstr);

		if (grfLaunch & DBGLAUNCH_NoDebug)
		{
			wil::com_ptr_nothrow<IVsUIShell> uiShell;
			hr = serviceProvider->QueryService(SID_SVsUIShell, &uiShell);
			if (SUCCEEDED(hr))
			{
				HWND parent;
				hr = uiShell->GetDialogOwnerHwnd(&parent); RETURN_IF_FAILED(hr);
				::MessageBox (parent, L"Starting without debugging is not yet implemented.\r\n\r\n"
					L"For now you can Start Debugging (F5 by default) and \"ignore\" the debugger.", L"Felix", 0);
				return S_OK;
			}

			return E_NOTIMPL;
		}
		else
		{
			WI_ASSERT(!grfLaunch); // TODO: read these flags and act on them

			wil::unique_bstr options;
			hr = MakeLaunchOptionsString(&options); RETURN_IF_FAILED(hr);

			auto portNameBstr = wil::make_bstr_nothrow(SingleDebugPortName); RETURN_IF_NULL_ALLOC(portNameBstr);
			VsDebugTargetInfo dti = { };
			dti.cbSize = sizeof(dti);
			dti.dlo = DLO_CreateProcess;
			dti.clsidCustom = Engine_Id;
			dti.bstrExe = exePathBstr.get();
			dti.clsidPortSupplier = PortSupplier_Id;
			dti.bstrPortName = portNameBstr.get();
			dti.bstrOptions = options.get();
			dti.grfLaunch = grfLaunch;
			dti.fSendStdoutToOutputWindow = TRUE;

			hr = debugger->LaunchDebugTargets (1, &dti);
			if (hr == OLE_E_PROMPTSAVECANCELLED)
				hr = E_ABORT;
			RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryDebugLaunch(VSDBGLAUNCHFLAGS grfLaunch, BOOL * pfCanLaunch) override
	{
		// if (!DebuggerRunning())...
		*pfCanLaunch = TRUE;
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsBuildableProjectCfg
	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE get_ProjectCfg(IVsProjectCfg ** ppIVsProjectCfg) override { RETURN_HR(E_NOTIMPL); }

	virtual HRESULT STDMETHODCALLTYPE AdviseBuildStatusCallback(IVsBuildStatusCallback * pIVsBuildStatusCallback, VSCOOKIE * pdwCookie) override
	{
		WI_ASSERT (GetCurrentThreadId() == _threadId);
		bool inserted = _buildStatusCallbacks.try_insert({ _buildStatusNextCookie, pIVsBuildStatusCallback }); RETURN_HR_IF(E_OUTOFMEMORY, !inserted);
		*pdwCookie = _buildStatusNextCookie;
		_buildStatusNextCookie++;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseBuildStatusCallback(VSCOOKIE dwCookie) override
	{
		WI_ASSERT (GetCurrentThreadId() == _threadId);
		auto it = _buildStatusCallbacks.find(dwCookie);
		if (it == _buildStatusCallbacks.end())
			return E_INVALIDARG;
		_buildStatusCallbacks.erase(it);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE StartBuild(IVsOutputWindowPane * pIVsOutputWindowPane, DWORD dwOptions) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE StartClean(IVsOutputWindowPane * pIVsOutputWindowPane, DWORD dwOptions) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE StartUpToDateCheck(IVsOutputWindowPane * pIVsOutputWindowPane, DWORD dwOptions) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryStatus(BOOL * pfBuildDone) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Stop(BOOL fSync) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	[[deprecated]]
	virtual HRESULT STDMETHODCALLTYPE Wait(DWORD dwMilliseconds, BOOL fTickWhenMessageQNotEmpty) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE QueryStartBuild(DWORD dwOptions, BOOL * pfSupported, BOOL * pfReady) override
	{
		if (pfSupported)
			*pfSupported = TRUE;
		if (pfReady)
			*pfReady = TRUE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryStartClean(DWORD dwOptions, BOOL * pfSupported, BOOL * pfReady) override
	{
		if (pfSupported)
			*pfSupported = TRUE;
		if (pfReady)
			*pfReady = TRUE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryStartUpToDateCheck(DWORD dwOptions, BOOL * pfSupported, BOOL * pfReady) override
	{
		if (pfSupported)
			*pfSupported = FALSE;
		if (pfReady)
			*pfReady = TRUE;
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsBuildableProjectCfg2
	virtual HRESULT STDMETHODCALLTYPE GetBuildCfgProperty(VSBLDCFGPROPID propid, VARIANT * pvar) override
	{
		if (propid == VSBLDCFGPROPID_SupportsMTBuild)
			return InitVariantFromBoolean (TRUE, pvar);

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE StartBuildEx(DWORD dwBuildId, IVsOutputWindowPane * pIVsOutputWindowPane, DWORD dwOptions) override
	{
		HRESULT hr;

		for (auto& p : _buildStatusCallbacks)
		{
			BOOL fContinue;
			hr = p.second->BuildBegin(&fContinue); RETURN_IF_FAILED(hr);
			if (!fContinue)
				RETURN_HR(OLECMDERR_E_CANCELED);
		}

		wil::unique_variant project_dir;
		hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &project_dir); RETURN_IF_FAILED(hr);
		if (project_dir.vt != VT_BSTR)
			return E_FAIL;

		wil::unique_bstr output_dir;
		hr = this->GetOutputDirectory(&output_dir); RETURN_IF_FAILED(hr);

		int win32err = SHCreateDirectoryExW (nullptr, output_dir.get(), nullptr);
		if (win32err != ERROR_SUCCESS && win32err != ERROR_ALREADY_EXISTS)
			RETURN_WIN32(win32err);

		const DWORD pathFlags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;

		wstring_builder cmdLine;

		wil::unique_hlocal_string moduleFilename;
		hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, moduleFilename); RETURN_IF_FAILED(hr);
		auto fnres = PathFindFileName(moduleFilename.get()); RETURN_HR_IF(CO_E_BAD_PATH, fnres == moduleFilename.get());
		cmdLine << L'\"';
		cmdLine.append(moduleFilename.get(), fnres - moduleFilename.get());
		cmdLine << "sjasmplus.exe" << L'\"';

		cmdLine << " --fullpath";

		auto addOutputPathParam = [&cmdLine, &output_dir, &project_dir](const char* paramName, const wchar_t* output_filename) -> HRESULT
			{
				wil::unique_hlocal_string outputFilePath;
				auto hr = PathAllocCombine (output_dir.get(), output_filename, pathFlags, &outputFilePath); RETURN_IF_FAILED(hr);
				auto outputFilePathRelativeUgly = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePathRelativeUgly);
				BOOL bRes = PathRelativePathToW (outputFilePathRelativeUgly.get(), project_dir.bstrVal, FILE_ATTRIBUTE_DIRECTORY, outputFilePath.get(), 0); RETURN_HR_IF(CS_E_INVALID_PATH, !bRes);
				size_t len = wcslen(outputFilePathRelativeUgly.get());
				auto outputFilePathRelative = wil::make_hlocal_string_nothrow(nullptr, len); RETURN_IF_NULL_ALLOC(outputFilePathRelative);
				hr = PathCchCanonicalizeEx (outputFilePathRelative.get(), len + 1, outputFilePathRelativeUgly.get(), pathFlags); RETURN_IF_FAILED(hr);
				cmdLine << paramName << outputFilePathRelative.get();
				return S_OK;
			};

		// --raw=...
		wil::unique_bstr output_filename;
		hr = GetOutputFileName(&output_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (" --raw=", output_filename.get()); RETURN_IF_FAILED(hr);

		// --sld=...
		wil::unique_bstr sld_filename;
		hr = GetSldFileName (&sld_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (" --sld=", sld_filename.get()); RETURN_IF_FAILED(hr);

		// --outprefix
		hr = addOutputPathParam (" --outprefix=", L""); RETURN_IF_FAILED(hr);

		// --lst
		if (_saveListing)
		{
			cmdLine << " --lst";
			if (_listingFilename && _listingFilename.get()[0])
				cmdLine << "=" << _listingFilename.get();
		}

		// input files
		wil::unique_variant childItemId;
		hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &childItemId); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, childItemId.vt != VT_VSITEMID);
		while (V_VSITEMID(&childItemId) != VSITEMID_NIL)
		{
			wil::unique_variant browseObjectVariant;
			hr = _hier->GetProperty(V_VSITEMID(&childItemId), VSHPROPID_BrowseObject, &browseObjectVariant);
			if (SUCCEEDED(hr) && (browseObjectVariant.vt == VT_DISPATCH))
			{
				com_ptr<IProjectFile> file;
				hr = browseObjectVariant.pdispVal->QueryInterface(&file);
				if (SUCCEEDED(hr))
				{
					BuildToolKind tool;
					hr = file->get_BuildTool(&tool);
					if (SUCCEEDED(hr) && (tool == BuildToolKind::Assembler))
					{
						wil::unique_bstr fileRelativePath;
						hr = file->get_Path(&fileRelativePath);
						if (SUCCEEDED(hr))
							cmdLine << ' ' << fileRelativePath.get();
					}
				}
			}

			hr = _hier->GetProperty(V_VSITEMID(&childItemId), VSHPROPID_NextSibling, &childItemId); RETURN_IF_FAILED(hr);
			RETURN_HR_IF(E_FAIL, childItemId.vt != VT_VSITEMID);
		}

		cmdLine << '\0';
		RETURN_HR_IF(E_OUTOFMEMORY, cmdLine.out_of_memory());

		// Create a pipe for the child process's STDOUT.
		wil::unique_handle stdoutReadHandle;
		wil::unique_handle stdoutWriteHandle;
		SECURITY_ATTRIBUTES saAttr = { .nLength = sizeof(SECURITY_ATTRIBUTES), .bInheritHandle = TRUE, };
		BOOL bres = CreatePipe(&stdoutReadHandle, &stdoutWriteHandle, &saAttr, 0); RETURN_IF_WIN32_BOOL_FALSE(bres);

		// Ensure the read handle to the pipe for STDOUT is not inherited.
		bres = SetHandleInformation(stdoutReadHandle.get(), HANDLE_FLAG_INHERIT, 0); RETURN_IF_WIN32_BOOL_FALSE(bres);

		STARTUPINFO startupInfo = { .cb = sizeof(startupInfo) };
		startupInfo.hStdError = stdoutWriteHandle.get();
		startupInfo.hStdOutput = stdoutWriteHandle.get();
		startupInfo.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
		startupInfo.dwFlags |= STARTF_USESTDHANDLES;

		pIVsOutputWindowPane->OutputString (cmdLine.data());
		pIVsOutputWindowPane->OutputString (L"\r\n");

		wil::unique_process_information process_info;
		bres = CreateProcess(NULL, const_cast<LPWSTR>(cmdLine.data()), NULL, NULL, TRUE, CREATE_NO_WINDOW | CREATE_UNICODE_ENVIRONMENT, NULL,
			project_dir.bstrVal, &startupInfo, &process_info); RETURN_IF_WIN32_BOOL_FALSE(bres);

		wil::unique_event_nothrow exit_request;
		hr = exit_request.create(wil::EventOptions::None); RETURN_IF_FAILED(hr);

		auto hwnd = wil::unique_hwnd (CreateWindowExW (0, wnd_class_name, L"", WS_CHILD, 0, 0, 0, 0, HWND_MESSAGE, 0, hinstance, this)); RETURN_LAST_ERROR_IF(!hwnd);
		SetWindowLongPtr (hwnd.get(), GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));

		auto thread_handle = wil::unique_handle (CreateThread (nullptr, 0, ReadBuildOutputThreadProcStatic, this, CREATE_SUSPENDED, nullptr)); RETURN_LAST_ERROR_IF_NULL(thread_handle);

		wil::com_ptr_nothrow<IVsOutputWindowPaneNoPump> output_window_no_pump;
		pIVsOutputWindowPane->QueryInterface(&output_window_no_pump); // we don't want error checking here

		_pending_build.reset(new (std::nothrow) pending_build_info_t{ }); RETURN_IF_NULL_ALLOC(_pending_build);
		_pending_build->output_window = pIVsOutputWindowPane;
		_pending_build->output_window_no_pump = std::move(output_window_no_pump);
		_pending_build->exit_request_event = std::move(exit_request);
		_pending_build->process_info = std::move(process_info);
		_pending_build->stdoutReadHandle = std::move(stdoutReadHandle);
		_pending_build->hwnd = std::move(hwnd);
		_pending_build->thread_handle = std::move(thread_handle);

		DWORD dwres = ResumeThread (_pending_build->thread_handle.get());
		if (dwres == (DWORD)-1)
		{
			DWORD le = GetLastError();
			_pending_build.reset();
			RETURN_WIN32(le);
		}

		WI_ASSERT(dwres == 1);

		_pending_build->thread_started_event.wait();

		return S_OK;
	}

	static DWORD WINAPI ReadBuildOutputThreadProcStatic (void* arg)
	{
		auto cfg = static_cast<ProjectConfig*>(arg);
		return cfg->ReadBuildOutputThreadProc();
	}

	DWORD ReadBuildOutputThreadProc()
	{
		_pending_build->thread_started_event.SetEvent();

		const HANDLE handles[2] = { _pending_build->exit_request_event.get(), _pending_build->stdoutReadHandle.get() };

		vector_nothrow<wchar_t> lineBuffer;

		while(true)
		{
			DWORD wait_res = WaitForMultipleObjects (2, handles, FALSE, INFINITE);
			if (wait_res == WAIT_OBJECT_0)
			{
				BOOL posted = PostMessageW (_pending_build->hwnd.get(), WM_BUILD_COMPLETE, 0, 0); LOG_IF_WIN32_BOOL_FALSE(posted);
				return S_OK;
			}
			else if (wait_res == WAIT_OBJECT_0 + 1)
			{
				// Read all available data and split into lines.
				while (true)
				{
					char buffer[100];
					DWORD bytes_read;
					BOOL bres = ReadFile (_pending_build->stdoutReadHandle.get(), buffer, (DWORD)sizeof(buffer), &bytes_read, nullptr);
					if (!bres)
					{
						DWORD le = GetLastError();
						if (le == ERROR_BROKEN_PIPE)
						{
							// The build process probably ended.
							DWORD exit_code_process = 1;
							bres = GetExitCodeProcess (_pending_build->process_info.hProcess, &exit_code_process); LOG_IF_WIN32_BOOL_FALSE(bres);
							bres = PostMessageW (_pending_build->hwnd.get(), WM_BUILD_COMPLETE, 0, 0); LOG_IF_WIN32_BOOL_FALSE(bres);
							return exit_code_process ? E_FAIL : S_OK;
						}

						LOG_WIN32(le);
						BOOL posted = PostMessageW (_pending_build->hwnd.get(), WM_BUILD_COMPLETE, 0, 0); LOG_IF_WIN32_BOOL_FALSE(posted);
						return HRESULT_FROM_WIN32(le);
					}

					// To keep things simple, reserve for the worst case: line ends at the end of the read buffer, plus nul-terminator
					bool reserved = lineBuffer.try_reserve(lineBuffer.size() + sizeof(buffer) + 1);
					if (!reserved)
					{
						BOOL posted = PostMessageW (_pending_build->hwnd.get(), WM_BUILD_COMPLETE, 0, 0); LOG_IF_WIN32_BOOL_FALSE(posted);
						RETURN_HR(E_OUTOFMEMORY);
					}

					for (uint32_t i = 0; i < bytes_read; i++)
					{
						lineBuffer.try_push_back(buffer[i]); // Not optimal, I know. I'll sort this out later.
						if (buffer[i] == '\x0A')
						{
							// We have one line. Forward it to the main thread.
							lineBuffer.try_push_back('\0');

							auto lock = _pending_build->string_queue_lock.lock_exclusive();
							bool pushed = _pending_build->string_queue.try_push_back(std::move(lineBuffer));
							if (!pushed)
							{
								BOOL posted = PostMessageW (_pending_build->hwnd.get(), WM_BUILD_COMPLETE, 0, 0); LOG_IF_WIN32_BOOL_FALSE(posted);
								RETURN_HR(E_OUTOFMEMORY);
							}

							BOOL posted = PostMessageW (_pending_build->hwnd.get(), WM_OUTPUT_LINE, 0, 0);
							if (!posted)
							{
								DWORD le = GetLastError();
								_pending_build->string_queue.remove(_pending_build->string_queue.end() - 1);
								RETURN_HR(HRESULT_FROM_WIN32(le));
							}
						}
					}
				}
			}
			else
			{
				DWORD le = GetLastError();
				LOG_WIN32(le);
				return HRESULT_FROM_WIN32(le);
			}
		}
	}
	#pragma endregion

	#pragma region ISpecifyPropertyPages
	virtual HRESULT STDMETHODCALLTYPE GetPages (CAUUID *pPages) override
	{
		pPages->pElems = (GUID*)CoTaskMemAlloc (2 * sizeof(GUID)); RETURN_IF_NULL_ALLOC(pPages->pElems);
		pPages->pElems[0] = GeneralPropertyPage_CLSID;
		pPages->pElems[1] = DebugPropertyPage_CLSID;
		pPages->cElems = 2;
		return S_OK;
	}
	#pragma endregion

	#pragma region IProjectConfig
	virtual HRESULT STDMETHODCALLTYPE get___id(BSTR *value) override
	{
		// For configurations, VS seems to request this and then ignore it.
		*value = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_AssemblerProperties (IProjectConfigAssemblerProperties** ppDispatch) override
	{
		com_ptr<IProjectConfigAssemblerProperties> props;
		auto hr = AssemblerPageProperties_CreateInstance (&props); RETURN_IF_FAILED(hr);
		hr = props->put_SaveListing (_saveListing ? VARIANT_TRUE : VARIANT_FALSE); RETURN_IF_FAILED(hr);
		hr = props->put_SaveListingFilename (_listingFilename.get()); RETURN_IF_FAILED(hr);
		*ppDispatch = props.detach();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_AssemblerProperties (IProjectConfigAssemblerProperties* gp) override
	{
		bool changed = false;

		VARIANT_BOOL saveListing;
		auto hr = gp->get_SaveListing(&saveListing); RETURN_IF_FAILED(hr);
		if (_saveListing != (saveListing == VARIANT_TRUE))
		{
			_saveListing = (saveListing == VARIANT_TRUE);
			changed = true;
		}

		wil::unique_bstr listingFilename;
		hr = gp->get_SaveListingFilename(&listingFilename);
		if (VarBstrCmp(_listingFilename.get(), listingFilename.get(), 0, 0) != VARCMP_EQ)
		{
			_listingFilename = std::move(listingFilename);
			changed = true;
		}

		if (changed)
		{
			com_ptr<IPropertyNotifySink> pns;
			hr = _hier->QueryInterface(&pns); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
				pns->OnChanged(dispidConfigurations);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_GeneralProperties (IProjectConfigAssemblerProperties* pDispatch) override
	{
		return put_AssemblerProperties(pDispatch);
	}

	virtual HRESULT STDMETHODCALLTYPE get_DebuggingProperties (IProjectConfigDebugProperties** ppDispatch) override
	{
		com_ptr<IProjectConfigDebugProperties> props;
		auto hr = DebuggingPageProperties_CreateInstance (&props); RETURN_IF_FAILED(hr);
		hr = props->put_LoadAddress(_loadAddress); RETURN_IF_FAILED(hr);
		hr = props->put_EntryPointAddress(_entryPointAddress); RETURN_IF_FAILED(hr);
		hr = props->put_LaunchType(_launchType); RETURN_IF_FAILED(hr);
		*ppDispatch = props.detach();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_DebuggingProperties (IProjectConfigDebugProperties* pDispatch) override
	{
		wil::com_ptr_nothrow<IProjectConfigDebugProperties> dp;
		auto hr = pDispatch->QueryInterface(&dp); RETURN_IF_FAILED(hr);

		bool changed = false;

		WORD entryPointAddress;
		hr = dp->get_EntryPointAddress(&entryPointAddress); RETURN_IF_FAILED(hr);
		if (_entryPointAddress != entryPointAddress)
		{
			_entryPointAddress = entryPointAddress;
			changed = true;
		}

		unsigned long loadAddress;
		hr = dp->get_LoadAddress(&loadAddress); RETURN_IF_FAILED(hr);
		if (_loadAddress != loadAddress)
		{
			_loadAddress = loadAddress;
			changed = true;
		}

		LaunchType launchType;
		hr = dp->get_LaunchType(&launchType); RETURN_IF_FAILED(hr);
		if (_launchType != launchType)
		{
			_launchType = launchType;
			changed = true;
		}

		if (changed)
		{
			com_ptr<IPropertyNotifySink> pns;
			hr = _hier->QueryInterface(&pns); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
				pns->OnChanged(dispidConfigurations);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_ConfigName (BSTR* pbstr) override
	{
		*pbstr = SysAllocString (_configName.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_ConfigName (BSTR value) override
	{
		// TODO: notifications
		auto newName = wil::make_bstr_nothrow(value); RETURN_IF_NULL_ALLOC(newName); 
		_configName = std::move(newName);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_PlatformName (BSTR* pbstr) override
	{
		*pbstr = SysAllocString (_platformName.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_PlatformName (BSTR value) override
	{
		// TODO: notifications
		auto newName = wil::make_bstr_nothrow(value); RETURN_IF_NULL_ALLOC(newName); 
		_platformName = std::move(newName);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetOutputDirectory (BSTR* pbstr) override
	{
		wil::unique_variant project_dir;
		auto hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &project_dir); RETURN_IF_FAILED(hr);
		if (project_dir.vt != VT_BSTR)
			return E_FAIL;

		size_t output_dir_cap = wcslen(project_dir.bstrVal) + 10 + wcslen(_configName.get() + 1);
		auto output_dir = wil::make_process_heap_string_nothrow(nullptr, output_dir_cap); RETURN_IF_NULL_ALLOC(output_dir);
		wcscpy_s (output_dir.get(), output_dir_cap, project_dir.bstrVal);
		wcscat_s (output_dir.get(), output_dir_cap, L"\\bin\\");
		wcscat_s (output_dir.get(), output_dir_cap, _configName.get());
		wcscat_s (output_dir.get(), output_dir_cap, L"\\");

		*pbstr = SysAllocString(output_dir.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetOutputFileName (BSTR* pbstr) override
	{
		*pbstr = SysAllocString(L"output.bin"); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSldFileName (BSTR* pbstr) override
	{
		*pbstr = SysAllocString(L"output.sld"); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}
	#pragma endregion

	#pragma region IXmlParent
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IUnknown* child, BSTR* xmlElementNameOut) override
	{
		if (dispidProperty == dispidAssemblerProperties)
		{
			*xmlElementNameOut = nullptr;
			return S_OK;
		}

		if (dispidProperty == dispidDebuggingProperties)
		{
			*xmlElementNameOut = nullptr;
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE CreateChild (DISPID dispidProperty, PCWSTR xmlElementName, IDispatch** childOut) override
	{
		if (dispidProperty == dispidAssemblerProperties || dispidProperty == dispidGeneralProperties)
			return AssemblerPageProperties_CreateInstance ((IProjectConfigAssemblerProperties**)childOut);

		if (dispidProperty == dispidDebuggingProperties)
			return DebuggingPageProperties_CreateInstance ((IProjectConfigDebugProperties**)childOut);

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT ProjectConfig_CreateInstance (IVsUIHierarchy* hier, IProjectConfig** to)
{
	auto p = com_ptr(new (std::nothrow) ProjectConfig()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(hier); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}

struct AssemblerPageProperties
	: IProjectConfigAssemblerProperties
	, IProvideClassInfo
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
{
	ULONG _refCount = 0;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	bool _saveListing = false;
	wil::unique_bstr _listingFilename;

	static HRESULT CreateInstance (IProjectConfigAssemblerProperties** to)
	{
		HRESULT hr;

		auto p = com_ptr(new (std::nothrow) AssemblerPageProperties()); RETURN_IF_NULL_ALLOC(p);

		hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(p, &p->_propNotifyCP); RETURN_IF_FAILED(hr);

		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_POINTER;

		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProjectConfigAssemblerProperties>(this, riid, ppvObject)
			|| TryQI<IProvideClassInfo>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		// These will never be implemented.
		if (   riid == IID_IManagedObject
			|| riid == IID_IInspectable
			|| riid == IID_ICustomTypeDescriptor
			|| riid == IID_IComponent
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_INoIdea3
			|| riid == IID_INoIdea4
			|| riid == IID_IPerPropertyBrowsing
//			|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_IProvideMultipleClassInfo
			|| riid == IID_IWeakReferenceSource
			)
			return E_NOINTERFACE;

		if (riid == IID_ICategorizeProperties)
			return E_NOINTERFACE;

		if (riid == IID_IProvidePropertyBuilder)
			return E_NOINTERFACE;

		if (riid == IID_ISpecifyPropertyPages)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IProjectConfigAssemblerProperties);

	#pragma region IProvideClassInfo
	virtual HRESULT STDMETHODCALLTYPE GetClassInfo (ITypeInfo **ppTI) override
	{
		return GetTypeInfo(0, 0, ppTI);
	}
	#pragma endregion

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL* pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidOutputFileType)
		{
			*fDefault = TRUE;
			return S_OK;
		}

		if (dispid == dispidSaveListing)
		{
			*fDefault = (_saveListing == false);
			return S_OK;
		}

		if (dispid == dispidListingFilename)
		{
			*fDefault = !_listingFilename || !_listingFilename.get()[0];
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetClassName (BSTR *pbstrClassName) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE CanResetPropertyValue (DISPID dispid, BOOL *pfCanReset) override
	{
		switch (dispid)
		{
			case dispidOutputFileType:
			case dispidSaveListing:
			case dispidListingFilename:
				*pfCanReset = TRUE;
				return S_OK;
			default:
				return E_NOTIMPL;
		}
	}

	virtual HRESULT STDMETHODCALLTYPE ResetPropertyValue (DISPID dispid) override
	{
		if (dispid == dispidSaveListing)
			return put_SaveListing(VARIANT_FALSE);

		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IConnectionPointContainer
	virtual HRESULT STDMETHODCALLTYPE EnumConnectionPoints (IEnumConnectionPoints **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE FindConnectionPoint (REFIID riid, IConnectionPoint **ppCP) override
	{
		if (riid == IID_IPropertyNotifySink)
		{
			*ppCP = _propNotifyCP;
			_propNotifyCP->AddRef();
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IProjectConfigAssemblerProperties
	virtual HRESULT STDMETHODCALLTYPE get___id(BSTR *value) override
	{
		// For configurations, this seems to be requested and then ignored.
		*value = nullptr;
		return S_OK;;
	}

	virtual HRESULT STDMETHODCALLTYPE get_OutputFileType (OutputFileType* value) override
	{
		*value = OutputFileType::Binary;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_OutputFileType (OutputFileType value) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_SaveListing (VARIANT_BOOL* save) override
	{
		*save = _saveListing ? VARIANT_TRUE: VARIANT_FALSE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_SaveListing (VARIANT_BOOL save) override
	{
		bool s = (save == VARIANT_TRUE);
		if (_saveListing != s)
		{
			_saveListing = s;
			_propNotifyCP->NotifyPropertyChanged(dispidSaveListing);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_SaveListingFilename (BSTR* pFilename) override
	{
		if (_listingFilename)
		{
			*pFilename = SysAllocString(_listingFilename.get()); RETURN_IF_NULL_ALLOC(*pFilename);
		}
		else
			*pFilename = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_SaveListingFilename (BSTR filename) override
	{
		if (VarBstrCmp(_listingFilename.get(), filename, 0, 0) != VARCMP_EQ)
		{
			auto fn = wil::make_bstr_nothrow(filename); RETURN_IF_NULL_ALLOC(fn);
			_listingFilename = std::move(fn);
			_propNotifyCP->NotifyPropertyChanged(dispidListingFilename);
		}

		return S_OK;
	}
	#pragma endregion
};

struct DebuggingPageProperties
	: IProjectConfigDebugProperties
	, IProvideClassInfo
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
{
	ULONG _refCount = 0;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	DWORD _loadAddress = LoadAddressDefaultValue;
	WORD _entryPointAddress = EntryPointAddressDefaultValue;
	LaunchType _launchType = LaunchTypeDefaultValue;

	static HRESULT CreateInstance (IProjectConfigDebugProperties** to)
	{
		HRESULT hr;

		com_ptr<DebuggingPageProperties> p = new (std::nothrow) DebuggingPageProperties(); RETURN_IF_NULL_ALLOC(p);
		hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(p, &p->_propNotifyCP); RETURN_IF_FAILED(hr);
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_POINTER;

		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProjectConfigDebugProperties>(this, riid, ppvObject)
			|| TryQI<IProvideClassInfo>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
		)
			return S_OK;

		// These will never be implemented.
		if (   riid == IID_IManagedObject
			|| riid == IID_IInspectable
			|| riid == IID_ICustomTypeDescriptor
			|| riid == IID_IComponent
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_INoIdea3
			|| riid == IID_INoIdea4
			|| riid == IID_IPerPropertyBrowsing
			//|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_IProvideMultipleClassInfo
			|| riid == IID_IWeakReferenceSource
			)
			return E_NOINTERFACE;

		else if (riid == IID_ICategorizeProperties)
			return E_NOINTERFACE;
		else if (riid == IID_IProvidePropertyBuilder)
			return E_NOINTERFACE;
		else if (riid == IID_IConnectionPointContainer)
			return E_NOINTERFACE;


		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IProjectConfigDebugProperties)

	#pragma region IProvideClassInfo
	virtual HRESULT STDMETHODCALLTYPE GetClassInfo (ITypeInfo **ppTI) override
	{
		return GetTypeInfo(0, 0, ppTI);
	}
	#pragma endregion

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL* pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidLoadAddress)
		{
			*fDefault = (_loadAddress == LoadAddressDefaultValue);
			return S_OK;
		}

		if (dispid == dispidEntryPointAddress)
		{
			*fDefault = (_entryPointAddress == EntryPointAddressDefaultValue);
			return S_OK;
		}

		if (dispid == dispidLaunchType)
		{
			*fDefault = (_launchType == LaunchTypeDefaultValue);
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetClassName (BSTR *pbstrClassName) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE CanResetPropertyValue (DISPID dispid, BOOL *pfCanReset) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE ResetPropertyValue (DISPID dispid) override { return E_NOTIMPL; }
	#pragma endregion
	
	#pragma region IConnectionPointContainer
	virtual HRESULT STDMETHODCALLTYPE EnumConnectionPoints (IEnumConnectionPoints **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE FindConnectionPoint (REFIID riid, IConnectionPoint **ppCP) override
	{
		if (riid == IID_IPropertyNotifySink)
		{
			*ppCP = _propNotifyCP;
			_propNotifyCP->AddRef();
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IZ80ProjectConfigDebugProperties
	virtual HRESULT STDMETHODCALLTYPE get___id(BSTR *value) override
	{
		// For configurations, this seems to be requested and then ignored.
		*value = nullptr;
		return S_OK;;
	}

	virtual HRESULT STDMETHODCALLTYPE get_LoadAddress (DWORD* value) override
	{
		*value = _loadAddress;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_LoadAddress (DWORD value) override
	{
		if (_loadAddress != value)
		{
			_loadAddress = value;
			_propNotifyCP->NotifyPropertyChanged(dispidLoadAddress);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_EntryPointAddress (WORD* value) override
	{
		*value = _entryPointAddress;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_EntryPointAddress (WORD value) override
	{
		if (_entryPointAddress != value)
		{
			_entryPointAddress = value;
			_propNotifyCP->NotifyPropertyChanged(dispidEntryPointAddress);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_LaunchType (enum LaunchType *value) override
	{
		*value = _launchType;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_LaunchType (enum LaunchType value) override
	{
		if (_launchType != value)
		{
			_launchType = value;
			_propNotifyCP->NotifyPropertyChanged(dispidLaunchType);
		}

		return S_OK;
	}
	#pragma endregion
};

HRESULT AssemblerPageProperties_CreateInstance (IProjectConfigAssemblerProperties** to)
{
	return AssemblerPageProperties::CreateInstance(to);
}

HRESULT DebuggingPageProperties_CreateInstance (IProjectConfigDebugProperties** to)
{
	return DebuggingPageProperties::CreateInstance(to);
}

