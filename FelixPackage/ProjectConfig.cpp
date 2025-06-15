
#include "pch.h"
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/com.h"
#include "shared/string_builder.h"
#include "dispids.h"
#include "guids.h"
#include "Z80Xml.h"
#include "../FelixPackageUi/resource.h"

// Useful doc: https://learn.microsoft.com/en-us/visualstudio/extensibility/internals/managing-configuration-options?view=vs-2022

static constexpr DWORD LoadAddressDefaultValue = 0x8000;
static constexpr wchar_t EntryPointAddressDefaultValue[] = L"32768";
static constexpr LaunchType LaunchTypeDefaultValue = LaunchType::PrintUsr;

struct ProjectConfig
	: IProjectConfig
	, IProjectConfigProperties
	, IVsDebuggableProjectCfg
	, IVsBuildableProjectCfg
	, IVsBuildableProjectCfg2
	, ISpecifyPropertyPages
	, IXmlParent
	, IProjectConfigBuilderCallback
	, IPropertyNotifySink
	//, public IVsProjectCfgDebugTargetSelection
	//, public IVsProjectCfgDebugTypeSelection
	, IProjectMacroResolver
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _hier;
	DWORD _threadId;
	wil::unique_bstr _configName;
	wil::unique_bstr _platformName;
	unordered_map_nothrow<VSCOOKIE, com_ptr<IVsBuildStatusCallback>> _buildStatusCallbacks;
	VSCOOKIE _buildStatusNextCookie = VSCOOKIE_NIL + 1;
	com_ptr<IProjectConfigAssemblerProperties> _assemblerProps;
	AdviseSinkToken _assemblerPropsAdviseToken;
	com_ptr<IProjectConfigDebugProperties> _debugProps;
	AdviseSinkToken _debugPropsAdviseToken;
	com_ptr<IProjectConfigPrePostBuildProperties> _preBuildProps;
	AdviseSinkToken _preBuildPropsAdviseToken;
	com_ptr<IProjectConfigPrePostBuildProperties> _postBuildProps;
	AdviseSinkToken _postBuildPropsAdviseToken;

	com_ptr<IProjectConfigBuilder> _pendingBuild;
	WeakRefToThis _weakRefToThis;

public:
	HRESULT InitInstance()
	{
		HRESULT hr;
		_threadId = GetCurrentThreadId();
		_platformName = wil::make_bstr_nothrow(L"ZX Spectrum 48K"); RETURN_IF_NULL_ALLOC(_platformName);
		
		hr = _weakRefToThis.InitInstance(static_cast<IProjectConfig*>(this)); RETURN_IF_FAILED(hr);

		hr = AssemblerPageProperties_CreateInstance(this, &_assemblerProps); RETURN_IF_FAILED(hr);
		hr = AdviseSink<IPropertyNotifySink>(_assemblerProps, _weakRefToThis, &_assemblerPropsAdviseToken); RETURN_IF_FAILED(hr);

		hr = DebuggingPageProperties_CreateInstance(&_debugProps); RETURN_IF_FAILED(hr);
		hr = AdviseSink<IPropertyNotifySink>(_debugProps, _weakRefToThis, &_debugPropsAdviseToken); RETURN_IF_FAILED(hr);

		hr = PrePostBuildPageProperties_CreateInstance(false, &_preBuildProps); RETURN_IF_FAILED(hr);
		hr = AdviseSink<IPropertyNotifySink>(_preBuildProps, _weakRefToThis, &_preBuildPropsAdviseToken); RETURN_IF_FAILED(hr);

		hr = PrePostBuildPageProperties_CreateInstance(true, &_postBuildProps); RETURN_IF_FAILED(hr);
		hr = AdviseSink<IPropertyNotifySink>(_postBuildProps, _weakRefToThis, &_postBuildPropsAdviseToken); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	~ProjectConfig()
	{
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_POINTER;

		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IVsDebuggableProjectCfg*>(this), riid, ppvObject)
			|| TryQI<IProjectConfig>(this, riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProjectConfigProperties>(this, riid, ppvObject)
			|| TryQI<IXmlParent>(this, riid, ppvObject)
			|| TryQI<IVsCfg>(this, riid, ppvObject)
			|| TryQI<IVsProjectCfg>(this, riid, ppvObject)
			|| TryQI<IVsDebuggableProjectCfg>(this, riid, ppvObject)
			|| TryQI<IVsBuildableProjectCfg>(this, riid, ppvObject)
			|| TryQI<IVsBuildableProjectCfg2>(this, riid, ppvObject)
			|| TryQI<ISpecifyPropertyPages>(this, riid, ppvObject)
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
			|| TryQI<IProjectMacroResolver>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

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

	IMPLEMENT_IDISPATCH(IID_IProjectConfigProperties);

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

		com_ptr<IVsHierarchy> hier;
		hr = _hier->QueryInterface(IID_PPV_ARGS(hier.addressof())); RETURN_IF_FAILED(hr);

		wil::unique_variant projectDir;
		hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr);

		RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);
		hr = opts->put_ProjectDir(projectDir.bstrVal); RETURN_IF_FAILED(hr);
		
		DWORD loadAddress;
		hr = _debugProps->get_LoadAddress(&loadAddress); RETURN_IF_FAILED(hr);
		wil::unique_bstr epAddressStr;
		hr = _assemblerProps->get_EntryPointAddress(&epAddressStr); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, !epAddressStr);
		wchar_t* endPtr;
		DWORD epAddress = wcstoul(epAddressStr.get(), &endPtr, 10); RETURN_HR_IF(E_NOTIMPL, endPtr != epAddressStr.get() + SysStringLen(epAddressStr.get()));
		opts->put_LoadAddress(loadAddress);
		opts->put_EntryPointAddress(epAddress);

		com_ptr<IStream> stream;
		hr = CreateStreamOnHGlobal (nullptr, TRUE, &stream); RETURN_IF_FAILED(hr);
		UINT UTF16CodePage = 1200;
		hr = SaveToXml (opts.get(), L"Options", 0, stream.get(), UTF16CodePage); RETURN_IF_FAILED(hr);

		return MakeBstrFromStreamOnHGlobal(stream, pOptionsString); 
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
		hr = _assemblerProps->GetOutputFileName(&output_filename); RETURN_IF_FAILED(hr);

		auto exe_path = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(exe_path);
		PathCombine (exe_path.get(), output_dir.get(), output_filename.get());
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
		RETURN_HR_IF(E_UNEXPECTED, !_pendingBuild);

		// In OnBuildComplete we're releasing our reference to _pendingBuild while code
		// is running on that object. Let's keep the object alive until after CancelBuild returns.
		auto keepAlive = com_ptr(_pendingBuild);

		return _pendingBuild->CancelBuild();
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

	// IProjectMacroResolver
	virtual HRESULT STDMETHODCALLTYPE ResolveMacro (const char* macroFrom, const char* macroTo, char** valueCoTaskMem) override
	{
		if (!strncmp(macroFrom, "LOAD_ADDR", 9))
		{
			unsigned long loadAddress;
			auto hr = _debugProps->get_LoadAddress(&loadAddress); RETURN_IF_FAILED(hr);
			char buffer[16];
			int ires = sprintf_s(buffer, "%u", loadAddress);
			auto value = wil::make_unique_ansistring_nothrow<wil::unique_cotaskmem_ansistring> (buffer, (size_t)ires); RETURN_IF_NULL_ALLOC(value);
			*valueCoTaskMem = value.release();
			return S_OK;
		}
		else if (!strncmp(macroFrom, "DEVICE", 6))
		{
			static const char name[] = "ZXSPECTRUM48";
			char* value = (char*)CoTaskMemAlloc(_countof(name)); RETURN_IF_NULL_ALLOC(value);
			strcpy(value, name);
			*valueCoTaskMem = value;
			return S_OK;
		}
		else if (!strncmp(macroFrom, "ENTRY_POINT_ADDR", 16))
		{
			wil::unique_bstr addr;
			auto hr = _assemblerProps->get_EntryPointAddress(&addr); RETURN_IF_FAILED(hr);
			int len = WideCharToMultiByte (CP_UTF8, 0, addr.get(), 1 + (int)SysStringLen(addr.get()), nullptr, 0, nullptr, nullptr); RETURN_LAST_ERROR_IF(len==0);
			auto s = wil::make_unique_ansistring_nothrow<wil::unique_cotaskmem_ansistring>(nullptr, len - 1); RETURN_IF_NULL_ALLOC(s);
			WideCharToMultiByte (CP_UTF8, 0, addr.get(), 1 + (int)SysStringLen(addr.get()), s.get(), len, nullptr, nullptr);
			*valueCoTaskMem = s.release();
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}

	#pragma region IVsBuildableProjectCfg2
	virtual HRESULT STDMETHODCALLTYPE GetBuildCfgProperty(VSBLDCFGPROPID propid, VARIANT * pvar) override
	{
		if (propid == VSBLDCFGPROPID_SupportsMTBuild)
			return InitVariantFromBoolean (TRUE, pvar);

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE StartBuildEx(DWORD dwBuildId, IVsOutputWindowPane* pIVsOutputWindowPane, DWORD dwOptions) override
	{
		HRESULT hr;

		// In VS2022, the "pIVsOutputWindowPane" passed to this function is a wrapper over
		// the real Output pane. This wrapper spams a CR/LR in every invocation of OutputTaskItemString/OutputTaskItemString/etc.
		com_ptr<IVsOutputWindow> ow;
		hr = serviceProvider->QueryService (SID_SVsOutputWindow, &ow); RETURN_IF_FAILED(hr);
		com_ptr<IVsOutputWindowPane> op;
		hr = ow->GetPane(GUID_BuildOutputWindowPane, &op); RETURN_IF_FAILED(hr);
		com_ptr<IVsOutputWindowPane2> op2;
		hr = op->QueryInterface(&op2); RETURN_IF_FAILED(hr);

		com_ptr<IProjectNode> project;
		hr = _hier->QueryInterface(IID_PPV_ARGS(project.addressof())); RETURN_IF_FAILED(hr);
		com_ptr<IVsShell> shell;
		hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);

		hr = MakeProjectConfigBuilder (project->AsHierarchy(), this, op2, &_pendingBuild); RETURN_IF_FAILED(hr);

		for (auto& cb : _buildStatusCallbacks)
		{
			BOOL fContinue = FALSE;
			hr = cb.second->BuildBegin(&fContinue); RETURN_IF_FAILED(hr);
			RETURN_HR_IF_EXPECTED(OLECMDERR_E_CANCELED, !fContinue);
		}

		hr = _pendingBuild->StartBuild(this);
		if (FAILED(hr))
		{
			for (auto& cb : _buildStatusCallbacks)
				cb.second->BuildEnd(FALSE);
			_pendingBuild = nullptr;
			return hr;
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IProjectConfigBuilderCallback
	virtual HRESULT OnBuildComplete (bool success) override
	{
		for (auto& cb : _buildStatusCallbacks)
			cb.second->BuildEnd(success);

		_pendingBuild = nullptr;
		return S_OK;
	}
	#pragma endregion

	#pragma region ISpecifyPropertyPages
	virtual HRESULT STDMETHODCALLTYPE GetPages (CAUUID *pPages) override
	{
		pPages->pElems = (GUID*)CoTaskMemAlloc (4 * sizeof(GUID)); RETURN_IF_NULL_ALLOC(pPages->pElems);
		pPages->pElems[0] = AssemblerPropertyPage_CLSID;
		pPages->pElems[1] = DebugPropertyPage_CLSID;
		pPages->pElems[2] = PreBuildPropertyPage_CLSID;
		pPages->pElems[3] = PostBuildPropertyPage_CLSID;
		pPages->cElems = 4;
		return S_OK;
	}
	#pragma endregion

	#pragma region IProjectConfig
	virtual HRESULT SetSite (IProjectNode* proj) override
	{
		return proj->QueryInterface(IID_PPV_ARGS(&_hier));
	}

	virtual HRESULT GetSite (REFIID riid, void** ppvObject) override
	{
		return _hier->QueryInterface(riid, ppvObject);
	}

	virtual HRESULT STDMETHODCALLTYPE GetOutputDirectory (BSTR* pbstr) override
	{
		com_ptr<IVsHierarchy> hier;
		auto hr = _hier->QueryInterface(IID_PPV_ARGS(hier.addressof())); RETURN_IF_FAILED(hr);
		wil::unique_variant project_dir;
		hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &project_dir); RETURN_IF_FAILED(hr);
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

	virtual IProjectConfigProperties* AsProjectConfigProperties() override { return this; }
	#pragma endregion

	#pragma region IProjectConfigProperties
	virtual HRESULT STDMETHODCALLTYPE get___id(BSTR *value) override
	{
		// For configurations, VS seems to request this and then ignore it.
		*value = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_AssemblerProperties (IProjectConfigAssemblerProperties** ppProps) override
	{
		return wil::com_copy_to_nothrow(_assemblerProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE put_GeneralProperties (IProjectConfigAssemblerProperties* pProps) override
	{
		_assemblerProps = pProps;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_DebuggingProperties (IProjectConfigDebugProperties** ppProps) override
	{
		return wil::com_copy_to_nothrow(_debugProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE get_PreBuildProperties (IProjectConfigPrePostBuildProperties **ppProps) override
	{
		return wil::com_copy_to_nothrow(_preBuildProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE get_PostBuildProperties (IProjectConfigPrePostBuildProperties **ppProps) override
	{
		return wil::com_copy_to_nothrow(_postBuildProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE get_ConfigName (BSTR* pbstr) override
	{
		return copy_bstr(_configName, pbstr);
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
		return copy_bstr(_platformName, pbstr);
	}

	virtual HRESULT STDMETHODCALLTYPE put_PlatformName (BSTR value) override
	{
		// TODO: notifications
		auto newName = wil::make_bstr_nothrow(value); RETURN_IF_NULL_ALLOC(newName); 
		_platformName = std::move(newName);
		return S_OK;
	}
	#pragma endregion

	#pragma region IXmlParent
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IDispatch* child, BSTR* xmlElementNameOut) override
	{
		switch (dispidProperty)
		{
			case dispidAssemblerProperties:
			case dispidDebuggingProperties:
			case dispidPreBuildProperties:
			case dispidPostBuildProperties:
				*xmlElementNameOut = nullptr;
				return S_OK;

			default:
				RETURN_HR(E_NOTIMPL);
		}
	}

	virtual HRESULT STDMETHODCALLTYPE CreateChild (DISPID dispidProperty, PCWSTR xmlElementName, IDispatch** childOut) override
	{
		if (dispidProperty == dispidAssemblerProperties || dispidProperty == dispidGeneralProperties)
			return AssemblerPageProperties_CreateInstance (this, (IProjectConfigAssemblerProperties**)childOut);

		if (dispidProperty == dispidDebuggingProperties)
			return DebuggingPageProperties_CreateInstance ((IProjectConfigDebugProperties**)childOut);

		if (dispidProperty == dispidPreBuildProperties)
			return PrePostBuildPageProperties_CreateInstance (false, (IProjectConfigPrePostBuildProperties**)childOut);

		if (dispidProperty == dispidPostBuildProperties)
			return PrePostBuildPageProperties_CreateInstance (true, (IProjectConfigPrePostBuildProperties**)childOut);

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IPropertyNotifySink
	virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
	{
		HRESULT hr;

		if (_hier)
		{
			if (   dispID == dispidLoadAddress
				|| dispID == dispidEntryPointAddress
				|| dispID == dispidPlatformName)
			{
				com_ptr<IProjectNode> proj;
				if (SUCCEEDED(_hier->QueryInterface(IID_PPV_ARGS(&proj))))
				{
					wil::unique_bstr name;
					hr = get_DisplayName(&name); RETURN_IF_FAILED(hr);
					GeneratePrePostIncludeFiles(proj, name.get(), this);
				}
			}
		
			com_ptr<IPropertyNotifySink> pns;
			hr = _hier->QueryInterface(&pns); RETURN_IF_FAILED(hr);
			pns->OnChanged(dispidConfigurations);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeProjectConfig (IProjectConfig** to)
{
	auto p = com_ptr(new (std::nothrow) ProjectConfig()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}

struct AssemblerPageProperties
	: IProjectConfigAssemblerProperties
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _config;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	OutputFileType _outputFileType = OutputFileType::Binary;
	wil::unique_bstr _entryPointAddress;
	bool _saveListing = false;
	wil::unique_bstr _listingFilename;

	HRESULT InitInstance (IProjectConfig* config)
	{
		HRESULT hr;
		hr = config->QueryInterface(IID_PPV_ARGS(_config.addressof())); RETURN_IF_FAILED(hr);
		hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(this, &_propNotifyCP); RETURN_IF_FAILED(hr);
		_entryPointAddress = wil::make_bstr_nothrow(EntryPointAddressDefaultValue); RETURN_IF_NULL_ALLOC(_entryPointAddress);
		return S_OK;
	}

	~AssemblerPageProperties()
	{
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


	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL* pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidOutputFileType)
		{
			*fDefault = (_outputFileType == OutputFileType::Binary);
			return S_OK;
		}

		if (dispid == dispidEntryPointAddress)
		{
			*fDefault = _entryPointAddress && !wcscmp(_entryPointAddress.get(), EntryPointAddressDefaultValue);
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
		if (dispid == dispidOutputFileType)
			return put_OutputFileType(OutputFileType::Binary);

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
		*value = _outputFileType;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_OutputFileType (OutputFileType value) override
	{
		if (_outputFileType != value)
		{
			_outputFileType = value;
			_propNotifyCP->NotifyPropertyChanged(dispidOutputFileType);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_EntryPointAddress (BSTR *pbstrAddress) override
	{
		if (!_entryPointAddress || !_entryPointAddress.get()[0])
			return (*pbstrAddress = nullptr), S_OK;

		auto res = SysAllocString(_entryPointAddress.get()); RETURN_IF_NULL_ALLOC(res);
		return (*pbstrAddress = res), S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_EntryPointAddress (BSTR bstrAddress) override
	{
		if (VarBstrCmp(_entryPointAddress.get(), bstrAddress, InvariantLCID, 0) != VARCMP_EQ)
		{
			BSTR newStr = nullptr;
			if (bstrAddress && bstrAddress[0])
			{
				newStr = SysAllocString(bstrAddress); RETURN_IF_NULL_ALLOC(newStr);
			}

			_entryPointAddress.reset(newStr);
			_propNotifyCP->NotifyPropertyChanged(dispidEntryPointAddress);
		}

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
};

struct DebuggingPageProperties
	: IProjectConfigDebugProperties
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
{
	ULONG _refCount = 0;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	DWORD _loadAddress = LoadAddressDefaultValue;
	LaunchType _launchType = LaunchTypeDefaultValue;

	static HRESULT CreateInstance (IProjectConfigDebugProperties** to)
	{
		HRESULT hr;

		com_ptr<DebuggingPageProperties> p = new (std::nothrow) DebuggingPageProperties(); RETURN_IF_NULL_ALLOC(p);
		hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(p, &p->_propNotifyCP); RETURN_IF_FAILED(hr);
		*to = p.detach();
		return S_OK;
	}

	~DebuggingPageProperties()
	{
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

	#pragma region IProjectConfigDebugProperties
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

FELIX_API HRESULT AssemblerPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigAssemblerProperties** to)
{
	auto p = com_ptr (new (std::nothrow) AssemblerPageProperties()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(config); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}

FELIX_API HRESULT DebuggingPageProperties_CreateInstance (IProjectConfigDebugProperties** to)
{
	return DebuggingPageProperties::CreateInstance(to);
}

struct PrePostBuildPageProperties
	: IProjectConfigPrePostBuildProperties
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
	, IProvidePropertyBuilder
{
	ULONG _refCount = 0;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	bool _post;
	wil::unique_bstr _commandLine;
	wil::unique_bstr _description;

	static HRESULT CreateInstance (bool post, IProjectConfigPrePostBuildProperties** to)
	{
		com_ptr<PrePostBuildPageProperties> p = new (std::nothrow) PrePostBuildPageProperties(); RETURN_IF_NULL_ALLOC(p);
		auto hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(p, &p->_propNotifyCP); RETURN_IF_FAILED(hr);
		p->_post = post;
		*to = p.detach();
		return S_OK;
	}

	~PrePostBuildPageProperties()
	{
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProjectConfigPrePostBuildProperties>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
			|| TryQI<IProvidePropertyBuilder>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IProjectConfigPrePostBuildProperties)

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL* pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override
	{
		if (dispid == dispidCommandLine)
		{
			if (pbstrLocalizeDescription)
			{
				wil::com_ptr_nothrow<IVsShell> shell;
				auto hr = serviceProvider->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);
				ULONG resid = _post ? IDS_POST_BUILD_CMD_LINE_DESCRIPTION : IDS_PRE_BUILD_CMD_LINE_DESCRIPTION;
				return shell->LoadPackageString(CLSID_FelixPackage, resid, pbstrLocalizeDescription);
			}

			return E_NOTIMPL;
		}

		if (dispid == dispidDescription)
		{
			if (pbstrLocalizeDescription)
			{
				wil::com_ptr_nothrow<IVsShell> shell;
				auto hr = serviceProvider->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);
				return shell->LoadPackageString(CLSID_FelixPackage, IDS_PRE_POST_BUILD_DESCRIPTION_DESCRIPTION, pbstrLocalizeDescription);
			}

			return E_NOTIMPL;
		}
		
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidCommandLine)
			return (*fDefault = !_commandLine || !_commandLine.get()[0]), S_OK;

		if (dispid == dispidDescription)
			return (*fDefault = !_description || !_description.get()[0]), S_OK;

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

	#pragma region IProvidePropertyBuilder
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToBuilder (LONG dispid, LONG* pdwCtlBldType, BSTR* pbstrGuidBldr, VARIANT_BOOL* pfRetVal) override
	{
		if (dispid == dispidCommandLine)
		{
			*pdwCtlBldType = CTLBLDTYPE_FINTERNALBUILDER;
			*pbstrGuidBldr = SysAllocString(GUID_CommandLineBuilderStr); RETURN_IF_NULL_ALLOC(*pbstrGuidBldr);
			*pfRetVal = VARIANT_TRUE;
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE ExecuteBuilder (LONG dispid, BSTR bstrGuidBldr, IDispatch* pdispApp,
		LONG_PTR hwndBldrOwner, VARIANT* pvarValue, VARIANT_BOOL* pfRetVal) override
	{
		if (!wcscmp(bstrGuidBldr, GUID_CommandLineBuilderStr))
		{
			RETURN_HR_IF(E_FAIL, pvarValue->vt != VT_BSTR && pvarValue->vt != VT_EMPTY);
			BSTR valueBefore = (pvarValue->vt == VT_BSTR && pvarValue->bstrVal) ? pvarValue->bstrVal : nullptr;
			wil::unique_bstr valueAfter;
			auto hr = ShowCommandLinePropertyBuilder((HWND)hwndBldrOwner, valueBefore, &valueAfter);
			if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
			{
				*pfRetVal = VARIANT_FALSE;
				return S_OK;
			}
			RETURN_IF_FAILED(hr);

			if (pvarValue->vt == VT_BSTR)
				SysFreeString(pvarValue->bstrVal);
			pvarValue->bstrVal = valueAfter.release();
			pvarValue->vt = VT_BSTR;
			*pfRetVal = VARIANT_TRUE;
			return S_OK;
		}

		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IProjectConfigPrePostBuildProperties
//	virtual HRESULT STDMETHODCALLTYPE get___id(BSTR *value) override
//	{
//		// For configurations, this seems to be requested and then ignored.
//		*value = nullptr;
//		return S_OK;
//	}

	virtual HRESULT STDMETHODCALLTYPE get_CommandLine (BSTR *value) override
	{
		if (_commandLine && _commandLine.get()[0])
		{
			*value = SysAllocStringLen(_commandLine.get(), SysStringLen(_commandLine.get())); RETURN_IF_NULL_ALLOC(*value);
		}
		else
			*value = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_CommandLine (BSTR value) override
	{
		if (VarBstrCmp(_commandLine.get(), value, 0, 0) != VARCMP_EQ)
		{
			if (value)
			{
				auto c = wil::make_bstr_nothrow(value); RETURN_IF_NULL_ALLOC(c);
				_commandLine = std::move(c);
			}
			else
				_commandLine = nullptr;
			_propNotifyCP->NotifyPropertyChanged(dispidCommandLine);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Description (BSTR *value) override
	{
		// Although a NULL BSTR has identical semantics as "", the Properties Window
		// handles a NULL BSTR by hiding the property, and "" by showing it empty.
		return (*value = SysAllocString(_description ? _description.get() : L"")) ? S_OK : E_OUTOFMEMORY;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Description (BSTR value) override
	{
		if (VarBstrCmp(_description.get(), value, 0, 0) != VARCMP_EQ)
		{
			if (value)
			{
				auto c = wil::make_bstr_nothrow(value); RETURN_IF_NULL_ALLOC(c);
				_description = std::move(c);
			}
			else
				_description = nullptr;
			_propNotifyCP->NotifyPropertyChanged(dispidDescription);
		}

		return S_OK;
	}
	#pragma endregion
};

FELIX_API HRESULT PrePostBuildPageProperties_CreateInstance (bool post, IProjectConfigPrePostBuildProperties** to)
{
	return PrePostBuildPageProperties::CreateInstance(post, to);
}
