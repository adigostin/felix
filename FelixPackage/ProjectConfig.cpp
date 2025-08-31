
#include "pch.h"
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/com.h"
#include "dispids.h"
#include "guids.h"
#include "Z80Xml.h"
#include "../FelixPackageUi/resource.h"
#include <string_view>

// Useful doc: https://learn.microsoft.com/en-us/visualstudio/extensibility/internals/managing-configuration-options?view=vs-2022

static constexpr DWORD BaseAddressDefaultValue = 0x8000;
static constexpr wchar_t EntryPointAddressDefaultValue[] = L"32768";

HRESULT GeneralPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigGeneralProperties** to);
HRESULT AssemblerPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigAssemblerProperties** to);
HRESULT DebuggingPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigDebugProperties** to);
HRESULT PrePostBuildPageProperties_CreateInstance (bool post, IProjectConfigPrePostBuildProperties** to);

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
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _hier;
	DWORD _threadId;
	wil::unique_process_heap_string _configName;
	wil::unique_process_heap_string _platformName;
	unordered_map_nothrow<VSCOOKIE, com_ptr<IVsBuildStatusCallback>> _buildStatusCallbacks;
	VSCOOKIE _buildStatusNextCookie = VSCOOKIE_NIL + 1;
	com_ptr<IProjectConfigGeneralProperties> _generalProps;
	AdviseSinkToken _generalPropsAdviseToken;
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
		_platformName = wil::make_process_heap_string_nothrow(L"ZX Spectrum 48K"); RETURN_IF_NULL_ALLOC(_platformName);
		
		hr = _weakRefToThis.InitInstance(static_cast<IProjectConfig*>(this)); RETURN_IF_FAILED(hr);

		hr = GeneralPageProperties_CreateInstance(this, &_generalProps); RETURN_IF_FAILED(hr);
		hr = AdviseSink<IPropertyNotifySink>(_generalProps, _weakRefToThis, &_generalPropsAdviseToken); RETURN_IF_FAILED(hr);

		hr = AssemblerPageProperties_CreateInstance(this, &_assemblerProps); RETURN_IF_FAILED(hr);
		hr = AdviseSink<IPropertyNotifySink>(_assemblerProps, _weakRefToThis, &_assemblerPropsAdviseToken); RETURN_IF_FAILED(hr);

		hr = DebuggingPageProperties_CreateInstance(this, &_debugProps); RETURN_IF_FAILED(hr);
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
			|| riid == IID_TypeDescriptor_IUnimplemented
			|| riid == IID_PropertyGrid_IUnimplemented
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
		wil::unique_process_heap_string s;
		auto hr = wil::str_printf_nothrow(s, L"%s|%s", _configName.get(), _platformName.get());
		*pbstrDisplayName = SysAllocString(s.get()); RETURN_IF_NULL_ALLOC(*pbstrDisplayName);
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

	HRESULT MakeLaunchOptionsString (IFelixSymbols* symbols, BSTR* pOptionsString)
	{
		com_ptr<IFelixLaunchOptions> opts;
		auto hr = MakeLaunchOptions(&opts); RETURN_IF_FAILED(hr);

		com_ptr<IVsHierarchy> hier;
		hr = _hier->QueryInterface(IID_PPV_ARGS(hier.addressof())); RETURN_IF_FAILED(hr);

		wil::unique_variant projectDir;
		hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr);

		RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);
		hr = opts->put_ProjectDir(projectDir.bstrVal); RETURN_IF_FAILED(hr);
		
		DWORD baseAddress;
		hr = _assemblerProps->get_BaseAddress(&baseAddress); RETURN_IF_FAILED(hr);
		opts->put_BaseAddress(baseAddress);

		wil::unique_bstr epAddressStr;
		hr = _assemblerProps->get_EntryPointAddress(&epAddressStr); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, !epAddressStr || !epAddressStr.get()[0]);
		if (isdigit(epAddressStr.get()[0]))
		{
			DWORD epAddress;
			hr = ParseNumber(epAddressStr.get(), &epAddress);
			if (hr != S_OK)
				return E_FAIL; // TODO: detailed error message
			opts->put_EntryPointAddress(epAddress);
		}
		else
		{
			UINT16 epAddress;
			hr = symbols->GetAddressFromSymbol(epAddressStr.get(), &epAddress);
			if (hr != S_OK)
				return E_FAIL;
			opts->put_EntryPointAddress(epAddress);
		}

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
		hr = _generalProps->get_OutputDirectory(&output_dir); RETURN_IF_FAILED(hr);
		wil::unique_bstr output_filename;
		hr = _generalProps->get_OutputFilename(&output_filename); RETURN_IF_FAILED(hr);

		wil::unique_bstr launchTarget;
		hr = _debugProps->get_LaunchTarget(&launchTarget); RETURN_IF_FAILED(hr);
		wil::unique_process_heap_string exePath;
		hr = ResolveMacros (launchTarget.get(), this, exePath); RETURN_IF_FAILED(hr);
		auto exePathBstr = wil::make_bstr_nothrow(exePath.get()); RETURN_IF_NULL_ALLOC(exePathBstr);

		auto sldFilename = wil::make_process_heap_string_nothrow(output_filename.get(), wcslen(output_filename.get()) + 10); RETURN_IF_NULL_ALLOC(sldFilename);
		PathRenameExtension(sldFilename.get(), L".sld");
		wil::unique_process_heap_string sldPath;
		hr = wil::str_concat_nothrow(sldPath, output_dir, sldFilename); RETURN_IF_FAILED(hr);
		hr = ResolveMacros (sldPath.get(), this, sldPath); RETURN_IF_FAILED(hr);
		com_ptr<IFelixSymbols> symbols;
		hr = MakeSldSymbols(sldPath.get(), &symbols); RETURN_IF_FAILED(hr);
		
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
			hr = MakeLaunchOptionsString(symbols, &options); RETURN_IF_FAILED(hr);

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

		hr = MakeProjectConfigBuilder (project, this, op2, &_pendingBuild); RETURN_IF_FAILED(hr);

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
		pPages->pElems = (GUID*)CoTaskMemAlloc (5 * sizeof(GUID)); RETURN_IF_NULL_ALLOC(pPages->pElems);
		pPages->pElems[0] = GeneralPropertyPage_CLSID;
		pPages->pElems[1] = AssemblerPropertyPage_CLSID;
		pPages->pElems[2] = DebugPropertyPage_CLSID;
		pPages->pElems[3] = PreBuildPropertyPage_CLSID;
		pPages->pElems[4] = PostBuildPropertyPage_CLSID;
		pPages->cElems = 5;
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

	virtual IProjectConfigGeneralProperties* GeneralProps() override { return _generalProps; }

	virtual IProjectConfigAssemblerProperties* AsmProps() override { return _assemblerProps; }

	virtual IProjectConfigProperties* AsProjectConfigProperties() override { return this; }
	
	virtual IVsProjectCfg* AsVsProjectConfig() override { return this; }
	#pragma endregion

	#pragma region IProjectConfigProperties
	virtual HRESULT STDMETHODCALLTYPE get___id(BSTR *value) override
	{
		// For configurations, VS seems to request this and then ignore it.
		*value = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_GeneralProperties (IProjectConfigGeneralProperties** ppProps) override
	{
		return wil::com_copy_to_nothrow(_generalProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE get_AssemblerProperties (IProjectConfigAssemblerProperties** ppProps) override
	{
		return wil::com_copy_to_nothrow(_assemblerProps, ppProps);
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
		*pbstr = SysAllocString(_configName.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_ConfigName (BSTR value) override
	{
		// TODO: notifications
		auto newName = wil::make_process_heap_string_nothrow(value); RETURN_IF_NULL_ALLOC(newName); 
		_configName = std::move(newName);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_PlatformName (BSTR* pbstr) override
	{
		*pbstr = SysAllocString(_platformName.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_PlatformName (BSTR value) override
	{
		// TODO: notifications
		auto newName = wil::make_process_heap_string_nothrow(value); RETURN_IF_NULL_ALLOC(newName); 
		_platformName = std::move(newName);
		return S_OK;
	}
	#pragma endregion

	#pragma region IXmlParent
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IDispatch* child, BSTR* xmlElementNameOut) override
	{
		switch (dispidProperty)
		{
			case dispidGeneralProperties:
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
		if (dispidProperty == dispidGeneralProperties)
			return GeneralPageProperties_CreateInstance (this, (IProjectConfigGeneralProperties**)childOut);

		if (dispidProperty == dispidAssemblerProperties)
			return AssemblerPageProperties_CreateInstance (this, (IProjectConfigAssemblerProperties**)childOut);

		if (dispidProperty == dispidDebuggingProperties)
			return DebuggingPageProperties_CreateInstance (this, (IProjectConfigDebugProperties**)childOut);

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
			if (   dispID == dispidBaseAddress
				|| dispID == dispidEntryPointAddress
				|| dispID == dispidOutputFileType
				|| dispID == dispidPlatformName)
			{
				com_ptr<IProjectNode> proj;
				if (SUCCEEDED(_hier->QueryInterface(IID_PPV_ARGS(&proj))))
				{
					hr = GeneratePrePostIncludeFiles(proj, this); RETURN_IF_FAILED(hr);
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

// ============================================================================

static HRESULT MapPropertyToBuilder_MacroResolver (LONG dispid, LONG* pdwCtlBldType, BSTR* pbstrGuidBldr, VARIANT_BOOL* pfRetVal)
{
	if (pdwCtlBldType)
		*pdwCtlBldType = CTLBLDTYPE_FINTERNALBUILDER;
	if (pbstrGuidBldr)
		*pbstrGuidBldr = SysAllocString(L"E84F7AF7-6EF0-43D3-8A4D-9D716C2B5596");
	*pfRetVal = VARIANT_TRUE;
	return S_OK;
}

static HRESULT ExecuteBuilder_MacroResolver (LONG dispid, VARIANT* pvarValue, VARIANT_BOOL* pfRetVal, IWeakRef* configWeakRef)
{
	if (pvarValue->vt == VT_EMPTY)
		return S_OK;

	RETURN_HR_IF(E_UNEXPECTED, pvarValue->vt != VT_BSTR || !pvarValue->bstrVal);

	com_ptr<IProjectConfig> config;
	auto hr = configWeakRef->QueryInterface(IID_PPV_ARGS(&config)); RETURN_IF_FAILED(hr);

	wil::unique_process_heap_string resolved;
	hr = ResolveMacros (pvarValue->bstrVal, config, resolved); RETURN_IF_FAILED(hr);

	com_ptr<IVsShell> shell;
	hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);
	wil::unique_bstr format;
	hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_PROPERTY_EVALUATES_TO, &format); RETURN_IF_FAILED(hr);
	wil::unique_process_heap_string message;
	hr = wil::str_printf_nothrow(message, format.get(), pvarValue->bstrVal, resolved.get());

	com_ptr<IVsUIShell> uiShell;
	hr = serviceProvider->QueryService(SID_SVsUIShell, IID_PPV_ARGS(&uiShell)); RETURN_IF_FAILED(hr);		
	LONG res;
	hr = uiShell->ShowMessageBox (0, IID_NULL, NULL, message.get(), NULL, 0, OLEMSGBUTTON_OK, OLEMSGDEFBUTTON_FIRST, OLEMSGICON_INFO, FALSE, &res); RETURN_IF_FAILED(hr);
	return S_OK;
}

// ============================================================================

static const wchar_t OutputNameDefaultValue[] = L"%PROJECT_NAME%";
static const OutputFileType OutputTypeDefaultValue = OutputFileType::Sna;
static const wchar_t OutputDirectoryDefaultValue[] = L"%PROJECT_DIR%Out\\%CONFIG_NAME%\\";

struct GeneralPageProperties
	: IProjectConfigGeneralProperties
	, IConnectionPointContainer
	, IVsPerPropertyBrowsing
	, IProvidePropertyBuilder
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _config;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	wil::unique_process_heap_string _outputName;
	OutputFileType _outputFileType = OutputTypeDefaultValue;
	wil::unique_process_heap_string _outputDirectory;

	HRESULT InitInstance (IProjectConfig* config)
	{
		HRESULT hr;
		_outputName = wil::make_process_heap_string_nothrow(OutputNameDefaultValue); RETURN_IF_NULL_ALLOC(_outputName);
		_outputDirectory = wil::make_process_heap_string_nothrow(OutputDirectoryDefaultValue); RETURN_IF_NULL_ALLOC(_outputDirectory);
		hr = config->QueryInterface(IID_PPV_ARGS(_config.addressof())); RETURN_IF_FAILED(hr);
		hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(this, &_propNotifyCP); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	~GeneralPageProperties()
	{
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProjectConfigGeneralProperties>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IProvidePropertyBuilder>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IProjectConfigGeneralProperties);

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

	#pragma region IProjectConfigGeneralProperties
	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		*value = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_ProjectName (BSTR* pbstrProjectName) override
	{
		com_ptr<IProjectConfig> config;
		auto hr = _config->QueryInterface(IID_PPV_ARGS(&config)); RETURN_IF_FAILED(hr);
		com_ptr<IVsHierarchy> hier;
		hr = config->GetSite(IID_PPV_ARGS(&hier)); RETURN_IF_FAILED(hr);
		wil::unique_variant name;
		hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &name); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_FAIL, name.vt != VT_BSTR);
		*pbstrProjectName = SysAllocString(name.release().bstrVal);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_OutputName (BSTR* pbstrTargetName) override
	{
		*pbstrTargetName = SysAllocString(_outputName.get()); RETURN_IF_NULL_ALLOC(*pbstrTargetName);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_OutputName (BSTR bstrTargetName) override
	{
		const wchar_t* tn = bstrTargetName ? bstrTargetName : L"";
		if (wcscmp(_outputName.get(), tn))
		{
			_outputName = wil::make_process_heap_string_nothrow(tn); RETURN_IF_NULL_ALLOC(_outputName);
			_propNotifyCP->NotifyPropertyChanged(dispidOutputName);
			_propNotifyCP->NotifyPropertyChanged(dispidOutputFilename);
		}
		
		return S_OK;
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
			_propNotifyCP->NotifyPropertyChanged(dispidOutputFilename);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_OutputFilename (BSTR *pbstrOutputFilename) override
	{
		wil::unique_process_heap_string fn;
		auto hr = wil::str_printf_nothrow(fn, L"%s%s", _outputName.get(), GetOutputExtensionFromOutputType(_outputFileType)); RETURN_IF_FAILED(hr);
		com_ptr<IProjectConfig> config;
		hr = _config->QueryInterface(IID_PPV_ARGS(&config)); RETURN_IF_FAILED(hr);
		wil::unique_process_heap_string resolved;
		hr = ResolveMacros(fn.get(), config, resolved); RETURN_IF_FAILED(hr);
		auto out = wil::make_bstr_nothrow(resolved.get()); RETURN_IF_NULL_ALLOC(out);
		*pbstrOutputFilename = out.release();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_OutputDirectory (BSTR *pbstrOutputDirectory) override
	{
		*pbstrOutputDirectory = SysAllocString(_outputDirectory.get()); RETURN_IF_NULL_ALLOC(*pbstrOutputDirectory);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_OutputDirectory (BSTR bstrOutputDirectory) override
	{
		if (wcscmp(_outputDirectory.get(), bstrOutputDirectory))
		{
			auto od = wil::make_process_heap_string_nothrow(bstrOutputDirectory); RETURN_IF_NULL_ALLOC(od);
			_outputDirectory = std::move(od);
			_propNotifyCP->NotifyPropertyChanged(dispidOutputDirectory);
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL* pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR* pbstrLocalizedName, BSTR* pbstrLocalizeDescription) override
	{
		HRESULT hr;

		if (dispid == dispidProjectName)
		{
			com_ptr<IVsShell> shell;
			hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(shell.addressof())); RETURN_IF_FAILED(hr);
			if (pbstrLocalizedName)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERAL_PROPS_PROJ_NAME_NAME, pbstrLocalizedName); LOG_IF_FAILED(hr);
			}

			if (pbstrLocalizeDescription)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERAL_PROPS_PROJ_NAME_DESCRIPTION, pbstrLocalizeDescription); LOG_IF_FAILED(hr);
			}

			return S_OK;
		}

		if (dispid == dispidOutputDirectory)
		{
			com_ptr<IVsShell> shell;
			hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(shell.addressof())); RETURN_IF_FAILED(hr);
			if (pbstrLocalizedName)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERAL_PROPS_OUTPUT_DIR_NAME, pbstrLocalizedName); LOG_IF_FAILED(hr);
			}

			if (pbstrLocalizeDescription)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERAL_PROPS_OUTPUT_DIR_DESC, pbstrLocalizeDescription); LOG_IF_FAILED(hr);
			}

			return S_OK;
		}

		if (dispid == dispidOutputName)
		{
			com_ptr<IVsShell> shell;
			hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(shell.addressof())); RETURN_IF_FAILED(hr);
			if (pbstrLocalizedName)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERAL_PROPS_OUTPUT_NAME_NAME, pbstrLocalizedName); LOG_IF_FAILED(hr);
			}

			if (pbstrLocalizeDescription)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERAL_PROPS_OUTPUT_NAME_DESCRIPTION, pbstrLocalizeDescription); LOG_IF_FAILED(hr);
			}

			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidProjectName)
			return (*fDefault = TRUE), S_OK;

		if (dispid == dispidOutputName)
		{
			*fDefault = !wcscmp(_outputName.get(), OutputNameDefaultValue);
			return S_OK;
		}

		if (dispid == dispidOutputFileType)
		{
			*fDefault = (_outputFileType == OutputTypeDefaultValue);
			return S_OK;
		}

		if (dispid == dispidOutputFilename)
			return (*fDefault = TRUE), S_OK;

		if (dispid == dispidOutputDirectory)
		{
			*fDefault = !wcscmp(_outputDirectory.get(), OutputDirectoryDefaultValue);
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetClassName (BSTR *pbstrClassName) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE CanResetPropertyValue (DISPID dispid, BOOL *pfCanReset) override
	{
		if (dispid == dispidOutputName)
			return (*pfCanReset = TRUE), S_OK;

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE ResetPropertyValue (DISPID dispid) override
	{
		if (dispid == dispidOutputName)
		{
			if (wcscmp(_outputName.get(), OutputNameDefaultValue))
			{
				auto tn = wil::make_process_heap_string_nothrow(OutputNameDefaultValue); RETURN_IF_NULL_ALLOC(tn);
				_outputName = std::move(tn);
				_propNotifyCP->NotifyPropertyChanged(dispidOutputName);
			}

			return S_OK;
		}

		if (dispid == dispidOutputFileType)
			return put_OutputFileType(OutputFileType::Binary);

		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IProvidePropertyBuilder
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToBuilder (LONG dispid, LONG* pdwCtlBldType, BSTR* pbstrGuidBldr, VARIANT_BOOL* pfRetVal) override
	{
		if (dispid == dispidOutputName || dispid == dispidOutputDirectory)
			return MapPropertyToBuilder_MacroResolver (dispid, pdwCtlBldType, pbstrGuidBldr, pfRetVal);

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE ExecuteBuilder (LONG dispid, BSTR bstrGuidBldr, IDispatch *pdispApp, LONG_PTR hwndBldrOwner, VARIANT *pvarValue, VARIANT_BOOL *pfRetVal) override
	{
		return ExecuteBuilder_MacroResolver (dispid, pvarValue, pfRetVal, _config);
	}
	#pragma endregion
};

static HRESULT GeneralPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigGeneralProperties** to)
{
	auto p = com_ptr(new (std::nothrow) GeneralPageProperties()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(config); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}

// ============================================================================

struct AssemblerPageProperties
	: IProjectConfigAssemblerProperties
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
	, IProvidePropertyBuilder
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _config;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	wil::unique_bstr _entryPointAddress;
	DWORD _baseAddress = BaseAddressDefaultValue;
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
			//|| TryQI<IVSMDPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
			|| TryQI<IProvidePropertyBuilder>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		// These will never be implemented - they're called from managed code, for example from Marshal.GetObjectForIUnknown()
		if (   riid == IID_IManagedObject
			|| riid == IID_IInspectable
			|| riid == IID_IComponent
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IWeakReferenceSource
			|| riid == IID_IProvideClassInfo
		)
			return E_NOINTERFACE;

		if (	riid == __uuidof(IXmlParent))
			return E_NOINTERFACE;

		if (   riid == IID_PropertyGrid_IUnimplemented
			|| riid == IID_TypeDescriptor_IUnimplemented
			|| riid == IID_ICustomTypeDescriptor
			|| riid == IID_IProvideMultipleClassInfo
			//|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_ISpecifyPropertyPages
		)
			return E_NOINTERFACE;

		if (riid == IID_IPerPropertyBrowsing)
			return E_NOINTERFACE;

		if (riid == IID_ICategorizeProperties)
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

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override
	{
		HRESULT hr;

		if (dispid == dispidEntryPointAddress)
		{
			com_ptr<IVsShell> shell;
			hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(shell.addressof())); RETURN_IF_FAILED(hr);
			if (pbstrLocalizedName)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_ASM_PROPS_ENTRY_POINT_ADDR_NAME, pbstrLocalizedName); LOG_IF_FAILED(hr);
			}

			if (pbstrLocalizeDescription)
			{
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_ASM_PROPS_ENTRY_POINT_ADDR_DESC, pbstrLocalizeDescription); LOG_IF_FAILED(hr);
			}

			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidBaseAddress)
		{
			*fDefault = (_baseAddress == BaseAddressDefaultValue);
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

	virtual HRESULT STDMETHODCALLTYPE get_BaseAddress (DWORD* value) override
	{
		*value = _baseAddress;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_BaseAddress (DWORD value) override
	{
		if (_baseAddress != value)
		{
			_baseAddress = value;
			_propNotifyCP->NotifyPropertyChanged(dispidBaseAddress);
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

	virtual HRESULT STDMETHODCALLTYPE get_ListingFilename (BSTR* pFilename) override
	{
		if (_listingFilename)
		{
			*pFilename = SysAllocString(_listingFilename.get()); RETURN_IF_NULL_ALLOC(*pFilename);
		}
		else
			*pFilename = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_ListingFilename (BSTR filename) override
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

	#pragma region IProvidePropertyBuilder
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToBuilder (LONG dispid, LONG* pdwCtlBldType, BSTR* pbstrGuidBldr, VARIANT_BOOL* pfRetVal) override
	{
		if (dispid == dispidListingFilename)
			return MapPropertyToBuilder_MacroResolver (dispid, pdwCtlBldType, pbstrGuidBldr, pfRetVal);

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE ExecuteBuilder (LONG dispid, BSTR bstrGuidBldr, IDispatch *pdispApp, LONG_PTR hwndBldrOwner, VARIANT *pvarValue, VARIANT_BOOL *pfRetVal) override
	{
		return ExecuteBuilder_MacroResolver (dispid, pvarValue, pfRetVal, _config);
	}
	#pragma endregion
};

static HRESULT AssemblerPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigAssemblerProperties** to)
{
	auto p = com_ptr (new (std::nothrow) AssemblerPageProperties()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(config); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}

// ============================================================================

static const LaunchType LaunchTypeDefaultValue = LaunchType::PrintUsr;
static const wchar_t LaunchTargetDefaultValue[] = L"%OUTPUT_DIR%%OUTPUT_FILENAME%";

struct DebuggingPageProperties
	: IProjectConfigDebugProperties
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
	, IProvidePropertyBuilder
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _config;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	LaunchType _launchType = LaunchTypeDefaultValue;
	wil::unique_process_heap_string _launchTarget;

	HRESULT InitInstance (IProjectConfig* config)
	{
		HRESULT hr;
		hr = config->QueryInterface(IID_PPV_ARGS(&_config)); RETURN_IF_FAILED(hr);
		hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(this, &_propNotifyCP); RETURN_IF_FAILED(hr);
		_launchTarget = wil::make_process_heap_string_nothrow(LaunchTargetDefaultValue); RETURN_IF_NULL_ALLOC(_launchTarget);
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
			|| TryQI<IProvidePropertyBuilder>(this, riid, ppvObject)
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
			|| riid == IID_TypeDescriptor_IUnimplemented
			|| riid == IID_PropertyGrid_IUnimplemented
			|| riid == IID_IPerPropertyBrowsing
			//|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_IProvideMultipleClassInfo
			|| riid == IID_IWeakReferenceSource
		)
			return E_NOINTERFACE;

		if (riid == IID_ICategorizeProperties)
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
		if (dispid == dispidLaunchTarget)
		{
			*fDefault = !wcscmp(_launchTarget.get(), LaunchTargetDefaultValue);
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

	virtual HRESULT STDMETHODCALLTYPE get_LaunchTarget(BSTR* pbstrLaunchTarget) override
	{
		auto b = SysAllocString(_launchTarget.get()); RETURN_IF_NULL_ALLOC(b);
		*pbstrLaunchTarget = b;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_LaunchTarget (BSTR bstrLaunchTarget) override
	{
		const wchar_t* from = bstrLaunchTarget ? bstrLaunchTarget : L"";
		if (wcscmp(_launchTarget.get(), from))
		{
			auto n = wil::make_process_heap_string_nothrow(from); RETURN_IF_NULL_ALLOC(n);
			_launchTarget = std::move(n);
			_propNotifyCP->NotifyPropertyChanged(dispidLaunchTarget);
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

	#pragma region IProvidePropertyBuilder
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToBuilder (LONG dispid, LONG* pdwCtlBldType, BSTR* pbstrGuidBldr, VARIANT_BOOL* pfRetVal) override
	{
		if (dispid == dispidLaunchTarget)
			return MapPropertyToBuilder_MacroResolver (dispid, pdwCtlBldType, pbstrGuidBldr, pfRetVal);

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE ExecuteBuilder (LONG dispid, BSTR bstrGuidBldr, IDispatch *pdispApp, LONG_PTR hwndBldrOwner, VARIANT *pvarValue, VARIANT_BOOL *pfRetVal) override
	{
		return ExecuteBuilder_MacroResolver (dispid, pvarValue, pfRetVal, _config);
	}
	#pragma endregion
};

static HRESULT DebuggingPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigDebugProperties** to)
{
	auto p = com_ptr(new (std::nothrow) DebuggingPageProperties()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(config); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}

// ============================================================================

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

static HRESULT PrePostBuildPageProperties_CreateInstance (bool post, IProjectConfigPrePostBuildProperties** to)
{
	return PrePostBuildPageProperties::CreateInstance(post, to);
}
