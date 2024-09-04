
#include "pch.h"
#include "FelixPackage.h"
#include "Guids.h"
#include "../FelixPackageUi/CommandIds.h"
#include "../FelixPackageUi/Resource.h"
#include "shared/com.h"
#include "shared/OtherGuids.h"
#include "Simulator/Simulator.h"
#include "DebugEngine/DebugEngine.h"
#include "sentry.h"

const wchar_t Z80AsmLanguageName[]  = L"Z80Asm";
const wchar_t SingleDebugPortName[] = L"Single Z80 Port";
const GUID Z80AsmLanguageGuid = { 0x598BC226, 0x2E96, 0x43AD, { 0xAD, 0x42, 0x67, 0xD9, 0xCC, 0x6F, 0x75, 0xF6 } };
const wchar_t SettingsCollection[] = L"FelixSettings";
const wchar_t SettingLoadSavePath[] = L"LoadSavePath";
static const wchar_t AlwaysReportSettingsName[] = L"AlwaysReportErrors";
static const wchar_t BinaryFilename[] = L"ROMs/Spectrum48K.rom";

// These must be kept in sync with the pkgdef line [$RootKey$\InstalledProducts\FelixPackage]
//static const wchar_t InstalledProductRegPath[] = L"InstalledProducts\\FelixPackage";
static const char SentryReleaseName[] = "0.9.7";

// {8F0D9E89-4C6C-4B63-83CF-1AA6B6E59BCB}
GUID SID_Simulator = { 0x8f0d9e89, 0x4c6c, 0x4b63, { 0x83, 0xcf, 0x1a, 0xa6, 0xb6, 0xe5, 0x9b, 0xcb } };

// I know global variables are bad, but there are places in code where
// it's impossible to reach the IServiceProvider received in IVsPackage::SetSite.
// One such example is the implementation of IPropertyPage::GetPageInfo.
wil::com_ptr_nothrow<IServiceProvider> serviceProvider;

class FelixPackageImpl : public IVsPackage, IVsSolutionEvents, IOleCommandTarget, IDebugEventCallback2, IServiceProvider
{
	ULONG _refCount = 0;
	wil::unique_hlocal_string _packageDir;
	wil::com_ptr_nothrow<IVsLanguageInfo> _z80AsmLanguageInfo;
	wil::com_ptr_nothrow<IServiceProvider> _sp;
	static inline HMODULE _uiLibrary = nullptr;
	VSCOOKIE _projectTypeRegistrationCookie = VSCOOKIE_NIL;
	VSCOOKIE _solutionEventsCookie = VSCOOKIE_NIL;
	wil::com_ptr_nothrow<IVsDebugger> _debugger;
	bool _advisedDebugEventCallback = false;
	VSCOOKIE _profferLanguageServiceCookie = VSCOOKIE_NIL;
	VSCOOKIE _profferZXSimulatorServiceCookie = VSCOOKIE_NIL;
	wil::com_ptr_nothrow<IVsWindowFrame> _simulatorWindowFrame;
	com_ptr<ISimulator> _simulator;
	wil::ThreadFailureCache _threadFailureCache;
	sentry_options_t *_sentryOptions = nullptr;

public:
	HRESULT InitInstance()
	{
		wil::unique_hlocal_string moduleFilename;
		auto hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, moduleFilename); RETURN_IF_FAILED(hr);
		auto fnres = PathFindFileName(moduleFilename.get()); RETURN_HR_IF(CO_E_BAD_PATH, fnres == moduleFilename.get());
		_packageDir = wil::make_hlocal_string_nothrow(moduleFilename.get(), fnres - moduleFilename.get()); RETURN_IF_NULL_ALLOC(_packageDir);

		hr = Z80AsmLanguageInfo_CreateInstance(&_z80AsmLanguageInfo); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IVsPackage*>(this), riid, ppvObject)
			|| TryQI<IVsPackage>(this, riid, ppvObject)
			|| TryQI<IVsSolutionEvents>(this, riid, ppvObject)
			|| TryQI<IOleCommandTarget>(this, riid, ppvObject)
			|| TryQI<IDebugEventCallback2>(this, riid, ppvObject)
			|| TryQI<IServiceProvider>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		if (riid == IID_IManagedObject
			|| riid == IID_IConvertible
			|| riid == IID_IDispatch
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IProvideClassInfo2
			|| riid == IID_IInspectable
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IVsSolutionEvents2
			|| riid == IID_IVsSolutionEvents3
			|| riid == IID_IVsSolutionEvents4
			|| riid == IID_IVsSolutionEvents5
			|| riid == IID_IVsSolutionEvents6
			|| riid == IID_IVsSolutionEvents7
			|| riid == IID_IVsSolutionEvents8
			|| riid == IID_ICustomCast
			|| riid == IID_IAsyncLoadablePackageInitialize
			|| riid == IID_IVsAsyncToolWindowFactoryProvider
			|| riid == IID_IVsToolWindowFactory2Private
			|| riid == IID_IVsToolWindowFactory
			|| riid == IID_IVsPersistSolutionProps
			|| riid == IID_IVsPersistSolutionOpts
			|| riid == IID_IVsPersistSolutionOpts2
		)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	HRESULT CreateZxSpectrumSimulator()
	{
		wil::com_ptr_nothrow<IProfferService> srpProffer;
		auto hr = _sp->QueryService (SID_SProfferService, &srpProffer); RETURN_IF_FAILED(hr);
		
		hr = MakeSimulator(_packageDir.get(), BinaryFilename, &_simulator); RETURN_IF_FAILED(hr);
		hr = srpProffer->ProfferService (SID_Simulator, this, &_profferZXSimulatorServiceCookie); RETURN_IF_FAILED(hr);

		_simulator->Resume(false);

		return S_OK;
	}

	void ReleaseZxSpectrumSimulator()
	{
		wil::com_ptr_nothrow<IProfferService> srpProffer;
		auto hr = _sp->QueryService (SID_SProfferService, &srpProffer); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
		{
			if (_profferZXSimulatorServiceCookie)
			{
				hr = srpProffer->RevokeService (_profferZXSimulatorServiceCookie); LOG_IF_FAILED(hr);
				_profferZXSimulatorServiceCookie = VSCOOKIE_NIL;
			}
		}

		_simulator = nullptr;
	}

	// From https://stackoverflow.com/a/34691914/451036
	// Implemented because the Windows-provided CaptureStackBackTrace stops at the first frame with managed code.
	// It's generally more useful for us to look further than that, for example we can find out if our function
	// Z80DebugEngine::ResumeProcess was called - through many frames with managed code - from ScreenWindowImpl::Loadfile
	// or from Z80ProjectConfig::DebugLaunch.
	static DWORD CaptureStackBackTrace3 (DWORD FramesToSkip, DWORD nFrames, PVOID* BackTrace, PDWORD pBackTraceHash)
	{
		CONTEXT ContextRecord;
		RtlCaptureContext(&ContextRecord);

		DWORD iFrame;
		for (iFrame = 0; iFrame < nFrames; iFrame++)
		{
			DWORD64 ImageBase;
			PRUNTIME_FUNCTION pFunctionEntry = RtlLookupFunctionEntry(ContextRecord.Rip, &ImageBase, NULL);

			if (pFunctionEntry == NULL)
				break;

			PVOID HandlerData;
			DWORD64 EstablisherFrame;
			RtlVirtualUnwind(UNW_FLAG_NHANDLER,
				ImageBase,
				ContextRecord.Rip,
				pFunctionEntry,
				&ContextRecord,
				&HandlerData,
				&EstablisherFrame,
				NULL);

			BackTrace[iFrame] = (PVOID)ContextRecord.Rip;
		}

		return iFrame;
	}

	static void __stdcall TelemetryCallback (bool alreadyReported, const wil::FailureInfo& failure) noexcept
	{
		HRESULT hr;

		if (!_uiLibrary)
			return;

		com_ptr<IVsSettingsManager> sm;
		hr = serviceProvider->QueryService(SID_SVsSettingsManager, &sm);
		if (FAILED(hr))
			return;

		com_ptr<IVsWritableSettingsStore> wss;
		hr = sm->GetWritableSettingsStore(SettingsScope_UserSettings, &wss);
		if (FAILED(hr))
			return;

		BOOL report = FALSE;
		hr = wss->GetBool(SettingsCollection, AlwaysReportSettingsName, &report);
		if (hr != S_OK || !report)
		{
			wil::com_ptr_nothrow<IVsUIShell> uiShell;
			hr = serviceProvider->QueryService(SID_SVsUIShell, &uiShell);
			if (SUCCEEDED(hr))
			{
				HWND parent;
				hr = uiShell->GetDialogOwnerHwnd(&parent);
				if (SUCCEEDED(hr))
				{
					INT_PTR res = DialogBoxParamW (_uiLibrary, MAKEINTRESOURCE(IDD_REPORT_ERROR), parent, TelemetryDialogProc, (LPARAM)(void*)&failure);
					if (res == IDYES)
					{
						if (SUCCEEDED(wss->CreateCollection(SettingsCollection)))
							wss->SetBool(SettingsCollection, AlwaysReportSettingsName, TRUE);
					}

					if (res == IDYES || res == IDOK)
						report = TRUE;
				}
			}
		}

		if (report && !::IsDebuggerPresent())
		{
			sentry_value_t event = sentry_value_new_event();

			static const char* const FailureTypeNames[] = { "THROW", "RETURN", "LOG", "FAIL_FAST" };
			char hrstr[16];
			sprintf_s(hrstr, "0x%08X", failure.hr);
			auto exc = sentry_value_new_exception (FailureTypeNames[(int)failure.type], hrstr);

			auto returnValues = wil::make_unique_hlocal_nothrow<void*[]>(256);
			if (!returnValues)
				return;
			DWORD captured = CaptureStackBackTrace3 (0, 256, returnValues.get(), nullptr);
			// Skip frames corresponding to framework code.
			DWORD userCodeIndex = 0;
			while (userCodeIndex < captured && returnValues.get()[userCodeIndex] != failure.returnAddress)
				userCodeIndex++;
			if (userCodeIndex == captured)
				userCodeIndex = 0; // We couldn't find it, so let's fall back to reporting all the frames
			auto stacktrace = sentry_value_new_stacktrace(returnValues.get() + userCodeIndex, captured - userCodeIndex);
			sentry_value_set_by_key(exc, "stacktrace", stacktrace);

			sentry_event_add_exception(event, exc);

			wchar_t message[2048];
			hr = wil::GetFailureLogString (message, ARRAYSIZE(message), failure);
			int len = WideCharToMultiByte (CP_UTF8, 0, message, -1, nullptr, 0, nullptr, nullptr);
			if (len)
			{
				auto s = wil::make_hlocal_ansistring_nothrow(nullptr, len - 1);
				if (s)
				{
					WideCharToMultiByte (CP_UTF8, 0, message, -1, s.get(), len, nullptr, nullptr);
					sentry_value_set_by_key (event, "message", sentry_value_new_string(s.get()));
				}
			}

			sentry_capture_event(event);
		}
	}

	static INT_PTR CALLBACK TelemetryDialogProc (HWND hwndDlg, UINT message, WPARAM wParam, LPARAM lParam)
	{
		if (message == WM_INITDIALOG)
		{
			// Get the owner window and dialog box rectangles. 

			com_ptr<IVsUIShell> shell;
			auto hr = serviceProvider->QueryService(SID_SVsUIShell, &shell);
			if (SUCCEEDED(hr))
			{
				HWND hwndOwner = GetParent(hwndDlg);
				shell->CenterDialogOnWindow(hwndDlg, hwndOwner);
			}

			auto& failure = *(const wil::FailureInfo*)(void*)lParam;
			wchar_t message[2048];
			wil::GetFailureLogString (message, ARRAYSIZE(message), failure);

			HWND edit = GetDlgItem(hwndDlg, IDC_EDIT_ERROR_REPORT);
			SetWindowText (edit, message);

			/*
			RECT rcOwner, rcDlg, rc;
			GetWindowRect(hwndOwner, &rcOwner); 
			GetWindowRect(hwndDlg, &rcDlg); 
			CopyRect(&rc, &rcOwner); 

			// Offset the owner and dialog box rectangles so that right and bottom 
			// values represent the width and height, and then offset the owner again 
			// to discard space taken up by the dialog box. 

			OffsetRect(&rcDlg, -rcDlg.left, -rcDlg.top); 
			OffsetRect(&rc, -rc.left, -rc.top); 
			OffsetRect(&rc, -rcDlg.right, -rcDlg.bottom); 

			// The new position is the sum of half the remaining space and the owner's 
			// original position. 

			SetWindowPos(hwndDlg, 
			HWND_TOP, 
			rcOwner.left + (rc.right / 2), 
			rcOwner.top + (rc.bottom / 2), 
			0, 0,          // Ignores size arguments. 
			SWP_NOSIZE); 
			*/
			return TRUE;
		}

		if (message == WM_COMMAND)
		{
			UINT cmd = LOWORD(wParam);
			if ((cmd == IDOK) || (cmd == IDCANCEL) || (cmd == IDYES))
			{
				EndDialog(hwndDlg, cmd);
				return TRUE;
			}

			return FALSE;
		}

		return FALSE;
	}

	static void GetAppId (BSTR* pBstr)
	{
		MIDL_INTERFACE("1EAA526A-0898-11d3-B868-00C04F79F802") IVsAppId : IUnknown
		{
			virtual STDMETHODIMP Dummy0() = 0; // int SetSite(IOleServiceProvider pSP);
			virtual STDMETHODIMP GetProperty (int propid /* VSAPROPID */, VARIANT* com_ptr) = 0;
			virtual STDMETHODIMP Dummy1() = 0; // int SetProperty(int propid, object var);
			virtual STDMETHODIMP Dummy2() = 0; // int GetGuidProperty(int propid, out Guid guid);
			virtual STDMETHODIMP Dummy3() = 0; // int SetGuidProperty(int propid, ref Guid rguid);
			virtual STDMETHODIMP Dummy4() = 0; // int Initialize();
		};

		// The IVsAppId API is not documented anywhere. let's give it a try because it gives us the minor version too (i.e., the "2" in 17.11.2).
		com_ptr<IVsAppId> appId;
		auto hr = serviceProvider->QueryService(__uuidof(IVsAppId), &appId);
		if (SUCCEEDED(hr))
		{
			wil::unique_variant ver;
			hr = appId->GetProperty (-8642, &ver);
			if (SUCCEEDED(hr) && (ver.vt == VT_BSTR))
			{
				*pBstr = ver.release().bstrVal;
				return;
			}
		}

		// Fall back to something documented
		com_ptr<IVsShell> shell;
		hr = serviceProvider->QueryService(SID_SVsShell, &shell);
		if (SUCCEEDED(hr))
		{
			wil::unique_variant ver;
			hr = shell->GetProperty (VSSPROPID_ReleaseVersion, &ver);
			if (SUCCEEDED(hr) && (ver.vt == VT_BSTR))
			{
				*pBstr = ver.release().bstrVal;
				return;
			}
		}

		*pBstr = SysAllocString(L"VS Ver Unknown");
	}

	void InitSentry()
	{
		HRESULT hr;
		/*
		com_ptr<ILocalRegistry4> lr;
		hr = _sp->QueryService(SID_SLocalRegistry, &lr); RETURN_IF_FAILED(hr);

		VSLOCALREGISTRYROOTHANDLE vsRootHandle;
		wil::unique_bstr regRoot;
		hr = lr->GetLocalRegistryRootEx (RegType_Configuration, &vsRootHandle, &regRoot); RETURN_IF_FAILED(hr);

		HKEY rootKey = (HKEY)(ULONG_PTR)(LONG)vsRootHandle;
		uint32_t keyNameLen = SysStringLen(regRoot.get()) + 1 + sizeof(InstalledProductRegPath) / 2;
		auto keyName = wil::make_hlocal_string_nothrow(nullptr, keyNameLen);
		hr = StringCchCat(keyName.get(), keyNameLen + 1, regRoot.get()); RETURN_IF_FAILED(hr);
		hr = StringCchCat(keyName.get(), keyNameLen + 1, L"\\"); RETURN_IF_FAILED(hr);
		hr = StringCchCat(keyName.get(), keyNameLen + 1, InstalledProductRegPath); RETURN_IF_FAILED(hr);
		wil::unique_hkey key;
		auto lresult = RegOpenKeyEx(rootKey, keyName.get(), 0, KEY_READ, &key); RETURN_IF_WIN32_ERROR(lresult);
		wchar_t value[20];
		DWORD valueSize = sizeof(value);
		lresult = RegGetValue (key.get(), NULL, L"PID", RRF_RT_REG_SZ, nullptr, value, &valueSize); RETURN_IF_WIN32_ERROR(lresult);
		*/
		char dbpath[MAX_PATH];
		hr = SHGetFolderPathA (nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, dbpath); LOG_IF_FAILED(hr);
		if (FAILED(hr))
			return;

		#pragma warning (push)
		#pragma warning (disable: 4995)
		LPSTR combineRes = PathCombineA (dbpath, dbpath, "sentry-native"); LOG_HR_IF_NULL(E_FAIL, combineRes);
		if (!combineRes)
			return;
		#pragma warning (pop)

		wil::SetResultTelemetryFallback(&TelemetryCallback);
		_sentryOptions = sentry_options_new(); LOG_IF_NULL_ALLOC(_sentryOptions);
		if (!_sentryOptions)
			return;
		sentry_options_set_dsn(_sentryOptions, "https://042ccb2ce64bea0c9e97e5f515153f35@o4506847414714368.ingest.us.sentry.io/4506849315127296");
		sentry_options_set_database_path(_sentryOptions, dbpath);
		sentry_options_set_release(_sentryOptions, SentryReleaseName);
		//sentry_options_set_debug(_sentryOptions, 1);
		//sentry_options_set_symbolize_stacktraces(_sentryOptions, 1);

		wil::unique_bstr ver;
		GetAppId (&ver); 
		char buffer[100];
		int ires = WideCharToMultiByte(CP_UTF8, 0, ver.get(), -1, buffer, sizeof(buffer), 0, nullptr);
		if (ires > 0)
			sentry_options_set_environment(_sentryOptions, buffer);

		int res = sentry_init(_sentryOptions);
	}

	#pragma region IVsPackage
	virtual HRESULT STDMETHODCALLTYPE SetSite (IServiceProvider *pSP) override
	{
		HRESULT hr;
		WI_ASSERT(!_sp);
		WI_ASSERT(!serviceProvider);
		
		_sp = pSP;
		serviceProvider = pSP;

		InitSentry();

		com_ptr<IVsShell> shell;
		hr = serviceProvider->QueryService(SID_SVsShell, &shell);
		hr = shell->LoadUILibrary(CLSID_FelixPackage, 0, (DWORD_PTR*)&_uiLibrary); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IVsSolution> solutionService;
		hr = pSP->QueryService (SID_SVsSolution, &solutionService); RETURN_IF_FAILED(hr);
		hr = solutionService->AdviseSolutionEvents (this, &_solutionEventsCookie); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IVsRegisterProjectTypes> regSvc;
		hr = pSP->QueryService (SID_SVsRegisterProjectTypes, &regSvc); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IVsProjectFactory> pf;
		hr = Z80ProjectFactory_CreateInstance (_sp.get(), &pf); RETURN_IF_FAILED(hr);

		VSCOOKIE cookie;
		hr = regSvc->RegisterProjectType (__uuidof(IZ80ProjectProperties), pf.get(), &cookie); RETURN_IF_FAILED(hr);
		_projectTypeRegistrationCookie = cookie;

		hr = pSP->QueryService (SID_SVsShellDebugger, &_debugger); RETURN_IF_FAILED(hr);
		hr = _debugger->AdviseDebugEventCallback(static_cast<IDebugEventCallback2*>(this)); RETURN_IF_FAILED(hr);
		_advisedDebugEventCallback = true;

		wil::com_ptr_nothrow<IProfferService> srpProffer;
		hr = pSP->QueryService (SID_SProfferService, &srpProffer); RETURN_IF_FAILED(hr);
		hr = srpProffer->ProfferService (Z80AsmLanguageGuid, this, &_profferLanguageServiceCookie); RETURN_IF_FAILED(hr);

		hr = CreateZxSpectrumSimulator(); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryClose (BOOL* pfCanClose) override
	{
		*pfCanClose = TRUE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Close() override
	{
		HRESULT hr;

		ReleaseZxSpectrumSimulator();

		wil::com_ptr_nothrow<IProfferService> srpProffer;
		hr = _sp->QueryService (SID_SProfferService, &srpProffer); LOG_IF_FAILED(hr);
		if (srpProffer)
		{
			if (_profferLanguageServiceCookie)
			{
				hr = srpProffer->RevokeService (_profferLanguageServiceCookie); LOG_IF_FAILED(hr);
				_profferLanguageServiceCookie = VSCOOKIE_NIL;
			}
		}

		if (_advisedDebugEventCallback)
		{
			hr = _debugger->UnadviseDebugEventCallback(static_cast<IDebugEventCallback2*>(this)); LOG_IF_FAILED(hr);
			_advisedDebugEventCallback = false;
		}
		_debugger = nullptr;

		if (_simulatorWindowFrame)
		{
			_simulatorWindowFrame->Hide();
			_simulatorWindowFrame = nullptr;
		}

		if (_projectTypeRegistrationCookie != VSCOOKIE_NIL)
		{
			wil::com_ptr_nothrow<IVsRegisterProjectTypes> regSvc;
			hr = _sp->QueryService (SID_SVsRegisterProjectTypes, &regSvc); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
			{
				regSvc->UnregisterProjectType(_projectTypeRegistrationCookie);
				_projectTypeRegistrationCookie = VSCOOKIE_NIL;
			}
		}

		if (_solutionEventsCookie != VSCOOKIE_NIL)
		{
			wil::com_ptr_nothrow<IVsSolution> solutionService;
			hr = _sp->QueryService (SID_SVsSolution, &solutionService); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
			{
				hr = solutionService->UnadviseSolutionEvents(_solutionEventsCookie); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr))
					_solutionEventsCookie = VSCOOKIE_NIL;
			}
		}

		::FreeLibrary(_uiLibrary);

		if (_sentryOptions)
		{
			sentry_close();
			sentry_options_free(_sentryOptions);
		}

		wil::SetResultTelemetryFallback(nullptr);

		_sp = nullptr;

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetAutomationObject (LPCOLESTR pszPropName, IDispatch **ppDisp) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE CreateTool (REFGUID rguidPersistenceSlot) override
	{
		if (rguidPersistenceSlot == CLSID_FelixPersistenceSlot)
			return CreateSimulatorToolWindow(_sp.get(), &_simulatorWindowFrame);

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE ResetDefaults (VSPKGRESETFLAGS grfFlags) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetPropertyPage (REFGUID rguidPage, VSPROPSHEETPAGE *ppage) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	HRESULT OpenFileOnProjectCreation(IVsHierarchy* pHierarchy)
	{
		wil::unique_variant first;
		auto hr = pHierarchy->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &first); RETURN_IF_FAILED(hr);

		com_ptr<IVsProject2> p2;
		hr = pHierarchy->QueryInterface(&p2); RETURN_IF_FAILED(hr);

		com_ptr<IVsWindowFrame> frame;
		hr = p2->OpenItem (V_VSITEMID(first.addressof()), LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &frame); RETURN_IF_FAILED(hr);

		hr = frame->Show(); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	#pragma region IVsSolutionEvents
	virtual HRESULT STDMETHODCALLTYPE OnAfterOpenProject(IVsHierarchy* pHierarchy, BOOL fAdded) override
	{
		com_ptr<IZ80ProjectProperties> proj;
		if (SUCCEEDED(pHierarchy->QueryInterface(&proj)))
		{
			if (!_simulatorWindowFrame)
			{
				auto hr = CreateSimulatorToolWindow (_sp.get(), &_simulatorWindowFrame); LOG_IF_FAILED(hr);
			}

			if (_simulatorWindowFrame)
			{
				auto hr = _simulatorWindowFrame->ShowNoActivate(); LOG_IF_FAILED(hr);
			}

			if (fAdded)
				OpenFileOnProjectCreation(pHierarchy);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnQueryCloseProject(IVsHierarchy* pHierarchy, BOOL fRemoving, BOOL* pfCancel) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnBeforeCloseProject(IVsHierarchy* pHierarchy, BOOL fRemoved) override
	{
		if (_simulatorWindowFrame)
		{
			auto hr = _simulatorWindowFrame->Hide(); LOG_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnAfterLoadProject(IVsHierarchy* pStubHierarchy, IVsHierarchy* pRealHierarchy) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnQueryUnloadProject(IVsHierarchy* pRealHierarchy, BOOL* pfCancel) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnBeforeUnloadProject(IVsHierarchy* pRealHierarchy, IVsHierarchy* pStubHierarchy) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnAfterOpenSolution(IUnknown* pUnkReserved, BOOL fNewSolution) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnQueryCloseSolution(IUnknown* pUnkReserved, BOOL* pfCancel) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnBeforeCloseSolution(IUnknown* pUnkReserved) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnAfterCloseSolution(IUnknown* pUnkReserved) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IOleCommandTarget
	virtual HRESULT STDMETHODCALLTYPE QueryStatus (const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText) override
	{
		if (*pguidCmdGroup == CLSID_FelixPackageCmdSet)
		{
			for (ULONG i = 0; i < cCmds; i++)
			{
				if (prgCmds[i].cmdID == cmdidSimulator)
					prgCmds[i].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
				else
					prgCmds[i].cmdf = 0;
			}

			return S_OK;
		}

		return OLECMDERR_E_UNKNOWNGROUP;
	}

	virtual HRESULT STDMETHODCALLTYPE Exec (const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
	{
		if (*pguidCmdGroup == CLSID_FelixPackageCmdSet)
		{
			if (nCmdID == cmdidSimulator)
			{
				if (!_simulatorWindowFrame)
				{
					auto hr = CreateSimulatorToolWindow(_sp.get(), &_simulatorWindowFrame); RETURN_IF_FAILED(hr);
				}

				if (_simulatorWindowFrame)
				{
					auto hr = _simulatorWindowFrame->Show(); RETURN_IF_FAILED(hr);
				}

				return S_OK;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		return OLECMDERR_E_UNKNOWNGROUP;
	}
	#pragma endregion

	static HRESULT CreateSimulatorToolWindow (IServiceProvider* sp, IVsWindowFrame** ppFrame)
	{
		HRESULT hr;

		wil::com_ptr_nothrow<IVsUIShell> shell;
		hr = sp->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IVsWindowPane> pane;
		hr = SimulatorWindowPane_CreateInstance (&pane); RETURN_IF_FAILED(hr);

		hr = shell->CreateToolWindow (CTW_fForceCreate | CTW_fActivateWithProject | CTW_fToolbarHost, 0, pane.get(), GUID_NULL,
			CLSID_FelixPersistenceSlot, GUID_NULL, nullptr, L"ZX Simulator", nullptr, ppFrame); RETURN_IF_FAILED(hr);

		VARIANT srpvt;
		srpvt.vt = VT_I4;
		srpvt.intVal = IDB_IMAGES;
		hr = (*ppFrame)->SetProperty(VSFPROPID_BitmapResource, srpvt); RETURN_IF_FAILED(hr);
		srpvt.intVal = 1;
		hr = (*ppFrame)->SetProperty(VSFPROPID_BitmapIndex, srpvt); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IDebugEventCallback2
	virtual HRESULT STDMETHODCALLTYPE Event (IDebugEngine2 *pEngine, IDebugProcess2 *pProcess, IDebugProgram2 *pProgram, IDebugThread2 *pThread, IDebugEvent2 *pEvent, REFIID riidEvent, DWORD dwAttrib) override
	{
		DWORD attrs;
		if (SUCCEEDED(pEvent->GetAttributes(&attrs)) && (attrs & EVENT_STOPPING))
		{
			// When starting debugging, VS switches to the Debug window layout. This window layout does not
			// include the Simulator tool window on a newly created project. This is bad UX for a new user.
			// We want to force show it every time debugging starts, _after_ VS has switched to the Debug window layout.
			// This last bit is tricky; the earliest point in time I could find when VS is already in the Debug window layout
			// is the handler of an EVENT_SYNC_STOP debug event. This may or may not work in future versions of VS;
			// we'll see about that.
			if (!_simulatorWindowFrame)
			{
				auto hr = CreateSimulatorToolWindow(_sp.get(), &_simulatorWindowFrame); RETURN_IF_FAILED(hr);
			}

			if (_simulatorWindowFrame)
			{
				auto hr = _simulatorWindowFrame->Show(); RETURN_IF_FAILED(hr);
			}
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IServiceProvider
	virtual HRESULT STDMETHODCALLTYPE QueryService (REFGUID guidService, REFIID riid, void** ppvObject) override
	{
		if (guidService == Z80AsmLanguageGuid)
			return _z80AsmLanguageInfo->QueryInterface(riid, ppvObject);

		if (guidService == SID_Simulator)
			return _simulator->QueryInterface(riid, ppvObject);

		RETURN_HR(E_NOINTERFACE);
	}
	#pragma endregion
};

HRESULT FelixPackage_CreateInstance (IVsPackage** out)
{
	wil::com_ptr_nothrow<FelixPackageImpl> p = new (std::nothrow) FelixPackageImpl(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*out = p.detach();
	return S_OK;
}

HRESULT GetDefaultProjectFileExtension (BSTR* ppExt)
{
	static wil::unique_hlocal_string cached;
	if (!cached)
	{
		com_ptr<ILocalRegistry4> lr;
		auto hr = serviceProvider->QueryService(SID_SLocalRegistry, &lr); RETURN_IF_FAILED(hr);

		VSLOCALREGISTRYROOTHANDLE vsRootHandle;
		wil::unique_bstr regRoot;
		hr = lr->GetLocalRegistryRootEx (RegType_Configuration, &vsRootHandle, &regRoot); RETURN_IF_FAILED(hr);
		HKEY rootKey = (HKEY)(ULONG_PTR)(LONG)vsRootHandle;

		static const wchar_t Projects[] = L"\\Projects\\";
		static const size_t GuidStrSize = 38;
		uint32_t keyNameLen = SysStringLen(regRoot.get()) + (sizeof(Projects) / 2 - 1) + GuidStrSize;
		auto keyName = wil::make_hlocal_string_nothrow(nullptr, keyNameLen);
		hr = StringCchCat(keyName.get(), keyNameLen + 1, regRoot.get()); RETURN_IF_FAILED(hr);
		STRSAFE_LPWSTR end;
		hr = StringCchCatEx(keyName.get(), keyNameLen + 1, Projects, &end, nullptr, 0); RETURN_IF_FAILED(hr);
		StringFromGUID2 (IID_IZ80ProjectProperties, end, GuidStrSize + 1);
		wil::unique_hkey key;
		auto lresult = RegOpenKeyEx(rootKey, keyName.get(), 0, KEY_READ, &key); RETURN_IF_WIN32_ERROR(lresult);

		wchar_t value[20];
		DWORD valueSize = sizeof(value);
		lresult = RegGetValue (key.get(), NULL, L"DefaultProjectExtension", RRF_RT_REG_SZ, nullptr, value, &valueSize); RETURN_IF_WIN32_ERROR(lresult);

		cached = wil::make_hlocal_string_nothrow(value, valueSize / 2 - 1); RETURN_IF_NULL_ALLOC(cached);
	}

	return (*ppExt = SysAllocString(cached.get())) ? S_OK : E_OUTOFMEMORY;
}

HRESULT SetErrorInfo1 (HRESULT errorHR, ULONG packageStringResId, LPCWSTR arg1)
{
	wil::com_ptr_nothrow<IVsShell> shell;
	auto hr = serviceProvider->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);

	wil::unique_bstr message;
	hr = shell->LoadPackageString(CLSID_FelixPackage, packageStringResId, &message); RETURN_IF_FAILED(hr);

	wil::unique_hlocal_string buffer;
	for (size_t sz = 100; ; sz *= 2)
	{
		buffer = wil::make_hlocal_string_nothrow(nullptr, sz); RETURN_IF_NULL_ALLOC(buffer);
		int ires = swprintf_s (buffer.get(), sz, message.get(), arg1); RETURN_HR_IF(E_FAIL, ires < 0);
		if (ires < sz)
			break;
	}

	com_ptr<IVsUIShell> uiShell;
	hr = serviceProvider->QueryService (SID_SVsUIShell, &uiShell); RETURN_IF_FAILED(hr);

	hr = uiShell->SetErrorInfo (errorHR, buffer.get(), 0, nullptr, nullptr); RETURN_IF_FAILED(hr);

	return S_OK;
}
