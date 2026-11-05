
#include "pch.h"
#include "DebugEngine.h"
#include "DebugEventBase.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
#include "../guids.h"
#include "../FelixPackage.h"
#include "../Z80Xml.h"
#include "../FelixPackageUi/resource.h"

class Z80DebugEngine : public IDebugEngine2, IDebugEngineLaunch2, ISimulatorEventNotifySink, IFelixLaunchOptionsProvider
	, ITapPlayNotifySink
{
	ULONG _refCount = 0;
	WeakRefToThis _weakRefToThis;
	wil::unique_bstr _registryRoot;
	WORD langId = 0;
	com_ptr<IDebugEventCallback2> _callback;
	com_ptr<IFelixProgram> _program;
	com_ptr<IFelixLaunchOptions> _launchOptions;
	SIM_BP_COOKIE _entryPointBreakpoint = 0;
	SIM_BP_COOKIE _exitPointBreakpoint = 0;
	AdviseSinkToken _simulatorEventsToken;
	AdviseSinkToken _tapPlaySinkToken;
	wil::unique_process_heap_string _launchExe;
	SpectrumVariant _platform = SpectrumVariantCount;

public:
	HRESULT InitInstance()
	{
		auto hr = _weakRefToThis.InitInstance(static_cast<IDebugEngine2*>(this)); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = NULL;

		if (   TryQI<IUnknown>(static_cast<IDebugEngine2*>(this), riid, ppvObject)
			|| TryQI<IDebugEngine2>(this, riid, ppvObject)
			|| TryQI<IDebugEngineLaunch2>(this, riid, ppvObject)
			|| TryQI<ISimulatorEventNotifySink>(this, riid, ppvObject)
			|| TryQI<IFelixLaunchOptionsProvider>(this, riid, ppvObject)
			|| TryQI<ITapPlayNotifySink>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

		#ifdef _DEBUG
		// Stuff which we'll never implement.
		if (   riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IStdMarshalInfo
			|| riid == IID_IdentityUnmarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IFastRundown
			|| riid == IID_IExternalConnection
			|| riid == IID_INoIdea6_DebugEngine
			|| riid == IID_INoIdea7
			|| riid == IID_INotARealInterface
			|| riid == IID_INoIdea9
			|| riid == IID_INoIdea10
			|| riid == IID_INoIdea11
			|| riid == IID_INoIdea14
			|| riid == IID_INoIdea15
		)
			return E_NOINTERFACE;

		// Stuff which we may implement
		if (   riid == IID_IDebugEngine110
			//|| riid == IID_IDebugEngine150
			|| riid == IID_IDebugEngine3
			|| riid == IID_IDebugEngineLaunch100
			|| riid == IID_IDebugProgramProvider2
			|| riid == IID_IDebugBreakpointFileUpdateNotification110
			|| riid == IID_IDebugEngineStepFilterManager90
			|| riid == IID_IDebugSymbolSettings100
			|| riid == IID_IDebugSymbolSettings170
			//|| riid == IID_IDebugVisualizerExtensionReceiver178
		)
			return E_NOINTERFACE;

		for (auto& i : HardcodedRundownIFsOfInterest)
			if (*i == riid)
				return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugEngine2
	virtual HRESULT __stdcall EnumPrograms(IEnumDebugPrograms2** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	// https://learn.microsoft.com/en-us/visualstudio/extensibility/debugger/attaching-to-the-program?view=visualstudio
	virtual HRESULT __stdcall Attach(IDebugProgram2** rgpPrograms, IDebugProgramNode2** rgpProgramNodes, DWORD celtPrograms, IDebugEventCallback2* pCallback, ATTACH_REASON dwReason) override
	{
		HRESULT hr;
		RETURN_HR_IF_NULL(E_POINTER, rgpPrograms);
		RETURN_HR_IF_NULL(E_POINTER, pCallback);
		RETURN_HR_IF(E_INVALIDARG, celtPrograms != 1 || !rgpPrograms[0]);
		RETURN_HR_IF(E_UNEXPECTED, _program);
		RETURN_HR_IF(E_UNEXPECTED, _callback);

		// If VS calls us from the Debug -> Attach to Process dialog, it's possible the package hasn't been loaded yet.
		// Happens when the Simulator window is visible in neither the Design layout nor in the Debug layout.
		if (!simulator)
		{
			com_ptr<IServiceProvider> sp;
			hr = GetGlobalServiceProvider(&sp); RETURN_IF_FAILED(hr);
			com_ptr<IVsShell> shell;
			hr = sp->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);
			com_ptr<IVsPackage> package;
			static const GUID felix = { 0x768BC57B, 0x42A8, 0x42AB, { 0xB3, 0x89, 0x45, 0x79, 0x46, 0xC4, 0xFC, 0x6A } };
			hr = shell->LoadPackage(felix, &package); RETURN_IF_FAILED(hr);
			RETURN_HR_IF(E_UNEXPECTED, !simulator);
		}

		// From https://learn.microsoft.com/en-us/visualstudio/extensibility/debugger/reference/idebugengine2-attach:
		// When this method is called, the DE needs to send these events in sequence:
		// 1. IDebugEngineCreateEvent2 (if it has not already been sent for a particular instance of the debug engine)
		// 2. IDebugProgramCreateEvent2
		// 3. IDebugLoadCompleteEvent2

		// From https://learn.microsoft.com/en-us/visualstudio/extensibility/debugger/attaching-and-detaching-to-a-program?view=visualstudio:
		// 2. The DE sends an IDebugEngineCreateEvent2 to the SDM with an EVENT_SYNC attribute.
		// 3. The DE sends an IDebugProgramCreateEvent2 to the SDM with an EVENT_SYNC attribute.
		// 4. The DE sends an IDebugLoadCompleteEvent2 to the SDM with an EVENT_SYNC_STOP attribute.

		com_ptr<IDebugProcess2> process;
		hr = rgpPrograms[0]->GetProcess(&process); RETURN_IF_FAILED(hr);

		_callback = pCallback;
		hr = rgpPrograms[0]->QueryInterface(&_program); RETURN_IF_FAILED(hr);

		HRESULT hrEngineAttach;
		hr = process->Attach(pCallback, const_cast<GUID*>(&Engine_Id), 1, &hrEngineAttach); RETURN_IF_FAILED(hr);
		RETURN_IF_FAILED(hrEngineAttach);

		auto resetOnFailure = wil::scope_exit([this,&process]
			{
				_simulatorEventsToken.reset();
				process->Detach();
				_program = nullptr;
				_callback = nullptr;
			});

		hr = AdviseSink<ISimulatorEventNotifySink>(simulator, _weakRefToThis, &_simulatorEventsToken); RETURN_IF_FAILED(hr);

		struct EngineCreateEvent : public EventBase<IDebugEngineCreateEvent2, EVENT_ASYNCHRONOUS>
		{
			com_ptr<IDebugEngine2> _engine;

			virtual HRESULT STDMETHODCALLTYPE GetEngine (IDebugEngine2 **pEngine) override
			{
				return copy_to(_engine, pEngine);
			}
		};

		auto ece = com_ptr (new (std::nothrow) EngineCreateEvent()); RETURN_IF_NULL_ALLOC(ece);
		ece->_engine = this;
		hr = ece->Send(pCallback, this, nullptr, nullptr); RETURN_IF_FAILED(hr);

		using ProgramCreateEvent = EventBase<IDebugProgramCreateEvent2, EVENT_ASYNCHRONOUS>;
		wil::com_ptr_nothrow<ProgramCreateEvent> pce = new (std::nothrow) ProgramCreateEvent(); RETURN_IF_NULL_ALLOC(pce);
		hr = pce->Send(pCallback, this, _program, nullptr); RETURN_IF_FAILED(hr);

		using ThreadCreateEvent = EventBase<IDebugThreadCreateEvent2, EVENT_ASYNCHRONOUS>;
		wil::com_ptr_nothrow<ThreadCreateEvent> tce = new (std::nothrow) ThreadCreateEvent(); RETURN_IF_NULL_ALLOC(tce);
		com_ptr<IEnumDebugThreads2> threads;
		hr = _program->EnumThreads(&threads); RETURN_IF_FAILED(hr);
		com_ptr<IDebugThread2> thread;
		hr = threads->Next(1, &thread, nullptr); RETURN_IF_FAILED(hr);
		hr = tce->Send(pCallback, this, _program, thread); RETURN_IF_FAILED(hr);

		hr = _program->EngineAttach(this, pCallback, dwReason); RETURN_IF_FAILED(hr);
		auto terminateOnFailure = wil::scope_exit([this] { _program->Terminate(); });
		if (dwReason == ATTACH_REASON_LAUNCH)
		{
			// ATTACH_REASON_LAUNCH indicates that the DE is attaching to the program because
			// the user launched the process that contains it.
			hr = LoadProgram(_launchExe.get()); RETURN_IF_FAILED(hr);
		}
		else if (dwReason == ATTACH_REASON_USER)
		{
			// ATTACH_REASON_USER indicates that the user has explicitly requested the DE
			// to attach to a program (or the process that contains a program).
			hr = SendLoadCompleteEvent(pCallback, this, _program, thread); RETURN_IF_FAILED(hr);
		}
		else
			RETURN_HR(E_NOTIMPL);

		terminateOnFailure.release();
		resetOnFailure.release();
		return S_OK;
	}

	virtual HRESULT __stdcall CreatePendingBreakpoint(IDebugBreakpointRequest2* pBPRequest, IDebugPendingBreakpoint2** ppPendingBP) override
	{
		BP_LOCATION_TYPE locationType;
		auto hr = pBPRequest->GetLocationType(&locationType); RETURN_IF_FAILED(hr);

		if ((locationType & BPLT_TYPE_MASK) == BPT_CODE)
		{
			switch (locationType & BPLT_LOCATION_TYPE_MASK)
			{
				case BPLT_CONTEXT:
				{
					BP_REQUEST_INFO info = { };
					hr = pBPRequest->GetRequestInfo (BPREQI_BPLOCATION, &info); RETURN_IF_FAILED(hr);
					RETURN_HR_IF(E_FAIL, info.bpLocation.bpLocationType != locationType);
					auto cc = com_ptr(info.bpLocation.bpLocation.bplocCodeContext.pCodeContext);
					info.bpLocation.bpLocation.bplocCodeContext.pCodeContext->Release();
					info.bpLocation.bpLocation.bplocCodeContext.pCodeContext = nullptr;
					com_ptr<IFelixCodeContext> z80cc;
					hr = cc->QueryInterface(&z80cc); RETURN_IF_FAILED(hr);
					RETURN_HR_IF(E_FAIL, z80cc->PhysicalMemorySpace());
					hr = MakeSimplePendingBreakpoint (_callback, this, _program, false, z80cc->Address(), ppPendingBP); RETURN_IF_FAILED(hr);
					return S_OK;
				}

				case BPLT_ADDRESS:
				{
					BP_REQUEST_INFO info = { };
					hr = pBPRequest->GetRequestInfo (BPREQI_BPLOCATION, &info); RETURN_IF_FAILED(hr);
					RETURN_HR_IF(E_FAIL, info.bpLocation.bpLocationType != locationType);

					uint64_t address;
					bool converted = false;
					auto str = info.bpLocation.bpLocation.bplocCodeAddress.bstrAddress;
					size_t len = wcslen(str);
					if ((str[0] >= '0' && str[0] <= '9')
						|| ((len > 1) && ((str[len - 1] & 0xDF) == 'H') && ((str[0] & 0xDF) >= 'A') && ((str[0] & 0xDF) <= 'F')))
					{
						bool hex = (len > 1) && ((str[len - 1] & 0xDF) == 'H');
						wchar_t* endPtr;
						address = wcstoull (str, &endPtr, hex ? 16 : 10);
						converted = (endPtr == &str[len - (hex ? 1 : 0)]);
					}

					auto& ca = info.bpLocation.bpLocation.bplocCodeAddress;
					if (ca.bstrContext)
						SysFreeString(ca.bstrContext);
					if (ca.bstrModuleUrl)
						SysFreeString(ca.bstrModuleUrl);
					if (ca.bstrFunction)
						SysFreeString(ca.bstrFunction);
					if (ca.bstrAddress)
						SysFreeString(ca.bstrAddress);

					RETURN_HR_IF(E_FAIL, !converted);

					hr = MakeSimplePendingBreakpoint (_callback, this, _program, false, address, ppPendingBP); RETURN_IF_FAILED(hr);
					return S_OK;
				}

				case BPLT_FILE_LINE:
				{
					BP_REQUEST_INFO info;
					hr = pBPRequest->GetRequestInfo (BPREQI_BPLOCATION, &info); RETURN_IF_FAILED(hr);
					if (info.bpLocation.bpLocationType != locationType)
						RETURN_HR(E_FAIL);

					wil::com_ptr_nothrow<IDebugDocumentPosition2> doc_pos;
					doc_pos.attach(info.bpLocation.bpLocation.bplocCodeFileLine.pDocPos);
					info.bpLocation.bpLocation.bplocCodeFileLine.pDocPos = nullptr;

					if (info.bpLocation.bpLocation.bplocCodeFileLine.bstrContext)
					{
						SysFreeString(info.bpLocation.bpLocation.bplocCodeFileLine.bstrContext);
						info.bpLocation.bpLocation.bplocCodeFileLine.bstrContext = nullptr;
					}

					wil::unique_bstr file_path;
					hr = doc_pos->GetFileName(&file_path); RETURN_IF_FAILED(hr);

					TEXT_POSITION begin, end;
					hr = doc_pos->GetRange(&begin, &end); RETURN_IF_FAILED(hr);

					hr = MakeSourceLinePendingBreakpoint (_callback, this, _program, pBPRequest, file_path.get(), begin.dwLine, ppPendingBP); RETURN_IF_FAILED(hr);
					return S_OK;
				}

				default:
					RETURN_HR(E_NOTIMPL);
			}
		}
		else if ((locationType & BPLT_TYPE_MASK) == BPT_DATA)
		{
			// We don't support data breakpoints for now.
			RETURN_HR(E_NOTIMPL);
		}
		else
		{
			// unknown type of breakpoint (neither code nor data)
			RETURN_HR(E_NOTIMPL);
		}
	}

	virtual HRESULT __stdcall SetException(EXCEPTION_INFO* pException) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT __stdcall RemoveSetException(EXCEPTION_INFO* pException) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall RemoveAllSetExceptions(const GUID& guidType) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT __stdcall GetEngineId(GUID* pguidEngine) override
	{
		*pguidEngine = Engine_Id;
		return S_OK;
	}

	// Informs a debug engine (DE) that the program specified has been atypically terminated and
	// that the DE should clean up all references to the program and send a program destroy event.
	virtual HRESULT __stdcall DestroyProgram(IDebugProgram2* pProgram) override
	{
		// called for example when we return an error code from ResumeProcess()
		WI_ASSERT(pProgram == _program.get());
		return TerminateProcess(nullptr);
	}

	virtual HRESULT __stdcall ContinueFromSynchronousEvent(IDebugEvent2* pEvent) override
	{
		HRESULT hr;

		wil::com_ptr_nothrow<IDebugProgramDestroyEvent2> pde;
		if (SUCCEEDED(pEvent->QueryInterface(&pde)))
		{
			com_ptr<IDebugProcess2> process;
			hr = _program->GetProcess(&process); RETURN_IF_FAILED(hr);
			hr = process->Detach(); RETURN_IF_FAILED(hr);

			_ASSERT(simulator->HasBreakpoints_HR() == S_FALSE);
			if (simulator->Running_HR() == S_FALSE)
			{
				hr = simulator->Resume(false); LOG_IF_FAILED(hr);
			}

			_callback = nullptr;
			_program = nullptr;
			_launchOptions = nullptr;

			return S_OK;
		}

		wil::com_ptr_nothrow<IDebugModuleLoadEvent2> mle;
		if (SUCCEEDED(pEvent->QueryInterface(&mle)))
			return S_OK;

		wil::com_ptr_nothrow<IDebugThreadDestroyEvent2> tde;
		if (SUCCEEDED(pEvent->QueryInterface(&tde)))
			return S_OK;

		com_ptr<IDebugLoadCompleteEvent2> lce;
		if (SUCCEEDED(pEvent->QueryInterface(lce.addressof())))
			return S_OK;

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall SetLocale(WORD wLangID) override
	{
		langId = wLangID;
		return S_OK;
	}

	virtual HRESULT __stdcall SetRegistryRoot(LPCOLESTR pszRegistryRoot) override
	{
		_registryRoot.reset (SysAllocString(pszRegistryRoot)); RETURN_IF_NULL_ALLOC(_registryRoot);
		return S_OK;
	}

	virtual HRESULT __stdcall SetMetric(LPCOLESTR pszMetric, VARIANT varValue) override
	{
		/*
		if (IsDebuggerPresent())
		{
			OutputDebugStringW (L"DE: SetMetric (\"");
			OutputDebugStringW (pszMetric);
			OutputDebugStringW (L"\" = ");
			VARIANT str;
			VariantInit (&str);
			VariantChangeType (&str, &varValue, 0, VT_BSTR);
			OutputDebugStringW (str.bstrVal);
			VariantClear(&str);
			OutputDebugStringW (L")\r\n");
		}
		*/
		return E_NOTIMPL;
	}

	virtual HRESULT __stdcall CauseBreak() override
	{
		if (_program)
			return _program->CauseBreak();
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	// From https://learn.microsoft.com/en-us/visualstudio/extensibility/debugger/sending-the-required-events?view=visualstudio:
	// "This event must be a stopping event."
	//
	// From https://learn.microsoft.com/en-us/visualstudio/extensibility/debugger/control-of-execution?view=visualstudio:
	// "The debug session typically responds to this first stopping event by calling the program's Continue method."
	struct FelixLoadCompleteEvent : IDebugEvent2, IDebugLoadCompleteEvent2, IFelixLoadCompleteEvent
	{
		ULONG _refCount = 0;

		#pragma region IUnknown
		virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(static_cast<IDebugEvent2*>(this), riid, ppvObject)
				|| TryQI<IDebugEvent2>(this, riid, ppvObject)
				|| TryQI<IDebugLoadCompleteEvent2>(this, riid, ppvObject)
				|| TryQI<IFelixLoadCompleteEvent>(this, riid, ppvObject))
				return S_OK;

			return E_NOINTERFACE;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		// IDebugEvent2
		virtual HRESULT __stdcall GetAttributes(DWORD* pdwAttrib) override
		{
			*pdwAttrib = EVENT_SYNC_STOP;
			return S_OK;
		}
	};
	static HRESULT SendLoadCompleteEvent (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, IDebugThread2* thread)
	{
		// Must be stopping as per https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/supported-event-types?view=vs-2019
		auto e = com_ptr(new (std::nothrow) FelixLoadCompleteEvent()); RETURN_IF_NULL_ALLOC(e);
		auto hr = callback->Event (engine, nullptr, program, thread, e, IID_IDebugLoadCompleteEvent2, EVENT_SYNC_STOP); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	static HRESULT SendEntryPointEvent (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, IDebugThread2* thread)
	{
		// Must be stopping as per https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/supported-event-types?view=vs-2019
		using EntryPointEvent = EventBase<IDebugEntryPointEvent2, EVENT_SYNC_STOP>;
		wil::com_ptr_nothrow<EntryPointEvent> e = new (std::nothrow) EntryPointEvent(); RETURN_IF_NULL_ALLOC(e);
		auto hr = e->Send (callback, engine, program, thread); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IDebugEngineLaunch2
	virtual HRESULT __stdcall LaunchSuspended (LPCOLESTR pszServer, IDebugPort2* pPort,
											   LPCOLESTR pszExe, LPCOLESTR pszArgs, LPCOLESTR pszDir,
											   BSTR bstrEnv, LPCOLESTR pszOptions, LAUNCH_FLAGS dwLaunchFlags,
											   DWORD hStdInput, DWORD hStdOutput, DWORD hStdError,
											   IDebugEventCallback2* pCallback, IDebugProcess2** ppProcess) override
	{
		HRESULT hr;

		WI_ASSERT(!_callback);
		WI_ASSERT(!_program);
		WI_ASSERT(!_launchOptions);

		_launchExe = wil::make_process_heap_string_nothrow(pszExe); RETURN_IF_NULL_ALLOC(_launchExe);

		_platform = SpectrumVariant48K;
		if (pszOptions)
		{
			auto stream = com_ptr(SHCreateMemStream((BYTE*)pszOptions, 2 * (UINT)wcslen(pszOptions))); RETURN_IF_NULL_ALLOC(stream);
			hr = MakeLaunchOptions(&_launchOptions); RETURN_IF_FAILED(hr);
			hr = LoadFromXml (_launchOptions.get(), nullptr, stream.get(), 1200); RETURN_IF_FAILED(hr);
			hr = _launchOptions->get_Platform(&_platform); RETURN_IF_FAILED(hr);
		}

		hr = simulator->Break(); RETURN_IF_FAILED(hr);
		hr = simulator->Reset(0, _platform); RETURN_IF_FAILED(hr);

		com_ptr<IDebugProcess2> process;
		pPort->GetProcess(AD_PROCESS_ID{ AD_PROCESS_ID_SYSTEM, SimulatorProcessId }, &process); RETURN_IF_FAILED(hr);

		return copy_to(process, ppProcess);
	}

	// https://docs.microsoft.com/en-us/previous-versions/bb145570%28v%3dvs.90%29

	virtual HRESULT __stdcall ResumeProcess(IDebugProcess2* pProcess) override
	{
		HRESULT hr;

		com_ptr<IFelixDebugProcess> felixProcess;
		hr = pProcess->QueryInterface(&felixProcess); RETURN_IF_FAILED(hr);
		hr = felixProcess->InitDebugProgram(_platform); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	HRESULT LoadProgram (LPCWSTR exePath)
	{
		HRESULT hr;
		com_ptr<IFelixProgram> mcoll;
		hr = _program->QueryInterface(&mcoll); RETURN_IF_FAILED(hr);

		wil::unique_bstr sourceBasePath;
		if (!_launchOptions || FAILED(_launchOptions->get_ProjectDir(&sourceBasePath)))
		{
			sourceBasePath = wil::make_bstr_nothrow(exePath); RETURN_IF_NULL_ALLOC(sourceBasePath);
			PathRemoveFileSpec(sourceBasePath.get());
		}

		auto fileExt = PathFindExtension(exePath);
		if (!_wcsicmp(fileExt, L".bin") || !_wcsicmp(fileExt, L".tap"))
		{
			com_ptr<IDebugModule2> romModule;
			hr = GetModuleAtAddress(_program, 0, &romModule); RETURN_IF_FAILED(hr);
			com_ptr<IFelixSymbols> romSymbols;
			hr = romModule.try_query<IZ80Module>()->GetSymbols(&romSymbols); RETURN_IF_FAILED(hr);
			hr = SimulateToEDITOR(romSymbols); RETURN_IF_FAILED(hr);

			auto terminateOnError = wil::scope_exit([this] { TerminateInternal(); });

			if (!_wcsicmp(fileExt, L".bin"))
			{
				hr = simulator->SetSpeed(100); RETURN_IF_FAILED(hr);
				hr = LoadBinary(romSymbols, exePath, sourceBasePath.get()); RETURN_IF_FAILED_EXPECTED(hr);
			}
			else //if (!_wcsicmp(fileExt, L".tap"))
			{
				hr = LoadTap(romSymbols, exePath, sourceBasePath.get()); RETURN_IF_FAILED_EXPECTED(hr);
			}

			terminateOnError.release();
		}
		else if (!_wcsicmp(fileExt, L".sna"))
		{
			hr = simulator->LoadFile(exePath);
			if (FAILED(hr))
			{
				// The LoadFile function generates rich error info via SetErrorInfo.
				// However, we are here deep in the call stack, and that error info will be lost
				// because some of our callers will call ::GetErrorInfo before returning control
				// to the GUI code. The GUI will eventually show a generic E_FAIL.
				// So let's show that detailed error info here; the user will see two error messages,
				// but that's not as bad as seeing a single one without detailed information.
				uiShell->ReportErrorInfo(hr);
				return hr;
			}

			com_ptr<IDebugModule2> ram_module;
			hr = MakeModule (0x4000, 0xC000, exePath, true, sourceBasePath.get(), &ram_module); RETURN_IF_FAILED(hr);
			hr = mcoll->AddModule(ram_module); RETURN_IF_FAILED(hr);

			hr = ProcessEntryPointBreakpointHit(); RETURN_IF_FAILED(hr);
		}
		else
		{
			hr = SetFelixErrorInfo(E_NOTIMPL, IDS_DONT_KNOW_HOW_TO_DEBUG_EXT_S, fileExt);
			uiShell->ReportErrorInfo(hr);
			return hr;
		}

		// Note AGO: I tried showing the simulator window here, but at this point the VS GUI is still
		// in the design mode layout. It will switch to the debug mode layout - with the simulator window
		// hidden - after this function returns.

		return S_OK;
	}

	static HRESULT ResolveEntryPointAddress (IDebugProgram2* program, IFelixLaunchOptions* _launchOptions, UINT16* pAddress)
	{
		HRESULT hr;

		wil::unique_bstr epAddressStr;
		hr = _launchOptions->get_EntryPointAddress(&epAddressStr); RETURN_IF_FAILED(hr);
		if (isdigit(epAddressStr.get()[0]))
		{
			DWORD addr;
			hr = ParseNumber(epAddressStr.get(), &addr);
			if (hr != S_OK || addr > 0xFFFF)
				return SetFelixErrorInfo(S_FALSE, IDS_WRONG_FORMAT_ENTRY_POINT_ADDRESS_S, epAddressStr.get());
			*pAddress = (UINT16)addr;
			return S_OK;
		}

		// symbol as entry point
		com_ptr<IEnumDebugModules2> modules;
		hr = program->EnumModules(&modules); RETURN_IF_FAILED(hr);
		while(true)
		{
			com_ptr<IDebugModule2> module;
			hr = modules->Next(1, &module, nullptr); RETURN_IF_FAILED(hr);
			if (hr == S_FALSE)
				return SetFelixErrorInfo(S_FALSE, IDS_CANNOT_RESOLVE_ENTRY_POINT_ADDRESS_S_S, epAddressStr.get(), L"");

			com_ptr<IZ80Module> fm;
			hr = module->QueryInterface(&fm); RETURN_IF_FAILED(hr);
			com_ptr<IFelixSymbols> symbols;
			hr = fm->GetSymbols(&symbols);
			if (SUCCEEDED(hr))
			{
				hr = symbols->GetAddressFromSymbol(epAddressStr.get(), pAddress);
				if (SUCCEEDED(hr))
					return S_OK;
			}
		}
	}

	HRESULT LoadBinary (IFelixSymbols* romSymbols, LPCWSTR exePath, LPCWSTR sourceBasePath)
	{
		HRESULT hr;

		DWORD baseAddress;
		hr = _launchOptions->get_BaseAddress(&baseAddress); RETURN_IF_FAILED(hr);

		DWORD loadedSize;
		hr = simulator->LoadBinary(exePath, baseAddress, &loadedSize);
		if (FAILED(hr))
			return uiShell->ReportErrorInfo(hr), hr;

		// Make a module as large as the loaded binary file.
		com_ptr<IFelixProgram> moduleColl;
		hr = _program->QueryInterface(&moduleColl); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IDebugModule2> exe_module;
		hr = MakeModule (baseAddress, loadedSize, exePath, true, sourceBasePath, &exe_module); RETURN_IF_FAILED(hr);
		hr = moduleColl->AddModule(exe_module.get()); RETURN_IF_FAILED(hr);

		// Now that we created all modules, let's try to resolve the entry point address string.
		// If we can't resolve the entry point, we don't have an address for PRINT USR, so we can't launch.
		UINT16 launchAddress;
		hr = ResolveEntryPointAddress(_program, _launchOptions, &launchAddress);
		if (hr != S_OK)
			return uiShell->ReportErrorInfo(hr), hr;

		// PRINT USR <LaunchAddress>
		char cmdLine[16];
		int cmdLineLen = sprintf_s (cmdLine, "\xF5\xC0%u\x0D\x80", launchAddress); RETURN_HR_IF(E_FAIL, cmdLineLen < 0);
		hr = SimulateBasicCommand(romSymbols, cmdLine); RETURN_IF_FAILED(hr);

		// Now put another breakpoint at the entry point of the Z80 program.
		WI_ASSERT(!_entryPointBreakpoint);
		hr = simulator->AddBreakpoint (BreakpointType::Code, false, launchAddress, &_entryPointBreakpoint); RETURN_IF_FAILED(hr);
		// Resume simulation so that the ZX Spectrum ROM parses our command and calls the Z80 program.
		hr = simulator->Resume(true); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	HRESULT LoadTap (IFelixSymbols* romSymbols, LPCWSTR exePath, LPCWSTR sourceBasePath)
	{
		HRESULT hr;

		// Make a module from the end of the ROM module to the end of the RAM, since we don't know how much we're going to load from tape.
		com_ptr<IDebugModule2> exe_module;
		hr = MakeModule (0x5CCB, 0x10000 - 0x5CCB, exePath, true, sourceBasePath, &exe_module); RETURN_IF_FAILED(hr);

		com_ptr<IFelixProgram> moduleColl;
		hr = _program->QueryInterface(&moduleColl); RETURN_IF_FAILED(hr);
		hr = moduleColl->AddModule(exe_module); RETURN_IF_FAILED(hr);

		// If we have launch options, we were probably launched from Start Debugging.
		// If we don't have launch options, we were probably launched from Simulator Window -> Debug File.
		if (_launchOptions)
		{
			RETURN_HR(E_NOTIMPL);
		}

		// LOAD "".
		hr = SimulateBasicCommand (romSymbols, "\xEF\x22\x22\x0D\x80"); RETURN_IF_FAILED(hr);

		// Resume simulation so that the ZX Spectrum ROM parses our command and calls the Z80 program.
		hr = simulator->Resume(true); RETURN_IF_FAILED(hr);

		hr = AdviseSink<ITapPlayNotifySink>(simulator, _weakRefToThis, &_tapPlaySinkToken); LOG_IF_FAILED(hr);

		hr = simulator->BeginPlayTapFile(exePath, TRUE);
		if (FAILED(hr))
			return uiShell->ReportErrorInfo(hr), hr;

		return S_OK;
	}

	HRESULT ProcessEntryPointBreakpointHit()
	{
		HRESULT hr;

		com_ptr<IDebugThread2> thread;
		{
			com_ptr<IEnumDebugThreads2> threads;
			hr = _program->EnumThreads(&threads); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
			{
				hr = threads->Next(1, &thread, nullptr); LOG_IF_FAILED(hr);
			}
		}

		hr = SendLoadCompleteEvent (_callback.get(), this, _program.get(), thread.get()); LOG_IF_FAILED(hr);

		hr = SendEntryPointEvent (_callback.get(), this, _program.get(), thread.get()); LOG_IF_FAILED(hr);

		// Let's put a breakpoint at the exit point.
		com_ptr<IDebugModule2> romModule;
		hr = GetModuleAtAddress(_program, 0, &romModule); RETURN_IF_FAILED(hr);
		com_ptr<IFelixSymbols> romSymbols;
		hr = romModule.try_query<IZ80Module>()->GetSymbols(&romSymbols); RETURN_IF_FAILED(hr);
		uint16_t addr;
		hr = romSymbols->GetAddressFromSymbol(L"MAIN-4", &addr); RETURN_IF_FAILED(hr);
		hr = simulator->AddBreakpoint(BreakpointType::Code, false, addr, &_exitPointBreakpoint); RETURN_IF_FAILED(hr);

		// The simulation paused execution before calling us, and we sent the entry point even which is a stopping event.
		// If the user started debugging with Start Debugging, VS will call our implementation of IDebugProgram2::Continue.
		// if the user started debugging with Step Into, VS will not call it.

		return S_OK;
	}

	void TerminateInternal()
	{
		// It's not right to call our own implementations, but I can't find another way to terminate the debug session.
		// I tried calling IDebugSession2::Terminate(), but it returns E_NOTIMPL.
		wil::com_ptr_nothrow<IDebugProcess2> process;
		auto hr = _program->GetProcess(&process); LOG_IF_FAILED(hr);
		hr = _program->Terminate(); LOG_IF_FAILED(hr);
		hr = TerminateProcess(process.get()); LOG_IF_FAILED(hr);
	}

	HRESULT ProcessExitPointBreakpointHit()
	{
		//z80_register_set regs;
		//hr = simulator->GetRegisters(&regs, (uint32_t)sizeof(regs));
		//uint16_t bc = SUCCEEDED(hr) ? regs.main.bc : 0xFFFF;

		TerminateInternal();

		// We'll resume execution of code from the ZX Spectrum ROM in ContinueFromSynchronousEvent().

		return S_OK;
	}

	virtual HRESULT __stdcall CanTerminateProcess(IDebugProcess2* pProcess) override
	{
		return S_OK;
	}

	struct ProgramDestroyEvent : public EventBase<IDebugProgramDestroyEvent2, EVENT_SYNCHRONOUS>
	{
		DWORD _exitCode;

		virtual HRESULT __stdcall GetExitCode(DWORD* exitCode) override
		{
			*exitCode = _exitCode;
			return S_OK;
		}
	};

	virtual HRESULT __stdcall TerminateProcess(IDebugProcess2* pProcess) override
	{
		_tapPlaySinkToken.reset();
		_simulatorEventsToken.reset();

		if (_entryPointBreakpoint)
		{
			auto hr = simulator->RemoveBreakpoint(_entryPointBreakpoint); LOG_IF_FAILED(hr);
			_entryPointBreakpoint = 0;
		}

		if (_exitPointBreakpoint)
		{
			auto hr = simulator->RemoveBreakpoint(_exitPointBreakpoint); LOG_IF_FAILED(hr);
			_exitPointBreakpoint = 0;
		}

		wil::com_ptr_nothrow<ProgramDestroyEvent> pde = new (std::nothrow) ProgramDestroyEvent(); RETURN_IF_NULL_ALLOC(pde);
		pde->_exitCode = 0;
		auto hr = pde->Send (_callback.get(), this, _program.get(), nullptr); LOG_IF_FAILED(hr);

		// We sent the Program Destroy event synchronously. This means execution continues on this thread,
		// this function returns, and sometime later the SDM will process the event and will call our
		// ContinueFromSynchronousEvent() with the event as parameter. In ContinueFromSynchronousEvent(),
		// we release the session's references without removing the program from discovery.
		// (Various tool windows might still keep references around. An example is the call stack window
		// which keeps many indirect references, via IEnumDebugFrameInfo2, until the user closes Visual Studio.
		// Another example is the Disassembly window.)

		return S_OK;
	}
	#pragma endregion

	#pragma region ISimulatorEventNotifySink
	virtual HRESULT STDMETHODCALLTYPE NotifySimulatorEvent (ISimulatorEvent* event, REFIID riidEvent) override
	{
		HRESULT hr;

		if (riidEvent == __uuidof(ISimulatorResumeEvent))
		{
			return S_OK;
		}

		if (riidEvent == __uuidof(ISimulatorBreakpointEvent))
		{
			com_ptr<ISimulatorBreakpointEvent> bpe;
			auto hr = event->QueryInterface(&bpe); RETURN_IF_FAILED(hr);
			for (ULONG i = 0; i < bpe->GetBreakpointCount(); i++)
			{
				SIM_BP_COOKIE bp;
				auto hr = bpe->GetBreakpointAt(i, &bp); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr))
				{
					if (bp == _entryPointBreakpoint)
					{
						hr = simulator->RemoveBreakpoint(_entryPointBreakpoint); LOG_IF_FAILED(hr);
						_entryPointBreakpoint = 0;
						ProcessEntryPointBreakpointHit();
					}
					else if (bp == _exitPointBreakpoint)
					{
						hr = simulator->RemoveBreakpoint(_exitPointBreakpoint); LOG_IF_FAILED(hr);
						_exitPointBreakpoint = 0;
						ProcessExitPointBreakpointHit();
					}
					else
					{ } // some other breakpoint, not interesting for us
				}
			}

			return S_OK;
		}

		return S_FALSE;
	}
	#pragma endregion

	#pragma region IFelixLaunchOptionsProvider
	virtual HRESULT GetLaunchOptions (IFelixLaunchOptions** ppOptions) override
	{
		return _launchOptions.copy_to(ppOptions);
	}
	#pragma endregion

	#pragma region ITapPlayNotifySink
	virtual HRESULT STDMETHODCALLTYPE NotifyTapPlayStarting() override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE NotifyTapPlayComplete() override
	{
		HRESULT hr;

		WI_ASSERT(_tapPlaySinkToken);
		_tapPlaySinkToken.reset();

		hr = simulator->SetSpeed(100); LOG_IF_FAILED(hr);
		hr = simulator->Break(); LOG_IF_FAILED(hr);
		hr = ProcessEntryPointBreakpointHit(); LOG_IF_FAILED(hr);
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeDebugEngine (IDebugEngine2** to)
{
	wil::com_ptr_nothrow<Z80DebugEngine> p = new (std::nothrow) Z80DebugEngine(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
