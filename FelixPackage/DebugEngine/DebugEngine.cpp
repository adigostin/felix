
#include "pch.h"
#include "DebugEngine.h"
#include "DebugEventBase.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
#include "../guids.h"
#include "../FelixPackage.h"
#include "../Z80Xml.h"

class Z80DebugEngine : public IDebugEngine2, IDebugEngineLaunch2, ISimulatorEventHandler, IFelixLaunchOptionsProvider
{
	ULONG _refCount = 0;
	wil::unique_bstr _registryRoot;
	WORD langId = 0;
	bool _attachCalled = false;
	com_ptr<IDebugEventCallback2> _callback;
	com_ptr<ISimulator> _simulator;
	com_ptr<IDebugPort2> _port;
	com_ptr<IDebugProgram2> _program;
	com_ptr<IFelixLaunchOptions> _launchOptions;
	com_ptr<IBreakpointManager> _bpman;
	SIM_BP_COOKIE _editorFunctionBreakpoint = 0;
	SIM_BP_COOKIE _entryPointBreakpoint = 0;
	SIM_BP_COOKIE _exitPointBreakpoint = 0;
	bool _advisingSimulatorEvents = false;
	static const uint16_t binFileAddr = 0x8000;

public:
	~Z80DebugEngine()
	{
		WI_ASSERT(!_advisingSimulatorEvents);
		WI_ASSERT(!_simulator);
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = NULL;

		if (   TryQI<IUnknown>(static_cast<IDebugEngine2*>(this), riid, ppvObject)
			|| TryQI<IDebugEngine2>(this, riid, ppvObject)
			|| TryQI<IDebugEngineLaunch2>(this, riid, ppvObject)
			|| TryQI<ISimulatorEventHandler>(this, riid, ppvObject)
			|| TryQI<IFelixLaunchOptionsProvider>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		if (riid == IID_IDebugEngine110)
			return E_NOTIMPL;
		if (riid == IID_IDebugEngine150)
			return E_NOTIMPL;
		if (riid == IID_IDebugSymbolSettings100)
			return E_NOTIMPL;

		if (     riid == IID_IDebugEngine3
			||   riid == IID_IDebugEngineLaunch100
			||   riid == IID_INoIdea6_DebugEngine
			||   riid == IID_IDebugProgramProvider2
			||   riid == IID_INoIdea7
			||   riid == IID_INotARealInterface
			||   riid == IID_INoIdea9
			||   riid == IID_INoIdea10
			||   riid == IID_INoIdea11
			||   riid == IID_INoIdea14
			||   riid == IID_INoIdea15
			||   riid == IID_IMarshal
			||   riid == IID_INoMarshal
			||   riid == IID_IStdMarshalInfo
			||   riid == IID_IdentityUnmarshal
			||   riid == IID_IAgileObject
			||   riid == IID_IFastRundown
			||   riid == IID_IExternalConnection
			||   riid == IID_IDebugBreakpointFileUpdateNotification110
			||   riid == IID_IDebugEngineStepFilterManager90
			||   riid == IID_IDebugVisualizerExtensionReceiver178
			||   riid == IID_IDebugSymbolSettings170
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

	virtual HRESULT __stdcall Attach(IDebugProgram2** rgpPrograms, IDebugProgramNode2** rgpProgramNodes, DWORD celtPrograms, IDebugEventCallback2* pCallback, ATTACH_REASON dwReason) override
	{
		WI_ASSERT(!_attachCalled);

		if (celtPrograms != 1)
			RETURN_HR(E_INVALIDARG);

		using ProgramCreateEvent = EventBase<IDebugProgramCreateEvent2, EVENT_ASYNCHRONOUS>;
		wil::com_ptr_nothrow<ProgramCreateEvent> pce = new (std::nothrow) ProgramCreateEvent(); RETURN_IF_NULL_ALLOC(pce);
		auto hr = pce->Send(pCallback, this, rgpPrograms[0], nullptr); RETURN_IF_FAILED(hr);

		_attachCalled = true;

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
					RETURN_HR(E_NOTIMPL);
///					BP_REQUEST_INFO info = { };
///					hr = pBPRequest->GetRequestInfo (BPREQI_BPLOCATION, &info); RETURN_IF_FAILED(hr);
///					if (info.bpLocation.bpLocationType != locationType)
///						RETURN_HR(E_FAIL);					
///					auto cc = wil::com_ptr_nothrow(info.bpLocation.bpLocation.bplocCodeContext.pCodeContext);
///					info.bpLocation.bpLocation.bplocCodeContext.pCodeContext->Release();
///					info.bpLocation.bpLocation.bplocCodeContext.pCodeContext = 0;
///					wil::com_ptr_nothrow<IZ80CodeContext> z80cc;
///					hr = cc->QueryInterface(&z80cc); RETURN_IF_FAILED(hr);
///					if (z80cc->PhysicalMemorySpace())
///						RETURN_HR(E_FAIL);
///					hr = MakeSimplePendingBreakpoint (_program.get(), (uint16_t)z80cc->Address(), ppPendingBP); RETURN_IF_FAILED(hr);
///					return S_OK;
				}

				case BPLT_ADDRESS:
				{
					BP_REQUEST_INFO info = { };
					hr = pBPRequest->GetRequestInfo (BPREQI_BPLOCATION, &info); RETURN_IF_FAILED(hr);
					if (info.bpLocation.bpLocationType != locationType)
						RETURN_HR(E_FAIL);

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

					if (!converted)
						RETURN_HR(E_FAIL);

					if (address >= 0x10000)
						RETURN_HR(E_FAIL);

					RETURN_HR(E_NOTIMPL);
					//hr = MakeSimplePendingBreakpoint (_program.get(), (uint16_t)address, ppPendingBP); RETURN_IF_FAILED(hr);
					//return S_OK;
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

					hr = MakeSourceLinePendingBreakpoint (_callback, this, _program, _bpman, pBPRequest, file_path.get(), begin.dwLine, ppPendingBP); RETURN_IF_FAILED(hr);
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
		*pguidEngine = Engine_CLSID;
		return S_OK;
	}

	virtual HRESULT __stdcall DestroyProgram(IDebugProgram2* pProgram) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall ContinueFromSynchronousEvent(IDebugEvent2* pEvent) override
	{
		wil::com_ptr_nothrow<IDebugProgramDestroyEvent2> pde;
		if (SUCCEEDED(pEvent->QueryInterface(&pde)))
		{
			// Now that we told the SDM that we destroyed the program and the SDM is done acting on the event,
			// we tell also the port about the program being gone. This should release even more references.

			wil::com_ptr_nothrow<IZ80DebugPort> z80Port;
			auto hr = _port->QueryInterface(&z80Port); RETURN_IF_FAILED(hr);
			
			hr = z80Port->SendProgramDestroyEventToSinks (_program.get(), 0); RETURN_IF_FAILED(hr);
			
			if (_advisingSimulatorEvents)
			{
				_simulator->UnadviseDebugEvents(this);
				_advisingSimulatorEvents = false;
			}

			WI_ASSERT(_simulator->HasBreakpoints_HR() == S_FALSE);
			if (_simulator->Running_HR() == S_FALSE)
			{
				auto hr = _simulator->Resume(false); LOG_IF_FAILED(hr);
			}

			_callback = nullptr;
			_simulator = nullptr;
			_program = nullptr;
			_launchOptions = nullptr;
			_bpman = nullptr;
			_port = nullptr;

			return S_OK;
		}

		wil::com_ptr_nothrow<IDebugModuleLoadEvent2> mle;
		if (SUCCEEDED(pEvent->QueryInterface(&mle)))
			return S_OK;

		wil::com_ptr_nothrow<IDebugThreadDestroyEvent2> tde;
		if (SUCCEEDED(pEvent->QueryInterface(&tde)))
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
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	static HRESULT SendLoadCompleteEvent (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, IDebugThread2* thread)
	{
		// Must be stopping as per https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/supported-event-types?view=vs-2019
		using LCE = EventBase<IDebugLoadCompleteEvent2, EVENT_SYNC_STOP>;
		auto e = com_ptr(new (std::nothrow) LCE()); RETURN_IF_NULL_ALLOC(e);
		auto hr = callback->Event(engine, nullptr, program, thread, e.get(), IID_IDebugLoadCompleteEvent2, EVENT_SYNC_STOP); RETURN_IF_FAILED(hr);
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
		WI_ASSERT(!_simulator);
		WI_ASSERT(!_program);
		WI_ASSERT(!_launchOptions);

		if (pszOptions)
		{
			auto stream = com_ptr(SHCreateMemStream((BYTE*)pszOptions, 2 * (UINT)wcslen(pszOptions))); RETURN_IF_NULL_ALLOC(stream);
			hr = MakeLaunchOptions(&_launchOptions); RETURN_IF_FAILED(hr);
			hr = LoadFromXml (_launchOptions.get(), nullptr, stream.get(), 1200); RETURN_IF_FAILED(hr);
		}
		//_callback = pCallback;
		//auto reset_callback_ptr = wil::scope_exit([this] { _callback = nullptr; });

		com_ptr<ISimulator> simulator;
		hr = serviceProvider->QueryService(SID_Simulator, &simulator); RETURN_IF_FAILED(hr);
		hr = simulator->Break(); RETURN_IF_FAILED(hr);
		hr = simulator->Reset(0); RETURN_IF_FAILED(hr);

		com_ptr<IDebugProcess2> process;
		hr = MakeDebugProcess (pPort, pszExe, this, simulator.get(), pCallback, &process); RETURN_IF_FAILED(hr);

		struct EngineCreateEvent : public EventBase<IDebugEngineCreateEvent2, EVENT_ASYNCHRONOUS>
		{
			wil::com_ptr_nothrow<IDebugEngine2> _engine;

			virtual HRESULT STDMETHODCALLTYPE GetEngine (IDebugEngine2 **pEngine) override
			{
				*pEngine = _engine.get();
				_engine->AddRef();
				return S_OK;
			}
		};

		wil::com_ptr_nothrow<EngineCreateEvent> ece = new (std::nothrow) EngineCreateEvent(); RETURN_IF_NULL_ALLOC(ece);
		ece->_engine = this;
		hr = ece->Send(pCallback, this, nullptr, nullptr); RETURN_IF_FAILED(hr);

		_callback = pCallback;
		_simulator = std::move(simulator);
		_port = pPort;

		*ppProcess = process.detach();
		return S_OK;
	}

	// https://docs.microsoft.com/en-us/previous-versions/bb145570%28v%3dvs.90%29

	virtual HRESULT __stdcall ResumeProcess(IDebugProcess2* pProcess) override
	{
		// TODO: cleanup member variables in case of failure in this function.

		com_ptr<IEnumDebugPrograms2> programs;
		auto hr = pProcess->EnumPrograms(&programs); RETURN_IF_FAILED(hr);
		hr = programs->Next(1, &_program, nullptr); RETURN_IF_FAILED(hr);
		auto resetProgram = wil::scope_exit([this] { _program = nullptr; });

		hr = MakeBreakpointManager(_callback, this, _program, _simulator, &_bpman); RETURN_IF_FAILED(hr);

		WI_ASSERT (!_attachCalled);
		wil::com_ptr_nothrow<IDebugPort2> port;
		hr = pProcess->GetPort(&port); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IZ80DebugPort> z80Port;
		hr = port->QueryInterface(&z80Port); RETURN_IF_FAILED(hr);
		hr = z80Port->SendProgramCreateEventToSinks (_program.get()); RETURN_IF_FAILED(hr);
		WI_ASSERT (_attachCalled);

		using ThreadCreateEvent = EventBase<IDebugThreadCreateEvent2, EVENT_ASYNCHRONOUS>;
		wil::com_ptr_nothrow<ThreadCreateEvent> tce = new (std::nothrow) ThreadCreateEvent(); RETURN_IF_NULL_ALLOC(tce);
		com_ptr<IEnumDebugThreads2> enumThreads;
		hr = _program->EnumThreads(&enumThreads); RETURN_IF_FAILED(hr);
		com_ptr<IDebugThread2> thread;
		hr = enumThreads->Next(1, &thread, nullptr); RETURN_IF_FAILED(hr);
		hr = tce->Send (_callback.get(), this, _program.get(), thread.get()); LOG_IF_FAILED(hr);

		// Create a debug module for the ROM.
		wil::unique_hlocal_string packageDir;
		hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, packageDir); RETURN_IF_FAILED(hr);
		auto fnres = PathFindFileName(packageDir.get()); RETURN_HR_IF(CO_E_BAD_PATH, fnres == packageDir.get());
		*fnres = 0;
		DWORD PathFlags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
		wil::unique_hlocal_string rom_path;
		hr = PathAllocCombine (packageDir.get(), L"ROMs/Spectrum48K.rom", PathFlags, &rom_path); RETURN_IF_FAILED(hr);
		wil::unique_hlocal_string rom_debug_info_path;
		hr = PathAllocCombine (packageDir.get(), L"ROMs/Spectrum48K.z80sym", PathFlags, &rom_debug_info_path); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IDebugModule2> romModule;
		hr = MakeModule (0, 0x5CCB, rom_path.get(), rom_debug_info_path.get(), false,
			this, _program.get(), _callback.get(), &romModule); RETURN_IF_FAILED(hr);
		com_ptr<IDebugModuleCollection> mcoll;
		hr = _program->QueryInterface(&mcoll); RETURN_IF_FAILED(hr);
		hr = mcoll->AddModule(romModule.get()); RETURN_IF_FAILED(hr);

		hr = _simulator->AdviseDebugEvents(this); RETURN_IF_FAILED(hr);
		_advisingSimulatorEvents = true;

		wil::unique_bstr exePath;
		hr = pProcess->GetName (GN_FILENAME, &exePath); RETURN_IF_FAILED(hr);
		if (!exePath)
			RETURN_HR(E_NO_EXE_FILENAME);

		/*
		if (!_wcsicmp(PathFindExtension(exePath.get()), L".sna"))
		{
			wil::com_ptr_nothrow<IStream> stream;
			auto hr = SHCreateStreamOnFileEx (exePath.get(), STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &stream); RETURN_IF_FAILED(hr);
			hr = _simulator->LoadSnapshot(stream.get()); RETURN_IF_FAILED(hr);
			wil::com_ptr_nothrow<IDebugModule2> ram_module;
			hr = MakeModule (0x4000, 0xC000, exePath.get(), nullptr, true, &ram_module); RETURN_IF_FAILED(hr);
			hr = program->AddModule(ram_module.get()); RETURN_IF_FAILED(hr);

			hr = SendLoadCompleteEvent (_callback.get(), this, program.get(), thread.get()); RETURN_IF_FAILED(hr);
			hr = SendEntryPointEvent (_callback.get(), this, program.get(), thread.get()); RETURN_IF_FAILED(hr);
		}
		else
		*/if (!_wcsicmp(PathFindExtension(exePath.get()), L".bin"))
		{
			// Simulate some instructions until the EDITOR function is called.
			// TODO: use timeout
			wil::com_ptr_nothrow<IZ80Module> rom;
			hr = romModule->QueryInterface(&rom); RETURN_IF_FAILED(hr);
			wil::com_ptr_nothrow<IZ80Symbols> romSymbols;
			hr = rom->GetSymbols(&romSymbols); RETURN_IF_FAILED(hr);
			UINT16 editorFunctionAddr;
			hr = romSymbols->GetAddressFromSymbol(L"EDITOR", &editorFunctionAddr); RETURN_IF_FAILED(hr);
			WI_ASSERT(!_entryPointBreakpoint);
			WI_ASSERT(!_editorFunctionBreakpoint);
			hr = _simulator->AddBreakpoint (BreakpointType::Code, false, editorFunctionAddr, &_editorFunctionBreakpoint); RETURN_IF_FAILED(hr);
			hr = _simulator->Reset(0); RETURN_IF_FAILED(hr);
			if (_simulator->Running_HR() == S_FALSE)
			{
				hr = _simulator->Resume(true); RETURN_IF_FAILED(hr);
			}
		}
		else
			RETURN_HR(E_UNKNOWN_FILE_EXTENSION);

		// Note AGO: I tried showing the simulator window here, but at this point the VS GUI is still
		// in the design mode layout. It will switch to the debug mode layout - with the simulator window
		// hidden - after this function returns.

		resetProgram.release();
		return S_OK;
	}

	/* From https://github.com/reclaimed/prettybasic/blob/master/doc/ZX%20Spectrum%2048K%20ROM%20Original%20Disassembly.asm
	; The memory.
	;
	; +---------+-----------+------------+--------------+-------------+--
	; | BASIC   |  Display  | Attributes | ZX Printer   |    System   | 
	; |  ROM    |   File    |    File    |   Buffer     |  Variables  | 
	; +---------+-----------+------------+--------------+-------------+--
	; ^         ^           ^            ^              ^             ^
	; $0000   $4000       $5800        $5B00          $5C00         $5CB6 = CHANS 
	;
	;
	;  --+----------+---+---------+-----------+---+------------+--+---+--
	;    | Channel  |$80|  BASIC  | Variables |$80| Edit Line  |NL|$80|
	;    |   Info   |   | Program |   Area    |   | or Command |  |   |
	;  --+----------+---+---------+-----------+---+------------+--+---+--
	;    ^              ^         ^               ^                   ^
	;  CHANS           PROG      VARS           E_LINE              WORKSP
	;
	;
	;                             ---5-->         <---2---  <--3---
	;  --+-------+--+------------+-------+-------+---------+-------+-+---+------+
	;    | INPUT |NL| Temporary  | Calc. | Spare | Machine | GOSUB |?|$3E| UDGs |
	;    | data  |  | Work Space | Stack |       |  Stack  | Stack | |   |      |
	;  --+-------+--+------------+-------+-------+---------+-------+-+---+------+
	;    ^                       ^       ^       ^                   ^   ^      ^
	;  WORKSP                  STKBOT  STKEND   sp               RAMTOP UDG  P_RAMT
	;               
	*/

	HRESULT ResolveZxSpectrumSymbol (LPCWSTR name, UINT16* address)
	{
		wil::com_ptr_nothrow<IEnumDebugModules2> modules;
		auto hr = _program->EnumModules(&modules); RETURN_IF_FAILED(hr);

		while(true)
		{
			com_ptr<IDebugModule2> module;
			hr = modules->Next(1, &module, nullptr); RETURN_IF_FAILED(hr);
			if (hr == S_FALSE)
				return E_SYMBOL_NOT_IN_SYMBOL_FILE;
	
			wil::com_ptr_nothrow<IZ80Module> fm;
			hr = module->QueryInterface(&fm); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
			{
				wil::com_ptr_nothrow<IZ80Symbols> romSymbols;
				hr = fm->GetSymbols(&romSymbols); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr))
				{
					hr = romSymbols->GetAddressFromSymbol(name, address);
					if (SUCCEEDED(hr))
						return S_OK;
				}
			}
		}
	}

	HRESULT ReadZxSpectrumSystemVar (LPCWSTR name, UINT16* value)
	{
		UINT16 addr;
		auto hr = ResolveZxSpectrumSymbol (name, &addr); RETURN_IF_FAILED(hr);
		hr = _simulator->ReadMemoryBus(addr, 2, value); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	HRESULT WriteZxSpectrumSystemVar (LPCWSTR name, UINT16 value)
	{
		UINT16 addr;
		auto hr = ResolveZxSpectrumSymbol (name, &addr); RETURN_IF_FAILED(hr);
		hr = _simulator->WriteMemoryBus(addr, 2, &value); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	
	HRESULT ProcessEditorFunctionBreakpointHit()
	{
		WI_ASSERT(_editorFunctionBreakpoint);
		auto hr = _simulator->RemoveBreakpoint(_editorFunctionBreakpoint); LOG_IF_FAILED(hr);
		_editorFunctionBreakpoint = 0;

		com_ptr<IDebugProcess2> process;
		hr = _program->GetProcess(&process); RETURN_IF_FAILED(hr);
		wil::unique_bstr exePath;
		hr = process->GetName (GN_FILENAME, &exePath); RETURN_IF_FAILED(hr);
		if (!exePath)
			RETURN_HR(E_NO_EXE_FILENAME);

		// Load the binary file.
		wil::com_ptr_nothrow<IStream> stream;
		hr = SHCreateStreamOnFileEx (exePath.get(), STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &stream); RETURN_IF_FAILED(hr);
		STATSTG stat;
		hr = stream->Stat (&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		hr = _simulator->LoadBinary(stream.get(), binFileAddr); RETURN_IF_FAILED(hr);
		auto debug_info_path = wil::make_process_heap_string_nothrow(exePath.get()); RETURN_IF_NULL_ALLOC(debug_info_path);
		hr = PathCchRenameExtension (debug_info_path.get(), wcslen(debug_info_path.get()) + 1, L".sld"); RETURN_IF_FAILED(hr);
		com_ptr<IDebugModuleCollection> moduleColl;
		hr = _program->QueryInterface(&moduleColl); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IDebugModule2> exe_module;
		hr = MakeModule (binFileAddr, stat.cbSize.LowPart, exePath.get(), debug_info_path.get(), true,
			this, _program.get(), _callback.get(), &exe_module); RETURN_IF_FAILED(hr);
		hr = moduleColl->AddModule(exe_module.get()); RETURN_IF_FAILED(hr);

		// Simulate what the EDITOR function would do when typing "PRINT USR 32768".

		// The command line is from E-LINE to WORKSP.
		UINT16 eline, worksp;
		hr = ReadZxSpectrumSystemVar(L"E-LINE", &eline); RETURN_IF_FAILED(hr);
		hr = ReadZxSpectrumSystemVar(L"WORKSP", &worksp); RETURN_IF_FAILED(hr);

		// For now let's assume that STKBOT and STKEND have the same value as WORKSP.
		// This is true since we just reset the processor and simulated to the EDITOR function.
		UINT16 stkbot, stkend;
		hr = ReadZxSpectrumSystemVar(L"STKBOT", &stkbot); RETURN_IF_FAILED(hr);
		hr = ReadZxSpectrumSystemVar(L"STKEND", &stkend); RETURN_IF_FAILED(hr);
		RETURN_HR_IF (E_FAIL, (stkbot != worksp) || (stkend != worksp));

		// We need to replace the command line with PRINT USR <addr>.
		char cmdLine[16];
		int cmdLineLen = sprintf_s (cmdLine, "\xF5\xC0%u\x0D\x80", binFileAddr); RETURN_HR_IF(E_FAIL, cmdLineLen < 0);
		hr = _simulator->WriteMemoryBus (eline, (uint16_t)cmdLineLen, cmdLine); RETURN_IF_FAILED(hr);
		hr = WriteZxSpectrumSystemVar (L"K-CUR", eline + cmdLineLen - 2); RETURN_IF_FAILED(hr);
		hr = WriteZxSpectrumSystemVar (L"WORKSP", eline + cmdLineLen); RETURN_IF_FAILED(hr);
		hr = WriteZxSpectrumSystemVar (L"STKBOT", eline + cmdLineLen); RETURN_IF_FAILED(hr);
		hr = WriteZxSpectrumSystemVar (L"STKEND", eline + cmdLineLen); RETURN_IF_FAILED(hr);

		// Jump to some RET instruction, say the one at 0F91h.
		uint8_t testRet;
		hr = _simulator->ReadMemoryBus(0x0F91, 1, &testRet); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, testRet != 0xC9);
		hr = _simulator->SetPC(0x0F91); RETURN_IF_FAILED(hr);
			
		// Now put another breakpoint at the entry point of the Z80 program.
		WI_ASSERT(!_entryPointBreakpoint);
		hr = _simulator->AddBreakpoint (BreakpointType::Code, false, binFileAddr, &_entryPointBreakpoint); RETURN_IF_FAILED(hr);
		// Resume simulation so that the ZX Spectrum ROM parses our command and calls the Z80 program.
		hr = _simulator->Resume(true); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	
	HRESULT ProcessEntryPointBreakpointHit()
	{
		WI_ASSERT(_entryPointBreakpoint);
		auto hr = _simulator->RemoveBreakpoint(_entryPointBreakpoint); LOG_IF_FAILED(hr);
		_entryPointBreakpoint = 0;

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
		uint16_t stackBC;
		hr = ResolveZxSpectrumSymbol (L"STACK-BC", &stackBC); RETURN_IF_FAILED(hr);
		hr = _simulator->AddBreakpoint(BreakpointType::Code, false, stackBC, &_exitPointBreakpoint); RETURN_IF_FAILED(hr);

		// The simulation paused execution before calling us, and we sent the entry point even which is a stopping event.
		// If the user started debugging with Start Debugging, VS will call our implementation of IDebugProgram2::Continue.
		// if the user started debugging with Step Into, VS will not call it.

		return S_OK;
	}

	HRESULT ProcessExitPointBreakpointHit()
	{
		WI_ASSERT(_exitPointBreakpoint);
		auto hr = _simulator->RemoveBreakpoint(_exitPointBreakpoint); LOG_IF_FAILED(hr);
		_exitPointBreakpoint = 0;

		z80_register_set regs;
		hr = _simulator->GetRegisters(&regs, (uint32_t)sizeof(regs));
		uint16_t bc = SUCCEEDED(hr) ? regs.main.bc : 0xFFFF;

		// It's not right to call our own implementations, but I can't find another way to terminate the debug session.
		// I tried calling IDebugSession2::Terminate(), but it returns E_NOTIMPL.
		wil::com_ptr_nothrow<IDebugProcess2> process;
		hr = _program->GetProcess(&process); LOG_IF_FAILED(hr);
		hr = _program->Terminate(); LOG_IF_FAILED(hr);
		hr = TerminateProcess(process.get()); LOG_IF_FAILED(hr);

		// We'll resume execution of code from the ZX Spectrum ROM in ContinueFromSynchronousEvent().

		return S_OK;
	}

	virtual HRESULT __stdcall CanTerminateProcess(IDebugProcess2* pProcess) override
	{
		return S_OK;
	}

	virtual HRESULT __stdcall TerminateProcess(IDebugProcess2* pProcess) override
	{
		if (_editorFunctionBreakpoint)
		{
			auto hr = _simulator->RemoveBreakpoint(_editorFunctionBreakpoint); LOG_IF_FAILED(hr);
			_editorFunctionBreakpoint = 0;
		}
		
		if (_entryPointBreakpoint)
		{
			auto hr = _simulator->RemoveBreakpoint(_entryPointBreakpoint); LOG_IF_FAILED(hr);
			_entryPointBreakpoint = 0;
		}

		if (_exitPointBreakpoint)
		{
			auto hr = _simulator->RemoveBreakpoint(_exitPointBreakpoint); LOG_IF_FAILED(hr);
			_exitPointBreakpoint = 0;
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

		wil::com_ptr_nothrow<ProgramDestroyEvent> pde = new (std::nothrow) ProgramDestroyEvent(); RETURN_IF_NULL_ALLOC(pde);
		pde->_exitCode = 0;
		auto hr = pde->Send (_callback.get(), this, _program.get(), nullptr); LOG_IF_FAILED(hr);

		// We sent the Program Destroy event synchronously. This means execution continues on this thread,
		// this function returns, and sometime later the SDM will process the event and will call our
		// ContinueFromSynchronousEvent() with the event as parameter. In ContinueFromSynchronousEvent(),
		// we tell the port that the program is gone and we release our reference to the program.
		// (Various tool windows might still keep references around. An example is the call stack window
		// which keeps many indirect references, via IEnumDebugFrameInfo2, until the user closes Visual Studio.
		// Another example is the Disassembly window.)

		return S_OK;
	}
	#pragma endregion

	#pragma region ISimulatorEventHandler
	virtual HRESULT STDMETHODCALLTYPE ProcessSimulatorEvent (ISimulatorEvent* event, REFIID riidEvent) override
	{
		if (riidEvent == __uuidof(ISimulatorSimulateOneEvent))
		{
			// Must be stopping as per https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/supported-event-types?view=vs-2019
			using StepCompleteEvent = EventBase<IDebugStepCompleteEvent2, EVENT_SYNC_STOP>;
			wil::com_ptr_nothrow<StepCompleteEvent> e = new (std::nothrow) StepCompleteEvent(); RETURN_IF_NULL_ALLOC(e);
			com_ptr<IDebugThread2> thread;
			{
				com_ptr<IEnumDebugThreads2> enumThreads;
				auto hr = _program->EnumThreads(&enumThreads); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr))
				{
					hr = enumThreads->Next(1, &thread, nullptr); LOG_IF_FAILED(hr);
				}
			}
			auto hr = e->Send (_callback.get(), this, _program.get(), thread.get()); RETURN_IF_FAILED(hr);
			return S_OK;
		}

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
					if (bp == _editorFunctionBreakpoint)
						ProcessEditorFunctionBreakpointHit();
					else if (bp == _entryPointBreakpoint)
						ProcessEntryPointBreakpointHit();
					else if (bp == _exitPointBreakpoint)
						ProcessExitPointBreakpointHit();
					else
					{ } // some other breakpoint, not interesting for us
				}
			}

			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IFelixLaunchOptionsProvider
	virtual HRESULT GetLaunchOptions (IFelixLaunchOptions** ppOptions) override
	{
		return _launchOptions.copy_to(ppOptions);
	}
	#pragma endregion
};

HRESULT MakeDebugEngine (IDebugEngine2** to)
{
	wil::com_ptr_nothrow<Z80DebugEngine> p = new (std::nothrow) Z80DebugEngine(); RETURN_IF_NULL_ALLOC(p);
	*to = p.detach();
	return S_OK;
}

HRESULT GetAddressFromSourceLocation (IDebugProgram2* program, LPCWSTR projectDirOrNull, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress)
{
	RETURN_HR_IF(E_POINTER, !pAddress);

	com_ptr<IEnumDebugModules2> modules;
	auto hr = program->EnumModules(&modules); RETURN_IF_FAILED(hr);

	com_ptr<IDebugModule2> module;
	while (SUCCEEDED(modules->Next(1, &module, nullptr)) && module)
	{
		hr = ::GetAddressFromSourceLocation(module.get(), projectDirOrNull, sourceLocationFilename, sourceLocationLineIndex, pAddress);
		if (SUCCEEDED(hr))
			return S_OK;
	}

	return E_ADDRESS_NOT_IN_SYMBOL_FILE;
}

HRESULT GetAddressFromSourceLocation (IDebugModule2* module, LPCWSTR projectDirOrNull, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress)
{
	wil::com_ptr_nothrow<IZ80Module> z80m;
	auto hr = module->QueryInterface(&z80m); RETURN_IF_FAILED_EXPECTED(hr);

	wil::com_ptr_nothrow<IZ80Symbols> symbols;
	hr = z80m->GetSymbols(&symbols); RETURN_IF_FAILED_EXPECTED(hr);

	if (symbols->HasSourceLocationInformation() != S_OK)
		return E_SRC_LOCATION_NOT_IN_SYMBOLS;

	wil::unique_process_heap_string filePathRelative;
	if (projectDirOrNull)
	{
		auto relative = wil::make_process_heap_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(relative);
		BOOL bRes = PathRelativePathToW (relative.get(), projectDirOrNull, FILE_ATTRIBUTE_DIRECTORY, sourceLocationFilename, 0); RETURN_HR_IF(CS_E_INVALID_PATH, !bRes);
		hr = PathAllocCanonicalize (relative.get(), PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS, &filePathRelative); RETURN_IF_FAILED(hr);
	}
	else
		RETURN_HR(E_NOTIMPL);

	uint16_t addr;
	hr = symbols->GetAddressFromSourceLocation(filePathRelative.get(), sourceLocationLineIndex, &addr); RETURN_IF_FAILED_EXPECTED(hr);

	*pAddress = addr;
	return S_OK;
}

HRESULT GetModuleAtAddress (IDebugProgram2* program, UINT64 address, OUT IDebugModule2** ppModule)
{
	*ppModule = nullptr;

	com_ptr<IEnumDebugModules2> modules;
	auto hr = program->EnumModules(&modules); RETURN_IF_FAILED(hr);
	com_ptr<IDebugModule2> module;
	while(SUCCEEDED(modules->Next(1, &module, nullptr)) && module)
	{
		MODULE_INFO mi;
		hr = module->GetInfo(MIF_LOADADDRESS | MIF_SIZE, &mi); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr) && (address >= mi.m_addrLoadAddress) && (address < mi.m_addrLoadAddress + mi.m_dwSize))
		{
			*ppModule = module.detach();
			return S_OK;
		}
	}

	return E_NO_MODULE_AT_THIS_ADDRESS;
}
