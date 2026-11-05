
#include "pch.h"
#include "shared/com.h"
#include "shared/OtherGuids.h"
#include "shared/vector_nothrow.h"
#include "shared/unordered_map_nothrow.h"
#include "../FelixPackageUi/resource.h"
#include "DebugEngine.h"
#include "DebugEventBase.h"
#include "../FelixPackage.h"

static const wchar_t SingleDebugProgramName[] = L"Z80 Program";

class DebugProgramImpl : public IFelixProgram, IDebugProgramNode2, ISimulatorEventNotifySink
{
	ULONG _refCount = 0;
	WeakRefToThis _weakRefToThis;
	GUID _programId;
	com_ptr<IWeakRef> _engine;
	com_ptr<IWeakRef> _process;
	vector_nothrow<com_ptr<IDebugModule2>> _modules;
	com_ptr<IDebugEventCallback2> _callback;
	com_ptr<IDebugThread2> _thread;
	AdviseSinkToken _eventNotifyToken;
	SIM_BP_COOKIE _stepOverOrOutBreakpoint = 0;
	unordered_map_nothrow<SIM_BP_COOKIE, com_ptr<IDebugBoundBreakpoint2>> _bps;

public:
	HRESULT InitInstance (IDebugProcess2* process)
	{
		HRESULT hr;
		hr = _weakRefToThis.InitInstance(static_cast<IDebugProgram2*>(this)); RETURN_IF_FAILED(hr);
		hr = CoCreateGuid(&_programId); RETURN_IF_FAILED(hr);
		hr = process->QueryInterface(&_process); RETURN_IF_FAILED(hr);
		hr = MakeThread(this, &_thread); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	~DebugProgramImpl()
	{
		WI_ASSERT(_bps.empty());
	}

	IUnknown* AsUnknown() { return static_cast<IDebugProgram2*>(this); }

	#pragma region IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(AsUnknown(), riid, ppvObject)
			|| TryQI<IDebugProgram2>(this, riid, ppvObject)
			|| TryQI<IDebugProgramNode2>(this, riid, ppvObject)
			|| TryQI<IFelixProgram>(this, riid, ppvObject)
			|| TryQI<ISimulatorEventNotifySink>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

		#ifdef _DEBUG
		if (   /*riid == IID_IDebugProgramEngines2
			|| */riid == IID_IDebugProgramEx2
			|| riid == IID_IDebugProgram3
			|| riid == IID_IDebugDisconnectableProgram166
			|| riid == IID_IDebugProgramEnhancedStep90
			|| riid == IID_IClientSecurity
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IdentityUnmarshal
			|| riid == IID_IStdMarshalInfo
			|| riid == IID_IFastRundown
			|| riid == IID_INoIdea7
			|| riid == IID_INotARealInterface
			|| riid == IID_INoIdea9
			|| riid == IID_INoIdea10
			|| riid == IID_INoIdea11
			|| riid == IID_INoIdea12
			|| riid == IID_INoIdea14
			|| riid == IID_INoIdea15
			|| riid == IID_INoIdea20
			|| riid == IID_IExternalConnection
			|| riid == IID_IDebugHistoricalProgram156
			|| riid == IID_IDebugReversibleEngineProgram160
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

	#pragma region IDebugProgram2
	HRESULT STDMETHODCALLTYPE EnumThreads(IEnumDebugThreads2** ppEnum) override
	{
		return make_single_entry_enumerator<IEnumDebugThreads2, IDebugThread2>(_thread, ppEnum);
	}

	HRESULT STDMETHODCALLTYPE GetName(BSTR* pbstrName) override
	{
		*pbstrName = SysAllocString (SingleDebugProgramName); RETURN_IF_NULL_ALLOC(*pbstrName);
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetProcess(IDebugProcess2** ppProcess) override
	{
		return _process->QueryInterface(ppProcess);
	}

	HRESULT STDMETHODCALLTYPE Terminate() override
	{
		WI_ASSERT(_callback && _eventNotifyToken);
		if (_stepOverOrOutBreakpoint)
		{
			auto hr = simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint); LOG_IF_FAILED(hr);
			_stepOverOrOutBreakpoint = 0;
		}

		_eventNotifyToken.reset();
		_callback.reset();
		_engine.reset();
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE Attach(IDebugEventCallback2* pCallback) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	HRESULT STDMETHODCALLTYPE CanDetach(void) override
	{
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE Detach(void) override
	{
		com_ptr<IDebugEngine2> engine;
		auto hr = _engine->QueryInterface(&engine); RETURN_IF_FAILED(hr);
		hr = Terminate(); RETURN_IF_FAILED(hr);
		return engine->DestroyProgram(this);
	}

	HRESULT STDMETHODCALLTYPE GetProgramId(GUID* pguidProgramId) override
	{
		// TODO: A debug engine (DE) must return the program identifier originally passed to the OnAttach or Attach methods. This allows identification of the program across debugger components.
		*pguidProgramId = _programId;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetDebugProperty(IDebugProperty2** ppProperty) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	HRESULT STDMETHODCALLTYPE Execute(void) override
	{
		// Despite the name, VS calls Execute() to resume execution of a program
		// after the program breaks into the debugger and the user clicks Continue.
		bool checkBreakpointsAtCurrentPC = false;
		auto hr = simulator->Resume(checkBreakpointsAtCurrentPC); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE Continue(IDebugThread2* pThread) override
	{
		// Despite the name, VS calls this function when the user starts debugging by clicking Start Debugging.
		// VS calls this function from its handler for the IDebugLoadCompleteEvent2 we sent.
		// (If the user stars debugging by clicking Step Into instead, VS does not call this function;
		// it assumes execution is stopped because we sent the Load Complete event as SYNC_STOP).

		if (simulator->Running_HR() == S_OK)
			return S_OK;
		bool checkBreakpointsAtCurrentPC = true;
		auto hr = simulator->Resume(checkBreakpointsAtCurrentPC); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE Step(IDebugThread2* pThread, STEPKIND sk, STEPUNIT step) override
	{
		HRESULT hr;

		if (step == STEP_INSTRUCTION
			|| step == STEP_STATEMENT)
		{
			if (sk == STEP_INTO)
			{
				hr = simulator->SimulateOne(); RETURN_IF_FAILED(hr);
				return S_OK;
			}
			else if (sk == STEP_OVER)
			{
				if (_stepOverOrOutBreakpoint)
				{
					hr = simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint); LOG_IF_FAILED(hr);
					_stepOverOrOutBreakpoint = 0;
				}

				uint16_t pc;
				hr = simulator->GetPC(&pc); RETURN_IF_FAILED(hr);

				uint32_t len = 0;
///				TryDecodeBlockInstruction (pc, &len, nullptr); // return value ignored on purpose; we're looking at "len" instead.

				if (!len)
				{
					uint8_t buffer[4];
					hr = simulator->ReadMemoryBus(pc, sizeof(buffer), buffer); RETURN_IF_FAILED(hr);

					if (buffer[0] == 0xCD // CALL nn
						||  (buffer[0] & 0xC7) == 0xC4) // CALL cc, nn
					{
						len = 3;
					}
					else if ((buffer[0] & 0xC7) == 0xC7) // RST p
					{
						len = 1;
					}
					else if ((buffer[0] == 0xED && buffer[1] == 0xB0) // LDIR
						|| (buffer[0] == 0xED && buffer[1] == 0xB8) // LDDR
						|| (buffer[0] == 0xED && buffer[1] == 0xB1) // CPIR
						|| (buffer[0] == 0xED && buffer[1] == 0xB9)) // CPDR
					{
						len = 2;
					}
					else if (buffer[0] == 0x76) // HALT
					{
						len = 1;
					}
				}

				if (len)
				{
					hr = simulator->AddBreakpoint (BreakpointType::Code, false, pc + len, &_stepOverOrOutBreakpoint); RETURN_IF_FAILED(hr);
					auto clearBP = wil::scope_exit([this]
						{
							simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint);
							_stepOverOrOutBreakpoint = 0;
						});

					bool check_breakpoints = false;
					hr = simulator->Resume(check_breakpoints); RETURN_IF_FAILED(hr);

					clearBP.release();
					return S_OK;
				}
				else
				{
					auto hr = simulator->SimulateOne(); RETURN_IF_FAILED(hr);
					return S_OK;
				}
			}
			else if (sk == STEP_OUT)
			{
				SetFelixErrorInfo(E_NOTIMPL, IDS_STEP_OUT_NOT_IMPL);
				uiShell->ReportErrorInfo(E_NOTIMPL);
				SendStepCompleteEvent();
				return S_OK;
			}
			else
				RETURN_HR(E_NOTIMPL);
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	HRESULT STDMETHODCALLTYPE CauseBreak() override
	{
		return simulator->Break();
	}

	HRESULT STDMETHODCALLTYPE GetEngineInfo(BSTR* pbstrEngine, GUID* pguidEngine) override
	{
		if (pbstrEngine)
			*pbstrEngine = SysAllocString(L"Z80 Debug Engine");
		if (pguidEngine)
			*pguidEngine = Engine_Id;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE EnumCodeContexts(IDebugDocumentPosition2* pDocPos, IEnumDebugCodeContexts2** ppEnum) override
	{
		wil::unique_bstr filePath;
		auto hr = pDocPos->GetFileName(&filePath); RETURN_IF_FAILED(hr);

		TEXT_POSITION begin, end;
		hr = pDocPos->GetRange (&begin, &end); RETURN_IF_FAILED(hr);

		UINT64 address;
		hr = ::GetAddressFromSourceLocation (this, filePath.get(), begin.dwLine, &address); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IDebugCodeContext2> cc;
		hr = MakeDebugContext (false, address, this, IID_PPV_ARGS(&cc)); RETURN_IF_FAILED(hr);

		hr = make_single_entry_enumerator<IEnumDebugCodeContexts2, IDebugCodeContext2>(cc.get(), ppEnum); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetMemoryBytes(IDebugMemoryBytes2** ppMemoryBytes) override
	{
		return MakeMemoryBytes (ppMemoryBytes);
	}

	HRESULT STDMETHODCALLTYPE GetDisassemblyStream(DISASSEMBLY_STREAM_SCOPE dwScope, IDebugCodeContext2* pCodeContext, IDebugDisassemblyStream2** ppDisassemblyStream) override
	{
		return MakeDisassemblyStream (dwScope, this, pCodeContext, ppDisassemblyStream);
	}

	HRESULT STDMETHODCALLTYPE EnumModules(IEnumDebugModules2** ppEnum) override
	{
		return enumerator<IEnumDebugModules2, IDebugModule2>::create_instance(_modules, ppEnum);
	}

	HRESULT STDMETHODCALLTYPE GetENCUpdate(IDebugENCUpdate** ppUpdate) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	HRESULT STDMETHODCALLTYPE EnumCodePaths(LPCOLESTR pszHint, IDebugCodeContext2* pStart, IDebugStackFrame2* pFrame, BOOL fSource, IEnumCodePaths2** ppEnum, IDebugCodeContext2** ppSafety) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE WriteDump(DUMPTYPE DumpType, LPCOLESTR pszDumpUrl) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IDebugProgramNode2
	virtual HRESULT STDMETHODCALLTYPE GetProgramName (BSTR *pbstrProgramName) override
	{
		RETURN_HR_IF_NULL(E_POINTER, pbstrProgramName);
		*pbstrProgramName = SysAllocString(SingleDebugProgramName); RETURN_IF_NULL_ALLOC(*pbstrProgramName);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetHostName (GETHOSTNAME_TYPE dwHostNameType, BSTR *pbstrHostName) override
	{
		RETURN_HR_IF_NULL(E_POINTER, pbstrHostName);
		*pbstrHostName = nullptr;
		if (dwHostNameType == GHN_FILE_NAME)
		{
			RETURN_HR(E_NOTIMPL);
			//wchar_t path[MAX_PATH];
			//DWORD length = GetModuleFileNameW(nullptr, path, ARRAYSIZE(path));
			//RETURN_LAST_ERROR_IF(length == 0 || length == ARRAYSIZE(path));
			//*pbstrHostName = SysAllocString(path); RETURN_IF_NULL_ALLOC(*pbstrHostName);
		}
		else
		{
			*pbstrHostName = SysAllocString(L"Visual Studio"); RETURN_IF_NULL_ALLOC(*pbstrHostName);
		}
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetHostPid (AD_PROCESS_ID *pHostProcessId) override
	{
		RETURN_HR_IF_NULL(E_POINTER, pHostProcessId);
		com_ptr<IDebugProcess2> process;
		auto hr = GetProcess(&process); RETURN_IF_FAILED(hr);
		return process->GetPhysicalProcessId(pHostProcessId);
	}

	virtual HRESULT STDMETHODCALLTYPE GetHostMachineName_V7 (BSTR *pbstrHostMachineName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Attach_V7 (IDebugProgram2 *pMDMProgram, IDebugEventCallback2 *pCallback, DWORD dwReason) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	//virtual HRESULT STDMETHODCALLTYPE GetEngineInfo (BSTR *pbstrEngine, GUID *pguidEngine) override { RETURN_HR(E_NOTIMPL); }

	virtual HRESULT STDMETHODCALLTYPE DetachDebugger_V7() override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	struct ModuleLoadEvent : public EventBase<IDebugModuleLoadEvent2, EVENT_ASYNCHRONOUS>
	{
		wil::com_ptr_nothrow<IDebugModule2> _module;
		BOOL _load;

		static HRESULT CreateInstance (IDebugModule2* module, BOOL load, ModuleLoadEvent** to)
		{
			wil::com_ptr_nothrow<ModuleLoadEvent> p = new (std::nothrow) ModuleLoadEvent(); RETURN_IF_NULL_ALLOC(p);
			p->_module = module;
			p->_load = load;
			*to = p.detach();
			return S_OK;
		}

		virtual HRESULT __stdcall GetModule(IDebugModule2** pModule, BSTR* pbstrDebugMessage, BOOL* pbLoad) override
		{
			if (pModule)
			{
				*pModule = _module.get();
				(*pModule)->AddRef();
			}

			if (pbstrDebugMessage)
				*pbstrDebugMessage = nullptr;

			if (pbLoad)
				*pbLoad = _load;

			return S_OK;
		}
	};

	#pragma region IFelixProgram
	virtual HRESULT STDMETHODCALLTYPE AddModule (IDebugModule2* module) override
	{
		bool pushed = _modules.try_push_back(module); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return NotifyModule(module, TRUE);
	}

	virtual HRESULT STDMETHODCALLTYPE UnloadModules() override
	{
		for (auto& module : _modules)
		{
			auto hr = NotifyModule(module.get(), FALSE); RETURN_IF_FAILED(hr);
		}
		_modules.clear();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE AddBreakpoint (IDebugBoundBreakpoint2* bp, BreakpointType type, bool physicalMemorySpace, UINT64 address) override
	{
		for (auto& existing : _bps)
			RETURN_HR_IF(E_INVALIDARG, existing.second == bp);

		bool reserved = _bps.try_reserve(_bps.size() + 1); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);
		SIM_BP_COOKIE cookie;
		auto hr = simulator->AddBreakpoint(type, physicalMemorySpace, address, &cookie); RETURN_IF_FAILED(hr);
		bool inserted = _bps.try_insert({ cookie, bp }); WI_ASSERT(inserted);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveBreakpoint (IDebugBoundBreakpoint2* bp) override
	{
		auto it = _bps.find_if([bp](auto& p) { return p.second == bp; }); RETURN_HR_IF(E_INVALIDARG, it == _bps.end());
		auto hr = simulator->RemoveBreakpoint(it->first); LOG_IF_FAILED(hr);
		_bps.erase(it);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE ContainsBreakpoint (IDebugBoundBreakpoint2* bp) override
	{
		auto it = _bps.find_if([bp](auto& p) { return p.second == bp; });
		return (it != _bps.end()) ? S_OK : S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE EngineAttach (IDebugEngine2* pEngine, IDebugEventCallback2* pCallback, ATTACH_REASON dwReason) override
	{
		HRESULT hr;
		WI_ASSERT(!_callback && !_eventNotifyToken);
		hr = pEngine->QueryInterface(IID_PPV_ARGS(&_engine)); RETURN_IF_FAILED(hr);
		_callback = pCallback;
		hr = AdviseSink<ISimulatorEventNotifySink>(simulator, AsUnknown(), &_eventNotifyToken); RETURN_IF_FAILED(hr);
		auto terminateOnFailure = wil::scope_exit([this] { Terminate(); });
		for (auto& module : _modules)
		{
			hr = NotifyModule(module.get(), TRUE); RETURN_IF_FAILED(hr);
		}
		terminateOnFailure.release();
		return S_OK;
	}
	#pragma endregion

	HRESULT NotifyModule(IDebugModule2* module, BOOL load)
	{
		if (!_callback)
			return S_OK;

		com_ptr<IDebugEventCallback2> callback;
		auto hr = _callback->QueryInterface(&callback); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<ModuleLoadEvent> event;
		hr = ModuleLoadEvent::CreateInstance(module, load, &event); RETURN_IF_FAILED(hr);
		hr = event->Send (callback.get(), nullptr, this, _thread.get()); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	HRESULT SendStepCompleteEvent()
	{
		// Must be stopping as per https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/supported-event-types?view=vs-2019
		using StepCompleteEvent = EventBase<IDebugStepCompleteEvent2, EVENT_SYNC_STOP>;
		com_ptr<StepCompleteEvent> e = new (std::nothrow) StepCompleteEvent(); RETURN_IF_NULL_ALLOC(e);
		com_ptr<IDebugEventCallback2> callback;
		auto hr = _callback->QueryInterface(&callback); RETURN_IF_FAILED(hr);
		hr = e->Send (callback.get(), nullptr, this, _thread.get()); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region ISimulatorEventNotifySink
	virtual HRESULT STDMETHODCALLTYPE NotifySimulatorEvent (ISimulatorEvent* event, REFIID riidEvent) override
	{
		HRESULT hr;

		if (riidEvent == __uuidof(ISimulatorBreakpointEvent))
		{
			com_ptr<ISimulatorBreakpointEvent> bpe;
			hr = event->QueryInterface(&bpe); RETURN_IF_FAILED(hr);
			vector_nothrow<com_ptr<IDebugBoundBreakpoint2>> bps;
			bool stepComplete = false;
			for (ULONG i = 0; i < bpe->GetBreakpointCount(); i++)
			{
				SIM_BP_COOKIE bp;
				auto hr = bpe->GetBreakpointAt(i, &bp); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr))
				{
					if (bp == _stepOverOrOutBreakpoint)
					{
						hr = simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint); LOG_IF_FAILED(hr);
						_stepOverOrOutBreakpoint = 0;
						stepComplete = true;
					}
					if (auto it = _bps.find(bp); it != _bps.end())
					{
						bool pushed = bps.try_push_back(it->second); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
					}
				}
			}

			if (stepComplete)
			{
				hr = SendStepCompleteEvent(); RETURN_IF_FAILED(hr);
			}

			if (!bps.empty())
			{
				struct BreakpointEvent : EventBase<IDebugBreakpointEvent2, EVENT_SYNC_STOP>
				{
					com_ptr<IEnumDebugBoundBreakpoints2> _boundBPs;

					virtual HRESULT STDMETHODCALLTYPE EnumBreakpoints (IEnumDebugBoundBreakpoints2** ppEnum) override
					{
						return _boundBPs.copy_to(ppEnum);
					}
				};

				auto breakpointEvent = com_ptr(new (std::nothrow) BreakpointEvent()); RETURN_IF_NULL_ALLOC(breakpointEvent);
				hr = enumerator<IEnumDebugBoundBreakpoints2, IDebugBoundBreakpoint2>::create_instance(std::move(bps), &breakpointEvent->_boundBPs); RETURN_IF_FAILED(hr);
				hr = breakpointEvent->Send(_callback, nullptr, this, _thread); RETURN_IF_FAILED(hr);
			}

			return S_OK;
		}

		if (riidEvent == __uuidof(ISimulatorBreakEvent))
		{
			using BreakEvent = EventBase<IDebugBreakEvent2, EVENT_SYNC_STOP>;
			com_ptr<BreakEvent> e = new (std::nothrow) BreakEvent(); RETURN_IF_NULL_ALLOC(e);
			com_ptr<IDebugEventCallback2> callback;
			auto hr = _callback->QueryInterface(&callback); RETURN_IF_FAILED(hr);
			hr = e->Send (callback, nullptr, this, _thread); RETURN_IF_FAILED(hr);
			return S_OK;
		}

		if (riidEvent == __uuidof(ISimulatorSimulateOneEvent))
			return SendStepCompleteEvent();

		return S_FALSE;
	}
	#pragma endregion
};

HRESULT MakeDebugProgram (IDebugProcess2* process, IDebugProgram2** ppProgram)
{
	auto p = com_ptr(new (std::nothrow) DebugProgramImpl()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance (process); RETURN_IF_FAILED(hr);
	*ppProgram = p.detach();
	return S_OK;
}
