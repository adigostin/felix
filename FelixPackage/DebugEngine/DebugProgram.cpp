
#include "pch.h"
#include "shared/com.h"
#include "shared/OtherGuids.h"
#include "shared/vector_nothrow.h"
#include "../guids.h"
#include "DebugEngine.h"
#include "DebugEventBase.h"
#include "../FelixPackage.h"

static const wchar_t SingleDebugProgramName[] = L"Z80 Program";

class DebugProgramProcessImpl : public IDebugProgram2, IDebugModuleCollection, ISimulatorEventHandler
{
	ULONG _refCount = 0;
	GUID _programId;
	com_ptr<IDebugProcess2> _process;
	vector_nothrow<com_ptr<IDebugModule2>> _modules;
	wil::com_ptr_nothrow<IDebugEngine2> _engine;
	wil::com_ptr_nothrow<IDebugEventCallback2> _callback;
	com_ptr<IDebugThread2> _thread;
	com_ptr<ISimulator> _simulator;
	bool _advisingSimulatorEvents = false;
	SIM_BP_COOKIE _stepOverOrOutBreakpoint = 0;

public:
	HRESULT InitInstance (IDebugProcess2* process, IDebugEngine2* engine, ISimulator* simulator, IDebugEventCallback2* callback)
	{
		HRESULT hr;

		RETURN_HR_IF_NULL(E_INVALIDARG, process);

		_process = process;

		hr = CoCreateGuid (&_programId); RETURN_IF_FAILED(hr);

		hr = MakeThread (engine, this, callback, &_thread); RETURN_IF_FAILED(hr);

		hr = simulator->AdviseDebugEvents(this); RETURN_IF_FAILED(hr);
		_advisingSimulatorEvents = true;

		_engine = engine;
		_callback = callback;
		_simulator = simulator;

		return S_OK;
	}

	#pragma region IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDebugProgram2*>(this), riid, ppvObject)
			|| TryQI<IDebugProgram2>(this, riid, ppvObject)
			|| TryQI<IDebugModuleCollection>(this, riid, ppvObject)
			|| TryQI<ISimulatorEventHandler>(this, riid, ppvObject)
		)
			return S_OK;
		
		#ifdef _DEBUG
		if (   riid == IID_IDebugProgramEngines2
			|| riid == IID_IDebugProgramEx2
			|| riid == IID_IDebugProgram3
			|| riid == IID_IDebugEngineProgram2
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
		*ppProcess = _process;
		_process->AddRef();
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE Terminate() override
	{
		if (_stepOverOrOutBreakpoint)
		{
			auto hr = _simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint); LOG_IF_FAILED(hr);
			_stepOverOrOutBreakpoint = 0;
		}

		if (_advisingSimulatorEvents)
		{
			_simulator->UnadviseDebugEvents(this);
			_advisingSimulatorEvents = false;
		}

		_engine = nullptr;
		_callback = nullptr;
		_simulator = nullptr;
		_thread = nullptr;
		_modules.clear();
		_process = nullptr;
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
		RETURN_HR(E_NOTIMPL);
	}

	HRESULT STDMETHODCALLTYPE GetProgramId(GUID* pguidProgramId) override
	{
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
		auto hr = _simulator->Resume(checkBreakpointsAtCurrentPC); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	
	HRESULT STDMETHODCALLTYPE Continue(IDebugThread2* pThread) override
	{
		// Despite the name, VS calls this function when the user starts debugging by clicking Start Debugging.
		// VS calls this function from its handler for the IDebugLoadCompleteEvent2 we sent.
		// (If the user stars debugging by clicking Step Into instead, VS does not call this function;
		// it assumes execution is stopped because we sent the Load Complete event as SYNC_STOP).

		bool checkBreakpointsAtCurrentPC = true;
		auto hr = _simulator->Resume(checkBreakpointsAtCurrentPC); RETURN_IF_FAILED(hr);
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
				hr = _simulator->SimulateOne(); RETURN_IF_FAILED(hr);
				return S_OK;
			}
			else if (sk == STEP_OVER)
			{
				if (_stepOverOrOutBreakpoint)
				{
					hr = _simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint); LOG_IF_FAILED(hr);
					_stepOverOrOutBreakpoint = 0;
				}

				uint16_t pc;
				hr = _simulator->GetPC(&pc); RETURN_IF_FAILED(hr);

				uint32_t len = 0;
///				TryDecodeBlockInstruction (pc, &len, nullptr); // return value ignored on purpose; we're looking at "len" instead.

				if (!len)
				{
					uint8_t buffer[4];
					hr = _simulator->ReadMemoryBus(pc, sizeof(buffer), buffer); RETURN_IF_FAILED(hr);

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
				}

				if (len)
				{
					hr = _simulator->AddBreakpoint (BreakpointType::Code, false, pc + len, &_stepOverOrOutBreakpoint); RETURN_IF_FAILED(hr);
					auto clearBP = wil::scope_exit([this]
						{ 
							_simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint);
							_stepOverOrOutBreakpoint = 0;
						});

					bool check_breakpoints = false;
					hr = _simulator->Resume(check_breakpoints); RETURN_IF_FAILED(hr);

					clearBP.release();
					return S_OK;
				}
				else
				{
					auto hr = _simulator->SimulateOne(); RETURN_IF_FAILED(hr);
					return S_OK;
				}
			}
			else
				RETURN_HR(E_NOTIMPL);
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	HRESULT STDMETHODCALLTYPE CauseBreak(void) override
	{
		RETURN_HR(E_NOTIMPL);
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

		wil::unique_bstr projectDir;
		if (auto op = wil::try_com_query_nothrow<IFelixLaunchOptionsProvider>(_engine))
		{
			com_ptr<IFelixLaunchOptions> options;
			if (SUCCEEDED(op->GetLaunchOptions(&options)))
				options->get_ProjectDir(&projectDir);
		}

		TEXT_POSITION begin, end;
		hr = pDocPos->GetRange (&begin, &end); RETURN_IF_FAILED(hr);

		UINT64 address;
		hr = ::GetAddressFromSourceLocation (this, projectDir.get(), filePath.get(), begin.dwLine, &address); RETURN_IF_FAILED(hr);

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

	#pragma region IDebugModuleCollection
	virtual HRESULT STDMETHODCALLTYPE AddModule (IDebugModule2* module) override
	{
		bool pushed = _modules.try_push_back(module); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

		wil::com_ptr_nothrow<ModuleLoadEvent> event;
		auto hr = ModuleLoadEvent::CreateInstance(module, TRUE, &event); RETURN_IF_FAILED(hr);
		hr = event->Send (_callback.get(), _engine.get(), this, _thread.get()); LOG_IF_FAILED(hr);

		// This is a best attempt at binding the breakpoints. We want to ignore the returned error code here.
///		for (auto& pb : _breakpoints)
///		{
///			wil::com_ptr_nothrow<IEnumDebugErrorBreakpoints2> errBPs;
///			if (SUCCEEDED(pb->EnumErrorBreakpoints(BPET_ALL, &errBPs)))
///			{
///				ULONG count;
///				if (SUCCEEDED(errBPs->GetCount(&count)) && count)
///					pb->Bind();
///			}
///		}

		return S_OK;
	}
	#pragma endregion

	HRESULT SendStepCompleteEvent()
	{
		// Must be stopping as per https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/supported-event-types?view=vs-2019
		using StepCompleteEvent = EventBase<IDebugStepCompleteEvent2, EVENT_SYNC_STOP>;
		com_ptr<StepCompleteEvent> e = new (std::nothrow) StepCompleteEvent(); RETURN_IF_NULL_ALLOC(e);
		auto hr = e->Send (_callback.get(), _engine.get(), this, _thread.get()); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region ISimulatorEventHandler
	virtual HRESULT STDMETHODCALLTYPE ProcessSimulatorEvent (ISimulatorEvent* event, REFIID riidEvent) override
	{
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
					if (bp == _stepOverOrOutBreakpoint)
					{
						auto hr = _simulator->RemoveBreakpoint(_stepOverOrOutBreakpoint); LOG_IF_FAILED(hr);
						_stepOverOrOutBreakpoint = 0;
						return SendStepCompleteEvent();
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
};

HRESULT MakeDebugProgram (IDebugProcess2* process, IDebugEngine2* engine, ISimulator* simulator, IDebugEventCallback2* callback, IDebugProgram2** ppProgram)
{
	auto p = com_ptr(new (std::nothrow) DebugProgramProcessImpl()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance (process, engine, simulator, callback); RETURN_IF_FAILED(hr);
	*ppProgram = p.detach();
	return S_OK;
}
