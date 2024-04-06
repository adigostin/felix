
#include "pch.h"
#include "DebugEngine.h"
#include "DebugEventBase.h"
#include "../FelixPackage.h"
#include "shared/com.h"
#include "shared/unordered_map_nothrow.h"

// https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/binding-breakpoints?view=vs-2022

class BreakpointManagerImpl : public IBreakpointManager, ISimulatorEventHandler
{
	ULONG _refCount = 0;
	com_ptr<IDebugEventCallback2> _callback;
	com_ptr<IDebugEngine2> _engine;
	com_ptr<IDebugProgram2> _program;
	com_ptr<ISimulator> _simulator;
	unordered_map_nothrow<SIM_BP_COOKIE, com_ptr<IDebugBoundBreakpoint2>> _bps;

public:
	static HRESULT CreateInstance (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, ISimulator* simulator, IBreakpointManager** ppManager)
	{
		auto p = com_ptr(new (std::nothrow) BreakpointManagerImpl()); RETURN_IF_NULL_ALLOC(p);
		p->_callback = callback;
		p->_engine = engine;
		p->_program = program;
		p->_simulator = simulator;
		*ppManager = p.detach();
		return S_OK;
	}

	~BreakpointManagerImpl()
	{
	}

	#pragma region IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);

		if (   TryQI<IUnknown>(static_cast<IBreakpointManager*>(this), riid, ppvObject)
			|| TryQI<IBreakpointManager>(this, riid, ppvObject)
			|| TryQI<ISimulatorEventHandler>(this, riid, ppvObject))
			return S_OK;

		*ppvObject = nullptr;
		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region ISimulatorEventHandler
	virtual HRESULT STDMETHODCALLTYPE ProcessSimulatorEvent (ISimulatorEvent* event, REFIID riidEvent) override
	{
		if (riidEvent == __uuidof(ISimulatorBreakpointEvent))
		{
			com_ptr<ISimulatorBreakpointEvent> bpe;
			auto hr = event->QueryInterface(&bpe); RETURN_IF_FAILED(hr);
			ULONG count = bpe->GetBreakpointCount();
			vector_nothrow<com_ptr<IDebugBoundBreakpoint2>> bps;
			for (ULONG i = 0; i < count; i++)
			{
				SIM_BP_COOKIE hitBp;
				hr = bpe->GetBreakpointAt(i, &hitBp); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr))
				{
					if (auto it = _bps.find(hitBp); it != _bps.end())
					{
						bool pushed = bps.try_push_back(it->second);  RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
					}
				}
			}

			if (bps.size())
			{
				wil::com_ptr_nothrow<IEnumDebugBoundBreakpoints2> boundBPsEnum;
				auto hr = enumerator<IEnumDebugBoundBreakpoints2, IDebugBoundBreakpoint2>::create_instance(std::move(bps), &boundBPsEnum); RETURN_IF_FAILED(hr);

				struct BreakpointEvent : public EventBase<IDebugBreakpointEvent2, EVENT_SYNC_STOP>
				{
					wil::com_ptr_nothrow<IEnumDebugBoundBreakpoints2> _boundBPs;

					virtual HRESULT STDMETHODCALLTYPE EnumBreakpoints (IEnumDebugBoundBreakpoints2 **ppEnum) override
					{
						auto hr = wil::com_copy_to_nothrow(_boundBPs, ppEnum); RETURN_IF_FAILED(hr);
						return S_OK;
					}
				};

				wil::com_ptr_nothrow<BreakpointEvent> event = new (std::nothrow) BreakpointEvent(); RETURN_IF_NULL_ALLOC(event);
				event->_boundBPs = std::move(boundBPsEnum);
				com_ptr<IEnumDebugThreads2> threads;
				hr = _program->EnumThreads(&threads); RETURN_IF_FAILED(hr);
				com_ptr<IDebugThread2> thread;
				hr = threads->Next(1, &thread, nullptr); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_FAIL, hr != S_OK);
				hr = event->Send (_callback, _engine, _program, thread); RETURN_IF_FAILED(hr);
			}

			return S_OK;
		}

		return S_FALSE;
	}
	#pragma endregion

	#pragma region IBreakpointManager
	virtual HRESULT AddBreakpoint (IDebugBoundBreakpoint2* bp, BreakpointType type, bool physicalMemorySpace, UINT64 address) override
	{
		HRESULT hr;

		for (auto& existing : _bps)
			RETURN_HR_IF(E_INVALIDARG, existing.second == bp);

		if (_bps.empty())
		{
			hr = _simulator->AdviseDebugEvents(this); RETURN_IF_FAILED(hr);
		}

		bool reserved = _bps.try_reserve(_bps.size() + 1);
		if (!reserved)
		{
			if (_bps.empty())
				_simulator->UnadviseDebugEvents(this);
			RETURN_HR(E_OUTOFMEMORY);
		}

		SIM_BP_COOKIE cookie;
		hr = _simulator->AddBreakpoint (type, physicalMemorySpace, address, &cookie);
		if (FAILED(hr))
		{
			if (_bps.empty())
				_simulator->UnadviseDebugEvents(this);
			RETURN_HR(hr);
		}

		bool inserted = _bps.try_insert({ cookie, bp }); WI_ASSERT(inserted);
		return S_OK;
	}

	virtual HRESULT RemoveBreakpoint (IDebugBoundBreakpoint2* bp) override
	{
		auto it = _bps.find_if([bp](auto& p) { return p.second == bp; }); RETURN_HR_IF(E_INVALIDARG, it == _bps.end());
		auto hr = _simulator->RemoveBreakpoint(it->first); LOG_IF_FAILED(hr);
		_bps.erase(it);

		if (_bps.empty())
		{
			hr = _simulator->UnadviseDebugEvents(this); LOG_IF_FAILED(hr);
		}

		return S_OK;
	}
	
	virtual HRESULT ContainsBreakpoint (IDebugBoundBreakpoint2* bp) override
	{
		auto it = _bps.find_if([bp](auto& p) { return p.second == bp; });
		return (it != _bps.end()) ? S_OK : S_FALSE;
	}
	#pragma endregion
};

HRESULT MakeBreakpointManager (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, ISimulator* simulator, IBreakpointManager** ppManager)
{
	return BreakpointManagerImpl::CreateInstance(callback, engine, program, simulator, ppManager);
};

// ============================================================================

class ErrorBreakpointResolution : IDebugErrorBreakpointResolution2
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IDebugProgram2> _program;
	BP_TYPE _bp_type;
	BP_ERROR_TYPE _errorType;

public:
	static HRESULT CreateInstance (IDebugProgram2* program, BP_TYPE bp_type, BP_ERROR_TYPE errorType, IDebugErrorBreakpointResolution2** to)
	{
		wil::com_ptr_nothrow<ErrorBreakpointResolution> p = new (std::nothrow) ErrorBreakpointResolution(); RETURN_IF_NULL_ALLOC(p);

		p->_program = program;
		p->_bp_type = bp_type;
		p->_errorType = errorType;

		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_INVALIDARG;

		*ppvObject = NULL;

		if (riid == IID_IUnknown)
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IDebugErrorBreakpointResolution2)
		{
			*ppvObject = static_cast<IDebugErrorBreakpointResolution2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugErrorBreakpointResolution2
	virtual HRESULT STDMETHODCALLTYPE GetBreakpointType (__RPC__out BP_TYPE *pBPType) override
	{
		*pBPType = _bp_type;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetResolutionInfo (BPERESI_FIELDS dwFields, __RPC__out BP_ERROR_RESOLUTION_INFO *pErrorResolutionInfo) override
	{
		if (dwFields & BPERESI_MESSAGE)
		{
			// TODO: free if this function fails
			pErrorResolutionInfo->bstrMessage = SysAllocString(L"IDebugErrorBreakpointResolution2 message"); RETURN_IF_NULL_ALLOC(pErrorResolutionInfo->bstrMessage);
			pErrorResolutionInfo->dwFields |= BPERESI_MESSAGE;
			dwFields &= ~BPERESI_MESSAGE;
		}

		if (dwFields & BPERESI_TYPE)
		{
			pErrorResolutionInfo->dwType = _errorType;
			pErrorResolutionInfo->dwFields |= BPERESI_TYPE;
			dwFields &= ~BPERESI_TYPE;
		}

		dwFields &= ~0x20; // BPERESI174_SUGGESTEDFIX

		LOG_HR_IF (E_NOTIMPL, !!dwFields);

		return S_OK;
	}
	#pragma endregion
};

// ============================================================================

struct BoundBreakpointImpl : IDebugBoundBreakpoint2, IDebugBreakpointResolution2
{
	ULONG _refCount = 0;
	com_ptr<IDebugPendingBreakpoint2> _parent;
	com_ptr<IDebugProgram2> _program;
	com_ptr<IBreakpointManager> _bpman;
	bool _physicalMemorySpace;
	UINT64 _address;
	DWORD _hitCount = 0;
	enum_BP_STATE _state = BPS_NONE;

	static HRESULT CreateInstance (IDebugPendingBreakpoint2* parent, IDebugProgram2* program, IBreakpointManager* bpman,
		bool physicalMemorySpace, UINT64 address, IDebugBoundBreakpoint2** to)
	{
		auto p = wil::com_ptr_nothrow(new (std::nothrow) BoundBreakpointImpl()); RETURN_IF_NULL_ALLOC(p);

		p->_parent = parent;
		p->_program = program;
		p->_bpman = bpman;
		p->_physicalMemorySpace = physicalMemorySpace;
		p->_address = address;

		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);

		if (TryQI<IUnknown>(static_cast<IDebugBoundBreakpoint2*>(this), riid, ppvObject)
			|| TryQI<IDebugBoundBreakpoint2>(this, riid, ppvObject)
			|| TryQI<IDebugBreakpointResolution2>(this, riid, ppvObject))
			return S_OK;

		*ppvObject = NULL;
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugBoundBreakpoint2
	virtual HRESULT STDMETHODCALLTYPE GetPendingBreakpoint (IDebugPendingBreakpoint2 **ppPendingBreakpoint) override
	{
		*ppPendingBreakpoint = _parent.get();
		(*ppPendingBreakpoint)->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetState (BP_STATE *pState) override
	{
		*pState = _state;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetHitCount (DWORD *pdwHitCount) override
	{
		if (_state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		*pdwHitCount = _hitCount;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetBreakpointResolution (IDebugBreakpointResolution2 **ppBPResolution) override
	{
		RETURN_HR_IF(E_BP_DELETED, _state == PBPS_DELETED);
		*ppBPResolution = this;
		AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Enable (BOOL fEnable) override
	{
		HRESULT hr;
		if (_state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		if (fEnable)
		{
			if (_state != BPS_ENABLED)
			{
				hr = _bpman->AddBreakpoint(this, BreakpointType::Code, _physicalMemorySpace, _address); RETURN_IF_FAILED(hr);
			}

			_state = BPS_ENABLED;
			return S_OK;
		}
		else if (!fEnable)
		{
			if (_state == BPS_ENABLED)
			{
				hr = _bpman->RemoveBreakpoint(this); RETURN_IF_FAILED(hr);
			}
			_state = BPS_DISABLED;
			return S_OK;
		}
		else
			RETURN_HR(S_FALSE);
	}

	virtual HRESULT STDMETHODCALLTYPE SetHitCount (DWORD dwHitCount) override
	{
		if (_state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		_hitCount = dwHitCount;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetCondition (BP_CONDITION bpCondition) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetPassCount (BP_PASSCOUNT bpPassCount) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Delete() override
	{
		if (_state == BPS_DELETED)
			return E_BP_DELETED;

		if (_state == BPS_ENABLED)
		{
			auto hr = _bpman->RemoveBreakpoint(this); LOG_IF_FAILED(hr);
			_parent = nullptr;
			_program = nullptr;
			_bpman = nullptr;
		}

		_state = BPS_DELETED;
		return S_OK;
	}
	#pragma endregion

	#pragma region IDebugBreakpointResolution2
	virtual HRESULT STDMETHODCALLTYPE GetBreakpointType (BP_TYPE *pBPType) noexcept override
	{
		*pBPType = BPT_CODE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetResolutionInfo (BPRESI_FIELDS dwFields, BP_RESOLUTION_INFO *pBPResolutionInfo) override
	{
		pBPResolutionInfo->dwFields = 0;

		if (dwFields & BPRESI_BPRESLOCATION)
		{
			pBPResolutionInfo->bpResLocation.bpType = BPT_CODE;
			auto hr = MakeDebugContext(false, _address, _program.get(), IID_PPV_ARGS(&pBPResolutionInfo->bpResLocation.bpResLocation.bpresCode.pCodeContext)); RETURN_IF_FAILED(hr);
			pBPResolutionInfo->dwFields |= BPRESI_BPRESLOCATION;
		}

		if (dwFields & BPRESI_PROGRAM)
		{
			pBPResolutionInfo->pProgram = _program.get();
			pBPResolutionInfo->pProgram->AddRef();
			pBPResolutionInfo->dwFields |= BPRESI_PROGRAM;
		}

		if (dwFields & BPRESI_THREAD)
		{
			com_ptr<IEnumDebugThreads2> threads;
			auto hr = _program->EnumThreads(&threads); RETURN_IF_FAILED(hr);
			com_ptr<IDebugThread2> thread;
			hr = threads->Next(1, &thread, nullptr); RETURN_IF_FAILED(hr);
			thread.copy_to(&pBPResolutionInfo->pThread);
			pBPResolutionInfo->dwFields |= BPRESI_THREAD;
		}

		return S_OK;
	}
	#pragma endregion
};

// ============================================================================

class ErrorBreakpoint : IDebugErrorBreakpoint2
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IDebugProgram2> _program;
	wil::com_ptr_nothrow<IDebugPendingBreakpoint2> _pending;
	BP_ERROR_TYPE _errorType;

public:
	static HRESULT CreateInstance (IDebugProgram2* program, IDebugPendingBreakpoint2* pending, BP_ERROR_TYPE errorType, IDebugErrorBreakpoint2** to)
	{
		wil::com_ptr_nothrow<ErrorBreakpoint> p = new (std::nothrow) ErrorBreakpoint(); RETURN_IF_NULL_ALLOC(p);

		p->_program = program;
		p->_pending = pending;
		p->_errorType = errorType;

		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_INVALIDARG;

		*ppvObject = NULL;

		if (riid == IID_IUnknown)
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IDebugErrorBreakpoint2)
		{
			*ppvObject = static_cast<IDebugErrorBreakpoint2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugErrorBreakpoint2
	virtual HRESULT STDMETHODCALLTYPE GetPendingBreakpoint (__RPC__deref_out_opt IDebugPendingBreakpoint2 **ppPendingBreakpoint) override
	{
		_pending.copy_to(ppPendingBreakpoint);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetBreakpointResolution (__RPC__deref_out_opt IDebugErrorBreakpointResolution2 **ppErrorResolution) override
	{
		wil::com_ptr_nothrow<IDebugBreakpointRequest2> bpreq;
		auto hr = _pending->GetBreakpointRequest(&bpreq); RETURN_IF_FAILED(hr);

		BP_LOCATION_TYPE loctype;
		hr = bpreq->GetLocationType(&loctype); RETURN_IF_FAILED(hr);

		BP_TYPE bp_type = (BP_TYPE)(loctype & BPLT_TYPE_MASK);
		hr = ErrorBreakpointResolution::CreateInstance (_program.get(), bp_type, _errorType, ppErrorResolution); RETURN_IF_FAILED(hr);

		return S_OK;
	}
	#pragma endregion
};

// ============================================================================

// https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/breakpoint-related-methods?view=vs-2022
struct BreakpointBoundEvent : EventBase<IDebugBreakpointBoundEvent2, EVENT_ASYNCHRONOUS>
{
	wil::com_ptr_nothrow<IDebugPendingBreakpoint2> _pending;
	wil::com_ptr_nothrow<IDebugBoundBreakpoint2> _bound;

	BreakpointBoundEvent (IDebugPendingBreakpoint2* pending, IDebugBoundBreakpoint2* bound)
		: _pending(pending), _bound(bound)
	{ }

	virtual HRESULT STDMETHODCALLTYPE GetPendingBreakpoint (IDebugPendingBreakpoint2 **ppPendingBP) override
	{
		*ppPendingBP = _pending.get();
		(*ppPendingBP)->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE EnumBoundBreakpoints (IEnumDebugBoundBreakpoints2 **ppEnum) override
	{
		auto hr = make_single_entry_enumerator<IEnumDebugBoundBreakpoints2, IDebugBoundBreakpoint2>(_bound.get(), ppEnum); RETURN_IF_FAILED(hr);
		return S_OK;
	}
};

struct BreakpointErrorEvent : EventBase<IDebugBreakpointErrorEvent2, EVENT_ASYNCHRONOUS>
{
	wil::com_ptr_nothrow<IDebugPendingBreakpoint2> _pending;
	wil::com_ptr_nothrow<IDebugErrorBreakpoint2> _err;

	BreakpointErrorEvent (IDebugPendingBreakpoint2* pending, IDebugErrorBreakpoint2* err)
		: _pending(pending), _err(err)
	{ }

	virtual HRESULT STDMETHODCALLTYPE GetErrorBreakpoint (__RPC__deref_out_opt IDebugErrorBreakpoint2 **ppErrorBP) override
	{
		_err.copy_to(ppErrorBP);
		return S_OK;
	}
};

// ============================================================================

struct SimplePendingBreakpoint : IDebugPendingBreakpoint2
{
	ULONG _refCount = 0;
	com_ptr<IDebugEventCallback2> _callback;
	com_ptr<IDebugEngine2> _engine;
	com_ptr<IDebugProgram2> _program;
	com_ptr<IBreakpointManager> _bpman;
	bool _physicalMemorySpace;
	UINT64 _address;
	PENDING_BP_STATE_INFO _stateInfo = { };
	wil::com_ptr_nothrow<IDebugBoundBreakpoint2> _boundBP;
	wil::com_ptr_nothrow<IDebugErrorBreakpoint2> _errorBP;

	static HRESULT CreateInstance (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, 
		IBreakpointManager* bpman, bool physicalMemorySpace, UINT64 address, IDebugPendingBreakpoint2** to)
	{
		wil::com_ptr_nothrow<SimplePendingBreakpoint> p = new (std::nothrow) SimplePendingBreakpoint(); RETURN_IF_NULL_ALLOC(p);
		p->_callback = callback;
		p->_engine = engine;
		p->_program = program;
		p->_bpman = bpman;
		p->_physicalMemorySpace = physicalMemorySpace;
		p->_address = address;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			RETURN_HR(E_POINTER);

		*ppvObject = NULL;

		if (riid == __uuidof(IUnknown))
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == __uuidof(IDebugPendingBreakpoint2))
		{
			*ppvObject = static_cast<IDebugPendingBreakpoint2*>(this);
			AddRef();
			return S_OK;
		}

		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugPendingBreakpoint2
	virtual HRESULT STDMETHODCALLTYPE CanBind (IEnumDebugErrorBreakpoints2 **ppErrorEnum) override
	{
		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Bind() override
	{
		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		if (_boundBP)
			RETURN_HR(S_FALSE);

		auto hr = BoundBreakpointImpl::CreateInstance (this, _program, _bpman, _physicalMemorySpace, _address, &_boundBP); RETURN_IF_FAILED(hr);
		
		if (_stateInfo.state == BPS_ENABLED)
			_boundBP->Enable(TRUE);
		
		auto p = wil::com_ptr_nothrow(new (std::nothrow) BreakpointBoundEvent (this, _boundBP.get()));
		if (!p)
			RETURN_HR(E_OUTOFMEMORY);
		hr = p->Send (_callback.get(), _engine.get(), _program.get(), nullptr); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetState (PENDING_BP_STATE_INFO* pState) override
	{
		*pState = _stateInfo;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetBreakpointRequest (IDebugBreakpointRequest2 **ppBPRequest) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Virtualize (BOOL fVirtualize) override
	{
		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		if (fVirtualize)
			_stateInfo.flags |= PBPSF_VIRTUALIZED;
		else
			_stateInfo.flags &= ~PBPSF_VIRTUALIZED;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Enable (BOOL fEnable) override
	{
		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		//if (_boundBP)
		//	_boundBP->Enable(fEnable);

		if (fEnable)
		{
			if (_stateInfo.state != PBPS_ENABLED)
			{
///				_stateInfo.state = PBPS_ENABLED;
///				_program->RegisterPendingBreakpoint(this);
			}
		}
		else
		{
///			if (_stateInfo.state == PBPS_ENABLED)
///				_program->UnregisterPendingBreakpoint(this);
///			_stateInfo.state = PBPS_DISABLED;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetCondition (BP_CONDITION bpCondition) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetPassCount (BP_PASSCOUNT bpPassCount) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumBoundBreakpoints (IEnumDebugBoundBreakpoints2 **ppEnum) override
	{
		return make_single_entry_enumerator<IEnumDebugBoundBreakpoints2, IDebugBoundBreakpoint2>(_boundBP.get(), ppEnum);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumErrorBreakpoints (BP_ERROR_TYPE bpErrorType, IEnumDebugErrorBreakpoints2 **ppEnum) override
	{
		RETURN_HR_IF(E_NOTIMPL, bpErrorType != BPET_ALL);
		auto hr = make_single_entry_enumerator<IEnumDebugErrorBreakpoints2, IDebugErrorBreakpoint2>(_errorBP.get(), ppEnum); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Delete() override
	{
		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

///		if (_stateInfo.state == PBPS_ENABLED)
///			_program->UnregisterPendingBreakpoint(this);

		if (_boundBP)
		{
			// Judging by the values in the BP_UNBOUND_REASON enum, a Breakpoint Unbound Event
			// need not be sent due to breakpoint deletion.
			//auto hr = SendBreakpointUnboundEvent (EVENT_SYNCHRONOUS, _callback, _engine, _boundBP, BPUR_UNKNOWN); RETURN_IF_FAILED(hr);

			BP_STATE s;
			if (SUCCEEDED(_boundBP->GetState(&s)) && (s != BPS_DELETED))
				_boundBP->Delete();
			_boundBP = nullptr;
		}

		_errorBP = nullptr;

		_stateInfo.state = PBPS_DELETED;
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeSimplePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program,
	IBreakpointManager* bpman, bool physicalMemorySpace, UINT64 address, IDebugPendingBreakpoint2** to)
{
	return SimplePendingBreakpoint::CreateInstance (callback, engine, program, bpman, physicalMemorySpace, address, to);
}

// ============================================================================

class SourceLinePendingBreakpoint : IDebugPendingBreakpoint2, IDebugEventCallback2
{
	ULONG _refCount = 0;
	com_ptr<IDebugEventCallback2> _callback;
	com_ptr<IDebugEngine2> _engine;
	com_ptr<IDebugProgram2> _program;
	com_ptr<IBreakpointManager> _bpman;
	wil::com_ptr_nothrow<IDebugBreakpointRequest2> _bp_request;
	wil::unique_process_heap_string _file;
	uint32_t _line_index;
	PENDING_BP_STATE_INFO _stateInfo = { };
	wil::com_ptr_nothrow<IDebugBoundBreakpoint2> _boundBP;
	wil::com_ptr_nothrow<IDebugErrorBreakpoint2> _errorBP;
	bool _advisedDebugEventCallback = false;
	wil::unique_bstr _projectDir;

public:
	static HRESULT CreateInstance (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, 
		IBreakpointManager* bpman, IDebugBreakpointRequest2* bp_request, const wchar_t* file, uint32_t line_index, IDebugPendingBreakpoint2** to)
	{
		RETURN_HR_IF_NULL(E_INVALIDARG, program);

		wil::com_ptr_nothrow<SourceLinePendingBreakpoint> p = new (std::nothrow) SourceLinePendingBreakpoint(); RETURN_IF_NULL_ALLOC(p);

		p->_callback = callback;
		p->_engine = engine;
		p->_program = program;
		p->_bpman = bpman;
		p->_bp_request = bp_request;
		p->_file = wil::make_process_heap_string_nothrow(file); RETURN_IF_NULL_ALLOC(p->_file);
		p->_line_index = line_index;

		if (auto op = wil::try_com_query_nothrow<IFelixLaunchOptionsProvider>(engine))
		{
			com_ptr<IFelixLaunchOptions> options;
			if (SUCCEEDED(op->GetLaunchOptions(&options)))
				options->get_ProjectDir(&p->_projectDir);
		}

		*to = p.detach();
		return S_OK;
	}

	~SourceLinePendingBreakpoint()
	{
		WI_ASSERT(!_advisedDebugEventCallback);
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = NULL;

		if (TryQI<IUnknown>(static_cast<IDebugPendingBreakpoint2*>(this), riid, ppvObject)
			|| TryQI<IDebugPendingBreakpoint2>(this, riid, ppvObject)
			|| TryQI<IDebugEventCallback2>(this, riid, ppvObject))
			return S_OK;

		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugPendingBreakpoint2
	virtual HRESULT STDMETHODCALLTYPE CanBind (__RPC__deref_out_opt IEnumDebugErrorBreakpoints2 **ppErrorEnum) override
	{
		*ppErrorEnum = nullptr;
		RETURN_HR_IF(E_BP_DELETED, _stateInfo.state == PBPS_DELETED);

		UINT64 address;
		auto hr = ::GetAddressFromSourceLocation (_program.get(), _projectDir.get(), _file.get(), _line_index, &address);
		if (SUCCEEDED(hr))
			return S_OK;

		com_ptr<IDebugErrorBreakpoint2> errorBP;
		hr = ErrorBreakpoint::CreateInstance (_program.get(), this, BPET_GENERAL_WARNING, &errorBP); RETURN_IF_FAILED(hr);
		hr = make_single_entry_enumerator(errorBP.get(), ppErrorEnum); RETURN_IF_FAILED(hr);

		return S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE Bind() override
	{
		RETURN_HR_IF(E_BP_DELETED, _stateInfo.state == PBPS_DELETED);

		// TODO: what do we want to happen if the pending breakpoint is already bound?
		// Simply return S_FALSE, or remove the bound breakpoint and attempt to bind the pending breakpoint again?
		// The second case is something we might want to happen after we unload a module.

WI_ASSERT(!_boundBP);
WI_ASSERT(!_errorBP);

		if (_boundBP)
			RETURN_HR(S_FALSE);

		bool physicalMemorySpace = false;
		UINT64 address;
		auto hr = ::GetAddressFromSourceLocation (_program.get(), _projectDir.get(), _file.get(), _line_index, &address);
		if (FAILED(hr))
		{
			hr = ErrorBreakpoint::CreateInstance (_program.get(), this, BPET_GENERAL_WARNING, &_errorBP); RETURN_IF_FAILED(hr);
			com_ptr<BreakpointErrorEvent> p = new (std::nothrow) BreakpointErrorEvent (this, _errorBP.get()); LOG_IF_NULL_ALLOC(p);
			if (p)
			{
				hr = p->Send (_callback.get(), _engine.get(), _program.get(), nullptr); LOG_IF_FAILED(hr);
			}

			return S_OK;
		}
		else
		{
			_errorBP = nullptr;
			hr = BoundBreakpointImpl::CreateInstance (this, _program, _bpman, physicalMemorySpace, address, &_boundBP); RETURN_IF_FAILED(hr);

			if (_stateInfo.state == PBPS_ENABLED)
				_boundBP->Enable(TRUE);

			wil::com_ptr_nothrow<BreakpointBoundEvent> p = new (std::nothrow) BreakpointBoundEvent (this, _boundBP.get()); LOG_IF_NULL_ALLOC(p);
			if (p)
			{
				hr = p->Send (_callback, _engine, _program, nullptr); LOG_IF_FAILED(hr);
			}

			return S_OK;
		}
	}

	virtual HRESULT STDMETHODCALLTYPE GetState (__RPC__out PENDING_BP_STATE_INFO *pState) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetBreakpointRequest (__RPC__deref_out_opt IDebugBreakpointRequest2 **ppBPRequest) override
	{
		auto hr = _bp_request.copy_to(ppBPRequest); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Virtualize (BOOL fVirtualize) override
	{
		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		if (fVirtualize)
			_stateInfo.flags |= PBPSF_VIRTUALIZED;
		else
			_stateInfo.flags &= ~PBPSF_VIRTUALIZED;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Enable (BOOL fEnable) override
	{
		HRESULT hr;

		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		//if (_boundBP)
		//	_boundBP->Enable(fEnable);

		if (fEnable)
		{
			if (_stateInfo.state != PBPS_ENABLED)
			{
				///_program->RegisterPendingBreakpoint(this);
				WI_ASSERT(!_advisedDebugEventCallback);
				com_ptr<IVsDebugger> debugger;
				hr = serviceProvider->QueryService (SID_SVsShellDebugger, &debugger); RETURN_IF_FAILED(hr);
				hr = debugger->AdviseDebugEventCallback(static_cast<IDebugEventCallback2*>(this)); RETURN_IF_FAILED(hr);
				_advisedDebugEventCallback = true;

				_stateInfo.state = PBPS_ENABLED;
			}
		}
		else
		{
			if (_stateInfo.state == PBPS_ENABLED)
			{
				///_program->UnregisterPendingBreakpoint(this);
				WI_ASSERT(_advisedDebugEventCallback);
				com_ptr<IVsDebugger> debugger;
				hr = serviceProvider->QueryService (SID_SVsShellDebugger, &debugger); RETURN_IF_FAILED(hr);
				hr = debugger->UnadviseDebugEventCallback(static_cast<IDebugEventCallback2*>(this)); RETURN_IF_FAILED(hr);
				_advisedDebugEventCallback = false;
			}
			_stateInfo.state = PBPS_DISABLED;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetCondition (BP_CONDITION bpCondition) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetPassCount (BP_PASSCOUNT bpPassCount) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumBoundBreakpoints (__RPC__deref_out_opt IEnumDebugBoundBreakpoints2 **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumErrorBreakpoints (BP_ERROR_TYPE bpErrorType, __RPC__deref_out_opt IEnumDebugErrorBreakpoints2 **ppEnum) override
	{
		RETURN_HR_IF(E_NOTIMPL, bpErrorType != BPET_ALL);
		auto hr = make_single_entry_enumerator<IEnumDebugErrorBreakpoints2, IDebugErrorBreakpoint2>(_errorBP.get(), ppEnum); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Delete() override
	{
		if (_stateInfo.state == PBPS_DELETED)
			RETURN_HR(E_BP_DELETED);

		if (_stateInfo.state == PBPS_ENABLED)
		{
			///_program->UnregisterPendingBreakpoint(this);
			WI_ASSERT(_advisedDebugEventCallback);
			com_ptr<IVsDebugger> debugger;
			auto hr = serviceProvider->QueryService (SID_SVsShellDebugger, &debugger); RETURN_IF_FAILED(hr);
			hr = debugger->UnadviseDebugEventCallback(static_cast<IDebugEventCallback2*>(this)); RETURN_IF_FAILED(hr);
			_advisedDebugEventCallback = false;
		}

		if (_boundBP)
		{
			// Judging by the values in the BP_UNBOUND_REASON enum, a Breakpoint Unbound Event
			// need not be sent due to breakpoint deletion.
			//auto hr = SendBreakpointUnboundEvent (EVENT_SYNCHRONOUS, callback, engine, _boundBP, BPUR_UNKNOWN); RETURN_IF_FAILED(hr);

			BP_STATE s;
			if (SUCCEEDED(_boundBP->GetState(&s)) && (s != BPS_DELETED))
				_boundBP->Delete();
			_boundBP = nullptr;
		}

		// VS2022 keeps a reference to objects of this type for a long time, even after debugging stops,
		// even after deleting the breakpoint from the Breakpoints window, even after closing the solution.
		// It eventually releases this reference when the user closes the VS application.
		// Let's release here every reference we hold, and hope VS will not "reactivate" a deleted breakpoint.
		_errorBP = nullptr;
		_callback = nullptr;
		_engine = nullptr;
		_program = nullptr;
		_bpman = nullptr;
		_bp_request = nullptr;
		_file.reset();
		_projectDir.reset();

		_stateInfo.state = PBPS_DELETED;
		return S_OK;
	}
	#pragma endregion

	#pragma region IDebugEventCallback2
	virtual HRESULT STDMETHODCALLTYPE Event (IDebugEngine2 *pEngine, IDebugProcess2 *pProcess,
		IDebugProgram2 *pProgram, IDebugThread2 *pThread, IDebugEvent2 *pEvent, REFIID riidEvent, DWORD dwAttrib) override
	{
		// When the user clicks Stop Debugging, VS calls our Delete function, and in that function
		// we call IVsDebugger::UnadviseDebugEventCallback; from that point on we shouldn't receive
		// any more debug events. VS, however, sends us some more events, that's why the check for PBPS_ENABLED.
		if (_stateInfo.state != PBPS_ENABLED)
			return S_OK;

		if (riidEvent == IID_IDebugModuleLoadEvent2)
		{
			com_ptr<IDebugModuleLoadEvent2> mle;
			auto hr = pEvent->QueryInterface(&mle); RETURN_IF_FAILED(hr);
			com_ptr<IDebugModule2> module;
			BOOL load;
			hr = mle->GetModule(&module, nullptr, &load); RETURN_IF_FAILED(hr);
			if (auto z80m = module.try_query<IZ80Module>())
			{
				com_ptr<IZ80Symbols> symbols;
				if (SUCCEEDED(z80m->GetSymbols(&symbols)) && (symbols->HasSourceLocationInformation() == S_OK))
				{
					if (load)
					{
						// If we have an error breakpoint, we try again to bind this pending breakpoint,
						// maybe we can resolve it this time in the just-loaded module.
						if (_errorBP)
						{
							bool physicalMemorySpace = false;
							UINT64 address;
							hr = ::GetAddressFromSourceLocation (module.get(), _projectDir.get(), _file.get(), _line_index, &address);
							if (SUCCEEDED(hr))
							{
								hr = BoundBreakpointImpl::CreateInstance (this, _program, _bpman, physicalMemorySpace, address, &_boundBP); RETURN_IF_FAILED(hr);
								_errorBP = nullptr;

								WI_ASSERT (_stateInfo.state == PBPS_ENABLED);
								_boundBP->Enable(TRUE);

								auto p = com_ptr(new (std::nothrow) BreakpointBoundEvent (this, _boundBP.get())); LOG_IF_NULL_ALLOC(p);
								if (p)
								{
									hr = p->Send (_callback.get(), _engine.get(), _program.get(), nullptr); LOG_IF_FAILED(hr);
								}

								return S_OK;
							}
						}
					}
					else
					{
						// If we have a bound breakpoint whose address is in the just-unloaded module,
						// we unbind it and turn it into an error breakpoint.
						WI_ASSERT(false);
					}
				}
			}
		}

		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeSourceLinePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program,
	IBreakpointManager* bpman,
	IDebugBreakpointRequest2* bp_request, const wchar_t* file, uint32_t line_index, IDebugPendingBreakpoint2** to)
{
	return SourceLinePendingBreakpoint::CreateInstance (callback, engine, program, bpman, bp_request, file, line_index, to);
}
