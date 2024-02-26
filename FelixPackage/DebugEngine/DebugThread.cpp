
#include "pch.h"
#include "shared/TryQI.h"
#include "shared/OtherGuids.h"
#include "DebugEngine.h"

class FelixThread : public IDebugThread2
{
	ULONG _refCount = 0;
	com_ptr<IDebugEngine2> _engine;
	com_ptr<IDebugProgram2> _program;

public:
	HRESULT InitInstance (IDebugEngine2* engine, IDebugProgram2* program, IDebugEventCallback2* callback)
	{
		_engine = engine;
		_program = program;
		return S_OK;
	}

	#pragma region IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDebugThread2*>(this), riid, ppvObject)
			|| TryQI<IDebugThread2>(this, riid, ppvObject))
			return S_OK;

		if (   riid == IID_IAgileObject
			|| riid == IID_IClientSecurity
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IdentityUnmarshal
			|| riid == IID_IStdMarshalInfo
			|| riid == IID_IFastRundown
			|| riid == IID_INotARealInterface
			|| riid == IID_INoIdea7
			|| riid == IID_INoIdea9
			|| riid == IID_INoIdea10
			|| riid == IID_INoIdea11
			|| riid == IID_INoIdea14
			|| riid == IID_INoIdea15
			|| riid == IID_INoIdea16
			|| riid == IID_IDebugComputeThread110
			|| riid == IID_IDebugThread120
			|| riid == IID_IDebugThread157
			|| riid == IID_IExternalConnection)
			return E_NOINTERFACE;
		
		for (auto& i : HardcodedRundownIFsOfInterest)
			if (*i == riid)
				return E_NOINTERFACE;

		RETURN_HR(E_NOTIMPL);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override
	{
		return ++_refCount;
	}

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		WI_ASSERT(_refCount);
		ULONG newRefCount = --_refCount;
		if (newRefCount == 0)
			delete this;
		return newRefCount;
	}
	#pragma endregion

	#pragma region IDebugThread2
	virtual HRESULT STDMETHODCALLTYPE EnumFrameInfo (FRAMEINFO_FLAGS dwFieldSpec, UINT nRadix, IEnumDebugFrameInfo2 **ppEnum) override
	{
		return MakeEnumDebugFrameInfo (dwFieldSpec, nRadix, this, ppEnum);
	}

	virtual HRESULT STDMETHODCALLTYPE GetName(BSTR* pbstrName) override
	{
		*pbstrName = SysAllocString(L"<thread name>");
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetThreadName (LPCOLESTR pszName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetProgram (IDebugProgram2 **ppProgram) override
	{
		*ppProgram = _program.get();
		_program->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE CanSetNextStatement (IDebugStackFrame2 *pStackFrame, IDebugCodeContext2 *pCodeContext) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetNextStatement (IDebugStackFrame2 *pStackFrame, IDebugCodeContext2 *pCodeContext) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetThreadId (DWORD *pdwThreadId) override
	{
		*pdwThreadId = 0;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Suspend (DWORD *pdwSuspendCount) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Resume (DWORD *pdwSuspendCount) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetThreadProperties (THREADPROPERTY_FIELDS dwFields, THREADPROPERTIES *ptp) override
	{
		if (dwFields & TPF_STATE)
		{
			ptp->dwThreadState = THREADSTATE_RUNNING;
			ptp->dwFields |= TPF_STATE;
			dwFields &= ~TPF_STATE;
		}

		LOG_HR_IF(E_NOTIMPL, (bool)dwFields);
		
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetLogicalThread (IDebugStackFrame2 *pStackFrame, IDebugLogicalThread2 **ppLogicalThread) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeThread (IDebugEngine2* engine, IDebugProgram2* program, IDebugEventCallback2* callback, IDebugThread2** ppThread)
{
	auto p = com_ptr(new (std::nothrow) FelixThread()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(engine, program, callback); RETURN_IF_FAILED(hr);
	*ppThread = p.detach();
	return S_OK;
}
