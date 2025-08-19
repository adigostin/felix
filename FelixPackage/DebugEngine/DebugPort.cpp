
#include "pch.h"
#include "DebugEngine.h"
#include "DebugEventBase.h"
#include "shared/com.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/OtherGuids.h"

class Z80DebugPort : public IZ80DebugPort, IConnectionPoint, IConnectionPointContainer
{
	ULONG _refCount = 0;
	wil::unique_bstr _port_name;
	GUID _portId;
	unordered_map_nothrow<DWORD, wil::com_ptr_nothrow<IDebugPortEvents2>> _callbacks;
	DWORD _lastCallbackCookie = 1;

public:
	friend HRESULT MakeDebugPort (const wchar_t* portName, const GUID& portId, IDebugPort2** port);
	friend void DestroyDebugPort (Z80DebugPort* p);
	
	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IZ80DebugPort*>(this), riid, ppvObject)
			|| TryQI<IDebugPort2>(this, riid, ppvObject)
			|| TryQI<IZ80DebugPort>(this, riid, ppvObject)
			|| TryQI<IConnectionPoint>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		if (riid == IID_IClientSecurity) return E_NOINTERFACE;
		if (riid == IID_IMarshal) return E_NOINTERFACE;
		if (riid == IID_INoMarshal) return E_NOINTERFACE;
		if (riid == IID_IStdMarshalInfo) return E_NOINTERFACE;
		if (riid == IID_IAgileObject) return E_NOINTERFACE;
		if (riid == IID_IFastRundown) return E_NOINTERFACE;
		if (riid == IID_CDebugPort) return E_NOINTERFACE;
		if (riid == IID_IDebugPortEx2) return E_NOINTERFACE; // This interface is normally private between the SDM and the custom port supplier.
		if (riid == IID_INoIdea7) return E_NOINTERFACE;
		if (riid == IID_IdentityUnmarshal) return E_NOINTERFACE;
		if (riid == IID_INotARealInterface) return E_NOINTERFACE;
		if (riid == IID_INoIdea9) return E_NOINTERFACE;
		if (riid == IID_INoIdea10) return E_NOINTERFACE;
		if (riid == IID_INoIdea11) return E_NOINTERFACE;
		if (riid == IID_INoIdea14) return E_NOINTERFACE;
		if (riid == IID_INoIdea15) return E_NOINTERFACE;
		for (auto& i : HardcodedRundownIFsOfInterest)
			if (*i == riid)
				return E_NOINTERFACE;
		if (riid == IID_IExternalConnection) return E_NOINTERFACE;
		if (riid == IID_ICallFactory) return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugPort2
	virtual HRESULT __stdcall GetPortName(BSTR* pbstrName) override
	{
		*pbstrName = SysAllocString (_port_name.get()); RETURN_IF_NULL_ALLOC(*pbstrName);
		return S_OK;
	}

	virtual HRESULT __stdcall GetPortId(GUID* pguidPort) override
	{
		*pguidPort = _portId;
		return S_OK;
	}

	virtual HRESULT __stdcall GetPortRequest(IDebugPortRequest2** ppRequest) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall GetPortSupplier(IDebugPortSupplier2** ppSupplier) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall GetProcess(AD_PROCESS_ID ProcessId, IDebugProcess2** ppProcess) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall EnumProcesses(IEnumDebugProcesses2** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
		//return MakeSingleEntryEnumerator<IEnumDebugProcesses2, IDebugProcess2>(_z80Program->AsProcess(), ppEnum);
	}
	#pragma endregion

	#pragma region IConnectionPoint
	virtual HRESULT __stdcall GetConnectionInterface(IID* pIID) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall GetConnectionPointContainer(IConnectionPointContainer** ppCPC) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall Advise(IUnknown* pUnkSink, DWORD* pdwCookie) override
	{
		wil::com_ptr_nothrow<IDebugPortEvents2> callback;
		auto hr = pUnkSink->QueryInterface(&callback); RETURN_IF_FAILED(hr);

		_lastCallbackCookie++;
		*pdwCookie = _lastCallbackCookie;
		bool inserted = _callbacks.try_insert ({ _lastCallbackCookie, std::move(callback) }); RETURN_HR_IF(E_OUTOFMEMORY, !inserted);
		return S_OK;
	}

	virtual HRESULT __stdcall Unadvise(DWORD dwCookie) override
	{
		auto it = _callbacks.find(dwCookie);
		RETURN_HR_IF (E_INVALIDARG, it == _callbacks.end());
		_callbacks.remove(it);
		return S_OK;
	}

	virtual HRESULT __stdcall EnumConnections(IEnumConnections** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IZ80DebugPort

	// From https://docs.microsoft.com/en-us/visualstudio/extensibility/debugger/required-port-supplier-interfaces?view=vs-2022
	// A port typically sends program create and destroy events in response to
	// the AddProgramNode and RemoveProgramNode methods, respectively.

	virtual HRESULT __stdcall SendProgramCreateEventToSinks (IDebugProgram2 *pProgram) override
	{
		using ProgramCreateEvent = EventBase<IDebugProgramCreateEvent2, EVENT_ASYNCHRONOUS>;
		wil::com_ptr_nothrow<ProgramCreateEvent> pce = new (std::nothrow) ProgramCreateEvent(); RETURN_IF_NULL_ALLOC(pce);

		com_ptr<IDebugProcess2> process;
		auto hr = pProgram->GetProcess(&process); RETURN_IF_FAILED(hr);

		for (auto& c : _callbacks)
		{
			auto hr = c.second->Event (nullptr, this, process.get(), pProgram, pce.get(), IID_IDebugProgramCreateEvent2); LOG_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT __stdcall SendProgramDestroyEventToSinks (IDebugProgram2 *pProgram, DWORD exitCode) override
	{
		struct ProgramDestroyEvent : public EventBase<IDebugProgramDestroyEvent2, EVENT_ASYNCHRONOUS>
		{
			DWORD _exitCode;

			virtual HRESULT __stdcall GetExitCode(DWORD* exitCode) override
			{
				*exitCode = _exitCode;
				return S_OK;
			}
		};

		wil::com_ptr_nothrow<ProgramDestroyEvent> pde = new (std::nothrow) ProgramDestroyEvent(); RETURN_IF_NULL_ALLOC(pde);
		pde->_exitCode = exitCode;
		
		for (auto& c : _callbacks)
		{
			auto hr = c.second->Event (nullptr, this, nullptr, pProgram, pde.get(), IID_IDebugProgramDestroyEvent2); LOG_IF_FAILED(hr);
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IConnectionPointContainer
	virtual HRESULT __stdcall EnumConnectionPoints(IEnumConnectionPoints** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP) override
	{
		if (ppCP == nullptr)
			return E_POINTER;

		*ppCP = nullptr;

		if (riid == IID_IDebugPortEvents2)
		{
			*ppCP = this;
			AddRef();
			return S_OK;
		}

		RETURN_HR(CONNECT_E_NOCONNECTION);
	}
	#pragma endregion
};

HRESULT MakeDebugPort (const wchar_t* portName, const GUID& portId, IDebugPort2** port)
{
	wil::com_ptr_nothrow<Z80DebugPort> p = new (std::nothrow) Z80DebugPort(); RETURN_IF_NULL_ALLOC(p);
	p->_port_name = wil::make_bstr_nothrow(portName); RETURN_IF_NULL_ALLOC(p->_port_name);
	p->_portId = portId;
	/*
	// Let's pretend we detected a process running on the "device" that we (the port) are "connected" to.
	auto hr = MakeSimulatedProgram(&p->_z80Program); _ASSERT(SUCCEEDED(hr));
	wil::com_ptr_nothrow<IDebugProgramPublisher2> publisher;
	hr = CoCreateInstance (CLSID_ProgramPublisher, nullptr, CLSCTX_INPROC_SERVER, __uuidof(publisher), (void**)&publisher); IfFailedAssertRet();
	hr = publisher->PublishProgramNode(p->_z80Program->AsProgramNode()); IfFailedAssertRet();
	hr = publisher->PublishProgram ({ 1, &Engine_Id }, L"Programu", p->_z80Program->AsProgram()); IfFailedAssertRet();
	*/
	*port = p.detach();
	return S_OK;
}

static void DestroyDebugPort (Z80DebugPort* p)
{
	/*
	wil::com_ptr_nothrow<IDebugProgramPublisher2> publisher;
	auto hr = CoCreateInstance (CLSID_ProgramPublisher, nullptr, CLSCTX_INPROC_SERVER, __uuidof(publisher), (void**)&publisher);
	if (SUCCEEDED(hr))
	{
		publisher->UnpublishProgram (p->_z80Program->AsProgram());
		publisher->UnpublishProgramNode(p->_z80Program->AsProgramNode());
	}
	*/
	delete p;
}

