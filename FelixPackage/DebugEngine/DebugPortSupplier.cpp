
#include "pch.h"
#include "DebugEngine.h"
#include "shared/enumerator.h"
#include "shared/TryQI.h"
#include "../guids.h"
#include "../FelixPackage.h"

class Z80DebugPortSupplier : public IDebugPortSupplier2, /*public IDebugPortSupplierEx2, */public IDebugPortSupplierDescription2
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IDebugCoreServer2> _server;
	wil::com_ptr_nothrow<IDebugPort2> _z80Port;

public:
	friend HRESULT MakeDebugPortSupplier (IDebugPortSupplier2** to);
	friend static void DestroyDebugPortSupplier (Z80DebugPortSupplier* ps);

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDebugPortSupplier2*>(this), riid, ppvObject)
			|| TryQI<IDebugPortSupplier2>(this, riid, ppvObject)
			|| TryQI<IDebugPortSupplierDescription2>(this, riid, ppvObject))
			return S_OK;

		if (   riid == IID_IDebugPortSupplierLocale2
			|| riid == IID_IClientSecurity
			|| riid == IID_IDebugPortSupplierEx2
			|| riid == IID_IDebugPortSupplier3
			|| riid == IID_IDebugPortSupplier169)
			return E_NOINTERFACE;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugPortSupplier2
	virtual HRESULT __stdcall GetPortSupplierName(BSTR* pbstrName) override
	{
		*pbstrName = SysAllocString(L"Z80PortSupplierName"); RETURN_IF_NULL_ALLOC(*pbstrName);
		return S_OK;
	}

	virtual HRESULT __stdcall GetPortSupplierId(GUID* pguidPortSupplier) override
	{
		*pguidPortSupplier = PortSupplier_Id;
		return S_OK;
	}

	virtual HRESULT __stdcall GetPort(REFGUID guidPort, IDebugPort2** ppPort) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall EnumPorts(IEnumDebugPorts2** ppEnum) override
	{
		return make_single_entry_enumerator<IEnumDebugPorts2, IDebugPort2>(_z80Port.get(), ppEnum);
	}

	virtual HRESULT __stdcall CanAddPort() override
	{
		return S_OK;
	}

	virtual HRESULT __stdcall AddPort (IDebugPortRequest2* pRequest, IDebugPort2** ppPort) override
	{
		HRESULT hr;

		*ppPort = nullptr;

		wil::unique_bstr portName;
		hr = pRequest->GetPortName(&portName); RETURN_IF_FAILED(hr);

		if (_z80Port)
		{
			// We have a port already. If the name matches, we return this port.
			// Don't know why VS asks us to add a port even when it's already there.
			if (!portName || wcscmp(portName.get(), SingleDebugPortName))
				RETURN_HR(E_UNEXPECTED);
		}
		else
		{
			hr = MakeDebugPort (portName.get(), SingleDebugPortId, &_z80Port); RETURN_IF_FAILED(hr);
		}

		*ppPort = _z80Port.get();
		(*ppPort)->AddRef();

		return S_OK;
	}

	virtual HRESULT __stdcall RemovePort(IDebugPort2* pPort) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

///	#pragma region IDebugPortSupplierEx2
///	virtual HRESULT STDMETHODCALLTYPE SetServer (IDebugCoreServer2 *pServer) override
///	{
///		_server = pServer;
///		return S_OK;
///	}
///	#pragma endregion

	#pragma region IDebugPortSupplierDescription2
	virtual HRESULT STDMETHODCALLTYPE GetDescription (PORT_SUPPLIER_DESCRIPTION_FLAGS* pdwFlags, BSTR *pbstrText) override
	{
		if (*pbstrText = SysAllocString(L"18738910731327867312"))
			return S_OK;

		return E_OUTOFMEMORY;
	}
	#pragma endregion
};

HRESULT MakeDebugPortSupplier (IDebugPortSupplier2** to)
{
	wil::com_ptr_nothrow<Z80DebugPortSupplier> p = new (std::nothrow) Z80DebugPortSupplier();
	if (!p)
	{
		*to = nullptr;
		return E_OUTOFMEMORY;
	}
	
	// Let's pretend that we (the supplier) have detected an attached port; we create a IDebugPort2 for it.
	// This is needed in case the user starts the IDE and without any solution open clicks Attach...
	//auto hr = MakeDebugPort(SingleDebugPortName, SingleDebugPortId, &p->_z80Port); IfFailedAssertRet();

	*to = p.detach();
	return S_OK;
}

static void DestroyDebugPortSupplier (Z80DebugPortSupplier* ps)
{
	delete ps;
}
