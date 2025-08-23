
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "Z80Xml.h"
#include "dispids.h"

class LaunchOptionsImpl
	: public IFelixLaunchOptions
{
	ULONG _refCount = 0;
	wil::unique_process_heap_string _projectDir;
	DWORD _baseAddress;
	DWORD _entryPointAddress;

public:
	HRESULT InitInstance()
	{
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF (E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IFelixLaunchOptions*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IFelixLaunchOptions>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IFelixLaunchOptions)

	#pragma region IFelixLaunchOptions
	virtual HRESULT STDMETHODCALLTYPE get_ProjectDir (BSTR *pbstr) override
	{
		*pbstr = SysAllocString(_projectDir.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_ProjectDir (BSTR pbstr) override
	{
		_projectDir = wil::make_unique_string_nothrow<wil::unique_process_heap_string>(pbstr); RETURN_IF_NULL_ALLOC(_projectDir);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_BaseAddress (DWORD *pdwAddress) override
	{
		*pdwAddress = _baseAddress;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_BaseAddress (DWORD dwAddress) override
	{
		_baseAddress = dwAddress;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_EntryPointAddress (DWORD *pdwAddress) override
	{
		*pdwAddress = _entryPointAddress;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_EntryPointAddress (DWORD dwAddress) override
	{
		_entryPointAddress = dwAddress;
		return S_OK;
	}

	#pragma endregion
};

HRESULT MakeLaunchOptions (IFelixLaunchOptions** ppOptions)
{
	auto p = com_ptr(new (std::nothrow) LaunchOptionsImpl()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*ppOptions = p.detach();
	return S_OK;
}
