
#pragma once
#include "TryQI.h"

template<typename I, std::enable_if_t<std::is_constructible_v<IUnknown*, I*>, int> = 0>
class Z80ClassFactory : public IClassFactory
{
	ULONG _refCount = 0;

	using create_t = HRESULT(*)(I** out);
	create_t const _create;

public:
	Z80ClassFactory (create_t create)
		: _create(create)
	{ }

	Z80ClassFactory (const Z80ClassFactory&) = delete;
	Z80ClassFactory& operator= (const Z80ClassFactory&) = delete;

	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
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
		else if (riid == __uuidof(IClassFactory))
		{
			*ppvObject = static_cast<IClassFactory*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }

	virtual HRESULT __stdcall CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override
	{
		*ppvObject = nullptr;

		wil::com_ptr_nothrow<I> p;
		auto hr = _create(&p); RETURN_IF_FAILED(hr);
		
		hr = p->QueryInterface(riid, ppvObject); RETURN_IF_FAILED(hr);

		return S_OK;;
	}

	virtual HRESULT __stdcall LockServer(BOOL fLock) override
	{
		RETURN_HR(E_NOTIMPL);
	}
};
