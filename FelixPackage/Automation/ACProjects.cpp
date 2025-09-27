
#include "pch.h"
#include "shared/com.h"
#include <dte.h>

struct ACProjects : VxDTE::Projects
{
	ULONG _refCount = 0;

	HRESULT InitInstance()
	{
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<VxDTE::Projects*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<VxDTE::Projects>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(VxDTE::Projects);

	#pragma region VxDTE::Projects
    virtual HRESULT STDMETHODCALLTYPE Item (VARIANT index, VxDTE::Project** lppcReturn) override
	{
		MessageBox(0, 0, 0, 0);
		RETURN_HR(E_NOTIMPL);
	}

    virtual HRESULT STDMETHODCALLTYPE get_Parent (VxDTE::DTE	**lppaReturn) override
	{
		RETURN_HR(E_NOTIMPL);
	}

    virtual HRESULT STDMETHODCALLTYPE get_Count (long *lplReturn) override
	{
		RETURN_HR(E_NOTIMPL);
	}

    virtual HRESULT STDMETHODCALLTYPE _NewEnum (IUnknown **lppiuReturn) override
	{
		RETURN_HR(E_NOTIMPL);
	}

    virtual HRESULT STDMETHODCALLTYPE get_DTE (VxDTE::DTE** lppaReturn) override
	{
		RETURN_HR(E_NOTIMPL);
	}

    virtual HRESULT STDMETHODCALLTYPE get_Properties (VxDTE::Properties **ppObject) override
	{
		RETURN_HR(E_NOTIMPL);
	}

    virtual HRESULT STDMETHODCALLTYPE get_Kind (BSTR *lpbstrReturn) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeDTEProjects (VxDTE::Projects** ppProjects)
{
	auto p = com_ptr(new (std::nothrow) ACProjects()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*ppProjects = p.detach();
	return S_OK;
}
