
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"

class Z80ProjectFactory : public IVsProjectFactory
{
	ULONG _refCount = 0;

	wil::com_ptr_nothrow<IServiceProvider> _sp;

public:
	static HRESULT CreateInstance (IServiceProvider* sp, IVsProjectFactory** to)
	{
		wil::com_ptr_nothrow<Z80ProjectFactory> p = new (std::nothrow) Z80ProjectFactory(); RETURN_IF_NULL_ALLOC(p);
		p->_sp = sp;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IVsProjectFactory>(this, riid, ppvObject))
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsProjectFactory
	virtual HRESULT __stdcall CanCreateProject(LPCOLESTR pszFilename, VSCREATEPROJFLAGS grfCreateFlags, BOOL * pfCanCreate) override
	{
		*pfCanCreate = TRUE;
		return S_OK;
	}

	virtual HRESULT __stdcall CreateProject(LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags, REFIID iidProject, void ** ppvProject, BOOL * pfCanceled) override
	{
		RETURN_HR_IF(E_POINTER, !ppvProject || !pfCanceled);
		*ppvProject = nullptr;
		*pfCanceled = TRUE;

		auto hr = MakeFelixProject (_sp.get(), pszFilename, pszLocation, pszName, grfCreateFlags, iidProject, ppvProject); RETURN_IF_FAILED(hr);
	
		*pfCanceled = FALSE;
		return S_OK;
	}

	virtual HRESULT __stdcall SetSite(IServiceProvider* pSP) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall Close() override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT Z80ProjectFactory_CreateInstance (IServiceProvider* sp, IVsProjectFactory** to)
{
	return Z80ProjectFactory::CreateInstance(sp, to);
}

