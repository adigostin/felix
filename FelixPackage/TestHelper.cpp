
#include "pch.h"
#include "shared/com.h"
#include "FelixPackage.h"
#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>

struct TestHelper : IFelixTestHelper
{
	ULONG _refCount = 0;

	HRESULT InitInstance()
	{
		return S_OK;
	}

	~TestHelper()
	{
		// For some reason execution never gets here, even if VS exits cleanly. I'll debug this some other time.
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IFelixTestHelper*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IFelixTestHelper>(this, riid, ppvObject)
		)
			return S_OK;

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion
	
	IMPLEMENT_IDISPATCH(IFelixTestHelper)
	
	#pragma region IFelixTestHelper
	virtual HRESULT STDMETHODCALLTYPE get_BscProjectsEvents (IDispatch **ppdisp) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE get_BscProjectItemsEvents (IDispatch *Filter, IDispatch **ppdisp) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseProjectFileChange (IDispatch* project, IVsFileChangeEvents *pFCE, VSCOOKIE *pvsCookie) override
	{
		com_ptr<IVsFileChangeEx> fileChange;
		auto hr = serviceProvider->QueryService(SID_SVsFileChangeEx, IID_PPV_ARGS(&fileChange)); RETURN_IF_FAILED(hr);

		com_ptr<IProjectNode> pn;
		hr = project->QueryInterface(IID_PPV_ARGS(&pn)); RETURN_IF_FAILED(hr);
		
		wil::unique_bstr projectMk;
		hr = pn->AsVsProject()->GetMkDocument(VSITEMID_ROOT, &projectMk); RETURN_IF_FAILED(hr);

		hr = fileChange->AdviseFileChange (projectMk.get(), VSFILECHG_Time, pFCE, pvsCookie); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseProjectFileChange (VSCOOKIE vsCookie) override
	{
		com_ptr<IVsFileChangeEx> fileChange;
		auto hr = serviceProvider->QueryService(SID_SVsFileChangeEx, IID_PPV_ARGS(&fileChange)); RETURN_IF_FAILED(hr);
		hr = fileChange->UnadviseFileChange(vsCookie); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	#pragma endregion
};

HRESULT TestHelper_CreateInstance (IFelixTestHelper** out)
{
	com_ptr<TestHelper> p = new (std::nothrow) TestHelper(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*out = p.detach();
	return S_OK;
}