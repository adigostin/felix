
#include "pch.h"
#include "FelixPackage.h"
#include "guids.h"
#include "shared/vector_nothrow.h"
#include "shared/TryQI.h"
#include <vsmanaged.h>

class PGPropertyPage : public IPropertyPage, IVsPropertyPage, IVsPropertyPageNotify
{
	ULONG _refCount = 0;
	GUID _pageGuid;
	vector_nothrow<wil::com_ptr_nothrow<IUnknown>> _objects;
	wil::com_ptr_nothrow<IVSMDPropertyGrid> _grid;
	bool _dirty = false;
	UINT _titleStringResId;

public:
	HRESULT InitInstance (UINT titleStringResId, REFGUID pageGuid)
	{
		_titleStringResId = titleStringResId;
		_pageGuid = pageGuid;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IPropertyPage*>(this), riid, ppvObject)
			|| TryQI<IPropertyPage>(this, riid, ppvObject)
			|| TryQI<IVsPropertyPage>(this, riid, ppvObject)
			|| TryQI<IVsPropertyPageNotify>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == IID_IVsPropertyPage2)
			return E_NOINTERFACE;

		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IPropertyPage
	virtual HRESULT STDMETHODCALLTYPE SetPageSite(IPropertyPageSite* pPageSite) override
	{
		// No use for the pointer we get here. I debugged through the QueryInterface implementation
		// of pPageSite (msenv.dll of VS2022) and it seems the only interfaces it checks for are
		// IUnknown, IPropertyPageSite and some IID_IVsPropertyPageSitePrivate.
		// I've seen C# samples that cast this to a System.IServiceProvider, but that's different from
		// the IServiceProvider defined in servprov.h.
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Activate(HWND hWndParent, LPCRECT pRect, BOOL bModal) override
	{
		wil::com_ptr_nothrow<IVSMDPropertyBrowser> browser;
		auto hr = serviceProvider->QueryService(SID_SVSMDPropertyBrowser, &browser); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IVSMDPropertyGrid> grid;
		hr = browser->CreatePropertyGrid(&grid); RETURN_IF_FAILED(hr);

		HWND hwnd;
		hr = grid->get_Handle(&hwnd); RETURN_IF_FAILED(hr);

		HWND op = ::SetParent (hwnd, hWndParent); RETURN_LAST_ERROR_IF(!op);

		BOOL bRes = ::MoveWindow (hwnd, pRect->left, pRect->top, pRect->right - pRect->left, pRect->bottom - pRect->top, TRUE);
		RETURN_LAST_ERROR_IF(!bRes);

		vector_nothrow<IUnknown*> temp;
		bool bres = temp.try_reserve(_objects.size()); RETURN_HR_IF(E_OUTOFMEMORY, !bres);
		for (auto& o : _objects)
			temp.try_push_back(o.get());
		hr = grid->SetSelectedObjects ((int)temp.size(), temp.data()); RETURN_IF_FAILED(hr);

		_grid = std::move(grid);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Deactivate() override
	{
		_grid->Dispose();
		_grid = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetPageInfo(PROPPAGEINFO* pPageInfo) override
	{
		pPageInfo->cb = sizeof(PROPPAGEINFO);
		wil::com_ptr_nothrow<IVsShell> shell;
		auto hr = serviceProvider->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);
		wil::unique_bstr title;
		hr = shell->LoadPackageString(CLSID_FelixPackage, _titleStringResId, &title); RETURN_IF_FAILED(hr);
		pPageInfo->pszTitle = wil::make_cotaskmem_string_nothrow(title.get()).release(); RETURN_IF_NULL_ALLOC(pPageInfo->pszTitle);
		pPageInfo->size = { 100, 100 };
		pPageInfo->pszDocString = wil::make_cotaskmem_string_nothrow(L"pszDocString").release(); RETURN_IF_NULL_ALLOC(pPageInfo->pszDocString);
		pPageInfo->pszHelpFile = wil::make_cotaskmem_string_nothrow(L"").release(); RETURN_IF_NULL_ALLOC(pPageInfo->pszHelpFile);
		pPageInfo->dwHelpContext = 0;
		return S_OK;
	}

	// The objects passed to this function are expected to implement ISettingsPageSelector.
	virtual HRESULT STDMETHODCALLTYPE SetObjects(ULONG cObjects, IUnknown** ppUnk) override
	{
		HRESULT hr;

		vector_nothrow<wil::com_ptr_nothrow<IUnknown>> pageObjects;
		for (ULONG i = 0; i < cObjects; i++)
		{
			wil::com_ptr_nothrow<IPropertyGridObjectSelector> pp;
			hr = ppUnk[i]->QueryInterface(&pp); RETURN_IF_FAILED(hr);
			wil::com_ptr_nothrow<IUnknown> sub;
			hr = pp->GetObjectForPropertyGrid(_pageGuid, &sub); RETURN_IF_FAILED(hr);
			bool pushed = pageObjects.try_push_back(std::move(sub)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}

		_objects = std::move(pageObjects);

		if (_grid)
		{
			hr = _grid->SetSelectedObjects ((int)_objects.size(), _objects.data()->addressof()); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Show(UINT nCmdShow) override
	{
		RETURN_HR_IF (E_FAIL, !_grid);

		HWND hwnd;
		auto hr = _grid->get_Handle(&hwnd); RETURN_IF_FAILED(hr);

		BOOL bRes = ::ShowWindow (hwnd, nCmdShow);
		if (!bRes)
			return HRESULT_FROM_WIN32(::GetLastError());

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Move(LPCRECT pRect) override
	{
		RETURN_HR_IF (E_FAIL, !_grid);

		HWND hwnd;
		auto hr = _grid->get_Handle(&hwnd); RETURN_IF_FAILED(hr);

		BOOL bRes = ::MoveWindow (hwnd, pRect->left, pRect->top, pRect->right - pRect->left, pRect->bottom - pRect->top, TRUE);
		RETURN_LAST_ERROR_IF(!bRes);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE IsPageDirty(void) override
	{
		return _dirty ? S_OK : S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE Apply() override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Help(LPCOLESTR pszHelpDir) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(MSG* pMsg) override
	{
		return E_NOTIMPL;
	}
	#pragma region

	#pragma region IVsPropertyPageNotify
	virtual HRESULT STDMETHODCALLTYPE OnShowPage (BOOL fPageActivated) override
	{
		//__debugbreak();
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsPropertyPage
	virtual HRESULT STDMETHODCALLTYPE get_CategoryTitle (UINT iLevel, BSTR *pbstrCategory) override
	{
		// If your property page does not have a category and you would prefer it to show
		// on the top level of the tree view directly under the appropriate top-level category,
		// then either implement IPropertyPage alone, or return E_NOTIMPL from this method.
		return E_NOTIMPL;
	}
	#pragma endregion
};

HRESULT MakePGPropertyPage (UINT titleStringResId, REFGUID pageGuid, IPropertyPage** to)
{
	com_ptr<PGPropertyPage> p = new (std::nothrow) PGPropertyPage(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(titleStringResId, pageGuid); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
