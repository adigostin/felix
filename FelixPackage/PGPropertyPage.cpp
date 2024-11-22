
#include "pch.h"
#include "FelixPackage.h"
#include "guids.h"
#include "shared/vector_nothrow.h"
#include "shared/com.h"
#include "Z80Xml.h"
#include <vsmanaged.h>

class ObjInfo
{
	com_ptr<IDispatch> _parent;
	DISPID _dispid;
	com_ptr<IDispatch> _child;
	DWORD _propNotifyCookie = 0;

public:
	ObjInfo() = default;

	HRESULT InitInstance (IUnknown* parent, DISPID dispid, IPropertyNotifySink* sink)
	{
		auto hr = parent->QueryInterface(&_parent); RETURN_IF_FAILED(hr);

		_dispid = dispid;

		DISPPARAMS params = { };
		wil::unique_variant result;
		EXCEPINFO exception;
		UINT uArgErr;
		hr = _parent->Invoke (dispid, IID_NULL, LANG_INVARIANT, DISPATCH_PROPERTYGET,
			&params, &result, &exception, &uArgErr); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_UNEXPECTED, result.vt != VT_DISPATCH);

		// In a property page we don't want to work directly on the selected object,
		// because the user might change some properties and then click Cancel.
		// So let's clone the selected object. See related comment in Apply() below.
		auto stream = com_ptr(SHCreateMemStream(nullptr, 0)); RETURN_IF_NULL_ALLOC(stream);
		hr = SaveToXml(result.pdispVal, L"Temp", stream); RETURN_IF_FAILED_EXPECTED(hr);
		hr = stream->Seek({ 0 }, STREAM_SEEK_SET, nullptr); RETURN_IF_FAILED(hr);
		com_ptr<IXmlParent> xmlParent;
		hr = _parent->QueryInterface(&xmlParent); RETURN_IF_FAILED(hr);
		hr = xmlParent->CreateChild(dispid, nullptr, &_child); RETURN_IF_FAILED(hr);
		hr = LoadFromXml(_child, L"Temp", stream); RETURN_IF_FAILED_EXPECTED(hr);

		if (auto cpc = wil::try_com_query_nothrow<IConnectionPointContainer>(_child))
		{
			com_ptr<IConnectionPoint> cp;
			auto hr = cpc->FindConnectionPoint(IID_IPropertyNotifySink, &cp);
			if (SUCCEEDED(hr))
				cp->Advise(sink, &_propNotifyCookie);
		}

		return S_OK;
	}
	
	ObjInfo (const ObjInfo&) = delete;
	ObjInfo& operator= (const ObjInfo&) = delete;

	ObjInfo (ObjInfo&& from)
	{
		std::swap (_parent, from._parent);
		std::swap (_dispid, from._dispid);
		std::swap (_child, from._child);
		std::swap (_propNotifyCookie, from._propNotifyCookie);
	}

	~ObjInfo()
	{
		if (_propNotifyCookie)
		{
			if (auto cpc = wil::try_com_query_nothrow<IConnectionPointContainer>(_child))
			{
				com_ptr<IConnectionPoint> cp;
				auto hr = cpc->FindConnectionPoint(IID_IPropertyNotifySink, &cp);
				if (SUCCEEDED(hr))
					cp->Unadvise(_propNotifyCookie);
			}
		}
	}

	HRESULT Apply()
	{
		// Rather that doing a DISPATCH_PROPERTYPUT passing in our clone, let's mirror the operations
		// we did in InitInstance: we save our clone to XML, and load the existing object from XML.
		// This has the advantages:
		//  - we don't change the value of the object property (might be a costly operation);
		//  - we can edit an object coming from a read-only property (one that doesn't have a put_Prop).

		DISPPARAMS params = { };
		wil::unique_variant result;
		EXCEPINFO exception;
		UINT uArgErr;
		auto hr = _parent->Invoke (_dispid, IID_NULL, LANG_INVARIANT, DISPATCH_PROPERTYGET,
			&params, &result, &exception, &uArgErr); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_UNEXPECTED, result.vt != VT_DISPATCH);

		auto stream = com_ptr(SHCreateMemStream(0, 0)); RETURN_IF_NULL_ALLOC(stream);
		hr = SaveToXml(_child, L"Temp", stream); RETURN_IF_FAILED_EXPECTED(hr);
		hr = stream->Seek({ 0 }, STREAM_SEEK_SET, nullptr); RETURN_IF_FAILED(hr);
		hr = LoadFromXml(result.pdispVal, L"Temp", stream); RETURN_IF_FAILED_EXPECTED(hr);

		// Note that the above won't work if the object creates and returns a clone from its
		// GET function. If we'll have such getters, the code above will need updating. As of
		// this writing, we don't have any such object. Let's check that this is indeed the case.
		#ifdef _DEBUG
		wil::unique_variant result2;
		hr = _parent->Invoke (_dispid, IID_NULL, LANG_INVARIANT, DISPATCH_PROPERTYGET,
			&params, &result2, &exception, &uArgErr); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_UNEXPECTED, result2.vt != VT_DISPATCH);
		WI_ASSERT(V_DISPATCH(&result) == V_DISPATCH(&result2));
		#endif

		return S_OK;
	}

	IDispatch* GetChild() { return _child; }
};

class PGPropertyPage : public IPropertyPage, IVsPropertyPage, IPropertyNotifySink
{
	ULONG _refCount = 0;
	GUID _pageGuid;
	DISPID _dispidChildObj;
	vector_nothrow<ObjInfo> _objects;
	com_ptr<IVSMDPropertyGrid> _grid;
	com_ptr<IPropertyPageSite> _pageSite;
	bool _pageDirty = false;
	UINT _titleStringResId;

public:
	HRESULT InitInstance (UINT titleStringResId, REFGUID pageGuid, DISPID dispidChildObj)
	{
		_titleStringResId = titleStringResId;
		_pageGuid = pageGuid;
		_dispidChildObj = dispidChildObj;
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
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == IID_IVsPropertyPage2)
			return E_NOINTERFACE;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	static HRESULT SetSelectedObjects (vector_nothrow<ObjInfo>& objects, IVSMDPropertyGrid* grid)
	{
		vector_nothrow<IUnknown*> temp;
		bool bres = temp.try_reserve(objects.size()); RETURN_HR_IF(E_OUTOFMEMORY, !bres);
		for (auto& o : objects)
			temp.try_push_back(o.GetChild());
		return grid->SetSelectedObjects ((int)temp.size(), temp.data());
	}

	#pragma region IPropertyPage
	virtual HRESULT STDMETHODCALLTYPE SetPageSite (IPropertyPageSite* pPageSite) override
	{
		// I debugged through the QueryInterface implementation
		// of pPageSite (msenv.dll of VS2022) and it seems the only interfaces it checks for are
		// IUnknown, IPropertyPageSite and some IID_IVsPropertyPageSitePrivate.
		// I've seen C# samples that cast this to a System.IServiceProvider, but that's different from
		// the IServiceProvider defined in servprov.h.

		_pageSite = pPageSite;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Activate(HWND hWndParent, LPCRECT pRect, BOOL bModal) override
	{
		com_ptr<IVSMDPropertyBrowser> browser;
		auto hr = serviceProvider->QueryService(SID_SVSMDPropertyBrowser, &browser); RETURN_IF_FAILED(hr);

		com_ptr<IVSMDPropertyGrid> grid;
		hr = browser->CreatePropertyGrid(&grid); RETURN_IF_FAILED(hr);

		HWND hwnd;
		hr = grid->get_Handle(&hwnd); RETURN_IF_FAILED(hr);

		HWND op = ::SetParent (hwnd, hWndParent); RETURN_LAST_ERROR_IF(!op);

		BOOL bRes = ::MoveWindow (hwnd, pRect->left, pRect->top, pRect->right - pRect->left, pRect->bottom - pRect->top, TRUE);
		RETURN_LAST_ERROR_IF(!bRes);

		hr = SetSelectedObjects (_objects, grid); RETURN_IF_FAILED(hr);

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

		_objects.clear();

		bool reserved = _objects.try_reserve(cObjects); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);

		for (ULONG i = 0; i < cObjects; i++)
		{
			com_ptr<IDispatch> parentObj;
			hr = ppUnk[i]->QueryInterface(&parentObj); RETURN_IF_FAILED(hr);
			
			ObjInfo oi;
			hr = oi.InitInstance (parentObj, _dispidChildObj, this); RETURN_IF_FAILED(hr);
			_objects.try_push_back (std::move(oi));
		}

		if (_grid)
		{
			hr = SetSelectedObjects (_objects, _grid); RETURN_IF_FAILED(hr);
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
		return _pageDirty ? S_OK : S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE Apply() override
	{
		for (auto& oi : _objects)
		{
			auto hr = oi.Apply(); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Help(LPCOLESTR pszHelpDir) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(MSG* pMsg) override
	{
		return E_NOTIMPL;
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

	#pragma region IPropertyNotifySink
	virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
	{
		_pageDirty = true;
		return _pageSite->OnStatusChange (PROPPAGESTATUS_DIRTY | PROPPAGESTATUS_VALIDATE);
	}

	virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakePGPropertyPage (UINT titleStringResId, REFGUID pageGuid, DISPID dispidChildObj, IPropertyPage** to)
{
	com_ptr<PGPropertyPage> p = new (std::nothrow) PGPropertyPage(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(titleStringResId, pageGuid, dispidChildObj); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
