
#include "pch.h"
#include "FelixPackage.h"
#include "guids.h"
#include "dispids.h"
#include "shared/vector_nothrow.h"
#include "shared/com.h"
#include "Z80Xml.h"
#include "../FelixPackageUi/resource.h"
#include <vsmanaged.h>

class ObjInfo
{
	com_ptr<IDispatch> _parent;
	DISPID _dispid;
	com_ptr<IDispatch> _child;
	DWORD _propNotifyCookie = 0;

public:
	ObjInfo() = default;

	HRESULT InitInstance (IUnknown* parentUnk, DISPID dispid, IPropertyNotifySink* sink)
	{
		HRESULT hr;

		com_ptr<IDispatch> parent;
		hr = parentUnk->QueryInterface(&parent); RETURN_IF_FAILED(hr);

		_parent = parent;
		_dispid = dispid;

		DISPPARAMS params = { };
		wil::unique_variant result;
		EXCEPINFO exception;
		UINT uArgErr;
		hr = _parent->Invoke (dispid, IID_NULL, LANG_INVARIANT, DISPATCH_PROPERTYGET,
			&params, &result, &exception, &uArgErr); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_UNEXPECTED, result.vt != VT_DISPATCH);
		RETURN_HR_IF(E_POINTER, V_DISPATCH(&result) == nullptr);
		_child = V_DISPATCH(&result);
		// In a property page we don't want to work directly on the selected object,
		// because the user might change some properties and then click Cancel.
		// So let's clone the selected object. See related comment in Apply() below.
		auto stream = com_ptr(SHCreateMemStream(nullptr, 0)); RETURN_IF_NULL_ALLOC(stream);
		hr = SaveToXml(result.pdispVal, L"Temp", SAVE_XML_FORCE_SERIALIZE_DEFAULTS, stream); RETURN_IF_FAILED_EXPECTED(hr);
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

		// We need to serialize also the properties with default values, to account for the case when
		// the user sets a property to its default value and clicks Apply/OK.
		auto stream = com_ptr(SHCreateMemStream(0, 0)); RETURN_IF_NULL_ALLOC(stream);
		hr = SaveToXml(_child, L"Temp", SAVE_XML_FORCE_SERIALIZE_DEFAULTS, stream); RETURN_IF_FAILED_EXPECTED(hr);
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

	IDispatch* GetParent() const { return _parent; }
	IDispatch* GetChild() const { return _child; }
};

static HRESULT SelectObjects (ULONG cObjects, IUnknown** ppUnk, DISPID dispid, 
	vector_nothrow<ObjInfo>& objects, IPropertyNotifySink* sink)
{
	objects.clear();

	bool reserved = objects.try_reserve(cObjects); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);

	for (ULONG i = 0; i < cObjects; i++)
	{
		ObjInfo oi;
		auto hr = oi.InitInstance (ppUnk[i], dispid, sink); RETURN_IF_FAILED(hr);
		objects.try_push_back (std::move(oi));
	}

	return S_OK;
}

static HRESULT SetSelectedObjects (const vector_nothrow<ObjInfo>& objects, IVSMDPropertyGrid* grid)
{
	vector_nothrow<IUnknown*> temp;
	bool bres = temp.try_reserve(objects.size()); RETURN_HR_IF(E_OUTOFMEMORY, !bres);
	for (auto& o : objects)
		temp.try_push_back(o.GetChild());
	return grid->SetSelectedObjects ((int)temp.size(), temp.data());
}

class PGPropertyPage : public IPropertyPage, IPropertyNotifySink
	//, IVsPropertyPage, IVsPropertyPage2, IVsPropertyPageNotify
{
	ULONG _refCount = 0;
	GUID _pageGuid;
	DISPID _dispidChildObj;
	vector_nothrow<ObjInfo> _objects;
	com_ptr<IVSMDPropertyGrid> _grid;
	com_ptr<IPropertyPageSite> _pageSite;
	bool _pageDirty = false;
	wil::unique_bstr title;

public:
	HRESULT InitInstance (UINT titleStringResId, REFGUID pageGuid, DISPID dispidChildObj)
	{
		HRESULT hr;
		com_ptr<IVsShell> shell;
		hr = serviceProvider->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);
		hr = shell->LoadPackageString(CLSID_FelixPackage, titleStringResId, &title); RETURN_IF_FAILED(hr);
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
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

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

		// When switching between pages, VS calls Activate but not Move. Need to arrange the page here too.
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
		pPageInfo->pszTitle = wil::make_cotaskmem_string_nothrow(title.get()).release(); RETURN_IF_NULL_ALLOC(pPageInfo->pszTitle);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetObjects(ULONG cObjects, IUnknown** ppUnk) override
	{
		auto hr = SelectObjects (cObjects, ppUnk, _dispidChildObj, _objects, this); RETURN_IF_FAILED(hr);

		if (_grid)
		{
			hr = SetSelectedObjects (_objects, _grid); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Show(UINT nCmdShow) override
	{
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
		// Pressing F1 will cause VS to execute this.
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(MSG* pMsg) override
	{
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

// ============================================================================

class AsmPropertyPage : public IPropertyPage, IPropertyNotifySink
	//, IVsPropertyPage, IVsPropertyPage2, IVsPropertyPageNotify
{
	ULONG _refCount = 0;
	vector_nothrow<ObjInfo> _objects;
	com_ptr<IVSMDPropertyGrid> _grid;
	HWND _staticHWnd = nullptr;
	HWND _editHWnd = nullptr;
	com_ptr<IPropertyPageSite> _pageSite;
	bool _pageDirty = false;
	wil::unique_bstr title;
	HFONT _font;
	LONG _staticHeight;
	LONG _editHeight;

public:
	HRESULT InitInstance()
	{
		HRESULT hr;
		com_ptr<IVsShell> shell;
		hr = serviceProvider->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);
		hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_ASSEMBLER_PROP_PAGE_TITLE, &title); RETURN_IF_FAILED(hr);

		com_ptr<IVsFontAndColorStorage> fcs;
		hr = serviceProvider->QueryService(SID_SVsFontAndColorStorage, IID_PPV_ARGS(fcs.addressof())); RETURN_IF_FAILED(hr);
		hr = fcs->OpenCategory(GUID_DialogsAndToolWindowsFC, FCSF_READONLY | FCSF_LOADDEFAULTS); RETURN_IF_FAILED(hr);
		LOGFONTW lf = { };
		hr = fcs->GetFont(&lf, NULL); RETURN_IF_FAILED(hr);
		fcs->CloseCategory();

		_font = CreateFontIndirectW(&lf);
		_staticHeight = 3 * abs(lf.lfHeight) / 2;
		_editHeight = 5 * abs(lf.lfHeight);

		return S_OK;
	}

	~AsmPropertyPage()
	{
		::DestroyWindow(_staticHWnd);
		::DestroyWindow(_editHWnd);
		::DeleteObject(_font);
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IPropertyPage*>(this), riid, ppvObject)
			|| TryQI<IPropertyPage>(this, riid, ppvObject)
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
			)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	HRESULT GetCommandLineText (BSTR* pText)
	{
		if (_objects.empty())
		{
			*pText = nullptr;
			return S_OK;
		}

		if (_objects.size() > 1)
		{
			*pText = SysAllocString(L"<multiple selection>");
			return S_OK;
		}

		com_ptr<IProjectConfig> config;
		auto hr = _objects.front().GetParent()->QueryInterface(&config); RETURN_IF_FAILED(hr);
		
		com_ptr<IVsHierarchy> hier;
		hr = config->GetSite(IID_PPV_ARGS(hier.addressof())); RETURN_IF_FAILED(hr);

		com_ptr<IProjectConfigAssemblerProperties> asmProps;
		hr = _objects.front().GetChild()->QueryInterface(&asmProps); RETURN_IF_FAILED(hr);

		return MakeSjasmCommandLine(hier, config, asmProps, pText);
	}

	#pragma region IPropertyPage
	virtual HRESULT STDMETHODCALLTYPE SetPageSite (IPropertyPageSite* pPageSite) override
	{
		_pageSite = pPageSite;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Activate(HWND hWndParent, LPCRECT pRect, BOOL bModal) override
	{
		com_ptr<IVSMDPropertyBrowser> browser;
		auto hr = serviceProvider->QueryService(SID_SVSMDPropertyBrowser, &browser); RETURN_IF_FAILED(hr);

		com_ptr<IVSMDPropertyGrid> grid;
		hr = browser->CreatePropertyGrid(&grid); RETURN_IF_FAILED(hr);

		HWND gridHWnd;
		hr = grid->get_Handle(&gridHWnd); RETURN_IF_FAILED(hr);

		HWND op = ::SetParent (gridHWnd, hWndParent); RETURN_LAST_ERROR_IF(!op);

		// When switching between pages, VS calls Activate but not Move. Need to arrange the page here too.
		BOOL bres = ::MoveWindow (gridHWnd, pRect->left, pRect->top, pRect->right - pRect->left, 
			pRect->bottom - pRect->top - _editHeight - _staticHeight - _staticHeight / 2, TRUE);
		RETURN_LAST_ERROR_IF(!bres);

		HWND staticHWnd = CreateWindowW (L"STATIC", L"Assembler Command Line", WS_VISIBLE | WS_CHILD,
			pRect->left, pRect->bottom - _editHeight - _staticHeight, pRect->right - pRect->left, _staticHeight,
			hWndParent, NULL, NULL, NULL); RETURN_LAST_ERROR_IF(!staticHWnd);
		::SendMessage (staticHWnd, WM_SETFONT, (WPARAM)_font, TRUE);

		wil::unique_bstr cmdLine;
		hr = GetCommandLineText (&cmdLine); RETURN_IF_FAILED(hr);
		HWND editHWnd = CreateWindowW (L"EDIT", cmdLine.get(), WS_VISIBLE | WS_CHILD | WS_BORDER | WS_VSCROLL | ES_READONLY | ES_AUTOVSCROLL | ES_MULTILINE,
			pRect->left, pRect->bottom - _editHeight, pRect->right - pRect->left, _editHeight,
			hWndParent, NULL, NULL, NULL); RETURN_LAST_ERROR_IF(!editHWnd);
		::SendMessage (editHWnd, WM_SETFONT, (WPARAM)_font, TRUE);

		hr = SetSelectedObjects (_objects, grid); RETURN_IF_FAILED(hr);

		_grid = std::move(grid);
		_staticHWnd = staticHWnd;
		_editHWnd = editHWnd;
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
		pPageInfo->pszTitle = wil::make_cotaskmem_string_nothrow(title.get()).release(); RETURN_IF_NULL_ALLOC(pPageInfo->pszTitle);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetObjects(ULONG cObjects, IUnknown** ppUnk) override
	{
		auto hr = SelectObjects (cObjects, ppUnk, dispidAssemblerProperties, _objects, this); RETURN_IF_FAILED(hr);

		if (_grid)
		{
			hr = SetSelectedObjects (_objects, _grid); RETURN_IF_FAILED(hr);
		}

		if (_editHWnd)
		{
			wil::unique_bstr cmdLine;
			hr = GetCommandLineText (&cmdLine); RETURN_IF_FAILED(hr);
			::SetWindowText (_editHWnd, cmdLine.get());
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Show(UINT nCmdShow) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Move(LPCRECT pRect) override
	{
		RETURN_HR_IF (E_FAIL, !_grid);

		HWND gridHWnd;
		auto hr = _grid->get_Handle(&gridHWnd); RETURN_IF_FAILED(hr);

		BOOL bres = ::MoveWindow (gridHWnd, pRect->left, pRect->top, pRect->right - pRect->left,
			pRect->bottom - pRect->top - _editHeight - _staticHeight - _staticHeight / 2, TRUE);
		RETURN_LAST_ERROR_IF(!bres);

		bres = ::MoveWindow (_staticHWnd, pRect->left, pRect->bottom - _editHeight - _staticHeight,
			pRect->right - pRect->left, _staticHeight, TRUE);

		bres = ::MoveWindow (_editHWnd, pRect->left, pRect->bottom - _editHeight,
			pRect->right - pRect->left, _editHeight, TRUE);
		RETURN_LAST_ERROR_IF(!bres);

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
		// Pressing F1 will cause VS to execute this.
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(MSG* pMsg) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IPropertyNotifySink
	virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
	{
		_pageDirty = true;
		_pageSite->OnStatusChange (PROPPAGESTATUS_DIRTY | PROPPAGESTATUS_VALIDATE);

		wil::unique_bstr cmdLine;
		GetCommandLineText (&cmdLine);
		::SetWindowText (_editHWnd, cmdLine.get());

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeAsmPropertyPage (IPropertyPage** to)
{
	auto p = com_ptr(new (std::nothrow) AsmPropertyPage()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
