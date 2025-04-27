
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockVsHierarchy : IVsHierarchy, IProjectNodeProperties, IProjectNode, IParentNode, IPropertyNotifySink
{
	ULONG _refCount = 0;
	wil::unique_bstr _projectDir;
	com_ptr<IChildNode> _firstChild;
	WeakRefToThis _weakRefToThis;
	VSITEMID _nextItemId = 1000;

	HRESULT InitInstance (const wchar_t* projectDir)
	{
		auto hr = _weakRefToThis.InitInstance(static_cast<IVsHierarchy*>(this)); RETURN_IF_FAILED(hr);
		_projectDir = wil::make_bstr_nothrow(projectDir); RETURN_IF_NULL_ALLOC_EXPECTED(_projectDir);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IVsHierarchy*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IVsHierarchy>(this, riid, ppvObject)
			|| TryQI<IProjectNodeProperties>(this, riid, ppvObject)
			|| TryQI<IProjectNode>(this, riid, ppvObject)
			|| TryQI<IParentNode>(this, riid, ppvObject)
			|| TryQI<INode>(this, riid, ppvObject)
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

		return E_NOINTERFACE;
	}
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(__uuidof(IProjectNodeProperties));

	#pragma region IProjectNodeProperties
	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Configurations (SAFEARRAY** configs) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Configurations (SAFEARRAY* sa) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Items (SAFEARRAY** items) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Items (SAFEARRAY* sa) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Guid (BSTR* value) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Guid (BSTR value) override
	{
		return E_NOTIMPL;
	}

    virtual HRESULT STDMETHODCALLTYPE get_AutoOpenFiles(BSTR *pbstrFilenames) override
	{
		return E_NOTIMPL;
	}
        
    virtual HRESULT STDMETHODCALLTYPE put_AutoOpenFiles (BSTR bstrFilenames) override
	{
		return E_NOTIMPL;
	}
        
    virtual HRESULT STDMETHODCALLTYPE GetAutoOpenFiles (BSTR *pbstrFilenames) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IVsHierarchy
	HRESULT STDMETHODCALLTYPE SetSite(IServiceProvider* pSP) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetSite(IServiceProvider** ppSP) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE QueryClose(BOOL* pfCanClose) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE Close(void) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetGuidProperty(VSITEMID itemid, VSHPROPID propid, GUID* pguid) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE SetGuidProperty(VSITEMID itemid, VSHPROPID propid, REFGUID rguid) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetProperty(VSITEMID itemid, VSHPROPID propid, VARIANT* pvar) override
	{
		if (itemid == VSITEMID_ROOT)
		{
			if (propid == VSHPROPID_Name)
				return InitVariantFromString (L"TestProject", pvar);

			if (propid == VSHPROPID_ProjectDir)
				return InitVariantFromString (_projectDir.get(), pvar);

			if (propid == VSHPROPID_FirstChild)
				return InitVariantFromVSITEMID (_firstChild ? _firstChild->GetItemId() : VSITEMID_NIL, pvar);

			return E_NOTIMPL;
		}

		for (auto p = _firstChild.get(); p; p = p->Next())
		{
			if (p->GetItemId() == itemid)
				return p->GetProperty(propid, pvar);
		}

		return E_INVALIDARG;
	}
	HRESULT STDMETHODCALLTYPE SetProperty(VSITEMID itemid, VSHPROPID propid, VARIANT var) override
	{
		//if (propid == VSHPROPID_ProjectDir)
		//{
		//	if (var.vt != VT_BSTR)
		//		return E_UNEXPECTED;
		//	_projectDir = wil::make_bstr_nothrow(var.bstrVal);
		//	return S_OK;
		//}

		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetNestedHierarchy(VSITEMID itemid, REFIID iidHierarchyNested, void** ppHierarchyNested, VSITEMID* pitemidNested) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetCanonicalName(VSITEMID itemid, BSTR* pbstrName) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE ParseCanonicalName(LPCOLESTR pszName, VSITEMID* pitemid) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE Unused0(void) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE AdviseHierarchyEvents(IVsHierarchyEvents* pEventSink, VSCOOKIE* pdwCookie) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE UnadviseHierarchyEvents(VSCOOKIE dwCookie) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE Unused1(void) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE Unused2(void) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE Unused3(void) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE Unused4(void) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IParentNode
	virtual VSITEMID STDMETHODCALLTYPE GetItemId() override { return VSITEMID_ROOT; }

	virtual IChildNode* STDMETHODCALLTYPE FirstChild() override { return _firstChild; }

	virtual void STDMETHODCALLTYPE SetFirstChild (IChildNode* firstChild) override
	{
		_firstChild = firstChild;
	}
	#pragma endregion

	#pragma region IProjectNode
	virtual VSITEMID STDMETHODCALLTYPE MakeItemId() override
	{
		return _nextItemId++;
	}

	virtual HRESULT STDMETHODCALLTYPE EnumHierarchyEventSinks(IEnumHierarchyEvents **ppSinks)
	{
		*ppSinks = nullptr;
		return S_FALSE;
	}
	#pragma endregion

	#pragma region IPropertyNotifySink
	virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion
};

com_ptr<IVsHierarchy> MakeMockVsHierarchy (const wchar_t* projectDir)
{
	auto p = com_ptr (new (std::nothrow) MockVsHierarchy());
	Assert::IsNotNull(p.get());
	auto hr = p->InitInstance(projectDir);
	Assert::IsTrue(SUCCEEDED(hr));
	return p;
}
