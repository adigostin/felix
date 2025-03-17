
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockVsHierarchy : IVsHierarchy, IZ80ProjectProperties, IProjectItemParent
{
	ULONG _refCount = 0;
	wil::unique_bstr _projectDir;
	com_ptr<IProjectItem> _firstChild;

	HRESULT InitInstance()
	{
		wchar_t tempPath[MAX_PATH+1];
		DWORD pathLen = GetTempPathW(_countof(tempPath), tempPath); RETURN_LAST_ERROR_IF_EXPECTED(pathLen == 0);
		_projectDir.reset(SysAllocStringLen(tempPath, pathLen)); RETURN_IF_NULL_ALLOC_EXPECTED(_projectDir);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IVsHierarchy*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IVsHierarchy>(this, riid, ppvObject)
			|| TryQI<IZ80ProjectProperties>(this, riid, ppvObject)
			|| TryQI<IProjectItemParent>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IZ80ProjectProperties);

	#pragma region IZ80ProjectProperties
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

	#pragma region IProjectItemParent
	virtual IProjectItem* STDMETHODCALLTYPE FirstChild() override { return _firstChild; }

	virtual void STDMETHODCALLTYPE SetFirstChild (IProjectItem* firstChild) override
	{
		_firstChild = firstChild;
	}
	#pragma endregion
};

com_ptr<IVsHierarchy> MakeMockVsHierarchy()
{
	auto p = com_ptr (new (std::nothrow) MockVsHierarchy());
	Assert::IsNotNull(p.get());
	auto hr = p->InitInstance();
	Assert::IsTrue(SUCCEEDED(hr));
	return p;
}
