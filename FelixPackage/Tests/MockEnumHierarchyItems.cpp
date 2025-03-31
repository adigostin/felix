
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockEnumHierarchyItems : IEnumHierarchyItems
{
	ULONG _refCount = 0;
	com_ptr<IVsHierarchy> _hier;
	VSEHI _grfItems;
	VSITEMID _current;

	HRESULT InitInstance (IVsHierarchy *pHierRoot, VSEHI grfItems, VSITEMID itemidRoot)
	{
		Assert::AreEqual<DWORD>(0, grfItems & VSEHI_Nest, L"Nested not supported");
		Assert::AreEqual<DWORD>(0, grfItems & VSEHI_Branch, L"Branch not supported");
		_hier = pHierRoot;
		_grfItems = grfItems;
		_current = itemidRoot;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { return E_NOTIMPL; }

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IEnumHierarchyItems
    virtual HRESULT STDMETHODCALLTYPE Next (ULONG celt, VSITEMSELECTION *rgelt, ULONG *pceltFetched) override
	{
		HRESULT hr;

		Assert::AreEqual<ULONG>(1, celt, L"Only 1 supported");

		wil::unique_variant firstChild;
		hr = _hier->GetProperty(_current, VSHPROPID_FirstChild, &firstChild);
		bool hasChildren = SUCCEEDED(hr) && (firstChild.vt == VT_VSITEMID) && (V_VSITEMID(&firstChild) != VSITEMID_NIL);
		if (hasChildren)
		{
			_current = V_VSITEMID(&firstChild);
			if ((_grfItems & VSEHI_OmitHier) == 0)
				rgelt->pHier = _hier;
			rgelt->itemid = _current;
			*pceltFetched = 1;
			return S_OK;
		}

		// No children.
		wil::unique_variant next;
		hr = _hier->GetProperty(_current, VSHPROPID_NextSibling, &next);
		bool hasNext = SUCCEEDED(hr) && (next.vt == VT_VSITEMID) && (V_VSITEMID(&next) != VSITEMID_NIL);
		if (hasNext)
		{
			_current = V_VSITEMID(&next);
			if ((_grfItems & VSEHI_OmitHier) == 0)
				rgelt->pHier = _hier;
			rgelt->itemid = _current;
			*pceltFetched = 1;
			return S_OK;
		}

		// no more children
		*pceltFetched = 0;
		return S_FALSE;
	}

    virtual HRESULT STDMETHODCALLTYPE Skip (ULONG celt) override
	{
		Assert::Fail();
	}

    virtual HRESULT STDMETHODCALLTYPE Reset() override
	{
		Assert::Fail();
	}

    virtual HRESULT STDMETHODCALLTYPE Clone (IEnumHierarchyItems **ppenum) override
	{
		Assert::Fail();
	}
    #pragma endregion
};

HRESULT MakeEnumHierarchyItems (IVsHierarchy *pHierRoot, VSEHI grfItems, VSITEMID itemidRoot, IEnumHierarchyItems** ppenum)
{
	auto p = com_ptr(new (std::nothrow) MockEnumHierarchyItems()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance (pHierRoot, grfItems, itemidRoot); RETURN_IF_FAILED(hr);
	*ppenum = p.detach();
	return S_OK;;
}

