
#include "pch.h"
#include "CppUnitTest.h"
#include "Z80Xml.h"
#include "shared/com.h"
#include "../Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockHierarchyEventSink : IVsHierarchyEvents
{
	ULONG _refCount = 0;
	PropChangedCallback _propChanged;
	
	MockHierarchyEventSink (PropChangedCallback propChanged)
		: _propChanged(std::move(propChanged))
	{ }

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IVsHierarchyEvents>(this, riid, ppvObject)
		)
			return S_OK;

		Assert::Fail();
	}
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsHierarchyEvents
    virtual HRESULT STDMETHODCALLTYPE OnItemAdded (VSITEMID itemidParent, VSITEMID itemidSiblingPrev, VSITEMID itemidAdded) override
	{
		return S_OK;
	}

    virtual HRESULT STDMETHODCALLTYPE OnItemsAppended (VSITEMID itemidParent) override
	{
		return S_OK;
	}

    virtual HRESULT STDMETHODCALLTYPE OnItemDeleted (VSITEMID itemid) override
	{
		return S_OK;
	}

    virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged (VSITEMID itemid, VSHPROPID propid, DWORD flags) override
	{
		if(_propChanged)
			_propChanged(itemid, propid, flags);
		return S_OK;
	}

    virtual HRESULT STDMETHODCALLTYPE OnInvalidateItems (VSITEMID itemidParent) override
	{
		return S_OK;
	}

    virtual HRESULT STDMETHODCALLTYPE OnInvalidateIcon (HICON hicon) override
	{
		return S_OK;
	}
    #pragma endregion
};

com_ptr<IVsHierarchyEvents> MakeMockHierarchyEventSink(PropChangedCallback propChanged)
{
	return com_ptr(new (std::nothrow) MockHierarchyEventSink(std::move(propChanged)));
}
