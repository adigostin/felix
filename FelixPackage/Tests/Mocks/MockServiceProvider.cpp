
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

HRESULT MakeEnumHierarchyItems (IVsHierarchy *pHierRoot, VSEHI grfItems, VSITEMID itemidRoot, IEnumHierarchyItems** ppenum);

struct MockServiceProvider : IServiceProvider, IVsEnumHierarchyItemsFactory
{
	ULONG _refCount = 0;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (TryQI<IVsEnumHierarchyItemsFactory>(this, riid, ppvObject))
			return S_OK;

		Assert::Fail();
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	virtual HRESULT STDMETHODCALLTYPE QueryService (REFGUID guidService, REFIID riid, void** ppvObject) override
	{
		if (guidService == SID_SVsEnumHierarchyItemsFactory)
			return QueryInterface(riid, ppvObject);

		Assert::Fail();
	}

	// IVsEnumHierarchyItemsFactory
	virtual HRESULT STDMETHODCALLTYPE EnumHierarchyItems (IVsHierarchy *pHierRoot,
		VSEHI grfItems, VSITEMID itemidRoot, IEnumHierarchyItems** ppenum) override
	{
		return MakeEnumHierarchyItems (pHierRoot, grfItems, itemidRoot, ppenum);
	}
};

com_ptr<IServiceProvider> MakeMockServiceProvider()
{
	return new (std::nothrow) MockServiceProvider();
}
