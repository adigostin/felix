
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockRDT : IVsRunningDocumentTable
{
	ULONG _refCount = 0;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IVsRunningDocumentTable>(this, riid, ppvObject)
			)
			return S_OK;

		Assert::Fail();
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsRunningDocumentTable
	virtual HRESULT STDMETHODCALLTYPE RegisterAndLockDocument(
		/* [in] */ VSRDTFLAGS grfRDTLockType,
		/* [in] */ __RPC__in LPCOLESTR pszMkDocument,
		/* [in] */ __RPC__in_opt IVsHierarchy* pHier,
		/* [in] */ VSITEMID itemid,
		/* [in] */ __RPC__in_opt IUnknown* punkDocData,
		/* [out] */ __RPC__out VSCOOKIE* pdwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE LockDocument(
		/* [in] */ VSRDTFLAGS grfRDTLockType,
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UnlockDocument(
		/* [in] */ VSRDTFLAGS grfRDTLockType,
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE FindAndLockDocument(
		/* [in] */ VSRDTFLAGS dwRDTLockType,
		/* [in] */ __RPC__in LPCOLESTR pszMkDocument,
		/* [out] */ __RPC__deref_out_opt IVsHierarchy** ppHier,
		/* [out] */ __RPC__out VSITEMID* pitemid,
		/* [out] */ __RPC__deref_out_opt IUnknown** ppunkDocData,
		/* [out] */ __RPC__out VSCOOKIE* pdwCookie) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE RenameDocument(
		/* [in] */ __RPC__in LPCOLESTR pszMkDocumentOld,
		/* [in] */ __RPC__in LPCOLESTR pszMkDocumentNew,
		/* [in] */ __RPC__in_opt IVsHierarchy* pHier,
		/* [in] */ VSITEMID itemidNew) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseRunningDocTableEvents(
		/* [in] */ __RPC__in_opt IVsRunningDocTableEvents* pSink,
		/* [out] */ __RPC__out VSCOOKIE* pdwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseRunningDocTableEvents(
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetDocumentInfo(
		/* [in] */ VSCOOKIE docCookie,
		/* [out] */ __RPC__out VSRDTFLAGS* pgrfRDTFlags,
		/* [out] */ __RPC__out DWORD* pdwReadLocks,
		/* [out] */ __RPC__out DWORD* pdwEditLocks,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrMkDocument,
		/* [out] */ __RPC__deref_out_opt IVsHierarchy** ppHier,
		/* [out] */ __RPC__out VSITEMID* pitemid,
		/* [out] */ __RPC__deref_out_opt IUnknown** ppunkDocData) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE NotifyDocumentChanged(
		/* [in] */ VSCOOKIE dwCookie,
		/* [in] */ VSRDTATTRIB grfDocChanged) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE NotifyOnAfterSave(
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetRunningDocumentsEnum(
		/* [out] */ __RPC__deref_out_opt IEnumRunningDocuments** ppenum) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SaveDocuments(
		/* [in] */ VSRDTSAVEOPTIONS grfSaveOpts,
		/* [in] */ __RPC__in_opt IVsHierarchy* pHier,
		/* [in] */ VSITEMID itemid,
		/* [in] */ VSCOOKIE docCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE NotifyOnBeforeSave(
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE RegisterDocumentLockHolder(
		/* [in] */ VSREGDOCLOCKHOLDER grfRDLH,
		/* [in] */ VSCOOKIE dwCookie,
		/* [in] */ __RPC__in_opt IVsDocumentLockHolder* pLockHolder,
		/* [out] */ __RPC__out VSCOOKIE* pdwLHCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UnregisterDocumentLockHolder(
		VSCOOKIE dwLHCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE ModifyDocumentFlags(
		VSCOOKIE docCookie,
		VSRDTFLAGS grfFlags,
		BOOL fSet) override
	{
		Assert::Fail(L"Not Implemented");
	}
	#pragma endregion
};

com_ptr<IVsRunningDocumentTable> MakeMockRDT()
{
	return com_ptr(new MockRDT());
}

