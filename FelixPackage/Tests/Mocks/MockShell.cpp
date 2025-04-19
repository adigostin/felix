
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockShell : IVsShell
{
	ULONG _refCount = 0;
	wil::unique_hmodule _ui;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IVsShell>(this, riid, ppvObject)
		)
			return S_OK;

		Assert::Fail();
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsShell
	virtual HRESULT STDMETHODCALLTYPE GetPackageEnum( 
		/* [out] */ __RPC__deref_out_opt IEnumPackages **ppEnum) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty( 
		/* [in] */ VSSPROPID propid,
		/* [out] */ __RPC__out VARIANT *pvar) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty( 
		/* [in] */ VSSPROPID propid,
		/* [in] */ VARIANT var) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseBroadcastMessages( 
		/* [in] */ __RPC__in_opt IVsBroadcastMessageEvents *pSink,
		/* [out] */ __RPC__out VSCOOKIE *pdwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseBroadcastMessages( 
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseShellPropertyChanges( 
		/* [in] */ __RPC__in_opt IVsShellPropertyEvents *pSink,
		/* [out] */ __RPC__out VSCOOKIE *pdwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseShellPropertyChanges( 
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE LoadPackage( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [retval][out] */ __RPC__deref_out_opt IVsPackage **ppPackage) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE LoadPackageString( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [in] */ ULONG resid,
		/* [retval][out] */ __RPC__deref_out_opt BSTR *pbstrOut) override
	{
		Assert::IsTrue(guidPackage == IID_IProjectNodeProperties);
		if (!_ui)
		{
			_ui.reset(LoadLibrary(L"FelixPackageUi.dll"));
			Assert::IsNotNull(_ui.get());
		}

		const wchar_t* str;
		int ires = LoadString(_ui.get(), resid, (LPWSTR)&str, 0);
		if (ires == 0)
		{
			DWORD gle = GetLastError();
			Assert::Fail();
		}

		BSTR s = SysAllocStringLen(str, (UINT)ires);
		Assert::IsNotNull(s);

		*pbstrOut = s;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE LoadUILibrary( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [in] */ DWORD dwExFlags,
		/* [retval][out] */ __RPC__out DWORD_PTR *phinstOut) override
	{
		Assert::IsTrue(guidPackage == IID_IProjectNodeProperties);
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE IsPackageInstalled( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [retval][out] */ __RPC__out BOOL *pfInstalled) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE IsPackageLoaded( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [retval][out] */ __RPC__deref_out_opt IVsPackage **ppPackage) override
	{
		Assert::Fail(L"Not Implemented");
	}
	#pragma endregion
};

com_ptr<IVsShell> MakeMockShell()
{
	return com_ptr(new MockShell());
}
