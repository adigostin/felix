
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockServiceProvider : IServiceProvider
{
	ULONG _refCount = 0;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		Assert::Fail();
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	virtual HRESULT STDMETHODCALLTYPE QueryService (REFGUID guidService, REFIID riid, void** ppvObject) override
	{
		if (guidService == SID_SVsUIShellOpenDocument)
			return E_NOTIMPL;

		if (guidService == SID_SVsRunningDocumentTable)
			return E_NOTIMPL;

		if (guidService == SID_SVsTrackProjectDocuments)
			return E_NOTIMPL;

		if (guidService == SID_SVsUIShell)
			return E_NOTIMPL;

		if (guidService == SID_SVsQueryEditQuerySave)
			return E_NOTIMPL;

		Assert::Fail();
	}
};

com_ptr<IServiceProvider> MakeMockServiceProvider()
{
	return new (std::nothrow) MockServiceProvider();
}
