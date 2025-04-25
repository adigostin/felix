
#include "pch.h"
#include "CppUnitTest.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

MIDL_INTERFACE("F761DCEE-D880-49B3-80CF-57D310DBF49B")
IMockServiceProvider : IUnknown
{
};

com_ptr<IVsSolution> MakeMockSolution();
com_ptr<IVsRunningDocumentTable> MakeMockRDT();
com_ptr<IVsShell> MakeMockShell();

struct MockServiceProvider : IServiceProvider, IMockServiceProvider
{
	ULONG _refCount = 0;
	com_ptr<IVsSolution> _solution = MakeMockSolution();
	com_ptr<IVsRunningDocumentTable> _rdt = MakeMockRDT();
	com_ptr<IVsShell> _shell = MakeMockShell();

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IMockServiceProvider>(this, riid, ppvObject)
		)
			return S_OK;

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
			return _rdt->QueryInterface(riid, ppvObject);

		if (guidService == SID_SVsTrackProjectDocuments)
			return E_NOTIMPL;

		if (guidService == SID_SVsShell)
			return _shell->QueryInterface(riid, ppvObject);

		if (guidService == SID_SVsUIShell)
			return E_NOTIMPL;

		if (guidService == SID_SVsQueryEditQuerySave)
			return E_NOTIMPL;

		if (guidService == SID_SVsSolution)
			return _solution->QueryInterface(riid, ppvObject);

		if (guidService == SID_SVsRegisterProjectTypes)
			return _solution->QueryInterface(riid, ppvObject);

		if (guidService == SID_SProfferService)
			return E_NOTIMPL;

		Assert::Fail();
	}
};

com_ptr<IServiceProvider> MakeMockServiceProvider()
{
	return new (std::nothrow) MockServiceProvider();
}
