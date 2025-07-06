
#include "pch.h"
#include "CppUnitTest.h"
#include "shared/com.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

static const GUID SID_SVsGeneralOutputWindowPane = { 0x65482c72, 0xdefa, 0x41b7, 0x90, 0x2c, 0x11, 0xc0, 0x91, 0x88, 0x9c, 0x83 };

MIDL_INTERFACE("F761DCEE-D880-49B3-80CF-57D310DBF49B")
IMockServiceProvider : IUnknown
{
};

com_ptr<IVsSolution> MakeMockSolution();
com_ptr<IVsRunningDocumentTable> MakeMockRDT();
com_ptr<IVsShell> MakeMockShell();

namespace FelixTests
{
	struct MockServiceProvider : IServiceProvider, IMockServiceProvider
	{
		ULONG _refCount = 0;
		com_ptr<IVsSolution> _solution = MakeMockSolution();
		com_ptr<IVsRunningDocumentTable> _rdt = MakeMockRDT();
		com_ptr<IVsShell> _shell = MakeMockShell();
		com_ptr<IVsOutputWindowPane2> _generalOutputWindowPane = MakeMockOutputWindowPane(nullptr);

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IMockServiceProvider>(this, riid, ppvObject)
			)
				return S_OK;

			Assert::Fail();
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override
		{
			return ++_refCount;
		}

		virtual ULONG STDMETHODCALLTYPE Release() override
		{
			return ReleaseST(this, _refCount);
		}
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

			if (guidService == SID_SVsSolutionBuildManager)
				return _solution->QueryInterface(riid, ppvObject);

			if (guidService == SID_SVsGeneralOutputWindowPane)
				return _generalOutputWindowPane->QueryInterface(riid, ppvObject);

			Assert::Fail();
		}
	};

	com_ptr<IServiceProvider> MakeMockServiceProvider()
	{
		return new (std::nothrow) MockServiceProvider();
	}
}
