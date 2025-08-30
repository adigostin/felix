
#include "pch.h"
#include "CppUnitTest.h"
#include "shared/com.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	struct MockFileChangeEx : IVsFileChangeEx
	{
		ULONG _refCount = 0;

		MockFileChangeEx()
		{
		}

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(this, riid, ppvObject)
				|| TryQI<IVsFileChangeEx>(this, riid, ppvObject)
				)
				return S_OK;

			Assert::Fail();
		}
		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IVsFileChangeEx
		virtual HRESULT STDMETHODCALLTYPE AdviseFileChange (LPCOLESTR pszMkDocument, VSFILECHANGEFLAGS grfFilter, IVsFileChangeEvents *pFCE, VSCOOKIE *pvsCookie) override
		{
			Assert::Fail();
		}

		virtual HRESULT STDMETHODCALLTYPE UnadviseFileChange (VSCOOKIE vsCookie) override
		{
			Assert::Fail();
		}

		virtual HRESULT STDMETHODCALLTYPE SyncFile (LPCOLESTR pszMkDocument) override
		{
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE IgnoreFile (VSCOOKIE vsCookie, LPCOLESTR pszMkDocument, BOOL fIgnore) override
		{
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE AdviseDirChange (LPCOLESTR pszDir, BOOL fWatchSubDir, IVsFileChangeEvents *pFCE, VSCOOKIE *pvsCookie) override
		{
			Assert::Fail();
		}

		virtual HRESULT STDMETHODCALLTYPE UnadviseDirChange (VSCOOKIE vsCookie) override
		{
			Assert::Fail();
		}
		#pragma endregion
	};

	com_ptr<IVsFileChangeEx> MakeMockFileChangeEx()
	{
		return com_ptr<IVsFileChangeEx>(new MockFileChangeEx());
	}
}
