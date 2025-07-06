
#include "pch.h"
#include "CppUnitTest.h"
#include "CppUnitTestLogger.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	struct MockOutputWindowPane : IVsOutputWindowPane, IVsOutputWindowPane2
	{
		ULONG _refCount = 0;
		com_ptr<IStream> _outputStreamUTF16;

		HRESULT InitInstance (IStream* outputStreamUTF16)
		{
			_outputStreamUTF16 = outputStreamUTF16;
			return S_OK;
		}

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(static_cast<IVsOutputWindowPane*>(this), riid, ppvObject)
				|| TryQI<IVsOutputWindowPane>(this, riid, ppvObject)
				|| TryQI<IVsOutputWindowPane2>(this, riid, ppvObject)
			)
				return S_OK;

			return E_NOINTERFACE;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IVsOutputWindowPane
		virtual HRESULT STDMETHODCALLTYPE OutputString (LPCOLESTR pszOutputString) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE Activate() override { return S_OK; }

		virtual HRESULT STDMETHODCALLTYPE Hide( void) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE Clear( void) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE FlushToTaskList( void) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE OutputTaskItemString( 
			/* [in] */ __RPC__in LPCOLESTR pszOutputString,
			/* [in] */ VSTASKPRIORITY nPriority,
			/* [in] */ VSTASKCATEGORY nCategory,
			/* [in] */ __RPC__in LPCOLESTR pszSubcategory,
			/* [in] */ VSTASKBITMAP nBitmap,
			/* [in] */ __RPC__in LPCOLESTR pszFilename,
			/* [in] */ ULONG nLineNum,
			/* [in] */ __RPC__in LPCOLESTR pszTaskItemText) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE OutputTaskItemStringEx( 
			/* [in] */ __RPC__in LPCOLESTR pszOutputString,
			/* [in] */ VSTASKPRIORITY nPriority,
			/* [in] */ VSTASKCATEGORY nCategory,
			/* [in] */ __RPC__in LPCOLESTR pszSubcategory,
			/* [in] */ VSTASKBITMAP nBitmap,
			/* [in] */ __RPC__in LPCOLESTR pszFilename,
			/* [in] */ ULONG nLineNum,
			/* [in] */ __RPC__in LPCOLESTR pszTaskItemText,
			/* [in] */ __RPC__in LPCOLESTR pszLookupKwd) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE GetName( 
			__RPC__deref_in_opt BSTR *pbstrPaneName) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE SetName( 
			__RPC__in LPCOLESTR pszPaneName) override { RETURN_HR(E_NOTIMPL); }

		virtual HRESULT STDMETHODCALLTYPE OutputStringThreadSafe( 
			/* [in] */ __RPC__in LPCOLESTR pszOutputString) override { RETURN_HR(E_NOTIMPL); }
		#pragma endregion

		#pragma region IVsOutputWindowPane2
		virtual HRESULT STDMETHODCALLTYPE OutputTaskItemStringEx2( 
			/* [in] */ __RPC__in LPCOLESTR pszOutputString,
			/* [in] */ VSTASKPRIORITY nPriority,
			/* [in] */ VSTASKCATEGORY nCategory,
			/* [in] */ __RPC__in LPCOLESTR pszSubcategory,
			/* [in] */ VSTASKBITMAP nBitmap,
			/* [in] */ __RPC__in LPCOLESTR pszFilename,
			/* [in] */ ULONG nLineNum,
			/* [in] */ ULONG nColumn,
			/* [in] */ __RPC__in LPCOLESTR pszProjectUniqueName,
			/* [in] */ __RPC__in LPCOLESTR pszTaskItemText,
			/* [in] */ __RPC__in LPCOLESTR pszLookupKwd) override
		{
			if (!_outputStreamUTF16)
			{
				size_t len = wcslen(pszOutputString);
				if (len >= 2 && pszOutputString[len - 2] == 0x0d && pszOutputString[len - 1] == 0x0A)
				{
					auto copy = wil::make_hlocal_string_nothrow(pszOutputString, len - 2);
					Logger::WriteMessage(copy.get());
				}
				else
					Logger::WriteMessage(pszOutputString);
				return S_OK;
			}

			return _outputStreamUTF16->Write(pszOutputString, (ULONG)wcslen(pszOutputString) * 2, NULL);
		}
		#pragma endregion
	};

	com_ptr<IVsOutputWindowPane2> MakeMockOutputWindowPane (IStream* outputStreamUTF16)
	{
		auto p = com_ptr (new (std::nothrow) MockOutputWindowPane());
		Assert::IsNotNull(p.get());
		auto hr = p->InitInstance(outputStreamUTF16);
		Assert::IsTrue(SUCCEEDED(hr));
		return p;
	}
}
