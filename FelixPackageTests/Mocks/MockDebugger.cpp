
#include "pch.h"
#include "CppUnitTest.h"
#include "shared/com.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	struct MockDebugger : IVsDebugger
	{
		ULONG _refCount = 0;

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
            if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<IVsDebugger>(this, riid, ppvObject))
				return S_OK;

			Assert::Fail();
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IVsDebugger
        virtual HRESULT STDMETHODCALLTYPE GetMode(DBGMODE *pdbgmode) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE AdviseDebuggerEvents (IVsDebuggerEvents *psink, VSCOOKIE *pdwCookie) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE UnadviseDebuggerEvents (VSCOOKIE dwCookie) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE GetDataTipValue (IVsTextLines *pTextBuf, const TextSpan *pTS, WCHAR *pszExpression, BSTR *pbstrValue) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE QueryStatusForTextPos (
            VsTextPos *pTextPos,
            const GUID *pguidCmdGroup,
            ULONG cCmds,
            OLECMD prgCmds[  ],
            OLECMDTEXT *pCmdText) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE ExecCmdForTextPos( 
            VsTextPos *pTextPos,
            const GUID *pguidCmdGroup,
            DWORD nCmdID,
            DWORD nCmdexecopt,
            VARIANT *pvaIn,
            VARIANT *pvaOut) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE AdviseDebugEventCallback (IUnknown *punkDebuggerEvents) override { return S_OK; }

        virtual HRESULT STDMETHODCALLTYPE UnadviseDebugEventCallback (IUnknown *punkDebuggerEvents) override { return S_OK; }

        virtual HRESULT STDMETHODCALLTYPE LaunchDebugTargets (ULONG cTargets, VsDebugTargetInfo *rgDebugTargetInfo) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE InsertBreakpointByName (REFGUID guidLanguage, LPCOLESTR pszCodeLocationText) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE RemoveBreakpointsByName (REFGUID guidLanguage, LPCOLESTR pszCodeLocationText) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE ToggleBreakpointByName (REFGUID guidLanguage, LPCOLESTR pszCodeLocationText) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE IsBreakpointOnName( 
            REFGUID guidLanguage,
            LPCOLESTR pszCodeLocationText,
            BOOL *pfIsBreakpoint) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE ParseFileRedirection( 
            LPOLESTR pszArgs,
            BSTR *pbstrArgsProcessed,
            HANDLE *phStdInput,
            HANDLE *phStdOutput,
            HANDLE *phStdError) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE GetENCUpdate (IUnknown **ppUpdate) override { Assert::Fail(); }

        virtual HRESULT STDMETHODCALLTYPE AllowEditsWhileDebugging (REFGUID guidLanguageService) override { Assert::Fail(); }
        #pragma endregion
	};

    com_ptr<IVsDebugger> MakeMockDebugger()
    {
        return com_ptr(new (std::nothrow) MockDebugger());
    }
}
