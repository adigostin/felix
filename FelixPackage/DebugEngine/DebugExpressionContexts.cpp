
#include "pch.h"
#include "DebugEngine.h"
#include "shared/com.h"
#include "guids.h"
#include "..\FelixPackageUi\resource.h"

class ExpressionContextNoSource : public IDebugExpressionContext2
{
	ULONG _refCount = 0;
	com_ptr<IDebugThread2> _thread;

public:
	HRESULT InitInstance (IDebugThread2* thread)
	{
		_thread = thread;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IDebugExpressionContext2>(this, riid, ppvObject))
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugExpressionContext2
	virtual HRESULT STDMETHODCALLTYPE GetName (BSTR *pbstrName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	static bool compare (LPCOLESTR pszCode, const char* str)
	{
		while (*pszCode && *str)
		{
			if ((*pszCode & 0xDF) != (*str & 0xDF))
				return false;
			pszCode++;
			str++;
		}

		return !*pszCode && !*str;
	}

	virtual HRESULT STDMETHODCALLTYPE ParseText (LPCOLESTR pszCode, PARSEFLAGS dwFlags, UINT nRadix,
		IDebugExpression2 **ppExpr, BSTR* pbstrError, UINT *pichError) override
	{
		HRESULT hr;

		*ppExpr = nullptr;
		*pbstrError = nullptr;
		*pichError = 0;

		for (uint8_t i = 0; i < (uint8_t)z80_reg16::count; i++)
		{
			if (compare(pszCode, z80_reg16_names[i]))
				return MakeRegisterExpression (_thread.get(), pszCode, (z80_reg16)i, ppExpr);
		}

		if (iswdigit(pszCode[0]))
		{
			DWORD val;
			hr = ParseNumber(pszCode, &val);
			if (hr != S_OK)
			{
				com_ptr<IVsShell> shell;
				serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell));
				wil::unique_bstr str;
				if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_UNK_NUM_FORMAT, &str)))
					*pbstrError = str.release();
				*pichError = 0;
				return E_FAIL;
			}

			com_ptr<IDebugProgram2> program;
			hr = _thread->GetProgram(&program); RETURN_IF_FAILED(hr);
			bool physicalMemorySpace = false;
			return MakeNumberExpression (pszCode, physicalMemorySpace, val, program, ppExpr);
		}

		// VS calls this function all the time when the user hovers the mouse in the editor
		// while debugging. Let's not load any detailed error message.
		//com_ptr<IVsShell> shell;
		//serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell));
		//wil::unique_bstr str;
		//if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_UNK_EXPRESSION, &str)))
		//	*pbstrError = str.release();
		//*pichError = 0;
		return E_FAIL;
	}
	#pragma endregion
};

HRESULT MakeExpressionContextNoSource (IDebugThread2* thread, IDebugExpressionContext2** to)
{
	auto p = com_ptr(new (std::nothrow) ExpressionContextNoSource()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(thread); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}