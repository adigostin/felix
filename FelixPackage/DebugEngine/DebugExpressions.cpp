
#include "pch.h"
#include "DebugEngine.h"
#include "shared/com.h"
#include "../FelixPackage.h"

class RegisterExpression : public IDebugExpression2
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IDebugThread2> _thread;
	wil::unique_bstr _originalText;
	z80_reg16 _reg;

public:
	HRESULT InitInstance (IDebugThread2* thread, LPCOLESTR originalText, z80_reg16 reg)
	{
		_thread = thread;
		_originalText = wil::make_bstr_nothrow(originalText); RETURN_IF_NULL_ALLOC(_originalText);
		_reg = reg;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IDebugExpression2>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugExpression2
	virtual HRESULT STDMETHODCALLTYPE EvaluateAsync (EVALFLAGS dwFlags, IDebugEventCallback2 *pExprCallback) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE Abort() override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE EvaluateSync (EVALFLAGS dwFlags, DWORD dwTimeout, IDebugEventCallback2 *pExprCallback, IDebugProperty2 **ppResult) override
	{
		com_ptr<ISimulator> simulator;
		auto hr = serviceProvider->QueryService(SID_Simulator, &simulator); RETURN_IF_FAILED(hr);
		com_ptr<IDebugProgram2> program;
		hr = _thread->GetProgram(&program); RETURN_IF_FAILED(hr);
		z80_register_set regs;
		hr = simulator->GetRegisters(&regs, sizeof(regs)); RETURN_IF_FAILED(hr);
		uint16_t value = regs.reg(_reg);
		hr = MakeNumberProperty (_originalText.get(), false, value, program, ppResult); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeRegisterExpression (IDebugThread2* thread, LPCOLESTR originalText, z80_reg16 reg, IDebugExpression2** to)
{
	auto p = com_ptr(new (std::nothrow) RegisterExpression()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(thread, originalText, reg); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}

// ============================================================================

class NumberProperty : IDebugProperty2
{
	ULONG _refCount = 0;
	wil::unique_bstr _originalText;
	bool _physicalMemorySpace;
	UINT64 _value;
	wil::com_ptr_nothrow<IDebugProgram2> _program;

public:
	static HRESULT CreateInstance (LPCOLESTR originalText, bool physicalMemorySpace, UINT64 value, IDebugProgram2* program, IDebugProperty2** to)
	{
		wil::com_ptr_nothrow<NumberProperty> p = new (std::nothrow) NumberProperty(); RETURN_IF_NULL_ALLOC(p);
		p->_originalText.reset (SysAllocString(originalText)); RETURN_IF_NULL_ALLOC(p->_originalText);
		p->_physicalMemorySpace = physicalMemorySpace;
		p->_value = value;
		p->_program = program;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_POINTER;

		*ppvObject = NULL;

		if (riid == __uuidof(IUnknown))
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == __uuidof(IDebugProperty2))
		{
			*ppvObject = static_cast<IDebugProperty2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugProperty2
	virtual HRESULT STDMETHODCALLTYPE GetPropertyInfo( 
		/* [in] */ DEBUGPROP_INFO_FLAGS dwFields,
		/* [in] */ DWORD dwRadix,
		/* [in] */ DWORD dwTimeout,
		/* [length_is][size_is][full][in] */ __RPC__in_ecount_part_opt(dwArgCount, dwArgCount) IDebugReference2 **rgpArgs,
		/* [in] */ DWORD dwArgCount,
		/* [out] */ __RPC__out DEBUG_PROPERTY_INFO *pPropertyInfo) override
	{
		pPropertyInfo->dwFields = 0;

		if (dwFields & DEBUGPROP_INFO_TYPE)
		{
			pPropertyInfo->bstrType = SysAllocString(L"Number");
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_TYPE;
		}

		if (dwFields & DEBUGPROP_INFO_VALUE)
		{
			pPropertyInfo->bstrValue = SysAllocString (_originalText.get());
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_VALUE;
		}

		if (dwFields & DEBUGPROP_INFO_ATTRIB)
		{
			pPropertyInfo->dwAttrib = DBG_ATTRIB_VALUE_READONLY;
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_ATTRIB;
		}

		if (dwFields & DEBUGPROP_INFO_PROP)
		{
			pPropertyInfo->pProperty = nullptr;
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_PROP;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsString( 
		/* [in] */ __RPC__in LPCOLESTR pszValue,
		/* [in] */ DWORD dwRadix,
		/* [in] */ DWORD dwTimeout) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsReference( 
		/* [length_is][size_is][full][in] */ __RPC__in_ecount_part_opt(dwArgCount, dwArgCount) IDebugReference2 **rgpArgs,
		/* [in] */ DWORD dwArgCount,
		/* [in] */ __RPC__in_opt IDebugReference2 *pValue,
		/* [in] */ DWORD dwTimeout) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumChildren( 
		/* [in] */ DEBUGPROP_INFO_FLAGS dwFields,
		/* [in] */ DWORD dwRadix,
		/* [in] */ __RPC__in REFGUID guidFilter,
		/* [in] */ DBG_ATTRIB_FLAGS dwAttribFilter,
		/* [full][in] */ __RPC__in_opt LPCOLESTR pszNameFilter,
		/* [in] */ DWORD dwTimeout,
		/* [out] */ __RPC__deref_out_opt IEnumDebugPropertyInfo2 **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetParent( 
		/* [out] */ __RPC__deref_out_opt IDebugProperty2 **ppParent) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetDerivedMostProperty( 
		/* [out] */ __RPC__deref_out_opt IDebugProperty2 **ppDerivedMost) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryBytes( 
		/* [out] */ __RPC__deref_out_opt IDebugMemoryBytes2 **ppMemoryBytes) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryContext (IDebugMemoryContext2 **ppMemory) override
	{
		return MakeDebugContext (_physicalMemorySpace, _value, _program.get(), IID_PPV_ARGS(ppMemory));
	}

	virtual HRESULT STDMETHODCALLTYPE GetSize( 
		/* [out] */ __RPC__out DWORD *pdwSize) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetReference( 
		/* [out] */ __RPC__deref_out_opt IDebugReference2 **ppReference) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetExtendedInfo( 
		/* [in] */ __RPC__in REFGUID guidExtendedInfo,
		/* [out] */ __RPC__out VARIANT *pExtendedInfo) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

class NumberExpression : IDebugExpression2
{
	ULONG _refCount = 0;
	bool _physicalMemorySpace;
	uint32_t _value;
	wil::unique_bstr _originalText;
	wil::com_ptr_nothrow<IDebugProgram2> _program;

public:
	static HRESULT CreateInstance (LPCOLESTR originalText, bool physicalMemorySpace, uint32_t value, IDebugProgram2* program, IDebugExpression2** to)
	{
		wil::com_ptr_nothrow<NumberExpression> p = new (std::nothrow) NumberExpression(); RETURN_IF_NULL_ALLOC(p);
		p->_originalText.reset(SysAllocString(originalText)); RETURN_IF_NULL_ALLOC(p->_originalText);
		p->_physicalMemorySpace = physicalMemorySpace;
		p->_value = value;
		p->_program = program;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_POINTER;

		*ppvObject = NULL;

		if (riid == __uuidof(IUnknown))
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == __uuidof(IDebugExpression2))
		{
			*ppvObject = static_cast<IDebugExpression2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugExpression2
	virtual HRESULT STDMETHODCALLTYPE EvaluateAsync (EVALFLAGS dwFlags, IDebugEventCallback2 *pExprCallback) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Abort() override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EvaluateSync (EVALFLAGS dwFlags, DWORD dwTimeout, IDebugEventCallback2 *pExprCallback, IDebugProperty2 **ppResult) override
	{
		return MakeNumberProperty (_originalText.get(), _physicalMemorySpace, _value, _program.get(), ppResult);
	}
	#pragma endregion
};

HRESULT MakeNumberExpression (LPCOLESTR originalText, bool physicalMemorySpace, uint32_t value, IDebugProgram2* program, IDebugExpression2** to)
{
	return NumberExpression::CreateInstance (originalText, physicalMemorySpace, value, program, to);
}

HRESULT MakeNumberProperty (LPCOLESTR originalText, bool physicalMemorySpace, UINT64 value, IDebugProgram2* program, IDebugProperty2** to)
{
	return NumberProperty::CreateInstance (originalText, physicalMemorySpace, value, program, to);
}
