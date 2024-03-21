
#include "pch.h"
#include "DebugEngine.h"
#include "../FelixPackage.h"
#include "shared/z80_register_set.h"
#include "shared/ula_register_set.h"
#include "shared/com.h"

enum class Reg { A, F, BC, DE, HL, SP, PC, IX, IY, I, R, IM, count };
static const wchar_t* const RegNames[] = { L"A", L"F", L"BC", L"DE", L"HL", L"SP", L"PC", L"IX", L"IY", L"I", L"R", L"IM" };

enum class ula_reg { frame_time, line_ticks, col_ticks, irq, count };
static inline const wchar_t* const ula_reg_names[] = { L"FRAME_TIME", L"LINE_TICKS", L"COL_TICKS", L"IRQ" };

BSTR ToString (uint8_t val)
{
	wchar_t str[8];
	swprintf_s(str, L"%02X", val);
	return SysAllocString(str);
}

BSTR ToString (uint16_t val)
{
	wchar_t str[16];
	swprintf_s(str, L"%04X", val);
	return SysAllocString(str);
}

template<typename reg_enum_t, typename register_set_t>
static HRESULT reg_to_string (reg_enum_t reg, const register_set_t& regs, BSTR* value);

template<>
static HRESULT reg_to_string (Reg reg, const z80_register_set& regs, BSTR* value)
{
	switch (reg)
	{
		case Reg::A:  *value = ToString(regs.main.a); break;
		case Reg::BC: *value = ToString(regs.main.bc); break;
		case Reg::DE: *value = ToString(regs.main.de); break;
		case Reg::HL: *value = ToString(regs.main.hl); break;
		case Reg::SP: *value = ToString(regs.sp); break;
		case Reg::PC: *value = ToString(regs.pc); break;
		case Reg::IX: *value = ToString(regs.ix); break;
		case Reg::IY: *value = ToString(regs.iy); break;
		case Reg::F:  *value = ToString(regs.main.f.val); break;
		case Reg::I:  *value = ToString(regs.i); break;
		case Reg::R:  *value = ToString(regs.r); break;
		case Reg::IM: *value = ToString(regs.im); break;
		default:
			RETURN_HR(E_NOTIMPL);
	}

	RETURN_IF_NULL_ALLOC(*value);
	
	return S_OK;
}

template<>
static HRESULT reg_to_string (ula_reg reg, const zx_spectrum_ula_regs& regs, BSTR* value)
{
	wchar_t buffer[32];
	if (reg == ula_reg::frame_time)
		swprintf_s(buffer, L"%d(dec)", regs.frame_time);
	else if (reg == ula_reg::line_ticks)
		swprintf_s(buffer, L"%d(dec)", regs.line_ticks);
	else if (reg == ula_reg::col_ticks)
		swprintf_s(buffer, L"%d(dec)", regs.col_ticks);
	else if (reg == ula_reg::irq)
		swprintf_s(buffer, regs.irq ? L"1" : L"0");
	else
		RETURN_HR(E_NOTIMPL);

	*value = SysAllocString(buffer); RETURN_IF_NULL_ALLOC(*value);
	return S_OK;
}

HRESULT string_to_u8 (const wchar_t* str, uint8_t& value)
{
	unsigned int v;
	int ires = swscanf_s (str, L"%02x", &v);
	if ((ires == 1) && (v < 256))
	{
		value = (uint8_t)v;
		return S_OK;
	}
	
	RETURN_HR(E_INVALIDARG);
}

HRESULT string_to_u16 (const wchar_t* str, uint16_t& value)
{
	unsigned int v;
	int ires = swscanf_s (str, L"%04x", &v);
	if ((ires == 1) && (v < 0x10000))
	{
		value = (uint16_t)v;
		return S_OK;
	}

	RETURN_HR(E_INVALIDARG);
}

class RegisterDebugProperty : IDebugProperty2
{
	wil::com_ptr_nothrow<IDebugThread2> _thread;
	Reg _reg;
	ULONG _refCount = 0;

public:
	static HRESULT CreateInstance (IDebugThread2* thread, Reg reg, IDebugProperty2** to)
	{
		wil::com_ptr_nothrow<RegisterDebugProperty> p = new (std::nothrow) RegisterDebugProperty(); RETURN_IF_NULL_ALLOC(p);
		p->_thread = thread;
		p->_reg = reg;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_INVALIDARG;

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
	virtual HRESULT STDMETHODCALLTYPE GetPropertyInfo(DEBUGPROP_INFO_FLAGS dwFields, DWORD dwRadix, DWORD dwTimeout, IDebugReference2** rgpArgs, DWORD dwArgCount, DEBUG_PROPERTY_INFO* pPropertyInfo) override
	{
		memset (pPropertyInfo, 0, sizeof(DEBUG_PROPERTY_INFO));

		com_ptr<ISimulator> simulator;
		auto hr = serviceProvider->QueryService(SID_Simulator, &simulator); RETURN_IF_FAILED(hr);
		z80_register_set regs;
		hr = simulator->GetRegisters(&regs, sizeof(regs)); RETURN_IF_FAILED(hr);

		auto requested = dwFields;
		if (requested & DEBUGPROP_INFO_VALUE)
		{
			hr = reg_to_string (_reg, regs, &pPropertyInfo->bstrValue); RETURN_IF_FAILED(hr);
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_VALUE;
			requested &= ~DEBUGPROP_INFO_VALUE;
		}

		if (requested & DEBUGPROP_INFO_TYPE)
		{
			pPropertyInfo->bstrType = SysAllocString(L"TYPE?!"); RETURN_IF_NULL_ALLOC(pPropertyInfo->bstrType);
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_TYPE;
			requested &= ~DEBUGPROP_INFO_TYPE;
		}

		if (requested & DEBUGPROP110_INFO_FORCE_REAL_FUNCEVAL)
		{
			pPropertyInfo->dwFields |= DEBUGPROP110_INFO_FORCE_REAL_FUNCEVAL;
			requested &= ~DEBUGPROP110_INFO_FORCE_REAL_FUNCEVAL;
		}

		if (requested && IsDebuggerPresent())
			__debugbreak();

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsString(LPCOLESTR pszValue, DWORD dwRadix, DWORD dwTimeout) override
	{
		uint16_t value;
		auto hr = string_to_u16 (pszValue, value); RETURN_IF_FAILED(hr);

		com_ptr<ISimulator> simulator;
		hr = serviceProvider->QueryService(SID_Simulator, &simulator); RETURN_IF_FAILED(hr);
		z80_register_set regs;
		hr = simulator->GetRegisters(&regs, sizeof(regs)); RETURN_IF_FAILED(hr);

		switch(_reg)
		{
			case Reg::A:  hr = string_to_u8(pszValue, regs.main.a); RETURN_IF_FAILED(hr); break;
			case Reg::BC: hr = string_to_u16(pszValue, regs.main.bc); RETURN_IF_FAILED(hr); break;
			case Reg::DE: hr = string_to_u16(pszValue, regs.main.de); RETURN_IF_FAILED(hr); break;
			case Reg::HL: hr = string_to_u16(pszValue, regs.main.hl); RETURN_IF_FAILED(hr); break;
			case Reg::SP: hr = string_to_u16(pszValue, regs.sp); RETURN_IF_FAILED(hr); break;
			case Reg::PC: hr = string_to_u16(pszValue, regs.pc); RETURN_IF_FAILED(hr); break;
			case Reg::IX: hr = string_to_u16(pszValue, regs.ix); RETURN_IF_FAILED(hr); break;
			case Reg::IY: hr = string_to_u16(pszValue, regs.iy); RETURN_IF_FAILED(hr); break;
			case Reg::F:  hr = string_to_u8(pszValue, regs.main.f.val); RETURN_IF_FAILED(hr); break;
			case Reg::I:  hr = string_to_u8(pszValue, regs.i); RETURN_IF_FAILED(hr); break;
			case Reg::R:  hr = string_to_u8(pszValue, regs.r); RETURN_IF_FAILED(hr); break;
			case Reg::IM: hr = string_to_u8(pszValue, regs.im); RETURN_IF_FAILED(hr); break;
			default:
				RETURN_HR(E_NOTIMPL);
		}

		hr = simulator->SetRegisters(&regs, sizeof(regs)); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsReference(IDebugReference2** rgpArgs, DWORD dwArgCount, IDebugReference2* pValue, DWORD dwTimeout) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumChildren(DEBUGPROP_INFO_FLAGS dwFields, DWORD dwRadix, REFGUID guidFilter, DBG_ATTRIB_FLAGS dwAttribFilter, LPCOLESTR pszNameFilter, DWORD dwTimeout, IEnumDebugPropertyInfo2** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetParent(IDebugProperty2** ppParent) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetDerivedMostProperty(IDebugProperty2** ppDerivedMost) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryBytes(IDebugMemoryBytes2** ppMemoryBytes) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryContext(IDebugMemoryContext2** ppMemory) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetSize(DWORD* pdwSize) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetReference(IDebugReference2** ppReference) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetExtendedInfo(REFGUID guidExtendedInfo, VARIANT* pExtendedInfo) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

class UlaRegisterDebugProperty : IDebugProperty2
{
	ula_reg _reg;
	ULONG _refCount = 0;

public:
	static HRESULT CreateInstance (ula_reg reg, IDebugProperty2** to)
	{
		wil::com_ptr_nothrow<UlaRegisterDebugProperty> p = new (std::nothrow) UlaRegisterDebugProperty(); RETURN_IF_NULL_ALLOC(p);
		p->_reg = reg;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_INVALIDARG;

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
	virtual HRESULT STDMETHODCALLTYPE GetPropertyInfo(DEBUGPROP_INFO_FLAGS dwFields, DWORD dwRadix, DWORD dwTimeout, IDebugReference2** rgpArgs, DWORD dwArgCount, DEBUG_PROPERTY_INFO* pPropertyInfo) override
	{
		HRESULT hr;

		memset (pPropertyInfo, 0, sizeof(DEBUG_PROPERTY_INFO));

		zx_spectrum_ula_regs regs = { };
///		hr = _simulator->GetUlaRegs(&regs); RETURN_IF_FAILED(hr);

		if (dwFields & DEBUGPROP_INFO_VALUE)
		{
			hr = reg_to_string (_reg, regs, &pPropertyInfo->bstrValue); RETURN_IF_FAILED(hr);
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_VALUE;
		}

		if (dwFields & DEBUGPROP_INFO_TYPE)
		{
			pPropertyInfo->bstrType = SysAllocString(L"TYPE?!"); RETURN_IF_NULL_ALLOC(pPropertyInfo->bstrType);
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_TYPE;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsString(LPCOLESTR pszValue, DWORD dwRadix, DWORD dwTimeout) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsReference(IDebugReference2** rgpArgs, DWORD dwArgCount, IDebugReference2* pValue, DWORD dwTimeout) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumChildren(DEBUGPROP_INFO_FLAGS dwFields, DWORD dwRadix, REFGUID guidFilter, DBG_ATTRIB_FLAGS dwAttribFilter, LPCOLESTR pszNameFilter, DWORD dwTimeout, IEnumDebugPropertyInfo2** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetParent(IDebugProperty2** ppParent) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetDerivedMostProperty(IDebugProperty2** ppDerivedMost) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryBytes(IDebugMemoryBytes2** ppMemoryBytes) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryContext(IDebugMemoryContext2** ppMemory) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetSize(DWORD* pdwSize) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetReference(IDebugReference2** ppReference) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetExtendedInfo(REFGUID guidExtendedInfo, VARIANT* pExtendedInfo) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

template<typename register_set_t, typename register_enum_t, const wchar_t* const* reg_names>
class Z80RegisterPropertyInfo : public IEnumDebugPropertyInfo2
{
	DEBUGPROP_INFO_FLAGS _dwFields;
	wil::com_ptr_nothrow<IDebugThread2> _thread;
	ULONG _refCount = 0;
	ULONG _nextIndex = 0;
	register_set_t _regs;

public:
	static HRESULT CreateInstance (IDebugThread2* thread, const register_set_t& regs, DEBUGPROP_INFO_FLAGS dwFields, IEnumDebugPropertyInfo2** to)
	{
		auto p = wil::com_ptr_nothrow<Z80RegisterPropertyInfo>(new (std::nothrow) Z80RegisterPropertyInfo()); RETURN_IF_NULL_ALLOC(p);
		p->_thread = thread;
		p->_regs = regs;
		p->_dwFields = dwFields;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_INVALIDARG;

		*ppvObject = NULL;

		if (riid == __uuidof(IUnknown))
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == __uuidof(IEnumDebugPropertyInfo2))
		{
			*ppvObject = static_cast<IEnumDebugPropertyInfo2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IEnumDebugPropertyInfo2
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG celt, DEBUG_PROPERTY_INFO* rgelt, ULONG* pceltFetched) override
	{
		*pceltFetched = 0;
		memset (rgelt, 0, celt * sizeof(DEBUG_PROPERTY_INFO));

		auto& i = *pceltFetched;
		for (i = 0; i < celt; i++)
		{
			if (_nextIndex == (ULONG)register_enum_t::count)
				return S_FALSE;

			rgelt[i].bstrName = SysAllocString(reg_names[_nextIndex]);
			rgelt[i].dwFields |= DEBUGPROP_INFO_NAME;

			auto hr = reg_to_string ((register_enum_t)_nextIndex, _regs, &rgelt[i].bstrValue); RETURN_IF_FAILED(hr);
			rgelt[i].dwFields |= DEBUGPROP_INFO_VALUE;

			//rgelt[i].bstrType = SysAllocString(L"WORD"); RETURN_IF_NULL_ALLOC(rgelt[i].bstrType);
			//rgelt[i].dwFields |= DEBUGPROP_INFO_TYPE;

			if constexpr (std::is_same_v<register_set_t, z80_register_set>)
				hr = RegisterDebugProperty::CreateInstance (_thread.get(), (register_enum_t)_nextIndex, &rgelt[i].pProperty);
			else if constexpr (std::is_same_v<register_set_t, zx_spectrum_ula_regs>)
				hr = UlaRegisterDebugProperty::CreateInstance ((register_enum_t)_nextIndex, &rgelt[i].pProperty);
			else
				hr = E_NOTIMPL;
			if (FAILED(hr))
			{
				// TODO: cleanup what's been allocated so far
				RETURN_HR(hr);
			}
			rgelt[i].dwFields |= DEBUGPROP_INFO_PROP;

			_nextIndex++;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Reset() override
	{
		_nextIndex = 0;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumDebugPropertyInfo2** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetCount(ULONG* pcelt) override
	{
		*pcelt = (ULONG)register_enum_t::count;
		return S_OK;
	}
	#pragma endregion
};

template<typename register_set_t, typename register_enum_t, const wchar_t* const* reg_names>
class Z80RegisterGroupProperty : public IDebugProperty2
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IDebugThread2> _thread;

public:
	static HRESULT CreateInstance (IDebugThread2* thread, IDebugProperty2** to)
	{
		auto p = wil::com_ptr_nothrow<Z80RegisterGroupProperty>(new (std::nothrow) Z80RegisterGroupProperty()); RETURN_IF_NULL_ALLOC(p);
		p->_thread = thread;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_INVALIDARG;

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

	// Inherited via IDebugProperty2
	virtual HRESULT STDMETHODCALLTYPE GetPropertyInfo(DEBUGPROP_INFO_FLAGS dwFields, DWORD dwRadix, DWORD dwTimeout, IDebugReference2** rgpArgs, DWORD dwArgCount, DEBUG_PROPERTY_INFO* pPropertyInfo) override
	{
		memset (pPropertyInfo, 0, sizeof(DEBUG_PROPERTY_INFO));

		auto requested = dwFields;
		if (requested & DEBUGPROP_INFO_TYPE)
		{
			pPropertyInfo->bstrType = SysAllocString(L"TYPE???");
			pPropertyInfo->dwFields |= DEBUGPROP_INFO_TYPE;
			requested &= ~DEBUGPROP_INFO_TYPE;
		}

		if (requested && IsDebuggerPresent())
			__debugbreak();

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsString(LPCOLESTR pszValue, DWORD dwRadix, DWORD dwTimeout) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetValueAsReference(IDebugReference2** rgpArgs, DWORD dwArgCount, IDebugReference2* pValue, DWORD dwTimeout) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumChildren(DEBUGPROP_INFO_FLAGS dwFields, DWORD dwRadix, REFGUID guidFilter, DBG_ATTRIB_FLAGS dwAttribFilter, LPCOLESTR pszNameFilter, DWORD dwTimeout, IEnumDebugPropertyInfo2** ppEnum) override
	{
		if (guidFilter != GUID_NULL)
			RETURN_HR(E_NOTIMPL);

		com_ptr<ISimulator> simulator;
		auto hr = serviceProvider->QueryService(SID_Simulator, &simulator); RETURN_IF_FAILED(hr);

		if constexpr (std::is_same_v<register_set_t, z80_register_set>)
		{
			z80_register_set regs;
			hr = simulator->GetRegisters(&regs, (uint32_t)sizeof(regs)); RETURN_IF_FAILED(hr);
			hr = Z80RegisterPropertyInfo<register_set_t, register_enum_t, reg_names>::CreateInstance(_thread.get(), regs, dwFields, ppEnum); RETURN_IF_FAILED(hr);
			return S_OK;
		}
		else if constexpr (std::is_same_v<register_set_t, zx_spectrum_ula_regs>)
		{
			zx_spectrum_ula_regs regs = { };
///			hr = simulator->GetUlaRegs(&regs); RETURN_IF_FAILED(hr);
			hr = Z80RegisterPropertyInfo<register_set_t, register_enum_t, reg_names>::CreateInstance(_thread.get(), regs, dwFields, ppEnum); RETURN_IF_FAILED(hr);
			return S_OK;
		}
		else
			RETURN_HR(E_NOTIMPL);

	}

	virtual HRESULT STDMETHODCALLTYPE GetParent(IDebugProperty2** ppParent) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetDerivedMostProperty(IDebugProperty2** ppDerivedMost) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryBytes(IDebugMemoryBytes2** ppMemoryBytes) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetMemoryContext(IDebugMemoryContext2** ppMemory) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetSize(DWORD* pdwSize) override
	{
		*pdwSize = 1;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetReference(IDebugReference2** ppReference) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetExtendedInfo(REFGUID guidExtendedInfo, VARIANT* pExtendedInfo) override
	{
		RETURN_HR(E_NOTIMPL);
	}
};

// ============================================================================

class Z80EnumRegisterGroupsPropertyInfo : public IEnumDebugPropertyInfo2
{
	wil::com_ptr_nothrow<IDebugThread2> _thread;
	DEBUGPROP_INFO_FLAGS _flags;
	ULONG _refCount = 0;
	ULONG _nextIndex = 0;

public:
	static HRESULT CreateInstance (IDebugThread2* thread, DEBUGPROP_INFO_FLAGS flags, IEnumDebugPropertyInfo2** to)
	{
		wil::com_ptr_nothrow<Z80EnumRegisterGroupsPropertyInfo> p = new (std::nothrow) Z80EnumRegisterGroupsPropertyInfo(); RETURN_IF_NULL_ALLOC(p);
		p->_thread = thread;
		p->_flags = flags;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_INVALIDARG;

		*ppvObject = NULL;

		if (riid == __uuidof(IUnknown))
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == __uuidof(IEnumDebugPropertyInfo2))
		{
			*ppvObject = static_cast<IEnumDebugPropertyInfo2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	static void Cleanup (ULONG celt, DEBUG_PROPERTY_INFO* rgelt)
	{
		// TODO: clean up what's been allocated so far.
	}

	#pragma region IEnumDebugPropertyInfo2
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG celt, DEBUG_PROPERTY_INFO* rgelt, ULONG* pceltFetched) override
	{
		memset (rgelt, 0, celt * sizeof(DEBUG_PROPERTY_INFO));
		*pceltFetched = 0;

		ULONG& i = *pceltFetched;
		while (true)
		{
			if (i == celt)
				return S_OK;

			rgelt[i].dwFields = 0;

			if (_nextIndex == 0)
			{
				// CPU regs
				if (_flags & DEBUGPROP_INFO_NAME)
				{
					rgelt[i].bstrName = SysAllocString(L"CPU");
					rgelt[i].dwFields |= DEBUGPROP_INFO_NAME;
				}

				if (_flags & DEBUGPROP_INFO_ATTRIB)
				{
					rgelt[i].dwAttrib = DBG_ATTRIB_OBJ_IS_EXPANDABLE | DBG_ATTRIB_VALUE_AUTOEXPANDED;
					rgelt[i].dwFields |= DEBUGPROP_INFO_ATTRIB;
				}

				if (_flags & DEBUGPROP_INFO_PROP)
				{
					auto hr = Z80RegisterGroupProperty<z80_register_set, Reg, RegNames>::CreateInstance(_thread.get(), &rgelt[i].pProperty);
					if (FAILED(hr))
					{
						Cleanup (celt, rgelt);
						RETURN_HR(hr);
					}

					rgelt[i].dwFields |= DEBUGPROP_INFO_PROP;
				}
			}
			else if (_nextIndex == 1)
			{
				// ULA regs
				if (_flags & DEBUGPROP_INFO_NAME)
				{
					rgelt[i].bstrName = SysAllocString(L"ULA");
					rgelt[i].dwFields |= DEBUGPROP_INFO_NAME;
				}

				if (_flags & DEBUGPROP_INFO_ATTRIB)
				{
					rgelt[i].dwAttrib = DBG_ATTRIB_VALUE_READONLY | DBG_ATTRIB_OBJ_IS_EXPANDABLE | DBG_ATTRIB_VALUE_AUTOEXPANDED;
					rgelt[i].dwFields |= DEBUGPROP_INFO_ATTRIB;
				}

				if (_flags & DEBUGPROP_INFO_PROP)
				{
					auto hr = Z80RegisterGroupProperty<zx_spectrum_ula_regs, ula_reg, ula_reg_names>::CreateInstance(_thread.get(), &rgelt[i].pProperty);
					if (FAILED(hr))
					{
						Cleanup (celt, rgelt);
						RETURN_HR(hr);
					}

					rgelt[i].dwFields |= DEBUGPROP_INFO_PROP;
				}
			}
			else
				return S_FALSE;

			i++;
			_nextIndex++;
		}
	}

	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Reset() override
	{
		_nextIndex = 0;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumDebugPropertyInfo2** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetCount(ULONG* pcelt) override
	{
		*pcelt = 2;
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeEnumRegisterGroupsPropertyInfo (IDebugThread2* thread, DEBUGPROP_INFO_FLAGS flags, IEnumDebugPropertyInfo2** to)
{
	return Z80EnumRegisterGroupsPropertyInfo::CreateInstance(thread, flags, to);
}