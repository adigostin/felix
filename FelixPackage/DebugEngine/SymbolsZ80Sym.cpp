
#include "pch.h"
#include "DebugEngine.h"
#include "../FelixPackage.h"
#include "shared/com.h"

struct Z80SymSymbols : IZ80Symbols
{
	ULONG _refCount = 0;
	wil::unique_hlocal_ansistring _text;

	static HRESULT CreateInstance (IDebugModule2* module, IZ80Symbols** to) noexcept
	{
		MODULE_INFO mi = { };
		auto hr = module->GetInfo (MIF_URLSYMBOLLOCATION, &mi); RETURN_IF_FAILED(hr);
		auto clear_mi = wil::scope_exit([&mi] {SysFreeString(mi.m_bstrUrlSymbolLocation); });

		wil::unique_hfile h (CreateFileW (mi.m_bstrUrlSymbolLocation, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr)); RETURN_LAST_ERROR_IF_EXPECTED(!h);
		DWORD file_size = GetFileSize (h.get(), nullptr);
		auto text = wil::make_hlocal_ansistring_nothrow(nullptr, file_size + 1); RETURN_IF_NULL_ALLOC(text);
		DWORD bytes_read;
		BOOL bres = ReadFile (h.get(), text.get(), file_size, &bytes_read, nullptr); RETURN_LAST_ERROR_IF(!bres);
		text.get()[file_size] = 0;

		auto p = wil::com_ptr_nothrow<Z80SymSymbols>(new (std::nothrow) Z80SymSymbols()); RETURN_IF_NULL_ALLOC(p);
		p->_text = std::move(text);
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
		else if (riid == __uuidof(IZ80Symbols))
		{
			*ppvObject = static_cast<IZ80Symbols*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	struct Entry
	{
		SymbolKind  kind; // when zero, means structure is not-a-value
		uint16_t    address;
		uint16_t    size; // when zero, the entry extends to the next entry, or to the module's end address
		const char* name_from;
		const char* name_to;
	};

	bool equals (const char* a, const char* a_to, LPCWSTR b)
	{
		while(true)
		{
			if (a == a_to)
				return *b == 0;
			if (*a != *b)
				return false;
			a++;
			b++;
		}
	}

	static HRESULT try_parse_line (const char*& ptr, Entry* entry) noexcept
	{
		const char* p = ptr;

		if ((p[0] != 'C' && p[0] != 'D') || (p[1] != ' ' && p[1] != 9))
			RETURN_HR(E_INVALID_Z80SYM_LINE);
		
		const char type = *p++;
		while (*p == ' ' || *p == 9)
		{
			p++;
			if (!*p)
				RETURN_HR(E_INVALID_Z80SYM_LINE);
		}

		char* endptr;
		uint16_t address = (uint16_t)strtoul(p, &endptr, 16);
		if (endptr == p)
			RETURN_HR(E_INVALID_Z80SYM_LINE);
		p = endptr;
		if (*p != ' ' && *p != 9)
			RETURN_HR(E_INVALID_Z80SYM_LINE);
		p++;
		while (*p == ' ' || *p == 9)
		{
			p++;
			if (!*p)
				RETURN_HR(E_INVALID_Z80SYM_LINE);
		}

		// Optional byte length present?
		uint16_t size = 0;
		if (isdigit(*p))
		{
			size = (uint16_t)strtoul(p, &endptr, 16);
			if (endptr == p)
				RETURN_HR(E_INVALID_Z80SYM_LINE);
			p = endptr;
			if (*p != ' ' && *p != 9)
				RETURN_HR(E_INVALID_Z80SYM_LINE);
			p++;
			while (*p == ' ' || *p == 9)
			{
				p++;
				if (!*p)
					RETURN_HR(E_INVALID_Z80SYM_LINE);
			}
		}

		// Symbol.
		if (!isalpha(p[0]))
			RETURN_HR(E_INVALID_Z80SYM_LINE);
		const char* name_from = p;
		while (isalnum(p[0]) || p[0] == '_' || p[0] == '-' || p[0] == '+' || p[0] == '$' || p[0] == '/' || p[0] == '?')
			p++;
		const char* name_to = p;
		if (*p != ' ' && *p != 9 && *p != 0x0d && *p != 0)
			RETURN_HR(E_INVALID_Z80SYM_LINE);
		while (*p == ' ' || *p == 9)
			p++;

		// Optional comment
		if (*p == ';')
		{
			p++;
			while (*p && *p != 0x0d)
				p++;
		}

		// EOL or EOF expected here.
		if (*p == 0x0d)
		{
			// skip all EOLs
			while (*p == 0x0d)
			{
				p++;
				if (*p == 0x0a)
					p++;
			}
		}
		else if (*p)
			RETURN_HR(E_INVALID_Z80SYM_LINE);

		entry->kind      = (type == 'C') ? SK_Code : SK_Data;
		entry->address   = address;
		entry->size      = size;
		entry->name_from = name_from;
		entry->name_to   = name_to;
		ptr = p;
		return S_OK;
	}

	#pragma region IZ80Symbols
	virtual HRESULT STDMETHODCALLTYPE HasSourceLocationInformation() override { return S_FALSE; }

	virtual HRESULT STDMETHODCALLTYPE GetSourceLocationFromAddress(
		__RPC__in uint16_t address,
		__RPC__deref_out BSTR* srcFilename,
		__RPC__out UINT32* srcLineIndex) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSourceLocation(
		__RPC__in LPCWSTR src_filename,
		__RPC__in uint32_t line_index,
		__RPC__out UINT16* address_out) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSymbolAtAddress(
		__RPC__in uint16_t address,
		__RPC__in SymbolKind searchKind,
		__RPC__deref_out_opt SymbolKind* foundKind,
		__RPC__deref_out_opt BSTR* foundSymbol,
		__RPC__deref_out_opt UINT16* foundOffset) override
	{
		HRESULT hr;
		
		if (searchKind == SK_None)
			RETURN_HR(E_INVALIDARG);

		static const char header[] = "Z80SYMVER1";
		if (strncmp (_text.get(), header, sizeof(header) - 1))
			RETURN_HR(E_UNRECOGNIZED_Z80SYM_VERSION);
		const char* p = _text.get() + sizeof(header) - 1;
		while (p[0] == 0x0d)
		{
			p++;
			if (p[0] == 0x0a)
				p++;
		}

		if (!foundOffset)
		{
			// Simple case: only check for an exact address match.
			while(*p)
			{
				Entry entry = { };
				hr = try_parse_line (p, &entry); RETURN_IF_FAILED(hr);
				if (entry.address == address)
				{
					if (searchKind & entry.kind)
					{
						if (foundKind)
							*foundKind = entry.kind;
						if (foundSymbol)
							RETURN_IF_FAILED(MakeBstrFromString(entry.name_from, entry.name_to, foundSymbol));
						return S_OK;
					}

					return E_ADDRESS_NOT_IN_SYMBOL_FILE;
				}

				if (entry.address > address)
					break;
			}

			return E_ADDRESS_NOT_IN_SYMBOL_FILE;
		}
		else
		{
			// Complex case: need to check if the address is within a symbol's address range
			Entry prevEntry = { };
			while (*p)
			{
				Entry entry;
				hr = try_parse_line (p, &entry); RETURN_IF_FAILED(hr);
				if (entry.address > address)
					break;
				prevEntry = entry;
			}

			if ((searchKind & prevEntry.kind)
				&& (!prevEntry.size || (address < prevEntry.address + prevEntry.size)))
			{
				if (foundKind)
					*foundKind = prevEntry.kind;
				if (foundSymbol)
					RETURN_IF_FAILED(MakeBstrFromString(prevEntry.name_from, prevEntry.name_to, foundSymbol));
				*foundOffset = address - prevEntry.address;
				return S_OK;
			}
	
			return E_ADDRESS_NOT_IN_SYMBOL_FILE;
		}
	}

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSymbol (LPCWSTR symbolName, UINT16* address) override
	{
		static const char header[] = "Z80SYMVER1";
		if (strncmp (_text.get(), header, sizeof(header) - 1))
			RETURN_HR(E_UNRECOGNIZED_Z80SYM_VERSION);
		const char* p = _text.get() + sizeof(header) - 1;
		while (p[0] == 0x0d)
		{
			p++;
			if (p[0] == 0x0a)
				p++;
		}

		while(*p)
		{
			Entry entry = { };
			auto hr = try_parse_line (p, &entry); RETURN_IF_FAILED(hr);
			if (equals(entry.name_from, entry.name_to, symbolName))
			{
				*address = entry.address;
				return S_OK;
			}
		}

		return E_SYMBOL_NOT_IN_SYMBOL_FILE;
	}
	#pragma endregion
};

HRESULT MakeZ80SymSymbols (IDebugModule2* module, IZ80Symbols** to)
{
	return Z80SymSymbols::CreateInstance (module, to);
}
