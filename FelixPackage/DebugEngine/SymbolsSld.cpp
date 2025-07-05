
#include "pch.h"
#include "DebugEngine.h"
#include "shared/com.h"

static const char header[] = "|SLD.data.version|1\r\n";

struct SldSymbols : IZ80Symbols
{
	ULONG _refCount = 0;
	wil::unique_hlocal_ansistring _text;

	static HRESULT CreateInstance (const wchar_t* symbolsFullPath, IZ80Symbols** to) noexcept
	{
		HANDLE hraw = CreateFileW (symbolsFullPath, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr); RETURN_LAST_ERROR_IF_EXPECTED(hraw == INVALID_HANDLE_VALUE);
		wil::unique_hfile h (hraw);
		DWORD file_size = GetFileSize (hraw, nullptr);
		auto text = wil::make_hlocal_ansistring_nothrow(nullptr, file_size + 1); RETURN_IF_NULL_ALLOC(text);
		DWORD bytes_read;
		BOOL bres = ReadFile (hraw, text.get(), file_size, &bytes_read, nullptr); RETURN_LAST_ERROR_IF(!bres);
		text.get()[file_size] = 0;

		if (strncmp (text.get(), header, sizeof(header) - 1))
			RETURN_HR(E_UNRECOGNIZED_SLD_VERSION);

		auto p = wil::com_ptr_nothrow<SldSymbols>(new (std::nothrow) SldSymbols()); RETURN_IF_NULL_ALLOC(p);
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

	union sld_entry
	{
		struct
		{
			const char* source_file;
			const char* source_line;
			const char* def_file;
			const char* def_line;
			const char* page;
			const char* value;
			const char* type;
			const char* data;
			const char* end_of_line;
		};

		const char* fields[8];
	};

	static bool try_parse_line (const char*& p, sld_entry& e)
	{
		const size_t expected_field_count = sizeof(sld_entry::fields) / sizeof(char*);
		for (size_t i = 0;;)
		{
			if (!*p)
				return false;
			e.fields[i] = p;
			i++;
			while (*p && (*p != '|') && (*p != '\r'))
				p++;
			if (*p == '|')
			{
				if (i == expected_field_count)
					// We've parsed the number of fields we were expecting, but there are more fields on this line. Most likely an error.
					return false;
				p++;
			}
			else if (*p == '\r')
			{
				e.end_of_line = p;
				p++;
				if (*p == '\n')
					p++;

				return i == expected_field_count;
			}
			else
			{
				// end of file
				e.end_of_line = p;
				return i == expected_field_count;
			}
		}
	}

	#pragma region IZ80Symbols
	virtual HRESULT STDMETHODCALLTYPE HasSourceLocationInformation() override { return S_OK; }

	virtual HRESULT STDMETHODCALLTYPE GetSourceLocationFromAddress(
		__RPC__in uint16_t address,
		__RPC__deref_out BSTR* srcFilename,
		__RPC__out UINT32* srcLineIndex) override
	{
		sld_entry entry;
		
		{
			const char* p = _text.get() + sizeof(header) - 1;
			const char* prev_code_line = nullptr;
			
			while(true)
			{
				if (!*p)
					return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

				auto this_line = p;

				bool bres = try_parse_line(p, entry); RETURN_HR_IF(E_INVALID_SLD_LINE, !bres);

				if (!strncmp (entry.type, "T|", 2))
				{
					uint16_t line_addr = (uint16_t)strtoul(entry.value, nullptr, 10);
					if (line_addr == address)
						break;

					if (line_addr > address)
					{
						if (!prev_code_line)
							return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
						bres = try_parse_line(prev_code_line, entry); WI_ASSERT(bres);
						break;
					}

					prev_code_line = this_line;
				}
			}
		}

		const char* filename_end = strchr(entry.source_file, '|');
		size_t filename_len = filename_end - entry.source_file;
		int ires = MultiByteToWideChar(CP_UTF8, 0, entry.source_file, (int)filename_len, nullptr, 0); RETURN_HR_IF(E_FAIL, !ires);
		auto filenamew = wil::make_hlocal_string_nothrow(nullptr, (size_t)ires); RETURN_IF_NULL_ALLOC(filenamew);
		ires = MultiByteToWideChar (CP_UTF8, 0, entry.source_file, (int)filename_len, filenamew.get(), (int)filename_len); RETURN_HR_IF(E_FAIL, !ires);
		*srcFilename = SysAllocStringLen (filenamew.get(), (UINT)filename_len); RETURN_IF_NULL_ALLOC(*srcFilename);
		*srcLineIndex = strtoul(entry.source_line, nullptr, 10) - 1;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSourceLocation (LPCWSTR src_filename, uint32_t line_index, UINT16* address_out) override
	{
		int filename_buffer_len = WideCharToMultiByte (CP_UTF8, 0, src_filename, -1, nullptr, 0, nullptr, nullptr); RETURN_HR_IF(E_FAIL, !filename_buffer_len);
		auto u8_filename = wil::make_hlocal_ansistring_nothrow (nullptr, filename_buffer_len - 1); RETURN_IF_NULL_ALLOC(u8_filename);
		int ires = WideCharToMultiByte (CP_UTF8, 0, src_filename, -1, u8_filename.get(), filename_buffer_len, nullptr, nullptr); RETURN_HR_IF(E_FAIL, ires != filename_buffer_len);

		sld_entry entry;

		{
			const char* p = _text.get() + sizeof(header) - 1;

			while(true)
			{
				if (!*p)
					return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

				bool bres = try_parse_line(p, entry); RETURN_HR_IF(E_INVALID_SLD_LINE, !bres);

				if (strncmp (entry.type, "T|", 2))
					continue;

				if (!strncmp (entry.source_file, u8_filename.get(), (size_t)filename_buffer_len - 1)
					&& isdigit(entry.source_line[0])
					&& (strtoul(entry.source_line, nullptr, 10) >= line_index + 1))
				{
					break;
				}
			}
		}

		*address_out = (uint16_t) strtoul (entry.value, nullptr, 10);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSymbolAtAddress (uint16_t address, SymbolKind searchKind, SymbolKind* foundKind, BSTR* foundSymbol, UINT16* foundOffset) override
	{
		// The .sld file doesn't make a difference between code and data.
		// For example "label: nop" generates the same line in .sld as "label: db 0".
		// That's why we're not checking the "searchKind" parameter.

		const char* p = _text.get() + sizeof(header) - 1;
		const char* prev_symbol_line = nullptr;
		sld_entry entry;
		while(true)
		{
			if (!*p)
				return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

			auto this_line = p;

			bool bres = try_parse_line(p, entry); RETURN_HR_IF(E_INVALID_SLD_LINE, !bres);

			if (entry.type[0] != 'F' || entry.type[1] != '|')
				continue;
			
			uint16_t line_addr = (uint16_t)strtoul(entry.value, nullptr, 10);
			if (line_addr == address)
				break;

			if (line_addr > address)
			{
				if (!foundOffset)
					return HRESULT_FROM_WIN32(ERROR_NOT_FOUND); // caller wants exact address matches only
				if (!prev_symbol_line)
					return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
				bres = try_parse_line(prev_symbol_line, entry); WI_ASSERT(bres);
				break;
			}

			prev_symbol_line = this_line;
		}

		if (foundKind)
			*foundKind = searchKind;
		if (foundSymbol)
			RETURN_IF_FAILED(MakeBstrFromString(entry.data, entry.end_of_line, foundSymbol));
		if (foundOffset)
			*foundOffset = address - (uint16_t)strtoul(entry.value, nullptr, 10);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSymbol (LPCWSTR symbolName, UINT16* address) override
	{
		size_t symbolNameLen = wcslen(symbolName);
		const char* p = _text.get() + sizeof(header) - 1;
		sld_entry entry;
		while(true)
		{
			if (!*p)
				return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

			bool bres = try_parse_line(p, entry); RETURN_HR_IF(E_INVALID_SLD_LINE, !bres);

			if (entry.type[0] != 'L')
				continue;

			auto from = entry.data;
			while(from != entry.end_of_line && *from != ',')
				from++;
			if (from == entry.end_of_line)
				continue;
			from++;

			auto sn = symbolName;
			while(*sn)
			{
				if (*sn != *from)
					break;
				sn++;
				from++;
			}

			if (*sn == 0 && *from == ',')
				break;
		}

		*address = strtoul(entry.value, nullptr, 10);
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeSldSymbols (const wchar_t* symbolsFullPath, IZ80Symbols** to)
{
	return SldSymbols::CreateInstance (symbolsFullPath, to);
}
