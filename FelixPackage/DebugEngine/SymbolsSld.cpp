
#include "pch.h"
#include "DebugEngine.h"

struct SldSymbols : IZ80Symbols
{
	ULONG _ref_count = 0;
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

	virtual ULONG STDMETHODCALLTYPE AddRef() override
	{
		return ++_ref_count;
	}

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		WI_ASSERT(_ref_count);
		_ref_count--;
		if (_ref_count)
			return _ref_count;
		delete this;
		return 0;
	}
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
				p++;
				if (*p == '\n')
					p++;

				return i == expected_field_count;
			}
			else
			{
				// end of file
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
		static const char header[] = "|SLD.data.version|1\r\n";
		if (strncmp (_text.get(), header, sizeof(header) - 1))
			RETURN_HR(E_UNRECOGNIZED_SLD_VERSION);

		sld_entry entry;
		
		{
			const char* p = _text.get() + sizeof(header) - 1;
			const char* prev_code_line = nullptr;
			
			while(true)
			{
				if (!*p)
					return E_ADDRESS_NOT_IN_SYMBOL_FILE;

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
							RETURN_HR(E_ADDRESS_NOT_IN_SYMBOL_FILE);
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

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSourceLocation(
		__RPC__in LPCWSTR src_filename,
		__RPC__in uint32_t line_index,
		__RPC__out UINT16* address_out) override
	{
		static const char header[] = "|SLD.data.version|1\r\n";
		if (strncmp (_text.get(), header, sizeof(header) - 1))
			RETURN_HR(E_UNRECOGNIZED_SLD_VERSION);

		int filename_buffer_len = WideCharToMultiByte (CP_UTF8, 0, src_filename, -1, nullptr, 0, nullptr, nullptr); RETURN_HR_IF(E_FAIL, !filename_buffer_len);
		auto u8_filename = wil::make_hlocal_ansistring_nothrow (nullptr, filename_buffer_len - 1); RETURN_IF_NULL_ALLOC(u8_filename);
		int ires = WideCharToMultiByte (CP_UTF8, 0, src_filename, -1, u8_filename.get(), filename_buffer_len, nullptr, nullptr); RETURN_HR_IF(E_FAIL, ires != filename_buffer_len);

		sld_entry entry;

		{
			const char* p = _text.get() + sizeof(header) - 1;

			while(true)
			{
				if (!*p)
					return E_ADDRESS_NOT_IN_SYMBOL_FILE;

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

	virtual HRESULT STDMETHODCALLTYPE GetSymbolAtAddress(
		__RPC__in uint16_t address,
		__RPC__in SymbolKind searchKind,
		__RPC__deref_out_opt SymbolKind* foundKind,
		__RPC__deref_out_opt BSTR* foundSymbol,
		__RPC__deref_out_opt UINT16* foundOffset) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSymbol (LPCWSTR symbolName, UINT16* address) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion
};

HRESULT MakeSldSymbols (IDebugModule2* module, IZ80Symbols** to)
{
	return SldSymbols::CreateInstance (module, to);
}
