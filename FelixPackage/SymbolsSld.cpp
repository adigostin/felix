
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include <string_view>

static const char header[] = "|SLD.data.version|1\r\n";

struct SldSymbols : IFelixSymbols
{
	ULONG _refCount = 0;
	wil::unique_process_heap_string _filename;
	wil::unique_hlocal_ansistring _text;
	wil::unique_process_heap_string _sourceBasePath;

	HRESULT InitInstance (LPCWSTR symbolsFullPath, LPCWSTR sourceBasePath)
	{
		_filename = wil::make_process_heap_string_nothrow(symbolsFullPath); RETURN_IF_NULL_ALLOC(_filename);
		if (sourceBasePath && sourceBasePath[0]) {
			size_t len = wcslen(sourceBasePath);
			_sourceBasePath = wil::make_process_heap_string_nothrow(sourceBasePath, len + 1); RETURN_IF_NULL_ALLOC(_sourceBasePath);
			if (_sourceBasePath.get()[len - 1] != '\\' && _sourceBasePath.get()[len - 1] != '/')
			{
				_sourceBasePath.get()[len] = '\\';
				_sourceBasePath.get()[len + 1] = 0;
			}
		}
		HANDLE hraw = CreateFileW (symbolsFullPath, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr); RETURN_LAST_ERROR_IF_EXPECTED(hraw == INVALID_HANDLE_VALUE);
		wil::unique_hfile h (hraw);
		DWORD file_size = GetFileSize (hraw, nullptr);
		auto text = wil::make_hlocal_ansistring_nothrow(nullptr, file_size + 1); RETURN_IF_NULL_ALLOC(text);
		DWORD bytes_read;
		BOOL bres = ReadFile (hraw, text.get(), file_size, &bytes_read, nullptr); RETURN_LAST_ERROR_IF(!bres);
		text.get()[file_size] = 0;
		if (strncmp (text.get(), header, sizeof(header) - 1))
			RETURN_HR(E_UNRECOGNIZED_SLD_VERSION);

		_text = std::move(text);

		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { RETURN_HR(E_NOTIMPL); }

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

	static bool try_parse_line (const char*& p_, sld_entry& e)
	{
		const char* p = p_;

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
				if (i != expected_field_count)
					return false;

				e.end_of_line = p;
				p++;
				if (*p == '\n')
					p++;
				p_ = p;
				return true;
			}
			else
			{
				// end of file
				if (i != expected_field_count)
					return false;

				e.end_of_line = p;
				p_ = p;
				return true;
			}
		}
	}

	#pragma region IFelixSymbols
	virtual HRESULT STDMETHODCALLTYPE HasSourceLocationInformation() override { return S_OK; }

	virtual HRESULT STDMETHODCALLTYPE GetSourceLocationFromAddress(
		_In_ uint16_t address,
		_Outptr_ BSTR* pbstrSourceFilename,
		_Out_ UINT32* pSourceLineIndex,
		_Out_opt_ UINT32* pSourceRangeLineIndex) override
	{
		RETURN_HR_IF_NULL(E_POINTER, pbstrSourceFilename);
		*pbstrSourceFilename = nullptr;

		RETURN_HR_IF(E_UNEXPECTED, !_sourceBasePath);

		const char* prevSldLine;
		const char* foundSldLine;

		{
			const char* p = _text.get() + sizeof(header) - 1;
			const char* prev = nullptr;
			const char* prevprev = nullptr;
			while(true)
			{
				if (!*p)
					return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

				auto this_line = p;

				sld_entry entry;
				bool bres = try_parse_line(p, entry);
				RETURN_HR_IF_MSG(HRESULT_FROM_WIN32(ERROR_BAD_FORMAT), !bres, "%.200S", p);

				if (!strncmp (entry.type, "T|", 2))
				{
					uint16_t line_addr = (uint16_t)strtoul(entry.value, nullptr, 10);
					if (line_addr == address)
					{
						// Searched address is at the start of this line.
						prevSldLine = prev;
						foundSldLine = this_line;
						break;
					}

					if (line_addr > address)
					{
						// Searched address is after the start of the previous entry, but before this entry.
						// It can be mid-instruction, or in a multiple-instruction line with instructions separated by ':'.
						// We currently don't handle multiple-instruction lines.
						if (!prev)
							return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
						prevSldLine = prevprev;
						foundSldLine = prev;
						break;
					}

					prevprev = prev;
					prev = this_line;
				}
			}
		}

		sld_entry entry;
		bool bres = try_parse_line(foundSldLine, entry); _ASSERT(bres);

		const char* filename_end = strchr(entry.source_file, '|');
		size_t filename_len = filename_end - entry.source_file;
		int ires = MultiByteToWideChar(CP_UTF8, 0, entry.source_file, (int)filename_len, nullptr, 0); RETURN_HR_IF(E_FAIL, !ires);
		auto filenamew = wil::make_hlocal_string_nothrow(nullptr, (size_t)ires); RETURN_IF_NULL_ALLOC(filenamew);
		ires = MultiByteToWideChar (CP_UTF8, 0, entry.source_file, (int)filename_len, filenamew.get(), (int)filename_len); RETURN_HR_IF(E_FAIL, !ires);

		wchar_t filePath[MAX_PATH];
		auto res = PathCombineW (filePath, _sourceBasePath.get(), filenamew.get()); RETURN_HR_IF_NULL(E_FAIL, res);
		*pbstrSourceFilename = SysAllocString(filePath); RETURN_IF_NULL_ALLOC(*pbstrSourceFilename);
		*pSourceLineIndex = strtoul(entry.source_line, nullptr, 10) - 1;
		if (pSourceRangeLineIndex)
		{
			if (!prevSldLine) {
				*pSourceRangeLineIndex = 0;
			}
			else
			{
				sld_entry prevEntry;
				bres = try_parse_line(prevSldLine, prevEntry); _ASSERT(bres);

				// If prev entry is from another source file, we use the beginning of the found file as source range start.
				auto foundfn = std::string_view { entry.source_file, strchr(entry.source_file, '|') };
				auto prevfn = std::string_view{ prevEntry.source_file, strchr(prevEntry.source_file, '|') };
				if (foundfn != prevfn)
					*pSourceRangeLineIndex = 0;
				else
				{
					UINT32 prevEntryLineNumber = strtoul(prevEntry.source_line, nullptr, 10);
					UINT32 prevEntryLineIndex = prevEntryLineNumber - 1;
					UINT32 thisRangeLineIndex = prevEntryLineIndex + 1;
					*pSourceRangeLineIndex = thisRangeLineIndex < *pSourceLineIndex ? thisRangeLineIndex : *pSourceLineIndex;
				}
			}
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSourceLocation (
		_In_z_ LPCWSTR srcFilename,
		_In_ uint32_t lineIndex,
		_Out_ UINT64* pAddress) override
	{
		HRESULT hr;

		// sjasmplus writes in the .sld filenames as follows:
		//  - If the assembled file is in the CWD or a subdir of it, the filename is relative to the CWD.
		//    (Our extension invokes sjasmplus with the project dir as CWD.)
		//  - If the assembled file is anywhere else, the filename is absolute.

		wil::unique_process_heap_string buffer;
		if (!_sourceBasePath)
		{
			// Search for the filename we received as parameter.
			_ASSERT(false); // Should we ever be here?!
		}
		else if (PathIsRelative(srcFilename))
		{
			_ASSERT(false); // Should we ever be here?!
		}
		else
		{
			wchar_t relative[MAX_PATH];
			BOOL bRes = PathRelativePathToW (relative, _sourceBasePath.get(), FILE_ATTRIBUTE_DIRECTORY, srcFilename, 0);
			if (!bRes)
			{
				// Probably different drives. Search for the filename we received as parameter.
			}
			else if (relative[0] == '.' && relative[1] == '.'
				&& (relative[2] == 0 || relative[2] == '\\' || relative[2] == '/'))
			{
				// Same drive but outside the baseDir. Search for the filename we received as parameter.
			}
			else
			{
				// In baseDir or a subdir of it. Use the relative path.
				buffer = wil::make_process_heap_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(buffer);
				bRes = PathCanonicalize (buffer.get(), relative); RETURN_IF_WIN32_BOOL_FALSE(bRes);
				srcFilename = buffer.get();
			}
		}

		int filename_buffer_len = WideCharToMultiByte (CP_UTF8, 0, srcFilename, -1, nullptr, 0, nullptr, nullptr); RETURN_HR_IF(E_FAIL, !filename_buffer_len);
		auto u8_filename = wil::make_hlocal_ansistring_nothrow (nullptr, filename_buffer_len - 1); RETURN_IF_NULL_ALLOC(u8_filename);
		int ires = WideCharToMultiByte (CP_UTF8, 0, srcFilename, -1, u8_filename.get(), filename_buffer_len, nullptr, nullptr); RETURN_HR_IF(E_FAIL, ires != filename_buffer_len);

		// sjasmplus converts backslashes to slashes.
		for (auto p = strchr(u8_filename.get(), '\\'); p; p = strchr(p, '\\'))
			*p = '/';

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
					&& (strtoul(entry.source_line, nullptr, 10) >= lineIndex + 1))
				{
					break;
				}
			}
		}

		*pAddress = (uint16_t) strtoul (entry.value, nullptr, 10);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSymbolFromAddress (_In_ uint16_t address, _In_ SymbolKind searchKind, _Out_opt_ SymbolKind* foundKind, _Outptr_opt_ BSTR* foundSymbol, _Out_opt_ UINT16* foundOffset) override
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

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSymbol (_In_z_ LPCWSTR symbolName, _Out_ UINT16* address) override
	{
		size_t symbolNameLen = wcslen(symbolName);
		const char* p = _text.get() + sizeof(header) - 1;
		sld_entry entry;
		while(true)
		{
			if (!*p)
				return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

			auto ls = p;
			bool bres = try_parse_line(p, entry); RETURN_HR_IF_MSG(E_INVALID_SLD_LINE, !bres, "%.200S", ls);

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

	virtual HRESULT STDMETHODCALLTYPE GetFilename (_Outptr_ BSTR* pbstrFilename) override
	{
		*pbstrFilename = SysAllocString(_filename.get()); RETURN_IF_NULL_ALLOC(*pbstrFilename);
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeSldSymbols (LPCWSTR symbolsFullPath, LPCWSTR sourceBasePath, IFelixSymbols** to)
{
	auto p = com_ptr(new (std::nothrow) SldSymbols()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(symbolsFullPath, sourceBasePath); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
