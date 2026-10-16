
#include "pch.h"
#include "shared/com.h"
#include "DebugEngine.h"

HRESULT GetAddressFromSourceLocation (IDebugProgram2* program, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress)
{
	RETURN_HR_IF(E_POINTER, !pAddress);

	com_ptr<IEnumDebugModules2> modules;
	auto hr = program->EnumModules(&modules); RETURN_IF_FAILED(hr);

	com_ptr<IDebugModule2> module;
	while (SUCCEEDED(modules->Next(1, &module, nullptr)) && module)
	{
		hr = ::GetAddressFromSourceLocation(module.get(), sourceLocationFilename, sourceLocationLineIndex, pAddress);
		if (SUCCEEDED(hr))
			return S_OK;
	}

	return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

HRESULT GetAddressFromSourceLocation (IDebugModule2* module, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress)
{
	wil::com_ptr_nothrow<IZ80Module> z80m;
	auto hr = module->QueryInterface(&z80m); RETURN_IF_FAILED_EXPECTED(hr);

	wil::com_ptr_nothrow<IFelixSymbols> symbols;
	hr = z80m->GetSymbols(&symbols); RETURN_IF_FAILED_EXPECTED(hr);

	if (symbols->HasSourceLocationInformation() != S_OK)
		return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

	return symbols->GetAddressFromSourceLocation(sourceLocationFilename, sourceLocationLineIndex, pAddress);
}

HRESULT GetSourceLocationFromAddress (IDebugProgram2* program, UINT64 address, BSTR* pbstrSourceFilePath, UINT32* pSourceLineIndex, _Out_opt_ UINT32* sourceRangeLineIndex)
{
	com_ptr<IDebugModule2> module;
	auto hr = ::GetModuleAtAddress(program, address, &module);
	if (FAILED(hr))
		return hr;

	com_ptr<IZ80Module> z80Module;
	hr = module->QueryInterface(&z80Module);
	if (FAILED(hr))
		return hr;

	com_ptr<IFelixSymbols> symbols;
	hr = z80Module->GetSymbols(&symbols);
	if (FAILED(hr))
		return hr;

	hr = symbols->HasSourceLocationInformation();
	if (FAILED(hr))
		return hr;
	if (hr != S_OK)
		return E_FAIL;

	RETURN_HR_IF(E_BOUNDS, address > UINT16_MAX);

	return symbols->GetSourceLocationFromAddress ((uint16_t)address, pbstrSourceFilePath, pSourceLineIndex, sourceRangeLineIndex);
}

HRESULT GetModuleAtAddress (IDebugProgram2* program, UINT64 address, IDebugModule2** ppModule)
{
	*ppModule = nullptr;

	com_ptr<IEnumDebugModules2> modules;
	auto hr = program->EnumModules(&modules); RETURN_IF_FAILED(hr);
	com_ptr<IDebugModule2> module;
	while(SUCCEEDED(modules->Next(1, &module, nullptr)) && module)
	{
		MODULE_INFO mi;
		hr = module->GetInfo(MIF_LOADADDRESS | MIF_SIZE, &mi); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr) && (address >= mi.m_addrLoadAddress) && (address < mi.m_addrLoadAddress + mi.m_dwSize))
		{
			*ppModule = module.detach();
			return S_OK;
		}
	}

	return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

HRESULT GetSymbolFromAddress(
	__RPC__in IDebugProgram2* program,
	__RPC__in uint16_t address,
	__RPC__in SymbolKind searchKind,
	__RPC__deref_out_opt SymbolKind* foundKind,
	__RPC__deref_out_opt BSTR* foundSymbol,
	__RPC__deref_out_opt UINT16* foundOffset,
	__RPC__deref_out_opt IDebugModule2** foundModule)
{
	com_ptr<IDebugModule2> m;
	auto hr = ::GetModuleAtAddress (program, address, &m); RETURN_IF_FAILED_EXPECTED(hr);

	com_ptr<IZ80Module> z80m;
	hr = m->QueryInterface(&z80m); RETURN_IF_FAILED(hr);

	com_ptr<IFelixSymbols> syms;
	hr = z80m->GetSymbols(&syms); RETURN_IF_FAILED_EXPECTED(hr);

	hr = syms->GetSymbolFromAddress(address, searchKind, foundKind, foundSymbol, foundOffset); RETURN_IF_FAILED_EXPECTED(hr);

	if (foundModule)
		*foundModule = m.detach();

	return S_OK;
}

