
#include "pch.h"
#include "DebugEngine.h"
#include "../FelixPackage.h"
#include "shared/string_builder.h"
#include "shared/OtherGuids.h"
#include "shared/TryQI.h"

struct Z80CodeContext : public IDebugCodeContext2, IFelixCodeContext
{
	ULONG _refCount = 0;
	bool _physicalMemorySpace;
	UINT64 _addr;
	wil::com_ptr_nothrow<IDebugProgram2> _program;

	static HRESULT create_instance (bool physicalMemorySpace, UINT64 uCodeLocationId, IDebugProgram2* program, REFIID riid, void** ppContext)
	{
		WI_ASSERT(!physicalMemorySpace); // Not supported in many places in this file
		wil::com_ptr_nothrow<Z80CodeContext> p = new (std::nothrow) Z80CodeContext(); RETURN_IF_NULL_ALLOC(p);
		p->_physicalMemorySpace = physicalMemorySpace;
		p->_addr = uCodeLocationId;
		p->_program = program;
		if (riid == IID_IDebugCodeContext2)
			*ppContext = static_cast<IDebugCodeContext2*>(p.detach());
		else if (riid == IID_IDebugMemoryContext2)
			*ppContext = static_cast<IDebugMemoryContext2*>(p.detach());
		else
			RETURN_HR(E_NOINTERFACE);
		return S_OK;
	}
	
	#pragma region IFelixCodeContext
	virtual bool PhysicalMemorySpace() const override { return _physicalMemorySpace; }

	virtual UINT64 Address() const override { return _addr; }
	#pragma endregion
	
	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = NULL;

		if (   TryQI<IUnknown>(static_cast<IDebugMemoryContext2*>(this), riid, ppvObject)
			|| TryQI<IDebugMemoryContext2>(this, riid, ppvObject)
			|| TryQI<IDebugCodeContext2>(this, riid, ppvObject)
			|| TryQI<IFelixCodeContext>(this, riid, ppvObject))
			return S_OK;

		if (   riid == IID_IDebugCodeContext100
			|| riid == IID_IDebugCodeContext150)
			return E_NOINTERFACE;

		if (   riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
		)
			return E_NOINTERFACE;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugMemoryContext2
	virtual HRESULT __stdcall GetName(BSTR* pbstrName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	HRESULT GetAddressStr (BSTR* bstr)
	{
		wstring_builder sb;
		if (!_physicalMemorySpace)
			sb << fixed_width_hex((uint16_t)_addr) << 'h';
		else
			sb << _addr;
		
		*bstr = SysAllocStringLen (sb.data(), sb.size()); RETURN_IF_NULL_ALLOC(*bstr);
		return S_OK;
	}

	virtual HRESULT __stdcall GetInfo(CONTEXT_INFO_FIELDS dwFields, CONTEXT_INFO* pInfo) override
	{
		pInfo->dwFields = 0;

		if (dwFields & CIF_MODULEURL)
		{
			pInfo->bstrModuleUrl = SysAllocString(L"<module url>");
			pInfo->dwFields |= CIF_MODULEURL;
			dwFields &= ~CIF_MODULEURL;
		}

		if (dwFields & CIF_ADDRESS)
		{
			auto hr = GetAddressStr (&pInfo->bstrAddress); RETURN_IF_FAILED(hr);
			pInfo->dwFields |= CIF_ADDRESS;
			dwFields &= ~CIF_ADDRESS;
		}

		if (dwFields & CIF_ADDRESSOFFSET)
		{
			pInfo->bstrAddressOffset = nullptr;
			pInfo->dwFields |= CIF_ADDRESSOFFSET;
			dwFields &= ~CIF_ADDRESSOFFSET;
		}

		if (dwFields & CIF_ADDRESSABSOLUTE)
		{
			auto hr = GetAddressStr (&pInfo->bstrAddressAbsolute); RETURN_IF_FAILED(hr);
			pInfo->dwFields |= CIF_ADDRESSABSOLUTE;
			dwFields &= ~CIF_ADDRESSABSOLUTE;
		}

		if (dwFields & CIF_FUNCTION)
		{
			pInfo->bstrFunction = SysAllocString(L"<func>");
			pInfo->dwFields |= CIF_FUNCTION;
			dwFields &= ~CIF_FUNCTION;
		}

		LOG_HR_IF(E_NOTIMPL, (bool)dwFields);

		return S_OK;
	}

	virtual HRESULT __stdcall Add(UINT64 dwCount, IDebugMemoryContext2** ppMemCxt) override
	{
		UINT64 newAddr;
		if (!_physicalMemorySpace)
			newAddr = (uint16_t)(_addr + dwCount);
		else
			newAddr = _addr + dwCount;

		wil::com_ptr_nothrow<IDebugCodeContext2> cc;
		auto hr = MakeDebugContext (_physicalMemorySpace, newAddr, _program.get(), IID_PPV_ARGS(&cc)); RETURN_IF_FAILED(hr);
		hr = cc->QueryInterface(ppMemCxt); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT __stdcall Subtract(UINT64 dwCount, IDebugMemoryContext2** ppMemCxt) override
	{
		UINT64 newAddr;
		if (!_physicalMemorySpace)
			newAddr = (uint16_t)(_addr - dwCount);
		else
			newAddr = _addr - dwCount;

		wil::com_ptr_nothrow<IDebugCodeContext2> cc;
		auto hr = MakeDebugContext (_physicalMemorySpace, newAddr, _program.get(), IID_PPV_ARGS(&cc)); RETURN_IF_FAILED(hr);
		hr = cc->QueryInterface(ppMemCxt); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT __stdcall Compare(CONTEXT_COMPARE compare, IDebugMemoryContext2** rgpMemoryContextSet, DWORD dwMemoryContextSetLen, DWORD* pdwMemoryContext) override
	{
		for (DWORD i = 0; i < dwMemoryContextSetLen; i++)
		{
			wil::com_ptr_nothrow<IFelixCodeContext> other;
			auto hr = rgpMemoryContextSet[i]->QueryInterface(&other); RETURN_IF_FAILED(hr);
			switch (compare)
			{
			case CONTEXT_EQUAL:
				if ((_physicalMemorySpace == other->PhysicalMemorySpace()) && (_addr == other->Address()))
				{
					*pdwMemoryContext = i;
					return S_OK;
				}
				break;

			case CONTEXT_SAME_MODULE:
				// Always equal since we only have one module.
				*pdwMemoryContext = i;
				return S_OK;

			default:
				RETURN_HR(E_NOTIMPL);
			}
		}

		return S_FALSE;
	}
	#pragma endregion

	#pragma region IDebugCodeContext2
	virtual HRESULT __stdcall GetDocumentContext(IDebugDocumentContext2** ppSrcCxt) override
	{
		// A debug engine should return a failure code such as E_FAIL when the out parameter is null
		// such as when the code context has no associated source position.

		*ppSrcCxt = nullptr;
		com_ptr<IDebugModule2> module;
		auto hr = GetModuleAtAddress(_program.get(), _addr, &module); RETURN_IF_FAILED_EXPECTED(hr);
		hr = MakeDocumentContext (module.get(), (uint16_t)_addr, ppSrcCxt); RETURN_IF_FAILED_EXPECTED(hr);
		return S_OK;
	}

	virtual HRESULT __stdcall GetLanguageInfo(BSTR* pbstrLanguage, GUID* pguidLanguage) override
	{
		*pbstrLanguage = SysAllocString(Z80AsmLanguageName); RETURN_IF_NULL_ALLOC(*pbstrLanguage);
		*pguidLanguage = Z80AsmLanguageGuid;
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeDebugContext (bool physicalMemorySpace, UINT64 uCodeLocationId, IDebugProgram2* program, REFIID riid, void** ppContext)
{
	return Z80CodeContext::create_instance (physicalMemorySpace, uCodeLocationId, program, riid, ppContext);
}
