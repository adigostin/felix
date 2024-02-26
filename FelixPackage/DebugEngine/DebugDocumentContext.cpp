
#include "pch.h"
#include "DebugEngine.h"
#include "shared/TryQI.h"

class DocumentContext : public IDebugDocumentContext2
{
	ULONG _refCount = 0;

	wil::unique_bstr _doc_path;
	uint32_t _line_number;

public:
	HRESULT InitInstance (IDebugModule2* module, uint16_t address)
	{
		HRESULT hr;

		wil::com_ptr_nothrow<IZ80Module> mz80;
		hr = module->QueryInterface(&mz80); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IZ80Symbols> symbols;
		hr = mz80->GetSymbols(&symbols); RETURN_IF_FAILED_EXPECTED(hr);
		wil::unique_bstr doc_path;
		uint32_t line_number;
		hr = symbols->GetSourceLocationFromAddress(address, &doc_path, &line_number); RETURN_IF_FAILED_EXPECTED(hr);

		//wil::com_ptr_nothrow<ISimulatedProgram> program;
		//hr = mz80->GetProgram(&program); RETURN_IF_FAILED(hr);

		_doc_path = std::move(doc_path);
		_line_number = line_number;
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
		else if (riid == __uuidof(IDebugDocumentContext2))
		{
			*ppvObject = static_cast<IDebugDocumentContext2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugDocumentContext2
	virtual HRESULT STDMETHODCALLTYPE GetDocument (IDebugDocument2 **ppDocument) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetName (GETNAME_TYPE gnType, BSTR* pbstrFileName) override
	{
		if (gnType == GN_FILENAME)
		{
			*pbstrFileName = SysAllocString(_doc_path.get()); RETURN_IF_NULL_ALLOC(*pbstrFileName);
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumCodeContexts (IEnumDebugCodeContexts2 **ppEnumCodeCxts) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetLanguageInfo (BSTR* pbstrLanguage, GUID* pguidLanguage) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetStatementRange (TEXT_POSITION *pBegPosition, TEXT_POSITION *pEndPosition) override
	{
		if (pBegPosition)
		{
			pBegPosition->dwLine = _line_number;
			pBegPosition->dwColumn = 0;
		}

		if (pEndPosition)
		{
			pEndPosition->dwLine = _line_number;
			pEndPosition->dwColumn = 100;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSourceRange (TEXT_POSITION *pBegPosition, TEXT_POSITION *pEndPosition) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Compare (DOCCONTEXT_COMPARE compare, IDebugDocumentContext2** rgpDocContextSet, DWORD dwDocContextSetLen, DWORD* pdwDocContext) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Seek (int nCount, IDebugDocumentContext2 **ppDocContext) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeDocumentContext (IDebugModule2* module, uint16_t address, IDebugDocumentContext2** to)
{
	auto p = com_ptr(new (std::nothrow) DocumentContext()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(module, address); RETURN_IF_FAILED_EXPECTED(hr);
	*to = p.detach();
	return S_OK;
}
