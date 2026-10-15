
#include "pch.h"
#include "DebugEngine.h"
#include "shared/com.h"

class DocumentContext : public IDebugDocumentContext2
{
	ULONG _refCount = 0;

	wil::unique_bstr _doc_path;
	uint32_t _lineIndex;
	uint32_t _sourceRangeLineIndex;

public:
	HRESULT InitInstance (IDebugProgram2* program, UINT64 address)
	{
		auto hr = GetSourceLocationFromAddress (program, address, &_doc_path, &_lineIndex, &_sourceRangeLineIndex);
		if (FAILED(hr))
			return hr;
		if (hr != S_OK)
			return E_FAIL;
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
			pBegPosition->dwLine = _lineIndex;
			pBegPosition->dwColumn = 0;
		}

		if (pEndPosition)
		{
			pEndPosition->dwLine = _lineIndex;
			pEndPosition->dwColumn = 100;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSourceRange (TEXT_POSITION *pBegPosition, TEXT_POSITION *pEndPosition) override
	{
		if (pBegPosition)
		{
			pBegPosition->dwLine = _sourceRangeLineIndex;
			pBegPosition->dwColumn = 0;
		}

		if (pEndPosition)
		{
			pEndPosition->dwLine = _lineIndex;
			pEndPosition->dwColumn = 100;
		}

		return S_OK;
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

HRESULT MakeDocumentContext (IDebugProgram2* program, UINT64 address, IDebugDocumentContext2** to)
{
	auto p = com_ptr(new (std::nothrow) DocumentContext()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(program, address); RETURN_IF_FAILED_EXPECTED(hr);
	*to = p.detach();
	return S_OK;
}
