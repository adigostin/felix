
#include "pch.h"
#include "shared/OtherGuids.h"
#include "shared/TryQI.h"
#include "guids.h"
#include "FelixPackage.h"

/*
class ColorableItem : IVsColorableItem, IVsHiColorItem
{
	ULONG _refCount = 0;

public:
	static HRESULT CreateInstance (IVsColorableItem** to)
	{
		wil::com_ptr_nothrow<ColorableItem> p = new (std::nothrow) ColorableItem(); RETURN_IF_NULL_ALLOC(p);
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			return E_POINTER;

		*ppvObject = NULL;

		if (riid == IID_IUnknown)
		{
			*ppvObject = static_cast<IVsColorableItem*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IVsColorableItem)
		{
			*ppvObject = static_cast<IVsColorableItem*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IVsHiColorItem)
		{
			*ppvObject = static_cast<IVsHiColorItem*>(this);
			AddRef();
			return S_OK;
		}

		if (riid == IID_IVsMergeableUIItem)
			return E_NOINTERFACE;

		if (riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_IConnectionPoint
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IMarshal
			|| riid == IID_IRpcOptions)
			return E_NOINTERFACE;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsColorableItem
	virtual HRESULT STDMETHODCALLTYPE GetDefaultColors (COLORINDEX *piForeground, COLORINDEX *piBackground) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetDefaultFontFlags (DWORD *pdwFontFlags) override
	{
		*pdwFontFlags = FF_DEFAULT;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetDisplayName (BSTR *pbstrName) override
	{
		*pbstrName = SysAllocString(L"AB12!@"); RETURN_IF_NULL_ALLOC(*pbstrName);
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsHiColorItem
	virtual HRESULT STDMETHODCALLTYPE GetColorData (VSCOLORDATA cdElement, COLORREF *pcrColor) override
	{
		if (cdElement == CD_FOREGROUND)
		{
			*pcrColor = 0x412712;
			return S_OK;
		}

		if (cdElement == CD_BACKGROUND)
		{
			*pcrColor = 0x098368;
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};
*/
class Colorizer : IVsColorizer
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IVsTextLines> _buffer;

	enum State { StateNone, StateStringSingleQuoted, StateStringDoubleQuoted, StateComment };

public:
	static HRESULT CreateInstance (IVsTextLines *pBuffer, IVsColorizer** ppColorizer)
	{
		wil::com_ptr_nothrow<Colorizer> p = new (std::nothrow) Colorizer(); RETURN_IF_NULL_ALLOC(p);
		p->_buffer = pBuffer;
		*ppColorizer = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<Colorizer*>(this), riid, ppvObject)
			|| TryQI<IVsColorizer>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == IID_IVsColorizer2)
			return E_NOINTERFACE;

		if (   riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_IConnectionPoint
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IMarshal
			|| riid == IID_IRpcOptions)
			return E_NOINTERFACE;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsColorizer
	virtual HRESULT STDMETHODCALLTYPE GetStateMaintenanceFlag (BOOL *pfFlag) override
	{
		*pfFlag = TRUE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetStartState (long *piStartState) override
	{
		*piStartState = StateNone;
		return S_OK;
	}

	virtual long STDMETHODCALLTYPE ColorizeLine (long iLine, long iLength, const WCHAR *pszText, long iState, ULONG *pAttributes) override
	{
		// Let's implement a simple parser that looks for strings and comments.
		// We'll implement a full-fledged parser later as part of an assembler.

		long i = 0;
		while (i < iLength)
		{
			if (pszText[i] == ';')
			{
				do
				{
					pAttributes[i] = COLITEM_COMMENT;
					i++;
				} while (i < iLength);
			}
			else if (pszText[i] == '\'' || pszText[i] == '\"')
			{
				wchar_t startChar = pszText[i];
				pAttributes[i] = COLITEM_STRING;
				i++;
				while (i < iLength)
				{
					if (pszText[i] == startChar)
					{
						pAttributes[i] = COLITEM_STRING;
						i++;
						break;
					}
					else if (pszText[i] == '\\')
					{
						pAttributes[i] = COLITEM_STRING;
						i++;
						if (i < iLength)
						{
							pAttributes[i] = COLITEM_STRING;
							i++;
						}
					}
					else
					{
						pAttributes[i] = COLITEM_STRING;
						i++;
					}
				}
			}
			else
			{
				pAttributes[i] = COLITEM_TEXT;
				i++;
			}
		}

		return StateNone; // We must return the colorizer's state at the end of the line.
	}

	virtual long STDMETHODCALLTYPE GetStateAtEndOfLine (long iLine, long iLength, const WCHAR *pText, long iState) override
	{
		return 0;
	}

	virtual void STDMETHODCALLTYPE CloseColorizer() override
	{
		_buffer = nullptr;
	}
	#pragma endregion
};

class Z80AsmLanguageInfo : IVsLanguageInfo/*, IVsProvideColorableItems*/, IVsLanguageDebugInfo
{
	ULONG _refCount = 0;
	//wil::com_ptr_nothrow<IVsColorableItem> _item;

public:
	static HRESULT CreateInstance (IVsLanguageInfo** to)
	{
		wil::com_ptr_nothrow<Z80AsmLanguageInfo> p = new (std::nothrow) Z80AsmLanguageInfo(); RETURN_IF_NULL_ALLOC(p);
		//auto hr = ColorableItem::CreateInstance (&p->_item); RETURN_IF_FAILED(hr);
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IVsLanguageInfo*>(this), riid, ppvObject)
			|| TryQI<IVsLanguageInfo>(this, riid, ppvObject)
			|| TryQI<IVsLanguageDebugInfo>(this, riid, ppvObject)
		)
			return S_OK;

		if (   riid == IID_IVsProvideColorableItems
			|| riid == IID_IVsLanguageDebugInfo2
			|| riid == IID_IVsLanguageDebugInfo3
			|| riid == IID_IVsLanguageTextOps
			|| riid == IID_IVsLanguageContextProvider
			|| riid == IID_IVsLanguageBlock
			|| riid == IID_IVsAutoOutliningClient
			|| riid == IID_IVsLanguageLineIndent
			|| riid == IID_IVsFormatFilterProvider
			|| riid == IID_IVsNavigableLocationResolver
			|| riid == IID_IVsLanguageDebugInfoRemap)
		{
			// These will be implemented eventually.
			return E_NOINTERFACE;
		}

		if (riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_IConnectionPoint
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IMarshal
			|| riid == IID_IRpcOptions)
		{
			// These will never be implemented.
			return E_NOINTERFACE;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsLanguageInfo
	virtual HRESULT STDMETHODCALLTYPE GetLanguageName (BSTR *bstrName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetFileExtensions (BSTR *pbstrExtensions) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetColorizer (IVsTextLines *pBuffer, IVsColorizer **ppColorizer) override
	{
		return Colorizer::CreateInstance (pBuffer, ppColorizer);
	}

	virtual HRESULT STDMETHODCALLTYPE GetCodeWindowManager (IVsCodeWindow *pCodeWin, IVsCodeWindowManager **ppCodeWinMgr) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion
	/*
	#pragma region IVsProvideColorableItems
	virtual HRESULT STDMETHODCALLTYPE GetItemCount (int* piCount) override
	{
		*piCount = 1;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetColorableItem (int iIndex, IVsColorableItem **ppItem) override
	{
		if (iIndex == 1)
			return wil::com_copy_to_nothrow(_item, ppItem);
			
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
	*/
	#pragma region IVsLanguageDebugInfo
	virtual HRESULT STDMETHODCALLTYPE GetProximityExpressions (IVsTextBuffer *pBuffer, long iLine, long iCol, long cLines, IVsEnumBSTR **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE ValidateBreakpointLocation (IVsTextBuffer *pBuffer, long iLine, long iCol, TextSpan *pCodeSpan) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetNameOfLocation (IVsTextBuffer *pBuffer, long iLine, long iCol, BSTR *pbstrName, long *piLineOffset) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetLocationOfName (LPCOLESTR pszName, BSTR *pbstrMkDoc, TextSpan *pspanLocation) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE ResolveName (LPCOLESTR pszName, DWORD dwFlags, IVsEnumDebugName **ppNames) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetLanguageID (IVsTextBuffer *pBuffer, long iLine, long iCol, GUID *pguidLanguageID) override
	{
		*pguidLanguageID = Z80AsmLanguageGuid;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE IsMappedLocation (IVsTextBuffer *pBuffer, long iLine, long iCol) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT Z80AsmLanguageInfo_CreateInstance (IVsLanguageInfo** to)
{
	return Z80AsmLanguageInfo::CreateInstance(to);
}
