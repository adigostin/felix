
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

		#ifdef _DEBUG
		if (   riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_IConnectionPoint
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IMarshal
			|| riid == IID_IRpcOptions)
			return E_NOINTERFACE;

		if (riid == IID_IVsColorizer2)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	struct S
	{
		const wchar_t* p;
		long len;
		ULONG* attrs;
		long i;

		bool IsLetter()
		{
			return isalpha(p[i]);
		}

		static bool IsIdStartChar (wchar_t ch)
		{
			return (ch >= '0') && (ch <= '9')
				|| (ch >= 'A') && (ch <= 'Z')
				|| (ch >= 'a') && (ch <= 'z');
		}

		static bool IsIdChar (wchar_t ch)
		{
			return IsIdStartChar(ch) || (ch == '_');
		}

		bool TryParseWhitespacesAndComments()
		{
			bool result = false;
			while (i < len)
			{
				if (p[i] == ';')
				{
					while (i < len)
						attrs[i++] = COLITEM_COMMENT;
					result = true;
					break;
				}
				else if (p[i] == ' ' || p[i] == 9)
				{
					attrs[i++] = COLITEM_TEXT;
					result = true;
				}
				else
					break;
			}

			return result;
		}

		bool TryParseString()
		{
			if (p[i] == '\'' || p[i] == '\"')
			{
				wchar_t startChar = p[i];
				attrs[i++] = COLITEM_STRING;
				while (i < len)
				{
					if (p[i] == startChar)
					{
						attrs[i++] = COLITEM_STRING;
						return true;
					}
					else if (p[i] == '\\')
					{
						attrs[i++] = COLITEM_STRING;
						if (i < len)
							attrs[i++] = COLITEM_STRING;
					}
					else
						attrs[i++] = COLITEM_STRING;
				}
			}

			return false;
		}

		static bool TryParseConditionCode (const WCHAR* p, long len, long* pii)
		{
			long ii = *pii;
			if ((ii + 1 <= len) && (ii + 1 == len || !IsIdChar(p[ii + 1])))
			{
				if (   p[ii] == 'Z' || p[ii] == 'z' || p[ii] == 'C' || p[ii] == 'c'
					|| p[ii] == 'P' || p[ii] == 'p' || p[ii] == 'M' || p[ii] == 'm')
				{
					(*pii)++;
					return true;
				}
			}

			if ((ii + 2 <= len) && (ii + 2 == len || !IsIdChar(p[ii + 2])))
			{
				if (( p[ii] == 'N' || p[ii] == 'n')
					&& (p[ii+1] == 'Z' || p[ii+1] == 'z' || p[ii+1] == 'C' || p[ii+1] == 'c'))
				{
					(*pii) += 2;
					return true;
				}

				if (( p[ii] == 'P' || p[ii] == 'p')
					&& (p[ii+1] == 'O' || p[ii+1] == 'o' || p[ii+1] == 'E' || p[ii+1] == 'e'))
				{
					(*pii) += 2;
					return true;
				}
			}

			return false;
		}

		// A lowercase 'c' is a match for a whitespace followed by a condition code.
		// Instructions with a condition code must come in this list before their simple version,
		// for example "JRc" must come before "JR".
		static inline const char* const Instructions[] = {
			"ADC", "ADD", "AND", "BIT", "CALL", "CCF", "CP", "CPD", "CPDR", "CPI", "CPIR", "CPL",
			"DAA", "DB", "DEC", "DI", "DJNZ", "EI", "EX", "EXX", "HALT", "IM", "IN",
			"INC", "IND", "INDR", "INI", "INIR", "JPc", "JP", "JRc", "JR",
			"LD", "LDD", "LDDR", "LDI", "LDIR", "NEG", "NOP",
			"OR", "OTDR", "OTIR", "OUT", "OUTD", "OUTI", "POP", "PUSH",
			"RES", "RETc", "RET", "RETI", "RETN", "RL", "RLA", "RLC", "RLCA", "RLD", "RR", "RRA", "RRC", "RRCA", "RRD", "RST",
			"SBC", "SCF", "SET", "SLA", "SLL", "SRA", "SRL", "SUB", "XOR" };

		bool TryParseInstruction()
		{
			// TODO: split in two lists, first one containing the most common instructions;
			// within each list, do binary search.
			for (const char* instr : Instructions)
			{
				long ii = i;
				bool match = true;
				while (*instr >= 'A' && *instr <= 'Z')
				{
					if (ii == len || (toupper(p[ii]) != *instr))
					{
						match = false;
						break;
					}
					ii++;
					instr++;
				}

				if (!match)
					continue;
				
				if (*instr == 0)
				{
					// Instruction without condition code.
					if (ii < len && IsIdChar(p[ii]))
						continue;

					while (i < ii)
						attrs[i++] = COLITEM_KEYWORD;
					return true;
				}

				if (*instr == 'c')
				{
					if (ii == len || (p[ii] != 32 && p[ii] != 9))
						continue;

					long instrTo = ii;
					while (ii < len && (p[ii] == 32 || p[ii] == 9))
						ii++;
					long whitespaceTo = ii;

					if (!TryParseConditionCode(p, len, &ii))
						continue;

					while (i < instrTo)
						attrs[i++] = COLITEM_KEYWORD;
					while (i < whitespaceTo)
						attrs[i++] = COLITEM_TEXT;
					while (i < ii)
						attrs[i++] = COLITEM_KEYWORD;
					return true;
				}
				else
					WI_ASSERT(false);
			}

			return false;
		}

		void ParseUnknownIdentifier()
		{
			while (isalnum(p[i]) || (p[i] == '_'))
				attrs[i++] = COLITEM_TEXT;
		}
	};

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
		S s = { .p = pszText, .len = iLength, .attrs = pAttributes, .i = 0 };
		bool parsedInstruction = false;
		while (s.i < iLength)
		{
			if (s.TryParseWhitespacesAndComments())
			{
			}
			else if (s.TryParseString())
			{
			}
			else if (s.IsLetter())
			{
				if (!parsedInstruction && s.TryParseInstruction())
				{
					parsedInstruction = true;
				}
				else
					s.ParseUnknownIdentifier();
			}
			else
				pAttributes[s.i++] = COLITEM_TEXT;
		}

		// TODO: after a // or ; comment, we must return StateNone. Currently we return StateComment

		return (iLength == 0) ? iState : (long)pAttributes[iLength - 1]; // We must return the colorizer's state at the end of the line.
	}

	virtual long STDMETHODCALLTYPE GetStateAtEndOfLine (long iLine, long iLength, const WCHAR *pText, long iState) override
	{
		return StateNone;
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
		return E_NOTIMPL;
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
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE ResolveName (LPCOLESTR pszName, DWORD dwFlags, IVsEnumDebugName **ppNames) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetLanguageID (IVsTextBuffer *pBuffer, long iLine, long iCol, GUID *pguidLanguageID) override
	{
		*pguidLanguageID = Z80AsmLanguageGuid;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE IsMappedLocation (IVsTextBuffer *pBuffer, long iLine, long iCol) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion
};

HRESULT Z80AsmLanguageInfo_CreateInstance (IVsLanguageInfo** to)
{
	return Z80AsmLanguageInfo::CreateInstance(to);
}
