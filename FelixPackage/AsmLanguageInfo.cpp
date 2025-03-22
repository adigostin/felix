
#include "pch.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
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
class Colorizer : IVsColorizer, IVsColorizer2, IVsTextLinesEvents
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IVsTextLines> _buffer;
	DWORD _textLinesEventsCookie = 0;

	enum State
	{
		StateNormal,
		StateStringSingleQuoted,
		StateStringDoubleQuoted,
		StateSingleLineComment,
		StateMultiLineComment,
	};

public:
	static HRESULT CreateInstance (IVsTextLines *pBuffer, IVsColorizer** ppColorizer)
	{
		com_ptr<Colorizer> p = new (std::nothrow) Colorizer(); RETURN_IF_NULL_ALLOC(p);
		p->_buffer = pBuffer;
		com_ptr<IConnectionPointContainer> cpc;
		auto hr = pBuffer->QueryInterface(&cpc); RETURN_IF_FAILED(hr);
		com_ptr<IConnectionPoint> cp;
		hr = cpc->FindConnectionPoint(IID_IVsTextLinesEvents, &cp); RETURN_IF_FAILED(hr);
		hr = cp->Advise(static_cast<IVsTextLinesEvents*>(p), &p->_textLinesEventsCookie); RETURN_IF_FAILED(hr);
		*ppColorizer = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IVsColorizer*>(this), riid, ppvObject)
			|| TryQI<IVsColorizer>(this, riid, ppvObject)
			|| TryQI<IVsColorizer2>(this, riid, ppvObject)
			|| TryQI<IVsTextLinesEvents>(this, riid, ppvObject)
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
		State state;
		const wchar_t* p;
		long len;
		ULONG* attrs;
		long i;

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

		void ParseMultilineComment()
		{
			WI_ASSERT(state == StateMultiLineComment);

			while (i < len)
			{
				if (i + 1 < len && p[i] == '*' && p[i+1] == '/')
				{
					if (attrs)
					{
						attrs[i] = COLITEM_COMMENT;
						attrs[i+1] = COLITEM_COMMENT;
					}
					i += 2;
					state = StateNormal;
					return;
				}

				if (attrs)
					attrs[i] = COLITEM_COMMENT;
				i++;
			}
		}

		bool TryParseWhitespacesAndComments()
		{
			bool parsed = false;
			while (i < len)
			{
				if (p[i] == ' ' || p[i] == 9)
				{
					if (attrs)
						attrs[i] = COLITEM_TEXT;
					i++;
					state = StateNormal;
					parsed = true;
				}
				else if (p[i] == ';' || (i + 1 < len && p[i] == '/' && p[i+1] == '/'))
				{
					if (attrs)
					{
						while (i < len)
							attrs[i++] = COLITEM_COMMENT;
					}
					else
						i = len;
					state = StateNormal;
					parsed = true;
				}
				else if (i + 1 < len && p[i] == '/' && p[i+1] == '*')
				{
					state = StateMultiLineComment;
					ParseMultilineComment();
					parsed = true;
				}
				else
					break;
			}

			return parsed;
		}

		bool TryParseString()
		{
			if (p[i] == '\'' || p[i] == '\"')
			{
				wchar_t startChar = p[i];
				if (attrs)
					attrs[i] = COLITEM_STRING;
				i++;
				while (i < len)
				{
					if (p[i] == startChar)
					{
						if (attrs)
							attrs[i] = COLITEM_STRING;
						i++;
						return true;
					}
					else if (p[i] == '\\')
					{
						if (attrs)
						{
							attrs[i++] = COLITEM_STRING;
							if (i < len)
								attrs[i++] = COLITEM_STRING;
						}
						else
							i = len;
					}
					else
					{
						if (attrs)
							attrs[i] = COLITEM_STRING;
						i++;
					}
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

					if (attrs)
					{
						while (i < ii)
							attrs[i++] = COLITEM_KEYWORD;
					}
					else
						i = ii;
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

					if (attrs)
					{
						while (i < instrTo)
							attrs[i++] = COLITEM_KEYWORD;
						while (i < whitespaceTo)
							attrs[i++] = COLITEM_TEXT;
						while (i < ii)
							attrs[i++] = COLITEM_KEYWORD;
					}
					else
						i = ii;
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
			{
				if (attrs)
					attrs[i] = COLITEM_TEXT;
				i++;
			}
		}

		void ParseLine()
		{
			if (state == StateMultiLineComment)
			{
				ParseMultilineComment();
				if (i == len)
					return;
			}

			bool parsedInstruction = false;
			while (i < len)
			{
				if (TryParseWhitespacesAndComments())
				{
				}
				else if (TryParseString())
				{
				}
				else if (isalpha(p[i]))
				{
					if (!parsedInstruction && TryParseInstruction())
					{
						parsedInstruction = true;
					}
					else
						ParseUnknownIdentifier();
				}
				else
				{
					if (attrs)
						attrs[i] = COLITEM_TEXT;
					i++;
				}
			}
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
		*piStartState = StateNormal;
		return S_OK;
	}

	virtual long STDMETHODCALLTYPE ColorizeLine (long iLine, long iLength, const WCHAR *pszText, long iState, ULONG *pAttributes) override
	{
		S s = { .state = (State)iState, .p = pszText, .len = iLength, .attrs = pAttributes, .i = 0 };
		s.ParseLine();
		pAttributes[iLength] = COLITEM_TEXT;
		return (long)s.state;
	}

	virtual long STDMETHODCALLTYPE GetStateAtEndOfLine (long iLine, long iLength, const WCHAR *pText, long iState) override
	{
		S s = { .state = (State)iState, .p = pText, .len = iLength, .attrs = nullptr, .i = 0 };
		s.ParseLine();
		return (long)s.state;
	}

	virtual void STDMETHODCALLTYPE CloseColorizer() override
	{
		if (_textLinesEventsCookie)
		{
			com_ptr<IConnectionPointContainer> cpc;
			auto hr = _buffer->QueryInterface(&cpc);
			if (SUCCEEDED(hr))
			{
				com_ptr<IConnectionPoint> cp;
				hr = cpc->FindConnectionPoint(IID_IVsTextLinesEvents, &cp);
				if (SUCCEEDED(hr))
				{
					cp->Unadvise(_textLinesEventsCookie);
					_textLinesEventsCookie = 0;
				}
			}
		}

		_buffer = nullptr;
	}
	#pragma endregion

	#pragma region IVsColorizer2
	virtual HRESULT STDMETHODCALLTYPE BeginColorization() override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE EndColorization() override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	// Later edit: as of 17.13.4, the workaround below doesn't seem to be necessary anymore.
	//
	// Note AGO: I had the problem that when the user was typing the beginning of
	// a multiline comment "/*", Visual Studio 2022 was calling ColorizeLine/GetStateAtEndOfLine for the
	// changed line (as expected), but not also for the following lines (unexpected).
	// The expected behavior was for VS to call ColorizeLine/GetStateAtEndOfLine also for
	// the subsequent lines, because (1) we return a new state value from ColorizeLine
	// and (2) we return TRUE in GetStateMaintenanceFlag.
	// 
	// I debugged through the decompiled code of VS2022 and found that VS does have code
	// that is meant to call ColorizeLine/GetStateAtEndOfLine for all subsequent lines
	// until a line is encountered for which our ColorizeLine/GetStateAtEndOfLine no longer
	// return a new state value (that's the end of the multiline comment).
	// That code, however, is not called by VS; it looks like it used to be called
	// at some point in past versions of VS, but then something changed and that code is now dead.
	// 
	// The code that seems to have introduced this bug is in Microsoft.VisualStudio.Editor.
	// Implementation.LanguageServiceClassificationTagger.ClassifyLine(). This function
	// modifies a variable called cachedStartLineStates to the state we return _after_
	// the text edit happens; if I skip in the debugger the code in this function that
	// modifies this variable, then everything appears to work as expected, and all subsequent
	// lines are colorized correctly by a function called FixForwardStateCache().
	//
	// As workaround, I listen for IVsTextLinesEvents and in OnChangeLineText I call ReColorizeLines
	// for many lines after the edited point. This should be enough for the vast majority of cases.
	// This worsens the editor performance, but so far it doesn't seem to be noticeable.

	#pragma region IVsTextLinesEvents
	virtual void STDMETHODCALLTYPE OnChangeLineText (const TextLineChange *pTextLineChange, BOOL fLast) override
	{
		//long lineCount;
		//auto hr = _buffer->GetLineCount(&lineCount);
		//if(FAILED(hr))
		//	return;
		//if (pTextLineChange->iNewEndLine + 1 < lineCount)
		//{
		//	// Recolorize the line just after what was inserted
		//	com_ptr<IVsTextColorState> tcs;
		//	auto hr = _buffer->QueryInterface(&tcs);
		//	if (SUCCEEDED(hr))
		//	{
		//		tcs->ReColorizeLines(pTextLineChange->iNewEndLine + 1,
		//			std::min(pTextLineChange->iNewEndLine + 100, lineCount - 1));
		//	}
		//}
	}

	virtual void STDMETHODCALLTYPE OnChangeLineAttributes (long iFirstLine, long iLastLine) override
	{
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
