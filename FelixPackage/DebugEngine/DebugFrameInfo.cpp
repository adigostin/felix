
#include "pch.h"
#include "DebugEngine.h"
#include "../FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/z80_register_set.h"
#include "shared/vector_nothrow.h"
#include "shared/string_builder.h"
#include "shared/com.h"

class Z80StackFrame : public IDebugStackFrame2
{
	ULONG _refCount = 0;
	wil::com_ptr_nothrow<IDebugThread2> _thread;
	wil::com_ptr_nothrow<IDebugModule2> _module;
	uint16_t _pc;

public:
	HRESULT InitInstance (IDebugThread2* thread, IDebugModule2* module, uint16_t pc)
	{
		_thread = thread;
		_module = module;
		_pc = pc;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDebugStackFrame2*>(this), riid, ppvObject)
			|| TryQI<IDebugStackFrame2>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		if (riid == __uuidof(IDebugStackFrame110)
			//|| riid == __uuidof(Microsoft::VisualStudio::Debugger::CallStack::IVsWrappedDkmStackFrame)
			|| riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_INoMarshal
			|| riid == IID_IMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == GUID{ 0x7365A6C9, 0x96A6, 0x45A5, { 0x9B, 0x01, 0xFF, 0x6D, 0xF3, 0x85, 0x80, 0xD2 } }
		)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugStackFrame2
	virtual HRESULT __stdcall GetCodeContext(IDebugCodeContext2** ppCodeCxt) override
	{
		// Returns an IDebugCodeContext2 object that represents the current instruction pointer in this stack frame.
		com_ptr<IDebugProgram2> program;
		auto hr = _thread->GetProgram(&program); RETURN_IF_FAILED(hr);
		bool physicalMemorySpace = false;
		return MakeDebugContext (physicalMemorySpace, _pc, program.get(), IID_IDebugCodeContext2, (void**)ppCodeCxt);
	}

	virtual HRESULT __stdcall GetDocumentContext(IDebugDocumentContext2** ppCxt) override
	{
		// This method is faster than calling the GetCodeContext method and then calling
		// the GetDocumentContext method on the code context. However, it is not guaranteed
		// that every debug engine (DE) will implement this method.
		return E_NOTIMPL;
	}

	virtual HRESULT __stdcall GetName(BSTR* pbstrName) override
	{
		*pbstrName = SysAllocString(L"<stack frame name>");
		return *pbstrName ? S_OK : E_OUTOFMEMORY;
	}

	virtual HRESULT __stdcall GetInfo(FRAMEINFO_FLAGS dwFieldSpec, UINT nRadix, FRAMEINFO* pFrameInfo) override
	{
		pFrameInfo->m_dwValidFields = 0;

		if (dwFieldSpec & FIF_DEBUG_MODULEP)
		{
			if (_module)
			{
				pFrameInfo->m_pModule = _module.get();
				pFrameInfo->m_pModule->AddRef();
				pFrameInfo->m_dwValidFields |= FIF_DEBUG_MODULEP;
			}

			dwFieldSpec &= ~FIF_DEBUG_MODULEP;
		}

		if (dwFieldSpec & FIF_STACKRANGE)
		{
			pFrameInfo->m_addrMin = 0xFFF0;
			pFrameInfo->m_addrMax = 0xFFFF;
			pFrameInfo->m_dwValidFields |= FIF_STACKRANGE;
			dwFieldSpec &= ~FIF_STACKRANGE;
		}

		if (dwFieldSpec & FIF_DEBUGINFO)
		{
			// I stepped through decompiled VS2022 code and I noticed that if I set m_fHasDebugInfo to FALSE,
			// VS tries to convert the module to some internal DotNet class named IDebugModule110;
			// it cannot do the conversion since our modules are native C++; then VS unconditionally generates 
			// a symbol filename with the ".pdb" extension and complains about it, for example "Spectrum48K.pdb".
			// 
			// (There may be other bad consequences when setting this uncoditionally to TRUE. I haven't noticed them yet.)
			pFrameInfo->m_fHasDebugInfo = TRUE;
			pFrameInfo->m_dwValidFields |= FIF_DEBUGINFO;
			dwFieldSpec &= ~FIF_DEBUGINFO;
		}

		if (dwFieldSpec & FIF_FLAGS)
		{
			pFrameInfo->m_dwFlags = 0;
			pFrameInfo->m_dwValidFields |= FIF_FLAGS;
			dwFieldSpec &= ~FIF_FLAGS;
		}

		if (dwFieldSpec & FIF_MODULE)
		{
			if (_module)
			{
				MODULE_INFO mi = { };
				auto hr = _module->GetInfo(MIF_NAME, &mi); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr) && (mi.dwValidFields & MIF_NAME))
				{
					pFrameInfo->m_bstrModule = mi.m_bstrName;
					mi.m_bstrName = nullptr;
					pFrameInfo->m_dwValidFields |= FIF_MODULE;
				}
			
			}
			dwFieldSpec &= ~FIF_MODULE;
		}

		if (dwFieldSpec & FIF_LANGUAGE)
		{
			pFrameInfo->m_bstrLanguage = SysAllocString(Z80AsmLanguageName);
			pFrameInfo->m_dwValidFields |= FIF_LANGUAGE;
			dwFieldSpec &= ~FIF_LANGUAGE;
		}

		if (dwFieldSpec & FIF_STALECODE)
		{
			pFrameInfo->m_fStaleCode = FALSE;
			pFrameInfo->m_dwValidFields |= FIF_STALECODE;
			dwFieldSpec &= ~FIF_STALECODE;
		}

		//if (dwFieldSpec)
		//	LOG_HR(E_NOTIMPL);

		return S_OK;
	}

	virtual HRESULT __stdcall GetPhysicalStackRange(UINT64* paddrMin, UINT64* paddrMax) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall GetExpressionContext(IDebugExpressionContext2** ppExprCxt) override
	{
		return MakeExpressionContextNoSource (_thread.get(), ppExprCxt);
	}

	virtual HRESULT __stdcall GetLanguageInfo(BSTR* pbstrLanguage, GUID* pguidLanguage) override
	{
		*pbstrLanguage = SysAllocString(Z80AsmLanguageName); RETURN_IF_NULL_ALLOC(*pbstrLanguage);
		*pguidLanguage = Z80AsmLanguageGuid;
		return S_OK;
	}

	virtual HRESULT __stdcall GetDebugProperty(IDebugProperty2** ppProperty) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT __stdcall EnumProperties(DEBUGPROP_INFO_FLAGS dwFields, UINT nRadix, REFGUID guidFilter, DWORD dwTimeout, ULONG* pcelt, IEnumDebugPropertyInfo2** ppEnum) override
	{
		*ppEnum = nullptr;
		*pcelt = 0;

		if (guidFilter == guidFilterRegisters)
		{
			wil::com_ptr_nothrow<IEnumDebugPropertyInfo2> p;
			auto hr = MakeEnumRegisterGroupsPropertyInfo (_thread.get(), dwFields, &p); RETURN_IF_FAILED(hr);
			hr = p->GetCount(pcelt); RETURN_IF_FAILED(hr);
			*ppEnum = p.detach();
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT __stdcall GetThread(IDebugThread2** ppThread) override
	{
		return wil::com_copy_to_nothrow (_thread, ppThread);
	}
	#pragma endregion
};

// ============================================================================

class Z80EnumDebugFrameInfo : public IEnumDebugFrameInfo2
{
	ULONG _refCount = 0;
	ULONG _nextIndex = 0;
	FRAMEINFO_FLAGS _dwFieldSpec;
	UINT _nRadix;
	wil::com_ptr_nothrow<IDebugThread2> _thread;

	struct Entry
	{
		uint16_t sp;
		struct
		{
			uint16_t pc;
			wil::com_ptr_nothrow<IDebugModule2> module;
			wil::unique_bstr symbol;
			uint16_t offset;
			wil::com_ptr_nothrow<IDebugDocumentContext2> documentContext;
		} returnAddress;
	};

	vector_nothrow<Entry> _entries;

public:
	HRESULT InitInstance (FRAMEINFO_FLAGS dwFieldSpec, UINT nRadix, IDebugThread2* thread)
	{
		HRESULT hr;

		// Let's create a simple call stack in which every stack location is a frame.
		// A lot of Z80 code does tricks with the SP register, so it's nearly impossible to build a proper call stack.
		
		vector_nothrow<Entry> entries;

		com_ptr<IDebugProgram2> program;
		hr = thread->GetProgram(&program); RETURN_IF_FAILED(hr);

		UINT16 stackStart;
		hr = simulator->GetStackStartAddress (&stackStart); RETURN_IF_FAILED(hr);

		z80_register_set regs;
		hr = simulator->GetRegisters(&regs, sizeof(regs)); RETURN_IF_FAILED(hr);
		
		uint16_t sp = regs.sp;
		uint16_t pc = regs.pc;
		for (uint32_t i = 0;;)
		{
			wil::unique_bstr symbol;
			uint16_t offset = 0;
			com_ptr<IDebugModule2> pcModule;
			hr = GetSymbolFromAddress (program, pc, SK_Code, nullptr, &symbol, &offset, &pcModule);

			bool pushed = entries.try_push_back (Entry{ sp, { pc, std::move(pcModule), std::move(symbol), offset, nullptr } }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

			if (sp == stackStart)
				// This was the "outermost" call (= frame with highest address in memory = shows up at the bottom of the Call Stack window)
				break;

			hr = simulator->ReadMemoryBus(sp, 2, &pc);
			if (FAILED(hr))
				break; // could happen with a computer with 32 KB memory when SP is pointing above that

			sp += 2;
			
			// Let's limit ourselves to 32 entries, in case the Z80 code does something weird
			// with the SP register and our algorithm here goes haywire.
			i++;
			if (i == 32)
				break;
		}

		_dwFieldSpec = dwFieldSpec;
		_nRadix = nRadix;
		_thread = thread;
		_entries = std::move(entries);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
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
		else if (riid == __uuidof(IEnumDebugFrameInfo2))
		{
			*ppvObject = static_cast<IEnumDebugFrameInfo2*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	HRESULT ProduceStackFrame (const Entry* entry, FRAMEINFO* fi)
	{
		fi->m_dwValidFields = 0;

		// TODO: set up lambda for cleaning up what was allocated, in case this function fails in the middle.
		//wil::scope_exit cleanup = [fi]{ };

		FRAMEINFO_FLAGS requested = _dwFieldSpec &
			~(FIF_ARGS_NO_TOSTRING | FIF_FILTER_NON_USER_CODE | FIF_STALECODE | FIF_FLAGS);

		if (_dwFieldSpec & FIF_FUNCNAME)
		{
			wstring_builder sb;
			if (entry->returnAddress.symbol)
			{
				sb << entry->returnAddress.symbol.get();
				if ((_dwFieldSpec & FIF_FUNCNAME_OFFSET) && entry->returnAddress.offset)
					sb << '+' << hex<uint16_t>(entry->returnAddress.offset);
			}
			else
				sb << hex<uint16_t>(entry->returnAddress.pc);

			fi->m_bstrFuncName = SysAllocStringLen(sb.data(), sb.size()); RETURN_IF_NULL_ALLOC(fi->m_bstrFuncName);
			fi->m_dwValidFields |= FIF_FUNCNAME;
			requested &= ~FIF_FUNCNAME;
		}

		constexpr DWORD FuncnameFlags = FIF_FUNCNAME_FORMAT | FIF_FUNCNAME_RETURNTYPE
			| FIF_FUNCNAME_ARGS | FIF_FUNCNAME_LANGUAGE | FIF_FUNCNAME_MODULE | FIF_FUNCNAME_LINES
			| FIF_FUNCNAME_OFFSET | FIF_FUNCNAME_ARGS_ALL;
		if (_dwFieldSpec & FuncnameFlags)
		{
			// We purposely ignore these.
			requested &= ~FuncnameFlags;
		}

		if (_dwFieldSpec & FIF_LANGUAGE)
		{
			fi->m_bstrLanguage = SysAllocString(Z80AsmLanguageName);
			fi->m_dwValidFields |= FIF_LANGUAGE;
			requested &= ~FIF_LANGUAGE;
		}

		if (_dwFieldSpec & FIF_FRAME)
		{
			auto frame = com_ptr(new (std::nothrow) Z80StackFrame()); RETURN_IF_NULL_ALLOC(frame);
			auto hr = frame->InitInstance(_thread.get(), entry->returnAddress.module.get(), entry->returnAddress.pc); RETURN_IF_FAILED(hr);
			fi->m_pFrame = frame.detach();
			fi->m_dwValidFields |= FIF_FRAME;
			requested &= ~FIF_FRAME;
		}

		if (_dwFieldSpec & FIF_DEBUG_MODULEP)
		{
			if (entry->returnAddress.module)
			{
				entry->returnAddress.module.copy_to(&fi->m_pModule);
				fi->m_dwValidFields |= FIF_DEBUG_MODULEP;
			}

			requested &= ~FIF_DEBUG_MODULEP;
		}

		//WI_ASSERT(!requested);

		return S_OK;
	}

	#pragma region IEnumDebugFrameInfo2
	virtual HRESULT __stdcall Next(ULONG celt, FRAMEINFO* rgelt, ULONG* pceltFetched) override
	{
		for (ULONG i = 0; i < celt; i++)
		{
			if (_nextIndex < _entries.size())
			{
				auto hr = ProduceStackFrame (&_entries[_nextIndex], &rgelt[i]); RETURN_IF_FAILED(hr);
			}
			else
			{
				if (pceltFetched)
					*pceltFetched = i;
				return S_FALSE;
			}

			_nextIndex++;
		}

		if (pceltFetched)
			*pceltFetched = celt;
		return S_OK;
	}

	virtual HRESULT __stdcall Skip(ULONG celt) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall Reset() override
	{
		_nextIndex = 0;
		return S_OK;
	}

	virtual HRESULT __stdcall Clone(IEnumDebugFrameInfo2** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall GetCount(ULONG* pcelt) override
	{
		*pcelt = _entries.size();
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeEnumDebugFrameInfo (FRAMEINFO_FLAGS dwFieldSpec, UINT nRadix, IDebugThread2* thread, IEnumDebugFrameInfo2** to)
{
	auto p = com_ptr(new (std::nothrow) Z80EnumDebugFrameInfo()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(dwFieldSpec, nRadix, thread); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}