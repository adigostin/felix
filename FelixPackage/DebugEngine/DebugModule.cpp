
#include "pch.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
#include "DebugEngine.h"
#include "DebugEventBase.h"

class Z80Module : public IDebugModule3, IZ80Module
{
	ULONG  _refCount = 0;
	UINT64 _address;
	DWORD  _size;
	wil::unique_process_heap_string _path;
	wil::unique_process_heap_string _symbolsPath;
	bool _user_code;
	com_ptr<IDebugEngine2> _engine;
	com_ptr<IDebugProgram2> _program;
	com_ptr<IDebugEventCallback2> _callback;

	// _symbols==NULL and SUCCEEDED(_symbolLoadResult) - loading has not been attempted yet during this debug sessions
	// _symbols==NULL and FAILED(_symbolLoadResult) - loading has been attempted and failed with the error code from _symbolLoadResult
	// _symbols!=NULL - loading has been attempted and was successful.
	com_ptr<IZ80Symbols> _symbols;
	HRESULT _symbolLoadResult = S_OK;

public:
	HRESULT InitInstance (UINT64 address, DWORD size, const wchar_t* path, const wchar_t* symbolsFilePath,
		bool user_code, IDebugEngine2* engine, IDebugProgram2* program, IDebugEventCallback2* callback)
	{
		_address = address;
		_size = size;
		_path = wil::make_process_heap_string_nothrow(path); RETURN_IF_NULL_ALLOC(_path);
		if (symbolsFilePath)
		{
			_symbolsPath = wil::make_process_heap_string_nothrow(symbolsFilePath); RETURN_IF_NULL_ALLOC(_symbolsPath);
		}
		_user_code = user_code;
		_engine = engine;
		_program = program;
		_callback = callback;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = NULL;

		if (   TryQI<IUnknown>(static_cast<IDebugModule2*>(this), riid, ppvObject)
			|| TryQI<IDebugModule2>(this, riid, ppvObject)
			|| TryQI<IDebugModule3>(this, riid, ppvObject)
			|| TryQI<IZ80Module>(this, riid, ppvObject))
			return S_OK;

		// These may be implemented eventually.
		if (   riid == IID_IDebugDumpModule100
			|| riid == IID_IDebugModuleInternal165
			|| riid == IID_IDebugModuleManagedInternal165
			|| riid == IID_IDebugModule170
			|| riid == IID_IDebugModule171
			|| riid == IID_IDebugModule174
		)
			return E_NOINTERFACE;

		// These will never be implemented.
		if (   riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IsModuleMSBranded
			|| riid == IID_IAppDomainInfo110
		)
			return E_NOINTERFACE;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	static void FreeModuleInfo (MODULE_INFO* mi)
	{
		struct mi_info
		{
			MODULE_INFO_FIELDS field;
			BSTR (MODULE_INFO::*ptr);
		};

		static const mi_info infos[] = 
		{
			{ MIF_NAME,              &MODULE_INFO::m_bstrName },
			{ MIF_URL,               &MODULE_INFO::m_bstrUrl },
			{ MIF_VERSION,           &MODULE_INFO::m_bstrVersion },
			{ MIF_DEBUGMESSAGE,      &MODULE_INFO::m_bstrDebugMessage },
			{ MIF_URLSYMBOLLOCATION, &MODULE_INFO::m_bstrUrlSymbolLocation }
		};

		for (auto& i : infos)
		{
			if (mi->dwValidFields & i.field)
			{
				if (mi->*(i.ptr))
				{
					SysFreeString(mi->*(i.ptr));
					mi->*(i.ptr) = nullptr;
				}

				mi->dwValidFields &= ~i.field;
			}
		}
	}

	#pragma region IDebugModule2
	virtual HRESULT __stdcall GetInfo(MODULE_INFO_FIELDS dwFields, MODULE_INFO* pInfo) override
	{
		pInfo->dwValidFields = 0;
		auto cleanup = wil::scope_exit([pInfo] { FreeModuleInfo(pInfo); });

		if (dwFields & MIF_FLAGS)
		{
			pInfo->m_dwModuleFlags = 0;
			pInfo->dwValidFields |= MIF_FLAGS;
			dwFields &= ~MIF_FLAGS;
		}

		if (dwFields & MIF_URLSYMBOLLOCATION)
		{
			if (_symbolsPath)
			{
				pInfo->m_bstrUrlSymbolLocation = SysAllocString(_symbolsPath.get()); RETURN_IF_NULL_ALLOC(pInfo->m_bstrUrlSymbolLocation);
				pInfo->dwValidFields |= MIF_URLSYMBOLLOCATION;
			}

			dwFields &= ~MIF_URLSYMBOLLOCATION;
		}

		if (dwFields & MIF_VERSION)
		{
			pInfo->m_bstrVersion = SysAllocString(L"<version>");
			pInfo->dwValidFields |= MIF_VERSION;
			dwFields &= ~MIF_VERSION;
		}

		if (dwFields & MIF_URL)
		{
			pInfo->m_bstrUrl = SysAllocString(_path.get());
			pInfo->dwValidFields |= MIF_URL;
			dwFields &= ~MIF_URL;
		}

		if (dwFields & MIF_NAME)
		{
			// Need a valid filename here. VS code calls some filesystem functions
			// that throw if this string is not a valid filename.
			pInfo->m_bstrName = SysAllocString(PathFindFileNameW(_path.get()));
			pInfo->dwValidFields |= MIF_NAME;
			dwFields &= ~MIF_NAME;
		}

		if (dwFields & MIF_TIMESTAMP)
		{
			pInfo->m_TimeStamp = { 0, 0 };
			pInfo->dwValidFields |= MIF_TIMESTAMP;
			dwFields &= ~MIF_TIMESTAMP;
		}

		if (dwFields & MIF_DEBUGMESSAGE)
		{
			pInfo->m_bstrDebugMessage = SysAllocString(L"<debug.message>");
			pInfo->dwValidFields |= MIF_DEBUGMESSAGE;
			dwFields &= ~MIF_DEBUGMESSAGE;
		}

		if (dwFields & MIF_LOADORDER)
		{
			pInfo->m_dwLoadOrder = 1;
			pInfo->dwValidFields |= MIF_LOADORDER;
			dwFields &= ~MIF_LOADORDER;
		}

		if (dwFields & MIF_SIZE)
		{
			pInfo->m_dwSize = _size;
			pInfo->dwValidFields |= MIF_SIZE;
			dwFields &= ~MIF_SIZE;
		}

		if (dwFields & MIF_PREFFEREDADDRESS)
		{
			pInfo->m_addrPreferredLoadAddress = _address;
			pInfo->dwValidFields |= MIF_PREFFEREDADDRESS;
			dwFields &= ~MIF_PREFFEREDADDRESS;
		}

		if (dwFields & MIF_LOADADDRESS)
		{
			pInfo->m_addrLoadAddress = _address;
			pInfo->dwValidFields |= MIF_LOADADDRESS;
			dwFields &= ~MIF_LOADADDRESS;
		}

		WI_ASSERT (!dwFields);

		cleanup.release();
		return S_OK;
	}

	virtual HRESULT __stdcall ReloadSymbols_Deprecated(LPCOLESTR pszUrlToSymbols, BSTR* pbstrDebugMessage) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IDebugModule3
	virtual HRESULT STDMETHODCALLTYPE GetSymbolInfo (SYMBOL_SEARCH_INFO_FIELDS dwFields, MODULE_SYMBOL_SEARCH_INFO *pInfo) override
	{
		pInfo->dwValidFields = 0;

		if (dwFields & SSIF_VERBOSE_SEARCH_INFO)
		{
			pInfo->bstrVerboseSearchInfo = SysAllocString(L"bstrVerboseSearchInfo"); RETURN_IF_NULL_ALLOC(pInfo->bstrVerboseSearchInfo);
			pInfo->dwValidFields |= SSIF_VERBOSE_SEARCH_INFO;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE LoadSymbols() override
	{
		// VS calls this when right-clicking a frame in the Call Stack window and choosing Load Symbols.
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsUserCode (BOOL *pfUser) override
	{
		*pfUser = _user_code;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetJustMyCodeState (BOOL fIsUserCode) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	struct SymbolSearchEvent : EventBase<IDebugSymbolSearchEvent2, EVENT_ASYNCHRONOUS>
	{
		com_ptr<IDebugModule3> _module;
		wil::unique_hlocal_string _debugMessage;
		MODULE_INFO_FLAGS _flags;

		SymbolSearchEvent (IDebugModule3* module, wil::unique_hlocal_string debugMessage, MODULE_INFO_FLAGS flags)
			: _module(module), _debugMessage(std::move(debugMessage)), _flags(flags)
		{ }

		virtual HRESULT STDMETHODCALLTYPE GetSymbolSearchInfo (IDebugModule3 **pModule, BSTR *pbstrDebugMessage, MODULE_INFO_FLAGS *pdwModuleInfoFlags) override
		{
			_module.copy_to(pModule);
			*pbstrDebugMessage = SysAllocString(_debugMessage.get()); RETURN_IF_NULL_ALLOC(*pbstrDebugMessage);
			*pdwModuleInfoFlags = _flags;
			return S_OK;
		}
	};

	#pragma region IZ80Module
	virtual HRESULT STDMETHODCALLTYPE GetSymbols (IZ80Symbols** symbols) override
	{
		if (_symbols)
		{
			*symbols = _symbols;
			_symbols->AddRef();
			return S_OK;
		}

		*symbols = nullptr;

		if (FAILED(_symbolLoadResult))
			// We tried already to load them and this was the result. Don't try again.
			return _symbolLoadResult;
		
		// Attempt to load them for the first time.
		if (!_symbolsPath)
		{
			size_t bufferLen = wcslen(_path.get()) + 10;
			auto symPath = wil::make_process_heap_string_nothrow (_path.get(), MAX_PATH); RETURN_IF_NULL_ALLOC(symPath);
			BOOL bres = PathRenameExtension (symPath.get(), L".sld");
			if (!bres)
				return _symbolLoadResult = CO_E_BAD_PATH;
			
			if (!::PathFileExists(symPath.get()))
			{
				symPath = wil::make_process_heap_string_nothrow (_path.get(), MAX_PATH); RETURN_IF_NULL_ALLOC(symPath);
				bres = PathRenameExtension (symPath.get(), L".z80sym");
				if (!bres)
					return _symbolLoadResult = CO_E_BAD_PATH;

				if (!::PathFileExists(symPath.get()))
					return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
			}

			_symbolsPath = std::move(symPath);
		}

		auto* ext = PathFindExtension(_symbolsPath.get());
		if (!wcscmp(ext, L".sld"))
		{
			_symbolLoadResult = MakeSldSymbols (this, &_symbols);
		}
		else if (!wcscmp(ext, L".z80sym"))
		{
			_symbolLoadResult = MakeZ80SymSymbols (this, &_symbols);
		}
		else
			_symbolLoadResult = E_UNRECOGNIZED_DEBUG_FILE_EXTENSION;

		if (FAILED(_symbolLoadResult))
		{
			if (auto errorMessage = wil::make_hlocal_string_nothrow(L"load error"))
				if (auto e = com_ptr(new (std::nothrow) SymbolSearchEvent(this, std::move(errorMessage), 0)))
					e->Send(_callback.get(), _engine.get(), _program.get(), nullptr);
			RETURN_HR(_symbolLoadResult);
		}

		if (auto errorMessage = wil::make_hlocal_string_nothrow(_symbolsPath.get()))
			if (auto e = com_ptr(new (std::nothrow) SymbolSearchEvent(this, std::move(errorMessage), MIF_SYMBOLS_LOADED)))
				e->Send(_callback.get(), _engine.get(), _program.get(), nullptr);

		_symbols.copy_to(symbols);
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeModule (UINT64 address, DWORD size, const wchar_t* path, const wchar_t* symbolsFilePath, bool user_code,
	IDebugEngine2* engine, IDebugProgram2* program, IDebugEventCallback2* callback, IDebugModule2** to) 
{
	auto p = com_ptr(new (std::nothrow) Z80Module()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance (address, size, path, symbolsFilePath, user_code, engine, program, callback); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
