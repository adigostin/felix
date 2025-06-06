
#pragma once
#include "vector_nothrow.h"

// ============================================================================

// Purpose of this wrapper is to allow implicit casting to the raw pointer.
// (Without this we'd have to write ".get()" all over the place.)
template <typename T>
struct com_ptr : wil::com_ptr_nothrow<T>
{
	com_ptr(T* x) noexcept : wil::com_ptr_nothrow<T>(x) { }
	using wil::com_ptr_nothrow<T>::com_ptr_nothrow;
	operator T*() { return wil::com_ptr_nothrow<T>::get(); }
};

// ============================================================================

template<typename ITo>
static bool TryQI (ITo* from, REFIID riid, void** ppvObject)
{
	if (__uuidof(from) == riid)
	{
		*ppvObject = from;
		from->AddRef();
		return true;
	}
	return false;
}

// Helper function meant to be called from the IUnknown::Release() of STA objects.
// When the last reference is released, calls the destructor while the object still has a reference count of 1.
template<typename T>
ULONG ReleaseST (T* _this, ULONG& refCount)
{
	WI_ASSERT(refCount);
	if (refCount > 1)
		return --refCount;
	delete _this;
	return 0;
}

// ============================================================================

template<typename I, std::enable_if_t<std::is_constructible_v<IUnknown*, I*>, int> = 0>
class Z80ClassFactory : public IClassFactory
{
	ULONG _refCount = 0;

	using create_t = HRESULT(*)(I** out);
	create_t const _create;

public:
	Z80ClassFactory (create_t create)
		: _create(create)
	{ }

	Z80ClassFactory (const Z80ClassFactory&) = delete;
	Z80ClassFactory& operator= (const Z80ClassFactory&) = delete;

	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			RETURN_HR(E_POINTER);

		*ppvObject = NULL;

		if (riid == __uuidof(IUnknown))
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == __uuidof(IClassFactory))
		{
			*ppvObject = static_cast<IClassFactory*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }

	virtual HRESULT __stdcall CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override
	{
		*ppvObject = nullptr;

		wil::com_ptr_nothrow<I> p;
		auto hr = _create(&p); RETURN_IF_FAILED(hr);

		hr = p->QueryInterface(riid, ppvObject); RETURN_IF_FAILED(hr);

		return S_OK;;
	}

	virtual HRESULT __stdcall LockServer(BOOL fLock) override
	{
		RETURN_HR(E_NOTIMPL);
	}
};

// ============================================================================

HRESULT MakeConnectionPointEnumerator (const CONNECTDATA* data, uint32_t size, IEnumConnections** ppEnum);

template<const IID& iid>
class ConnectionPointImpl : public IConnectionPoint
{
	ULONG _refCount = 0;
	IConnectionPointContainer* _cont = nullptr;
	vector_nothrow<CONNECTDATA> _cps;
	DWORD _nextCookie = 1;

public:
	static HRESULT CreateInstance (IConnectionPointContainer* cont, ConnectionPointImpl<iid>** cp)
	{
		com_ptr<ConnectionPointImpl> p = new (std::nothrow) ConnectionPointImpl(); RETURN_IF_NULL_ALLOC(p);
		p->_cont = cont;
		*cp = p.detach();
		return S_OK;
	}

	~ConnectionPointImpl()
	{
		for (uint32_t i = 0; i < _cps.size(); i++)
			_cps[i].pUnk->Release();
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);

		if (riid == IID_IUnknown)
		{
			*ppvObject = static_cast<IConnectionPoint*>(this);
			AddRef();
			return S_OK;
		}

		if (riid == IID_IConnectionPoint)
		{
			*ppvObject = static_cast<IConnectionPoint*>(this);
			AddRef();
			return S_OK;
		}

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IConnectionPoint
	virtual HRESULT STDMETHODCALLTYPE GetConnectionInterface (IID *pIID) override
	{
		*pIID = iid;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetConnectionPointContainer (IConnectionPointContainer** ppCPC) override
	{
		*ppCPC = _cont;
		_cont->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Advise (IUnknown* pUnkSink, DWORD* pdwCookie) override
	{
		bool pushed = _cps.try_push_back({ pUnkSink, _nextCookie }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pUnkSink->AddRef();
		*pdwCookie = _nextCookie++;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Unadvise (DWORD dwCookie) override
	{
		auto it = _cps.find_if([dwCookie](const CONNECTDATA& cd) { return cd.dwCookie == dwCookie; });
		RETURN_HR_IF(E_POINTER, it == _cps.end());
		it->pUnk->Release();
		_cps.erase(it);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE EnumConnections (IEnumConnections **ppEnum) override
	{
		return MakeConnectionPointEnumerator(_cps.data(), _cps.size(), ppEnum);
	}
	#pragma endregion

	void NotifyPropertyChanged (DISPID dispID) const
	{
		for (auto& c : _cps)
		{
			if (auto sink = wil::try_com_query_nothrow<IPropertyNotifySink>(c.pUnk))
				sink->OnChanged(dispID);
		}
	}
};

class EnumConnectionsImpl : public IEnumConnections
{
	ULONG _refCount = 0;
	ULONG _next = 0;
	vector_nothrow<CONNECTDATA> _cps;

public:
	HRESULT InitInstace (const CONNECTDATA* data, uint32_t size)
	{
		bool reserved = _cps.try_reserve(size); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);
		for (uint32_t i = 0; i < size; i++)
		{
			_cps.try_push_back(data[i]);
			data[i].pUnk->AddRef();
		}

		return S_OK;
	}

	~EnumConnectionsImpl()
	{
		for (uint32_t i = 0; i < _cps.size(); i++)
			_cps[i].pUnk->Release();
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);

		if (riid == IID_IEnumConnections)
		{
			*ppvObject = static_cast<IEnumConnections*>(this);
			AddRef();
			return S_OK;
		}

		*ppvObject = nullptr;
		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IEnumConnections
	virtual HRESULT STDMETHODCALLTYPE Next (ULONG cConnections, LPCONNECTDATA rgcd, ULONG* pcFetched) override
	{
		if (_next == _cps.size())
		{
			if (pcFetched)
				*pcFetched = 0;
			return S_FALSE;
		}

		if (cConnections != 1)
			RETURN_HR(E_NOTIMPL);

		*rgcd = _cps[_next];
		rgcd->pUnk->AddRef();
		_next++;

		if (pcFetched)
			*pcFetched = 1;

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Skip (ULONG cConnections) override { RETURN_HR(E_NOTIMPL); }

	virtual HRESULT STDMETHODCALLTYPE Reset() override { RETURN_HR(E_NOTIMPL); }

	virtual HRESULT STDMETHODCALLTYPE Clone (IEnumConnections **ppEnum) override { RETURN_HR(E_NOTIMPL); }
	#pragma endregion
};

inline HRESULT MakeConnectionPointEnumerator (const CONNECTDATA* data, uint32_t size, IEnumConnections** ppEnum)
{
	com_ptr<EnumConnectionsImpl> p = new (std::nothrow) EnumConnectionsImpl(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstace(data, size); RETURN_IF_FAILED(hr);
	*ppEnum = p.detach();
	return S_OK;
}

// ============================================================================

template<typename IEnumeratorType, typename IEntryType, typename IFromType = IEntryType>
class enumerator : public IEnumeratorType
{
	vector_nothrow<com_ptr<IEntryType>> _entries;
	ULONG _refCount = 0;
	ULONG _nextIndex = 0;

public:
	static HRESULT create_instance (IFromType* const* entries, size_t count, IEnumeratorType** to) noexcept
	{
		auto p = com_ptr(new (std::nothrow) enumerator()); RETURN_IF_NULL_ALLOC(p);

		for (size_t i = 0; i < count; i++)
		{
			com_ptr<IEntryType> e;
			auto hr = wil::com_query_to_nothrow(entries[i], &e); RETURN_IF_FAILED(hr);
			bool added = p->_entries.try_push_back(std::move(e)); RETURN_HR_IF(E_OUTOFMEMORY, !added);
		}

		*to = p.detach();
		return S_OK;
	}

	static HRESULT create_instance (const wil::com_ptr_nothrow<IFromType>* entries, size_t count, IEnumeratorType** to) noexcept
	{
		auto p = wil::com_ptr_nothrow(new (std::nothrow) enumerator()); RETURN_IF_NULL_ALLOC(p);
		for (size_t i = 0; i < count; i++)
		{
			wil::com_ptr_nothrow<IEntryType> e;
			auto hr = wil::com_query_to_nothrow(entries[i], &e); RETURN_IF_FAILED(hr);
			bool added = p->_entries.try_push_back(std::move(e)); RETURN_HR_IF(E_OUTOFMEMORY, !added);
		}
		*to = p.detach();
		return S_OK;
	}

	static HRESULT create_instance (vector_nothrow<com_ptr<IEntryType>>&& entries, IEnumeratorType** to) noexcept
	{
		auto p = com_ptr(new (std::nothrow) enumerator()); RETURN_IF_NULL_ALLOC(p);
		p->_entries = std::move(entries);
		*to = p.detach();
		return S_OK;
	}

	static HRESULT create_instance (const vector_nothrow<com_ptr<IEntryType>>& entries, IEnumeratorType** to)
	{
		auto p = com_ptr(new (std::nothrow) enumerator()); RETURN_IF_NULL_ALLOC(p);
		bool reserved = p->_entries.try_reserve(entries.size()); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);
		for (uint32_t i = 0; i < entries.size(); i++)
			p->_entries.try_push_back(entries[i]);
		*to = p.detach();
		return S_OK;
	}

	enumerator() = default;

	enumerator (const enumerator&) = delete;
	enumerator& operator= (const enumerator&) = delete;

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (!ppvObject)
			RETURN_HR(E_POINTER);

		*ppvObject = NULL;

		if (riid == __uuidof(IUnknown))
		{
			*ppvObject = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == __uuidof(IEnumeratorType))
		{
			*ppvObject = static_cast<IEnumeratorType*>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IClientSecurity)
			return E_NOINTERFACE;

		//BreakIntoDebugger();
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	virtual HRESULT __stdcall Next (ULONG celt, IEntryType** rgelt, ULONG* pceltFetched) override
	{
		if (pceltFetched)
			*pceltFetched = 0;

		if (celt == 0)
			return S_OK;

		ULONG i = 0;
		while(i < celt && _nextIndex < _entries.size())
		{
			rgelt[i] = _entries[_nextIndex].get();
			rgelt[i]->AddRef();
			i++;
			_nextIndex++;
		}

		if (pceltFetched)
			*pceltFetched = i;

		return (i == celt) ? S_OK : S_FALSE;
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

	virtual HRESULT __stdcall Clone(IEnumeratorType** ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall GetCount(ULONG* pcelt) override
	{
		*pcelt = (ULONG)_entries.size();
		return S_OK;
	}
};

template<typename IEnumeratorType, typename IEntryType, typename IFromType = IEntryType>
HRESULT make_single_entry_enumerator (IEntryType* singleEntryOrNull, IEnumeratorType** ppEnum) noexcept
{
	return enumerator<IEnumeratorType, IEntryType, IFromType>::create_instance(&singleEntryOrNull, singleEntryOrNull ? 1 : 0, ppEnum);
}

// ============================================================================

HRESULT inline SetErrorInfo (HRESULT errorHR, LPCWSTR messageFormat, ...)
{
	if (IsDebuggerPresent())
		__debugbreak();

	va_list argptr;
	va_start (argptr, messageFormat);
	wchar_t message[256];
	vswprintf_s(message, messageFormat, argptr);
	va_end(argptr);

	com_ptr<ICreateErrorInfo> cei;
	if (SUCCEEDED(::CreateErrorInfo(&cei)))
	{
		if (SUCCEEDED(cei->SetDescription((LPOLESTR)message)))
		{
			com_ptr<IErrorInfo> ei;
			if (SUCCEEDED(cei->QueryInterface(&ei)))
			{
				::SetErrorInfo(0, ei.get());
			}
		}
	}

	return errorHR;
}

#define IMPLEMENT_IDISPATCH(DISP_IID) static ITypeInfo* GetTypeInfo() { \
		static com_ptr<ITypeInfo> _typeInfo; \
		if (!_typeInfo) { \
			wil::unique_process_heap_string filename; \
			auto hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, filename); FAIL_FAST_IF_FAILED(hr); \
			com_ptr<ITypeLib> _typeLib; \
			hr = LoadTypeLibEx (filename.get(), REGKIND_NONE, &_typeLib); FAIL_FAST_IF_FAILED(hr); \
			hr = _typeLib->GetTypeInfoOfGuid(DISP_IID, &_typeInfo); FAIL_FAST_IF_FAILED(hr); \
		} \
		return _typeInfo.get(); \
	} \
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override final { *pctinfo = 1; return S_OK; } \
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override final { \
		*ppTInfo = GetTypeInfo(); \
		(*ppTInfo)->AddRef(); \
		return S_OK; \
	} \
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override final { \
		return DispGetIDsOfNames (GetTypeInfo(), rgszNames, cNames, rgDispId); \
	} \
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override final { \
		return DispInvoke (this, GetTypeInfo(), dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr); \
	} \
