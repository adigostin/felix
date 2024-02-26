
#pragma once
#include "vector_nothrow.h"

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
