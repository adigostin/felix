
#pragma once
#include "vector_nothrow.h"

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

	const vector_nothrow<CONNECTDATA>& GetConnections() const { return _cps; }
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

HRESULT MakeConnectionPointEnumerator (const CONNECTDATA* data, uint32_t size, IEnumConnections** ppEnum)
{
	com_ptr<EnumConnectionsImpl> p = new (std::nothrow) EnumConnectionsImpl(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstace(data, size); RETURN_IF_FAILED(hr);
	*ppEnum = p.detach();
	return S_OK;
}
