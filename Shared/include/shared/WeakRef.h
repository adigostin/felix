
#pragma once

// Copied from WeakReference.h from the Windows SDK, with IUnknown instead of IInspectable.
struct DECLSPEC_NOVTABLE IWeakRef : IUnknown
{
	// Returns S_FALSE and sets objectRef to NULL when the reference cannot be resolved.
	// (This is different from IWeakReference::Resolve from the Windows SDK, which returns S_OK in this case.)
	virtual HRESULT STDMETHODCALLTYPE Resolve (REFIID riid, void** objectRef) = 0;

	template <typename T>
	HRESULT Resolve (T** objectRef)
	{
		static_assert(__is_base_of(IUnknown, T));
		return Resolve(__uuidof(T), (void**)objectRef);
	}
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{14CE9B12-D6EB-453C-AEA5-CA8486028FF5}") IWeakRefSource : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE GetWeakRef (IWeakRef **weakRef) = 0;
};

// This class eases implementation of IWeakRefSource. It must be used from a single thread.
// Classes that implement IWeakRefSource can have a member variable of type WeakRefToThis,
// for example "WeakRefToThis _weakRef;", and their implementation of GetWeakReference
// can simply do "return _weakRef.GetOrCreate(this, weakReference);".
class WeakRefToThis
{
	struct Impl : IWeakRef
	{
		ULONG _refCount = 0;
		IWeakRefSource* _ptr;

		Impl (IWeakRefSource* ptr) : _ptr(ptr) { }

		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { RETURN_HR(E_NOTIMPL); }

		virtual ULONG STDMETHODCALLTYPE AddRef() override
		{
			return ++_refCount;
		}

		virtual ULONG STDMETHODCALLTYPE Release() override
		{
			WI_ASSERT(_refCount);
			auto newRefCount = --_refCount;
			if (newRefCount == 0)
				delete this;
			return newRefCount;
		}

		virtual HRESULT STDMETHODCALLTYPE Resolve (REFIID riid, void** objectRef) override
		{
			RETURN_HR_IF(E_POINTER, !objectRef);
			if (!_ptr)
			{
				*objectRef = nullptr;
				return S_FALSE;
			}

			return _ptr->QueryInterface(riid, objectRef);
		}
	};

	Impl* wr = nullptr;

public:
	~WeakRefToThis()
	{
		Reset();
	}

	HRESULT GetOrCreate (IWeakRefSource* source, IWeakRef** weakReference)
	{
		if (!wr)
		{
			wr = new (std::nothrow) Impl(source); RETURN_IF_NULL_ALLOC(wr);
			wr->AddRef();
		}

		*weakReference = wr;
		wr->AddRef();
		return S_OK;
	}

	void Reset()
	{
		if (wr)
		{
			WI_ASSERT(wr->_ptr);
			wr->_ptr = nullptr;
			wr->Release();
			wr = nullptr;
		}
	}
};

inline HRESULT ToWeak (IUnknown* from, IWeakRef** ppWeak) noexcept
{
	wil::com_ptr_nothrow<IWeakRefSource> refSource;
	HRESULT hr = from->QueryInterface(IID_PPV_ARGS(refSource.addressof())); RETURN_IF_FAILED(hr);
	return refSource->GetWeakRef(ppWeak);
}
