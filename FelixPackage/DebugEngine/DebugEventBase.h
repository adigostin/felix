
#pragma once
#include "shared/TryQI.h"

template<typename ISpecific, EVENTATTRIBUTES attributes>
class EventBase : public IDebugEvent2, public ISpecific
{
	ULONG _refCount = 0;

public:
	EventBase() = default;

	EventBase (const EventBase&) = delete;
	EventBase& operator= (const EventBase&) = delete;

	virtual ~EventBase() = default;

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);

		if (   TryQI<IUnknown>(static_cast<IDebugEvent2*>(this), riid, ppvObject)
			|| TryQI<IDebugEvent2>(this, riid, ppvObject)
			|| TryQI<ISpecific>(this, riid, ppvObject))
			return S_OK;

		*ppvObject = NULL;
		return E_NOINTERFACE;
	}

	virtual ULONG __stdcall AddRef() override
	{
		return InterlockedIncrement(&_refCount);
	}

	virtual ULONG __stdcall Release() override
	{
		WI_ASSERT(_refCount);
		ULONG newRefCount = InterlockedDecrement(&_refCount);
		if (newRefCount == 0)
			delete this;
		return newRefCount;
	}
	#pragma endregion

	// IDebugEvent2
	virtual HRESULT __stdcall GetAttributes(DWORD* pdwAttrib) override
	{
		*pdwAttrib = attributes;
		return S_OK;
	}

	HRESULT Send (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, IDebugThread2* thread)
	{
		auto hr = callback->Event (engine, nullptr, program, thread, this, __uuidof(ISpecific), attributes); RETURN_IF_FAILED(hr);
		return S_OK;
	}
};
