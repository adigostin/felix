
#include "pch.h"
#include "DebugEngine.h"
#include "shared/enumerator.h"
#include "shared/TryQI.h"
#include "../FelixPackage.h"

class MemoryBytes : public IDebugMemoryBytes2
{
	ULONG _refCount = 0;
	com_ptr<ISimulator> _simulator;

public:
	HRESULT InitInstance()
	{
		auto hr = serviceProvider->QueryService(SID_Simulator, &_simulator); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppvObject);
		*ppvObject = nullptr;

		if (TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IDebugMemoryBytes2>(this, riid, ppvObject))
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugMemoryBytes2
	virtual HRESULT STDMETHODCALLTYPE ReadAt (IDebugMemoryContext2 *pStartContext, DWORD dwCount, BYTE* rgbMemory, DWORD* pdwRead, DWORD *pdwUnreadable) override
	{
		wil::com_ptr_nothrow<IFelixCodeContext> cc;
		auto hr = pStartContext->QueryInterface(&cc); RETURN_IF_FAILED(hr);
		
		if (!cc->PhysicalMemorySpace())
		{
			// We were given a CPU memory space, which is all readable.
			hr = _simulator->ReadMemoryBus ((uint16_t)cc->Address(), (uint16_t)dwCount, rgbMemory); RETURN_IF_FAILED(hr);
			*pdwRead = dwCount;
			if (pdwUnreadable)
				*pdwUnreadable = 0;
			return S_OK;
		}
		else
		{
			RETURN_HR(E_NOTIMPL);
			/*
			if (cc->Address() < _simulator->MemorySize())
			{
				// Some bytes are readable.
				UINT64 readableSize = _simulator->MemorySize() - cc->Address();
				if (dwCount < readableSize)
				{
					// The entire area the caller requested is readable.
					hr = _simulator->ReadMemory ((UINT32)cc->Address(), dwCount, rgbMemory); IfFailedAssertRet();
					*pdwRead = dwCount;
					if (pdwUnreadable)
						*pdwUnreadable = 0;
					return S_OK;
				}
				else
				{
					// Only part is readable.
					hr = _simulator->ReadMemory ((UINT32)cc->Address(), (DWORD)readableSize, rgbMemory); IfFailedAssertRet();
					*pdwRead = (DWORD)readableSize;
					if (pdwUnreadable)
						*pdwUnreadable = dwCount - (DWORD)readableSize;
					return S_OK;
				}
			}
			else
			{
				// Starting address is outside the available memory.
				UINT64 unreadableCount = (UINT64)0 - (UINT64)cc->Address();
				if (unreadableCount > dwCount)
					unreadableCount = dwCount;
				*pdwRead = 0;
				if (pdwUnreadable)
					*pdwUnreadable = (DWORD)unreadableCount;
				return S_OK;
			}
			*/
		}
	}

	virtual HRESULT STDMETHODCALLTYPE WriteAt (IDebugMemoryContext2 *pStartContext, DWORD dwCount, BYTE* rgbMemory) override
	{
		wil::com_ptr_nothrow<IFelixCodeContext> cc;
		auto hr = pStartContext->QueryInterface(&cc); RETURN_IF_FAILED(hr);

		if (!cc->PhysicalMemorySpace())
		{
			auto hr = _simulator->WriteMemoryBus ((uint16_t)cc->Address(), (uint16_t)dwCount, rgbMemory); RETURN_IF_FAILED(hr);
			return S_OK;
		}
		else
		{
			RETURN_HR(E_NOTIMPL);
		}
	}

	virtual HRESULT STDMETHODCALLTYPE GetSize (UINT64 *pqwSize) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeMemoryBytes (IDebugMemoryBytes2** to)
{
	com_ptr<MemoryBytes> p = new (std::nothrow) MemoryBytes(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
