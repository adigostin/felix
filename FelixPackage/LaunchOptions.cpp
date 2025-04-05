
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "Z80Xml.h"
#include "dispids.h"

class LaunchOptionsImpl
	: public IFelixLaunchOptions
	, IXmlParent
{
	ULONG _refCount = 0;
	wil::unique_process_heap_string _projectDir;
	com_ptr<IProjectConfigDebugProperties> _debuggingProperties;

public:
	HRESULT InitInstance()
	{
		auto hr = DebuggingPageProperties_CreateInstance (&_debuggingProperties); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF (E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IFelixLaunchOptions*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IFelixLaunchOptions>(this, riid, ppvObject)
			|| TryQI<IXmlParent>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IFelixLaunchOptions)

	#pragma region IFelixLaunchOptions
	virtual HRESULT STDMETHODCALLTYPE get_ProjectDir (BSTR *pbstr) override
	{
		*pbstr = SysAllocString(_projectDir.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_ProjectDir (BSTR pbstr) override
	{
		_projectDir = wil::make_unique_string_nothrow<wil::unique_process_heap_string>(pbstr); RETURN_IF_NULL_ALLOC(_projectDir);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_DebuggingProperties (IProjectConfigDebugProperties** ppDispatch) override
	{
		*ppDispatch = _debuggingProperties;
		if (_debuggingProperties)
			_debuggingProperties->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_DebuggingProperties (IProjectConfigDebugProperties* pDispatch) override
	{
		_debuggingProperties = pDispatch;
		return S_OK;
	}
	#pragma endregion

	#pragma region IXmlParent
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IDispatch* child, BSTR* xmlElementNameOut) override
	{
		if (dispidProperty == dispidDebuggingProperties)
		{
			*xmlElementNameOut = nullptr;
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE CreateChild (DISPID dispidProperty, PCWSTR xmlElementName, IDispatch** childOut) override
	{
		if (dispidProperty == dispidDebuggingProperties)
		{
			com_ptr<IProjectConfigDebugProperties> pp;
			auto hr = DebuggingPageProperties_CreateInstance(&pp); RETURN_IF_FAILED(hr);
			*childOut = pp.detach();
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeLaunchOptions (IFelixLaunchOptions** ppOptions)
{
	auto p = com_ptr(new (std::nothrow) LaunchOptionsImpl()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*ppOptions = p.detach();
	return S_OK;
}
