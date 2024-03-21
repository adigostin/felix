
#include "pch.h"
#include "FelixPackage_h.h"
#include "shared/com.h"
#include "Z80Xml.h"

class LaunchOptionsImpl : public IFelixLaunchOptions
{
	ULONG _refCount = 0;
	static inline wil::com_ptr_nothrow<ITypeLib> _typeLib;
	static inline wil::com_ptr_nothrow<ITypeInfo> _typeInfo;
	wil::unique_process_heap_string _projectDir;

public:
	HRESULT InitInstance()
	{
		HRESULT hr;

		if (!_typeLib)
		{
			wil::unique_process_heap_string filename;
			hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, filename); RETURN_IF_FAILED(hr);
			hr = LoadTypeLibEx (filename.get(), REGKIND_NONE, &_typeLib); RETURN_IF_FAILED(hr);
		}

		if (!_typeInfo)
		{
			hr = _typeLib->GetTypeInfoOfGuid(IID_IFelixLaunchOptions, &_typeInfo); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF (E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IFelixLaunchOptions>(this, riid, ppvObject))
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override
	{
		return ++_refCount;
	}

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		WI_ASSERT(_refCount);
		if (_refCount > 1)
			return --_refCount;
		delete this;
		return 0;
	}
	#pragma endregion

	#pragma region IDispatch
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override
	{
		*pctinfo = 1;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override
	{
		_typeInfo.copy_to(ppTInfo);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override
	{
		if (cNames == 1 && !wcscmp(rgszNames[0], L"ExtenderCATID"))
			return DISP_E_UNKNOWNNAME; // For this one name we don't want any error logging

		auto hr = DispGetIDsOfNames (_typeInfo.get(), rgszNames, cNames, rgDispId); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override
	{
		auto hr = DispInvoke (static_cast<IFelixLaunchOptions*>(this), _typeInfo.get(), dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	#pragma endregion

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
	#pragma endregion
};

HRESULT MakeLaunchOptions (IFelixLaunchOptions** ppOptions)
{
	auto p = com_ptr(new (std::nothrow) LaunchOptionsImpl()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*ppOptions = p.detach();
	return S_OK;
}
