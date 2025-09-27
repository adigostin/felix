
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/OtherGuids.h"
#include "dispids.h"
#include <vsmanaged.h>

struct CustomBuildToolProperties
	: ICustomBuildToolProperties
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
{
	ULONG _refCount = 0;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	wil::unique_bstr _commandLine;
	wil::unique_bstr _description;
	wil::unique_bstr _outputs;

	HRESULT InitInstance()
	{
		auto hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(this, &_propNotifyCP); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<ICustomBuildToolProperties>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		if (   riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IManagedObject
			|| riid == IID_IEnumerable
			|| riid == IID_IConvertible
			|| riid == IID_IInspectable
			|| riid == IID_IWeakReferenceSource
			|| riid == IID_ICustomTypeDescriptor
			|| riid == IID_IComponent
			|| riid == IID_TypeDescriptor_IUnimplemented
			|| riid == IID_IPerPropertyBrowsing
			|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_ISpecifyPropertyPages
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IProvideMultipleClassInfo
		)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(ICustomBuildToolProperties);

	#pragma region ICustomBuildToolProperties
	virtual HRESULT STDMETHODCALLTYPE get_CommandLine (BSTR *value) override
	{
		// Although a NULL BSTR has identical semantics as "", the Properties Window
		// handles a NULL BSTR by hiding the property, and "" by showing it empty.
		return (*value = SysAllocString(_commandLine ? _commandLine.get() : L"")) ? S_OK : E_OUTOFMEMORY;
	}

	virtual HRESULT STDMETHODCALLTYPE put_CommandLine (BSTR value) override
	{
		if (VarBstrCmp(_commandLine.get(), value, 0, 0) != VARCMP_EQ)
		{
			_commandLine = (value && value[0]) ? wil::make_bstr_nothrow(value) : nullptr;
			_propNotifyCP->NotifyPropertyChanged(dispidCommandLine);
		}

		return S_OK;
	}
	
	virtual HRESULT STDMETHODCALLTYPE get_Description (BSTR *value) override
	{
		// Although a NULL BSTR has identical semantics as "", the Properties Window
		// handles a NULL BSTR by hiding the property, and "" by showing it empty.
		return (*value = SysAllocString(_description ? _description.get() : L"")) ? S_OK : E_OUTOFMEMORY;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Description (BSTR value) override
	{
		if (VarBstrCmp(_description.get(), value, 0, 0) != VARCMP_EQ)
		{
			_description = (value && value[0]) ? wil::make_bstr_nothrow(value) : nullptr;
			_propNotifyCP->NotifyPropertyChanged(dispidDescription);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Outputs (BSTR *value) override
	{
		// Although a NULL BSTR has identical semantics as "", the Properties Window
		// handles a NULL BSTR by hiding the property, and "" by showing it empty.
		return (*value = SysAllocString(_outputs ? _outputs.get() : L"")) ? S_OK : E_OUTOFMEMORY;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Outputs (BSTR value) override
	{
		if (VarBstrCmp(_outputs.get(), value, 0, 0) != VARCMP_EQ)
		{
			_outputs = (value && value[0]) ? wil::make_bstr_nothrow(value) : nullptr;
			_propNotifyCP->NotifyPropertyChanged(dispidOutputs);
		}

		return S_OK;
	}
	#pragma endregion


	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL *pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidCommandLine)
			return *fDefault = !_commandLine, S_OK;
		if (dispid == dispidDescription)
			return *fDefault = !_description, S_OK;
		if (dispid == dispidOutputs)
			return *fDefault = !_outputs, S_OK;
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetClassName (BSTR *pbstrClassName) override
	{
		// Shown by VS as value string for the expandable property.
		*pbstrClassName = SysAllocString(L"");
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE CanResetPropertyValue (DISPID dispid, BOOL* pfCanReset) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE ResetPropertyValue (DISPID dispid) override { return E_NOTIMPL; }
	#pragma endregion

	#pragma region IConnectionPointContainer
	virtual HRESULT STDMETHODCALLTYPE EnumConnectionPoints (IEnumConnectionPoints **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE FindConnectionPoint (REFIID riid, IConnectionPoint **ppCP) override
	{
		if (riid == IID_IPropertyNotifySink)
			return wil::com_query_to_nothrow(_propNotifyCP, ppCP);

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

};

FELIX_API HRESULT MakeCustomBuildToolProperties (ICustomBuildToolProperties** to)
{
	auto p = com_ptr(new (std::nothrow) CustomBuildToolProperties()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
