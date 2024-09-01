#pragma once

// Error codes returned from XML-related functions

#define Z80_E_ID_PROPERTY_MISSING             MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x201)
#define Z80_E_ID_PROPERTY_EMPTY               MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x202)
#define Z80_E_ID_PROPERTY_NOT_BSTR            MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x203)

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("45B35EF7-DC2B-4EE3-BB44-EC25D607BFCE") IXmlParent : IUnknown
{
	// For an object property (not a value and not a collection) that has a single possible implementation,
	// this function can return a NULL name and S_OK to signal that the XML serializer should serialize
	// the child object directly on the XML element that corresponds to the property. This improves the readability
	// of the XML file. Normally, for such properties, the serializer creates something like this:
	// <PropertyName>                                  <= name of the property that comes from the parent's type info
	//   <xmlElementNameOut ... child attributes ...>  <= name that comes from this function
	// </PropertyName>
	// If this function returns a NULL name, the serializer can simplify this to:
	// <PropertyName ... child attributes ...>
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IUnknown* child, BSTR* xmlElementNameOut) = 0;

	// Creates a child object assignable to the property "dispidProperty", from the given xmlElementName, with default values for its properties.
	virtual HRESULT STDMETHODCALLTYPE CreateChild (DISPID dispidProperty, PCWSTR xmlElementName, IDispatch** childOut) = 0;

	// This function chooses whether a property needs to be serialized to XML or not.
	// Return S_OK to serialize the property. Return S_FALSE to not serialize it. Return an error code to abort serialization.
	virtual HRESULT STDMETHODCALLTYPE NeedSerialization (DISPID dispidProperty) = 0;
};

HRESULT SaveToXml (IDispatch* obj, PCWSTR elementName, IStream* stream, UINT nEncodingCodePage = CP_UTF8);

HRESULT LoadFromXml (IDispatch* obj, _In_opt_ PCWSTR expectedElementName, IStream* stream, UINT nEncodingCodePage = CP_UTF8);
