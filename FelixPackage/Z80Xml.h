#pragma once

// We use DLL exports only to be able to call various functions from unit test projects.
#ifdef FELIX_EXPORTS
#define FELIX_API __declspec(dllexport)
#else
#define FELIX_API __declspec(dllimport)
#endif

// Note AGO: Reason for this interface is the following: the XML loader needs a way to create objects
// out of the XML elements it encounters. The "traditional" way is to register somehow factories
// for each type of object, and pass these factories to the XML loader. This is already complicated,
// and moreover brings the restriction that every class must be constructible without parameters.
// (Declaring and using parameters to factories is overly complicated and not worth exploring.)
//
// After a lot of trial and error, I ended up with a model in which any serializable object which
// owns other serializable objects in its properties implements this interface. While saving to XML,
// the XML saving code calls GetChildXmlElementName. While loading from XML, the XML loading code
// calls CreateChild.
struct DECLSPEC_NOVTABLE DECLSPEC_UUID("45B35EF7-DC2B-4EE3-BB44-EC25D607BFCE") IXmlParent : IUnknown
{
	// For an object property (not a value and not a collection) that has a single possible implementation,
	// this function can return a NULL name and S_OK to signal that the XML serializer should serialize
	// the child object directly on the XML element that corresponds to the property. This improves the readability
	// of the XML file. Normally, for such properties, the serializer creates something like this:
	// <ParentClassName>
	//   ...
	//   <PropertyName>                                     <= name of the property that comes from the parent's type info
	//     <ChildObjectClassName ... child attributes ...>  <= name that comes from this function
	//   </PropertyName>
	// </ParentClassName>
	// If this function returns a NULL name, the serializer can simplify this to:
	// <ParentClassName>
	//   ...
	//   <PropertyName ... child attributes ...>
	// </ParentClassName>
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IDispatch* child, BSTR* xmlElementNameOut) = 0;

	// Creates a child object assignable to the property "dispidProperty", from the given xmlElementName, with default values for its properties.
	// This function must be implemented also for read-only properties, because the code behind property pages
	// calls this function to create copies of objects being edited (copies are needed in in case the user
	// changes some properties and then clicks Cancel).
	virtual HRESULT STDMETHODCALLTYPE CreateChild (DISPID dispidProperty, PCWSTR xmlElementName, IDispatch** childOut) = 0;
};

typedef enum
{
	SAVE_XML_FORCE_SERIALIZE_DEFAULTS = 1,
} SaveXmlFlags;

FELIX_API HRESULT SaveToXml (IDispatch* obj, PCWSTR elementName, DWORD flags, IStream* stream, UINT nEncodingCodePage = CP_UTF8);

FELIX_API HRESULT LoadFromXml (IDispatch* obj, _In_opt_ PCWSTR expectedElementName, IStream* stream, UINT nEncodingCodePage = CP_UTF8);
