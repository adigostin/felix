
#include "pch.h"
#include "shared/com.h"
#include "Z80Xml.h"

using unique_safearray = wil::unique_any<SAFEARRAY*, decltype(SafeArrayDestroy), &SafeArrayDestroy>;

static const auto InvariantLCID = LocaleNameToLCID(LOCALE_NAME_INVARIANT, 0);

static HRESULT GetNameFromEnumValue (ITypeInfo* ti, const TYPEATTR* typeAttr, LONG value, BSTR* nameOut)
{
	RETURN_HR_IF(E_INVALIDARG, typeAttr->typekind != TKIND_ENUM);

	for (WORD i = 0; i < typeAttr->cVars; i++)
	{
		VARDESC* varDesc;
		auto hr = ti->GetVarDesc(i, &varDesc); RETURN_IF_FAILED(hr);
		auto releaseVarDesc = wil::scope_exit([ti, varDesc] { ti->ReleaseVarDesc(varDesc); });
		RETURN_HR_IF(E_INVALIDARG, varDesc->lpvarValue->vt != VT_I4);
		if (varDesc->lpvarValue->lVal == value)
		{
			UINT cNames;
			hr = hr = ti->GetNames(varDesc->memid, nameOut, 1, &cNames); RETURN_IF_FAILED(hr);
			RETURN_HR_IF(E_FAIL, cNames != 1);
			return S_OK;
		}
	}

	RETURN_HR(DISP_E_UNKNOWNNAME);
}

static HRESULT GetEnumValueFromName (ITypeInfo* ti, const TYPEATTR* typeAttr, LPCWSTR name, LONG* valueOut)
{
	RETURN_HR_IF(E_INVALIDARG, typeAttr->typekind != TKIND_ENUM);

	for (WORD i = 0; i < typeAttr->cVars; i++)
	{
		VARDESC* varDesc;
		auto hr = ti->GetVarDesc(i, &varDesc); RETURN_IF_FAILED(hr);
		auto releaseVarDesc = wil::scope_exit([ti, varDesc] { ti->ReleaseVarDesc(varDesc); });
		RETURN_HR_IF(E_INVALIDARG, varDesc->lpvarValue->vt != VT_I4);
		wil::unique_bstr n;
		UINT cNames;
		hr = ti->GetNames(varDesc->memid, &n, 1, &cNames); RETURN_IF_FAILED(hr);
		if (!wcscmp(n.get(), name))
		{
			*valueOut = varDesc->lpvarValue->lVal;
			return S_OK;
		}
	}

	RETURN_HR(DISP_E_UNKNOWNNAME);
}

template<typename ValueType>
struct Property
{
	DISPID dispid;
	wil::unique_bstr name;
	ValueType value;
};

using ValueProperty = Property<wil::unique_bstr>;
using ObjectProperty = Property<wil::com_ptr_nothrow<IUnknown>>;
using ObjectCollectionProperty = Property<vector_nothrow<wil::com_ptr_nothrow<IUnknown>>>;

// VS runs out of memory if we templatize SaveToXmlInternal and have it call itself. So let's reinvent the wheel...
struct EnsureElementCreated
{
	bool insideElement = false;
	const wchar_t* const elementName;
	IXmlWriterLite* writer;
	EnsureElementCreated* outer;

	EnsureElementCreated (const wchar_t* elementName, IXmlWriterLite* writer, EnsureElementCreated* outer)
		: elementName(elementName), writer(writer), outer(outer)
	{ }

	HRESULT CreateStartElement()
	{
		if (!insideElement)
		{
			if (outer)
			{
				auto hr = outer->CreateStartElement(); RETURN_IF_FAILED(hr);
			}

			auto hr = writer->WriteStartElement(elementName, (UINT)wcslen(elementName)); RETURN_IF_FAILED(hr);
			insideElement = true;
		}

		return S_OK;
	};

	HRESULT CreateEndElement()
	{
		if (insideElement)
		{
			auto hr = writer->WriteEndElement(elementName, (UINT)wcslen(elementName)); RETURN_IF_FAILED(hr);
			insideElement = false;
		}

		return S_OK;
	}

	~EnsureElementCreated()
	{
		WI_ASSERT(!insideElement);
	}
};

static HRESULT SaveToXmlInternal (IUnknown* obj, PCWSTR elementName, IXmlWriterLite* writer, EnsureElementCreated* ensureOuterElementCreated)
{
	HRESULT hr;

	EnsureElementCreated ensureElementCreated (elementName, writer, ensureOuterElementCreated);

	wil::com_ptr_nothrow<IDispatch> objAsDispatch;
	if (SUCCEEDED(obj->QueryInterface(&objAsDispatch)))
	{
		wil::com_ptr_nothrow<IXmlParent> objAsXmlParent;
		hr = obj->QueryInterface(&objAsXmlParent); RETURN_HR_IF(hr, FAILED(hr) && (hr != E_NOINTERFACE));

		wil::com_ptr_nothrow<ITypeInfo> typeInfo;
		auto hr = objAsDispatch->GetTypeInfo(0, 0x0409, &typeInfo); RETURN_IF_FAILED(hr);
		TYPEATTR* typeAttr;
		hr = typeInfo->GetTypeAttr(&typeAttr); RETURN_IF_FAILED(hr);
		auto releaseTypeAttr = wil::scope_exit([typeInfo, typeAttr] { typeInfo->ReleaseTypeAttr(typeAttr); });
	
		wil::com_ptr_nothrow<IVsPerPropertyBrowsing> ppb;
		obj->QueryInterface(&ppb);

		vector_nothrow<ValueProperty> attributes;
		vector_nothrow<ObjectProperty> childObjects;
		vector_nothrow<ObjectCollectionProperty> childCollections;

		for (WORD i = 0; i < typeAttr->cFuncs; i++)
		{
			FUNCDESC* fd;
			hr = typeInfo->GetFuncDesc(i, &fd); RETURN_IF_FAILED(hr);
			if (fd->memid == DISPID_VALUE)
				continue;
			auto releaseFundDesc = wil::scope_exit([ti=typeInfo.get(), fd] { ti->ReleaseFuncDesc(fd); });
			if (fd->invkind == INVOKE_PROPERTYGET)
			{
				wil::unique_bstr name;
				UINT cNames;
				hr = typeInfo->GetNames(fd->memid, &name, 1, &cNames); RETURN_IF_FAILED(hr);

				BOOL hasDefaultValue = FALSE;
				if (ppb)
					ppb->HasDefaultValue(fd->memid, &hasDefaultValue);

				if (!hasDefaultValue)
				{
					DISPPARAMS params = { };
					wil::unique_variant result;
					EXCEPINFO exception;
					UINT uArgErr;
					hr = typeInfo->Invoke(obj, fd->memid, DISPATCH_PROPERTYGET, &params, &result, &exception, &uArgErr); RETURN_IF_FAILED(hr);

					switch (fd->elemdescFunc.tdesc.vt)
					{
						case VT_UI1:
						case VT_UI2:
						case VT_UI4:
						case VT_I1:
						case VT_I2:
						case VT_I4:
							hr = VariantChangeTypeEx (&result, &result, InvariantLCID, 0, VT_BSTR); RETURN_IF_FAILED(hr);
							[[fallthrough]];
						case VT_BSTR:
							// Write it to XML only if not empty.
							if (SysStringLen(V_BSTR(&result)))
							{
								auto value = wil::make_bstr_nothrow(V_BSTR(&result)); RETURN_IF_NULL_ALLOC(value);
								bool pushed = attributes.try_push_back(ValueProperty{ fd->memid, std::move(name), std::move(value) }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
							}
							break;

						case VT_USERDEFINED:
						{
							wil::com_ptr_nothrow<ITypeInfo> refTypeInfo;
							hr = typeInfo->GetRefTypeInfo(fd->elemdescFunc.tdesc.hreftype, &refTypeInfo); RETURN_IF_FAILED(hr);
							TYPEATTR* refTypeAttr;
							hr = refTypeInfo->GetTypeAttr(&refTypeAttr); RETURN_IF_FAILED(hr);
							auto releaseRefTypeAttr = wil::scope_exit([ti=refTypeInfo.get(), refTypeAttr] { ti->ReleaseTypeAttr(refTypeAttr); });
							if (refTypeAttr->typekind == TKIND_ENUM)
							{
								wil::unique_bstr value;
								hr = GetNameFromEnumValue (refTypeInfo.get(), refTypeAttr, V_I4(&result), &value); RETURN_IF_FAILED(hr);
								bool pushed = attributes.try_push_back(ValueProperty{ fd->memid, std::move(name), std::move(value) }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
							}
							else
								RETURN_HR(E_NOTIMPL);
							break;
						}

						case VT_DISPATCH:
						{
							// child object
							bool pushed = childObjects.try_push_back(ObjectProperty{ fd->memid, std::move(name), V_DISPATCH(&result) }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
							break;
						}

						case VT_SAFEARRAY:
						{
							SAFEARRAY* sa = V_ARRAY(&result);
							VARTYPE vt;
							hr = SafeArrayGetVartype(sa, &vt); RETURN_IF_FAILED(hr);
							RETURN_HR_IF(E_NOTIMPL, vt != VT_UNKNOWN);
							UINT dim = SafeArrayGetDim(sa);
							RETURN_HR_IF(E_NOTIMPL, dim != 1);
							LONG lbound;
							hr = SafeArrayGetLBound(sa, 1, &lbound); RETURN_IF_FAILED(hr);
							RETURN_HR_IF(E_NOTIMPL, lbound != 0);
							LONG ubound;
							hr = SafeArrayGetUBound(sa, 1, &ubound); RETURN_IF_FAILED(hr);

							bool pushed = childCollections.try_push_back(ObjectCollectionProperty{ fd->memid, std::move(name), { } }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

							for (LONG i = 0; i <= ubound; i++)
							{
								wil::com_ptr_nothrow<IUnknown> obj;
								hr = SafeArrayGetElement(sa, &i, &obj); RETURN_IF_FAILED(hr);
								pushed = childCollections.back().value.try_push_back(std::move(obj)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
							}
							break;
						}

						case VT_BOOL:
						{
							bool val = (V_BOOL(&result) == VARIANT_TRUE);
							auto value = wil::make_bstr_nothrow(val ? L"True" : L"False"); RETURN_IF_NULL_ALLOC(value);
							bool pushed = attributes.try_push_back(ValueProperty{ fd->memid, std::move(name), std::move(value) }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
							break;
						}

						default:
							RETURN_HR(E_NOTIMPL);
					}
				}
			}
		}

		if (attributes.size())
		{
			hr = ensureElementCreated.CreateStartElement(); RETURN_IF_FAILED(hr);

			for (auto& attr : attributes)
			{
				hr = writer->WriteAttributeString(
					attr.name.get(), SysStringLen(attr.name.get()), 
					attr.value.get(), SysStringLen(attr.value.get())); RETURN_IF_FAILED(hr);
			}
		}

		if (childObjects.size() || childCollections.size())
		{
			wil::com_ptr_nothrow<IXmlParent> objAsParent;
			hr = obj->QueryInterface(&objAsParent); RETURN_IF_FAILED(hr);

			for (auto& child : childObjects)
			{
				wil::unique_bstr childXmlElementName;
				hr = objAsParent->GetChildXmlElementName(child.dispid, child.value.get(), &childXmlElementName); RETURN_IF_FAILED(hr);
				if (childXmlElementName)
				{
					RETURN_HR(E_NOTIMPL);
				}
				else
				{
					hr = SaveToXmlInternal (child.value.get(), child.name.get(), writer, &ensureElementCreated); RETURN_IF_FAILED(hr);
				}
			}
	
			for (auto& coll : childCollections)
			{
				EnsureElementCreated ensureCollectionElementCreated (coll.name.get(), writer, &ensureElementCreated);

				for (auto& child : coll.value)
				{
					wil::unique_bstr xmlElementName;
					hr = objAsParent->GetChildXmlElementName(coll.dispid, child.get(), &xmlElementName); RETURN_IF_FAILED(hr);
					hr = SaveToXmlInternal (child.get(), xmlElementName.get(), writer, &ensureCollectionElementCreated); RETURN_IF_FAILED(hr);
				}

				hr = ensureCollectionElementCreated.CreateEndElement(); RETURN_IF_FAILED(hr);
			}
		}
	}

	hr = ensureElementCreated.CreateEndElement(); RETURN_IF_FAILED(hr);

	return S_OK;
}

HRESULT SaveToXml (IDispatch* obj, PCWSTR elementName, IStream* to, UINT nEncodingCodePage)
{
	wil::com_ptr_nothrow<IXmlWriterLite> writer;
	auto hr = CreateXmlWriter(IID_PPV_ARGS(&writer), nullptr); RETURN_IF_FAILED(hr);
	if (nEncodingCodePage != CP_UTF8)
	{
		com_ptr<IXmlWriterOutput> output;
		hr = CreateXmlWriterOutputWithEncodingCodePage (to, nullptr, nEncodingCodePage, &output); RETURN_IF_FAILED(hr);
		hr = writer->SetOutput(output.get()); RETURN_IF_FAILED(hr);
	}
	else
	{
		hr = writer->SetOutput(to); RETURN_IF_FAILED(hr);
	}
	hr = writer->SetProperty(XmlWriterProperty_Indent, TRUE);
	hr = writer->WriteStartDocument(XmlStandalone_Omit); RETURN_IF_FAILED(hr);
	hr = SaveToXmlInternal (obj, elementName, writer.get(), nullptr); RETURN_IF_FAILED(hr);
	hr = writer->WriteEndDocument(); RETURN_IF_FAILED(hr);

	return S_OK;
}

// ============================================================================

static HRESULT LoadFromXmlInternal (IXmlReader* reader, PCWSTR elementName, IDispatch* obj);

static HRESULT LoadCollection (IXmlReader* reader, const wchar_t* collectionElemName, IDispatch* obj, MEMBERID memid, SAFEARRAY** to)
{
	*to = nullptr;

	wil::com_ptr_nothrow<IXmlParent> objAsParent;
	auto hr = obj->QueryInterface(&objAsParent); RETURN_IF_FAILED(hr);

	vector_nothrow<wil::com_ptr_nothrow<IDispatch>> children;

	while(true)
	{
		XmlNodeType nodeType;
		hr = reader->Read(&nodeType); RETURN_IF_FAILED(hr);
		if (nodeType == XmlNodeType_Whitespace || nodeType == XmlNodeType_Comment)
		{
		}
		else if (nodeType == XmlNodeType_Element)
		{
			LPCWSTR entryName;
			hr = reader->GetLocalName(&entryName, nullptr); RETURN_IF_FAILED(hr);
			wil::com_ptr_nothrow<IDispatch> child;
			hr = objAsParent->CreateChild(memid, entryName, &child); RETURN_IF_FAILED(hr);
			hr = LoadFromXmlInternal (reader, entryName, child.get()); RETURN_IF_FAILED(hr);
			bool pushed = children.try_push_back (std::move(child)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}
		else if (nodeType == XmlNodeType_EndElement)
		{
			LPCWSTR endElementName;
			hr = reader->GetLocalName(&endElementName, nullptr); RETURN_IF_FAILED(hr);
			RETURN_HR_IF(E_UNEXPECTED, wcscmp(collectionElemName, endElementName));
			SAFEARRAYBOUND sabound = { .cElements = (ULONG)children.size(), .lLbound = 0 };
			SAFEARRAY* sa = SafeArrayCreate (VT_UNKNOWN, 1, &sabound); RETURN_IF_NULL_ALLOC(sa);
			auto freeSafeArray = wil::scope_exit([&sa] { SafeArrayDestroy(sa); sa = nullptr; });
			for (LONG i = 0; i < (LONG)children.size(); i++)
			{
				hr = SafeArrayPutElement (sa, &i, children[i].get()); RETURN_IF_FAILED(hr);
			}
			freeSafeArray.release();
			*to = sa;
			return S_OK;
		}
		else
			RETURN_HR(E_NOTIMPL);
	}
}

static HRESULT FindPutFunction (MEMBERID memid, ITypeInfo* typeInfo, TYPEATTR* typeAttr, VARTYPE* pvt)
{
	for (WORD i = 0; i < typeAttr->cFuncs; i++)
	{
		FUNCDESC* fd;
		auto hr = typeInfo->GetFuncDesc(i, &fd); RETURN_IF_FAILED(hr);
		auto releaseFundDesc = wil::scope_exit([typeInfo, fd] { typeInfo->ReleaseFuncDesc(fd); });
		if ((fd->invkind == INVOKE_PROPERTYPUT) && (fd->memid == memid))
		{
			RETURN_HR_IF(E_FAIL, fd->cParams != 1);
			*pvt = fd->lprgelemdescParam[0].tdesc.vt;
			return S_OK;
		}
	}

	RETURN_HR(DISP_E_MEMBERNOTFOUND);
}

static HRESULT LoadFromXmlInternal (IXmlReader* reader, PCWSTR elementName, IDispatch* obj)
{
	com_ptr<ITypeInfo> typeInfo;
	auto hr = obj->GetTypeInfo(0, 0x0409, &typeInfo); RETURN_IF_FAILED(hr);
	TYPEATTR* typeAttr;
	hr = typeInfo->GetTypeAttr(&typeAttr); RETURN_IF_FAILED(hr);
	auto releaseTypeAttr = wil::scope_exit([typeInfo, typeAttr] { typeInfo->ReleaseTypeAttr(typeAttr); });

	hr = reader->MoveToFirstAttribute(); RETURN_IF_FAILED(hr);
	while (hr == S_OK)
	{
		LPCWSTR attrName, attrValue;
		UINT attrNameLen, attrValueLen;
		hr = reader->GetLocalName(&attrName, &attrNameLen); RETURN_IF_FAILED(hr);
		hr = reader->GetValue(&attrValue, &attrValueLen); RETURN_IF_FAILED(hr);

		MEMBERID memid;
		hr = typeInfo->GetIDsOfNames(&const_cast<LPOLESTR&>(attrName), 1, &memid); RETURN_IF_FAILED(hr);
		FUNCDESC* fd = nullptr;
		for (WORD i = 0; i < typeAttr->cFuncs && !fd; i++)
		{
			hr = typeInfo->GetFuncDesc(i, &fd); RETURN_IF_FAILED(hr);
			if ((fd->invkind == INVOKE_PROPERTYPUT) && (fd->memid == memid))
			{
				RETURN_HR_IF(E_FAIL, fd->cParams != 1);
				break;
			}
			
			typeInfo->ReleaseFuncDesc(fd);
			fd = nullptr;
		}

		RETURN_HR_IF_NULL(DISP_E_MEMBERNOTFOUND, fd);
		auto releaseFD = wil::scope_exit([ti=typeInfo.get(), fd] { ti->ReleaseFuncDesc(fd); });

		wil::unique_variant valueVariant;
		VARTYPE vt = fd->lprgelemdescParam[0].tdesc.vt;
		switch (vt)
		{
			case VT_UI1:
			case VT_UI2:
			case VT_UI4:
			case VT_I1:
			case VT_I2:
			case VT_I4:
			{
				hr = InitVariantFromString(attrValue, &valueVariant); RETURN_IF_FAILED(hr);
				hr = VariantChangeTypeEx (&valueVariant, &valueVariant, InvariantLCID, 0, vt); RETURN_IF_FAILED(hr);
				break;
			}

			case VT_BSTR:
				hr = InitVariantFromString(attrValue, &valueVariant); RETURN_IF_FAILED(hr);
				break;

			case VT_BOOL:
				hr = InitVariantFromBoolean(!wcscmp(attrValue, L"True"), &valueVariant); RETURN_IF_FAILED(hr);
				break;

			case VT_USERDEFINED:
			{
				com_ptr<ITypeInfo> refTypeInfo;
				hr = typeInfo->GetRefTypeInfo (fd->lprgelemdescParam[0].tdesc.hreftype, &refTypeInfo); RETURN_IF_FAILED(hr);
				TYPEATTR* refTypeAttr;
				hr = refTypeInfo->GetTypeAttr(&refTypeAttr); RETURN_IF_FAILED(hr);
				auto releaseRefTypeAttr = wil::scope_exit([ti=refTypeInfo.get(), refTypeAttr] { ti->ReleaseTypeAttr(refTypeAttr); });
				if (refTypeAttr->typekind == TKIND_ENUM)
				{
					LONG value;
					hr = GetEnumValueFromName (refTypeInfo, refTypeAttr, attrValue, &value); RETURN_IF_FAILED(hr);
					hr = InitVariantFromInt32 (value, &valueVariant); RETURN_IF_FAILED(hr);
				}
				else
					RETURN_HR(E_NOTIMPL);
				break;
			}
			default:
				RETURN_HR(E_NOTIMPL);
		}

		DISPID named = DISPID_PROPERTYPUT;
		DISPPARAMS params = { .rgvarg = &valueVariant, .rgdispidNamedArgs=&named, .cArgs = 1, .cNamedArgs = 1 };
		wil::unique_variant result; // TODO: get rid of this
		EXCEPINFO exception;
		UINT uArgErr;
		hr = typeInfo->Invoke(obj, memid, DISPATCH_PROPERTYPUT, &params, &result, &exception, &uArgErr); RETURN_IF_FAILED(hr);

		hr = reader->MoveToNextAttribute(); RETURN_IF_FAILED(hr);
	}

	reader->MoveToElement();

	if (!reader->IsEmptyElement())
	{
		com_ptr<IXmlParent> objAsParent;
		hr = obj->QueryInterface(&objAsParent); RETURN_IF_FAILED(hr);

		// Try to read child elements.
		while (true)
		{
			XmlNodeType nodeType;
			hr = reader->Read(&nodeType); RETURN_IF_FAILED(hr);
			if (nodeType == XmlNodeType_Whitespace)
			{
			}
			else if (nodeType == XmlNodeType_Element)
			{
				LPCWSTR childElemName = nullptr;
				hr = reader->GetLocalName(&childElemName, nullptr); RETURN_IF_FAILED(hr);

				MEMBERID memid;
				auto hr = objAsParent->GetIDOfName(typeInfo, childElemName, &memid); RETURN_IF_FAILED(hr);
				VARTYPE vt;
				hr = FindPutFunction(memid, typeInfo.get(), typeAttr, &vt); RETURN_IF_FAILED(hr);

				if (vt == VT_DISPATCH)
				{
					com_ptr<IDispatch> child;
					hr = objAsParent->CreateChild(memid, childElemName, &child); RETURN_IF_FAILED(hr);
					hr = LoadFromXmlInternal (reader, childElemName, child.get()); RETURN_IF_FAILED(hr);
					wil::unique_variant value;
					hr = InitVariantFromDispatch(child.get(), &value); RETURN_IF_FAILED(hr);
					DISPID named = DISPID_PROPERTYPUT;
					DISPPARAMS params = { .rgvarg = &value, .rgdispidNamedArgs=&named, .cArgs = 1, .cNamedArgs = 1 };
					//wil::unique_variant result;
					EXCEPINFO exception;
					UINT uArgErr;
					hr = typeInfo->Invoke (obj, memid, DISPATCH_PROPERTYPUT, &params, nullptr, &exception, &uArgErr); RETURN_IF_FAILED(hr);
				}
				else if (vt == VT_SAFEARRAY)
				{
					SAFEARRAY* sa = nullptr;
					hr = LoadCollection(reader, childElemName, obj, memid, &sa); RETURN_IF_FAILED(hr);
					wil::unique_variant value;
					value.vt = VT_ARRAY | VT_DISPATCH;
					value.parray = sa;
					sa = nullptr;
					DISPID named = DISPID_PROPERTYPUT;
					DISPPARAMS params = { .rgvarg = &value, .rgdispidNamedArgs=&named, .cArgs = 1, .cNamedArgs = 1 };
					wil::unique_variant result;
					EXCEPINFO exception;
					UINT uArgErr;
					hr = typeInfo->Invoke (obj, memid, DISPATCH_PROPERTYPUT, &params, &result, &exception, &uArgErr); RETURN_IF_FAILED(hr);
				}
				else
				{
					RETURN_HR(E_NOTIMPL);
				}
			}
			else if (nodeType == XmlNodeType_EndElement)
			{
				LPCWSTR endElemName;
				hr = reader->GetLocalName(&endElemName, nullptr); RETURN_IF_FAILED(hr);
				RETURN_HR_IF(E_UNEXPECTED, wcscmp(elementName, endElemName));
				break;
			}
			else
			{
				WI_ASSERT(false); // TODO
			}
		}
	}

	return S_OK;
}

HRESULT LoadFromXml (IDispatch* obj, _In_opt_ PCWSTR expectedElementName, IStream* stream, UINT nEncodingCodePage)
{
	wil::com_ptr_nothrow<IXmlReader> reader;
	auto hr = CreateXmlReader(IID_PPV_ARGS(&reader), nullptr); RETURN_IF_FAILED(hr);
	if (nEncodingCodePage != CP_UTF8)
	{
		com_ptr<IXmlReaderInput> input;
		hr = CreateXmlReaderInputWithEncodingCodePage (stream, nullptr, nEncodingCodePage, TRUE, nullptr, &input); RETURN_IF_FAILED(hr);
		hr = reader->SetInput(stream); RETURN_IF_FAILED(hr);
	}
	else
	{
		hr = reader->SetInput(stream); RETURN_IF_FAILED(hr);
	}

	XmlNodeType nodeType;
	hr = reader->Read(&nodeType); RETURN_IF_FAILED(hr);
	if (nodeType != XmlNodeType_XmlDeclaration)
		RETURN_HR((HRESULT)WC_E_XMLDECL);
	while(SUCCEEDED(reader->Read(&nodeType)) && (nodeType == XmlNodeType_Whitespace))
		;

	RETURN_HR_IF((HRESULT)WC_E_DECLELEMENT, nodeType != XmlNodeType_Element);
	hr = reader->MoveToElement(); RETURN_IF_FAILED(hr);
	LPCWSTR elemName;
	hr = reader->GetLocalName(&elemName, nullptr); RETURN_IF_FAILED(hr);
	RETURN_HR_IF((HRESULT)WC_E_DECLELEMENT, expectedElementName && wcscmp(elemName, expectedElementName));

	return LoadFromXmlInternal(reader.get(), elemName, obj);
}
