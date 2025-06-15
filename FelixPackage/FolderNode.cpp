
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/inplace_function.h"
#include "Z80Xml.h"
#include "dispids.h"
#include "../FelixPackageUi/resource.h"
#include <KnownMonikers.h>

using namespace Microsoft::VisualStudio::Imaging;

struct FolderNode : IFolderNode, IParentNode, IFolderNodeProperties, IXmlParent, IVsPerPropertyBrowsing
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _parent;
	VSITEMID _itemId = VSITEMID_NIL;
	com_ptr<IChildNode> _next;
	com_ptr<IChildNode> _firstChild;
	wil::unique_bstr _name; // directory name, no path components needed
	WeakRefToThis _weakRefToThis;

public:
	HRESULT InitInstance()
	{
		auto hr = _weakRefToThis.InitInstance(static_cast<IFolderNode*>(this)); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	~FolderNode()
	{
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IFolderNodeProperties*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IFolderNodeProperties>(this, riid, ppvObject)
			|| TryQI<IFolderNode>(this, riid, ppvObject)
			|| TryQI<IChildNode>(this, riid, ppvObject)
			|| TryQI<IParentNode>(this, riid, ppvObject)
			|| TryQI<INode>(this, riid, ppvObject)
			|| TryQI<IXmlParent>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IFolderNodeProperties);

	#pragma region IFolderNodeProperties
	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		// Shown by VS at the top of the Properties Window.
		if (!_name || !_name.get()[0])
			return (*value = nullptr), S_OK;			
		*value = SysAllocString(_name.get()); RETURN_IF_NULL_ALLOC(*value);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Name (BSTR *value) override
	{
		if (!_name || !_name.get()[0])
			return (*value = nullptr), S_OK;			
		*value = SysAllocString(_name.get()); RETURN_IF_NULL_ALLOC(*value);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Name (BSTR value) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL);
		RETURN_HR_IF(E_INVALIDARG, !value || !value[0]);

		_name = wil::make_bstr_nothrow(value); RETURN_IF_NULL_ALLOC(_name);
		// TODO: property change notifications
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Items (SAFEARRAY** items) override
	{
		return GetItems(this, items);
	}

	virtual HRESULT STDMETHODCALLTYPE put_Items (SAFEARRAY* sa) override
	{
		return PutItems(sa, this);
	}
	#pragma endregion

	#pragma region IChildNode
	virtual VSITEMID STDMETHODCALLTYPE GetItemId() override
	{
		return _itemId;
	}

	virtual HRESULT STDMETHODCALLTYPE SetItemId (IParentNode* parent, VSITEMID id) override
	{
		RETURN_HR_IF(E_INVALIDARG, id == VSITEMID_NIL);
		RETURN_HR_IF(E_INVALIDARG, !parent);
		RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL);
		RETURN_HR_IF(E_UNEXPECTED, _parent);
		auto hr = parent->QueryInterface(IID_PPV_ARGS(_parent.addressof())); RETURN_IF_FAILED(hr);
		_itemId = id;
		return S_OK;
	}

	virtual HRESULT ClearItemId() override
	{
		RETURN_HR_IF(E_UNEXPECTED, _itemId == VSITEMID_NIL);
		RETURN_HR_IF(E_UNEXPECTED, !_parent);
		_itemId = VSITEMID_NIL;
		_parent = nullptr;
		return S_OK;
	}

	virtual HRESULT GetParent (IParentNode** ppParent) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !_parent);
		return _parent->QueryInterface(IID_PPV_ARGS(ppParent));
	}

	virtual IChildNode *STDMETHODCALLTYPE Next() override
	{
		return _next;
	}

	virtual void STDMETHODCALLTYPE SetNext (IChildNode *next) override
	{
		_next = next;
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty (VSHPROPID propid, VARIANT *pvar) override
	{
		HRESULT hr;

		RETURN_HR_IF(E_UNEXPECTED, !_parent); // callable only while in a hierarchy

		if (propid == VSHPROPID_Parent) // -1000
		{
			com_ptr<INode> parent;
			hr = _parent->QueryInterface(IID_PPV_ARGS(&parent)); RETURN_IF_FAILED(hr);
			return InitVariantFromInt32 (parent->GetItemId(), pvar);
		}

		if (   propid == VSHPROPID_FirstChild // -1001
			|| propid == VSHPROPID_FirstVisibleChild) // -2041
			return InitVariantFromInt32 (_firstChild ? _firstChild->GetItemId() : VSITEMID_NIL, pvar);

		if (   propid == VSHPROPID_NextSibling // -1002
			|| propid == VSHPROPID_NextVisibleSibling) // -2042
			return InitVariantFromInt32 (_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

		if (   propid == VSHPROPID_SaveName // -2002
			|| propid == VSHPROPID_Caption // -2003
			|| propid == VSHPROPID_Name // -2012
			|| propid == VSHPROPID_EditLabel // -2026
		)
			return InitVariantFromString(_name.get(), pvar);

		if (propid == VSHPROPID_IconIndex) // -2005
			return E_NOTIMPL;

		if (propid == VSHPROPID_Expandable) // -2006
			return InitVariantFromBoolean (_firstChild ? TRUE : FALSE, pvar);

		if (propid == VSHPROPID_ExpandByDefault) // -2011
			return InitVariantFromBoolean (FALSE, pvar);

		if (propid == VSHPROPID_IconHandle) // -2013
			return E_NOTIMPL;

		if (propid == VSHPROPID_BrowseObject) // -2018
			return InitVariantFromDispatch(this, pvar);

		if (propid == VSHPROPID_SupportsIconMonikers) // -2159
			return InitVariantFromBoolean (TRUE, pvar);

		if (propid == VSHPROPID_IconMonikerId) // -2161
			return InitVariantFromInt32(KnownImageIds::FolderClosed, pvar);

		if (propid == VSHPROPID_HasEnumerationSideEffects) // -2062
			return InitVariantFromBoolean (FALSE, pvar);

		if (propid == VSHPROPID_OpenFolderIconMonikerId) // -2163
			return InitVariantFromInt32(KnownImageIds::FolderOpened, pvar);

		#ifdef _DEBUG
		if (   propid == VSHPROPID_OpenFolderIconHandle // -2014
			|| propid == VSHPROPID_OpenFolderIconIndex // -2015
			|| propid == VSHPROPID_StateIconIndex // -2029
			|| propid == VSHPROPID_ItemDocCookie // -2034
			|| propid == VSHPROPID_IsHiddenItem // -2043
			|| propid == VSHPROPID_OverlayIconIndex	// -2048
			|| propid == VSHPROPID_ExternalItem // -2103
			|| propid == VSHPROPID_ProvisionalViewingStatus // -2112
			|| propid == -2177 // VSHPROPID_PreserveExpandCollapseState

		)
			return E_NOTIMPL;

		RETURN_HR(E_NOTIMPL);
		#else
		return E_NOTIMPL;
		#endif
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty (VSHPROPID propid, REFVARIANT var) override
	{
		HRESULT hr;

		RETURN_HR_IF(E_UNEXPECTED, !_parent); // callable only while in a hierarchy

		if (propid == VSHPROPID_Parent)
			RETURN_HR(E_UNEXPECTED); // Set this via SetItemId

		if (propid == VSHPROPID_EditLabel)
		{
			RETURN_HR_IF(E_INVALIDARG, var.vt != VT_BSTR);
			hr = RenameNode(var.bstrVal); RETURN_IF_FAILED_EXPECTED(hr);
			// If the above call reordered nodes, we need to select ourselves again.
			com_ptr<IVsUIHierarchyWindow> uiWindow;
			if (SUCCEEDED(GetHierarchyWindow(uiWindow.addressof())))
			{
				com_ptr<IVsUIHierarchy> hier;
				if (SUCCEEDED(FindHier(static_cast<IParentNode*>(this), IID_PPV_ARGS(&hier))))
					uiWindow->ExpandItem (hier, _itemId, EXPF_SelectItem);
			}

			return S_OK;
		}

		if (propid == VSHPROPID_SaveName || propid == VSHPROPID_Name)
		{
			// These two properties are meant to be set once, after this object is created.
			// Afterwards is should be renamed only using VSHPROPID_EditLabel.
			RETURN_HR_IF(E_UNEXPECTED, _name && _name.get()[0]);

			RETURN_HR_IF(E_INVALIDARG, var.vt != VT_BSTR);

			// The folder item is supposed to be created only for directories present in the file system.
			// But we can't check that since our caller might set the Name property before adding us to the hierarchy.

			_name = wil::make_bstr_nothrow(var.bstrVal); RETURN_IF_NULL_ALLOC(_name);

			// Also no need for notification since these properties are only set once.
			return S_OK;
		}

		#ifdef _DEBUG
		RETURN_HR(E_NOTIMPL);
		#else
		return E_NOTIMPL;
		#endif
	}

	virtual HRESULT STDMETHODCALLTYPE GetGuidProperty (VSHPROPID propid, GUID *pguid) override
	{
		if (propid == VSHPROPID_TypeGuid) // -1004
			return (*pguid = GUID_ItemType_PhysicalFolder), S_OK;

		if (   propid == VSHPROPID_OpenFolderIconMonikerGuid // -2162
			|| propid == VSHPROPID_IconMonikerGuid // -2160
		)
			return (*pguid = KnownImageIds::ImageCatalogGuid), S_OK;

		#ifdef _DEBUG
		RETURN_HR(E_NOTIMPL);
		#else
		return E_NOTIMPL;
		#endif
	}

	virtual HRESULT STDMETHODCALLTYPE SetGuidProperty (VSHPROPID propid, REFGUID rguid) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetCanonicalName (BSTR *pbstrName) override
	{
		return E_NOTIMPL; // we'll never support this for folders
	}

	virtual HRESULT STDMETHODCALLTYPE IsItemDirty (IUnknown *punkDocData, BOOL *pfDirty) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE QueryStatusCommand (const GUID *pguidCmdGroup, OLECMD* pCmd, OLECMDTEXT *pCmdText) override
	{
		return OLECMDERR_E_NOTSUPPORTED;
	}

	virtual HRESULT STDMETHODCALLTYPE ExecCommand (const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
	{
		HRESULT hr;

		if (*pguidCmdGroup == GUID_VsUIHierarchyWindowCmds)
		{
			if (nCmdID == UIHWCMDID_RightClick)
			{
				POINTS pts;
				memcpy (&pts, &pvaIn->uintVal, 4);

				com_ptr<IVsUIShell> shell;
				hr = serviceProvider->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);
				hr = shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_FOLDERNODE, pts, nullptr); RETURN_IF_FAILED(hr);
				return S_OK;
			}

			if (nCmdID == UIHWCMDID_DoubleClick || nCmdID == UIHWCMDID_EnterKey)
			{
				// VC++ expands or colappses the node
				return OLECMDERR_E_NOTSUPPORTED;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		return OLECMDERR_E_NOTSUPPORTED;
	}
	#pragma endregion

	#pragma region IParentNode
	virtual IChildNode *STDMETHODCALLTYPE FirstChild() override
	{
		return _firstChild;
	}

	virtual void STDMETHODCALLTYPE SetFirstChild (IChildNode *child) override
	{
		_firstChild = child;
	}
	#pragma endregion

	#pragma region IXmlParent : IUnknown
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IDispatch* child, BSTR* xmlElementNameOut) override
	{
		if (dispidProperty == dispidItems)
		{
			if (auto x = wil::try_com_query_nothrow<IFileNode>(child))
			{
				*xmlElementNameOut = SysAllocString(FileElementName); RETURN_IF_NULL_ALLOC(*xmlElementNameOut);
				return S_OK;
			}
			else if (auto x = wil::try_com_query_nothrow<IFolderNode>(child))
			{
				*xmlElementNameOut = SysAllocString(FolderElementName); RETURN_IF_NULL_ALLOC(*xmlElementNameOut);
				return S_OK;
			}

			RETURN_HR(E_NOTIMPL);
		}

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE CreateChild (DISPID dispidProperty, PCWSTR xmlElementName, IDispatch** childOut) override
	{
		HRESULT hr;

		if (dispidProperty == dispidItems)
		{
			if (!wcscmp(xmlElementName, FileElementName))
			{
				wil::com_ptr_nothrow<IFileNode> file;
				hr = MakeFileNode(&file); RETURN_IF_FAILED(hr);
				com_ptr<IFileNodeProperties> fileProps;
				hr = file->QueryInterface(&fileProps); RETURN_IF_FAILED(hr);
				*childOut = fileProps.detach();
				return S_OK;
			}
			else if (!wcscmp(xmlElementName, FolderElementName))
			{
				com_ptr<IFolderNode> folder;
				hr = MakeFolderNode(&folder); RETURN_IF_FAILED(hr);
				com_ptr<IFolderNodeProperties> folderProps;
				hr = folder->QueryInterface(IID_PPV_ARGS(&folderProps)); RETURN_IF_FAILED(hr);
				*childOut = folderProps.detach();
				return S_OK;
			}

			RETURN_HR(E_NOTIMPL);
		}

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IFolderNode
	virtual IParentNode* AsParentNode() override { return this; }
	#pragma endregion

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL *pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidFolderName)
		{
			*fDefault = FALSE;
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override
	{
		if (dispid == dispidFolderName)
		{
			// Too complicated to allow the user to edit this in the Properties window.
			*fReadOnly = TRUE;
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetClassName (BSTR *pbstrClassName) override
	{
		// Shown by VS at the top of the Properties Window.
		*pbstrClassName = SysAllocString(L"Folder");
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE CanResetPropertyValue (DISPID dispid, BOOL* pfCanReset) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE ResetPropertyValue (DISPID dispid) override { return E_NOTIMPL; }
	#pragma endregion

	HRESULT SortAfterRename (IVsHierarchy* hier, IParentNode* parent)
	{
		HRESULT hr;

		auto first = wil::try_com_query_nothrow<IFolderNode>(parent->FirstChild());
		if (first->Next() && wil::try_com_query_nothrow<IFolderNode>(first->Next()))
		{
			// We're not the only folder. Need to attempt reordering.
			com_ptr<IVsHierarchyEvents> hierEvents;
			hr = hier->QueryInterface(IID_PPV_ARGS(&hierEvents)); RETURN_IF_FAILED(hr);

			wil::unique_variant n;

			// Special cases at the beginning of the list: Are we the first folder? 
			if (this == first)
			{
				// Yes. Do we need to be moved somewhere after the next one?
				if (SUCCEEDED(_next->GetProperty(VSHPROPID_SaveName, &n))
					&& V_VT(&n) == VT_BSTR
					&& VarBstrCmp(_name.get(), V_BSTR(&n), 0, 0) == VARCMP_GT)
				{
					// Yes
					auto moveAfter = _next.get();
					com_ptr<IFolderNode> nf;
					while (moveAfter->Next()
						&& (nf = wil::try_com_query_nothrow<IFolderNode>(moveAfter->Next()))
						&& SUCCEEDED(nf->GetProperty(VSHPROPID_SaveName, &n))
						&& V_VT(&n) == VT_BSTR
						&& VarBstrCmp(_name.get(), V_BSTR(&n), 0, 0) == VARCMP_GT)
					{
						moveAfter = nf.get();
					}

					parent->SetFirstChild(_next);
					_next = moveAfter->Next();
					moveAfter->SetNext(this);

					hierEvents->OnInvalidateItems(parent->GetItemId());
				}
			}
			// If we're not the first folder, do we need to become the first?
			else if (SUCCEEDED(first->GetProperty(VSHPROPID_SaveName, &n))
				&& V_VT(&n) == VT_BSTR
				&& VarBstrCmp(_name.get(), V_BSTR(&n), 0, 0) == VARCMP_LT)
			{
				// Yes
				auto prev = parent->FirstChild();
				while(prev->Next() != this)
					prev = prev->Next();

				prev->SetNext(_next);
				_next = parent->FirstChild();
				parent->SetFirstChild(this);

				hierEvents->OnInvalidateItems(parent->GetItemId());
			}
			else
			{
				auto moveAfter = parent->FirstChild();
				com_ptr<IFolderNode> nf;
				while (moveAfter->Next()
					&& (nf = wil::try_com_query_nothrow<IFolderNode>(moveAfter->Next()))
					&& SUCCEEDED(nf->GetProperty(VSHPROPID_SaveName, &n))
					&& V_VT(&n) == VT_BSTR
					&& VarBstrCmp(_name.get(), V_BSTR(&n), 0, 0) != VARCMP_LT)
				{
					moveAfter = moveAfter->Next();
				}

				if (moveAfter != this)
				{
					auto prev = parent->FirstChild();
					while(prev->Next() != this)
						prev = prev->Next();

					prev->SetNext(_next);
					_next = moveAfter->Next();
					moveAfter->SetNext(this);

					hierEvents->OnInvalidateItems(parent->GetItemId());
				}
			}
		}
		else
			WI_ASSERT(first == this);

		return S_OK;
	}

	HRESULT RenameNode (const wchar_t* proposedName)
	{
		HRESULT hr;

		com_ptr<IParentNode> parent;
		_parent->QueryInterface(&parent);

		auto newName = wil::make_bstr_nothrow(proposedName); RETURN_IF_NULL_ALLOC(newName);
		LUtilFixFilename(newName.get());

		// Do we have a sibling with the new name?
		for (auto c = parent->FirstChild(); c; c = c->Next())
		{
			wil::unique_variant n;
			if (SUCCEEDED(c->GetProperty(VSHPROPID_SaveName, &n))
				&& V_VT(&n) == VT_BSTR
				&& VarBstrCmp(newName.get(), V_BSTR(&n), 0, 0) == VARCMP_EQ)
			{
				return SetErrorInfo0(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), IDS_NAME_ALREADY_EXISTS);
			}
		}

		com_ptr<IVsHierarchy> hier;
		hr = FindHier(static_cast<IChildNode*>(this), IID_PPV_ARGS(&hier)); RETURN_IF_FAILED(hr);

		hr = QueryEditProjectFile(hier); RETURN_IF_FAILED_EXPECTED(hr);

		// TODO: there's a lot that needs to be called in IVsTrackProjectDocumentsEvents2

		wil::unique_process_heap_string oldFullPath;
		hr = GetPathOf(this, oldFullPath); RETURN_IF_FAILED(hr);

		wil::unique_process_heap_string newFullPath;
		hr = GetPathTo(this, newFullPath); RETURN_IF_FAILED(hr);
		hr = wil::str_concat_nothrow (newFullPath, L"\\", newName); RETURN_IF_FAILED(hr);

		// Make a list of all open documents that are our descendants, and their paths, before folder renaming.
		vector_nothrow<std::pair<com_ptr<IFileNode>, wil::unique_process_heap_string>> openDocs;
		com_ptr<IVsRunningDocumentTable> rdt;
		if (SUCCEEDED(serviceProvider->QueryService(SID_SVsRunningDocumentTable, IID_PPV_ARGS(rdt.addressof()))))
		{
			stdext::inplace_function<HRESULT(IParentNode*)> enumDescendants;
			enumDescendants = [&enumDescendants, &openDocs, rdt=rdt.get()](IParentNode* parent) -> HRESULT
				{
					for (auto c = parent->FirstChild(); c; c = c->Next())
					{
						if (auto file = wil::try_com_query_nothrow<IFileNode>(c))
						{
							// TODO: call IVsTrackProjectDocuments2::OnQueryRenameFile here, call OnAfterRenameFile after renaming.

							wil::unique_variant docCookie;
							if (SUCCEEDED(file->GetProperty (VSHPROPID_ItemDocCookie, &docCookie))
								&& (docCookie.vt == VT_VSCOOKIE)
								&& (V_VSCOOKIE(&docCookie) != VSDOCCOOKIE_NIL))
							{
								wil::unique_process_heap_string path;
								if (SUCCEEDED(GetPathOf(file, path)))
									openDocs.try_push_back({ std::move(file), std::move(path) });
							}
						}
						else if (auto cAsParent = wil::try_com_query_nothrow<IParentNode>(c))
						{
							auto hr = enumDescendants(cAsParent); RETURN_IF_FAILED(hr);
						}
					}

					return S_OK;
				};
			hr = enumDescendants(this); RETURN_IF_FAILED(hr);
		}

		BOOL bres = MoveFile(oldFullPath.get(), newFullPath.get());
		if (!bres)
			return HRESULT_FROM_WIN32(GetLastError());
		auto oldName = std::move(_name);
		_name = std::move(newName);

		// We finished the renaming; now try to reorder the nodes and to send notifications.
		// we don't fail the renaming if these operations fail.

		// For every open file under the renamed folder, send notification about changed moniker.
		for (auto& od : openDocs)
		{
			wil::unique_process_heap_string newPath;
			if (SUCCEEDED(GetPathOf(od.first, newPath)))
				rdt->RenameDocument(od.second.get(), newPath.get(), HIERARCHY_DONTCHANGE, VSITEMID_NIL);
		}

		SortAfterRename(hier, parent);

		com_ptr<IVsHierarchyEvents> sink;
		hr = hier->QueryInterface(IID_PPV_ARGS(&sink)); RETURN_IF_FAILED(hr);
		sink->OnPropertyChanged (_itemId, VSHPROPID_SaveName, 0);
		sink->OnPropertyChanged (_itemId, VSHPROPID_Caption, 0);
		sink->OnPropertyChanged (_itemId, VSHPROPID_Name, 0);
		sink->OnPropertyChanged (_itemId, VSHPROPID_EditLabel, 0);
		sink->OnPropertyChanged (_itemId, VSHPROPID_DescriptiveName, 0);

		// Make sure the property browser is updated.
		com_ptr<IVsUIShell> uiShell;
		hr = serviceProvider->QueryService (SID_SVsUIShell, &uiShell);
		if (SUCCEEDED(hr))
			uiShell->RefreshPropertyBrowser(DISPID_UNKNOWN); // refresh all properties

		hier.try_query<IPropertyNotifySink>()->OnChanged(dispidItems);

		return S_OK;
	}
};

HRESULT MakeFolderNode (IFolderNode** ppFolder)
{
	auto p = com_ptr(new (std::nothrow) FolderNode()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*ppFolder = p.detach();
	return S_OK;
}

