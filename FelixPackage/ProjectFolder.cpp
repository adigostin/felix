
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include <KnownMonikers.h>

using namespace Microsoft::VisualStudio::Imaging;

struct ProjectFolder : IFolderNode, IParentNode, IFolderNodeProperties
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _hier;
	VSITEMID _itemId = VSITEMID_NIL;
	VSITEMID _parentItemId = VSITEMID_NIL;
	com_ptr<IChildNode> _next;
	com_ptr<IChildNode> _firstChild;
	wil::unique_bstr _name; // directory name, no path components needed

public:
	HRESULT InitInstance()
	{
		return S_OK;
	}

	~ProjectFolder()
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
			)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IFolderNodeProperties);

	#pragma region IFolderNode
	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		// Shown by VS at the top of the Properties Window.
		if (!_name || !_name.get()[0])
			return (*value = nullptr), S_OK;			
		*value = SysAllocString(_name.get()); RETURN_IF_NULL_ALLOC(*value);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_FolderName (BSTR *value) override
	{
		if (!_name || !_name.get()[0])
			return (*value = nullptr), S_OK;			
		*value = SysAllocString(_name.get()); RETURN_IF_NULL_ALLOC(*value);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_FolderName (LPCOLESTR value) override
	{
		RETURN_HR_IF(E_INVALIDARG, !value || !value[0]);

		_name = wil::make_bstr_nothrow(value); RETURN_IF_NULL_ALLOC(_name);
		// TODO: property change notifications
		return S_OK;
	}
	#pragma endregion

	#pragma region IChildNode
    virtual VSITEMID STDMETHODCALLTYPE GetItemId() override
	{
		return _itemId;
	}

	virtual HRESULT STDMETHODCALLTYPE SetItemId (IRootNode* root, VSITEMID id) override
	{
		if (id != VSITEMID_NIL)
		{
			// setting it
			RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL);
			RETURN_HR_IF(E_UNEXPECTED, _hier);
		}
		else
		{
			// clearing it
			RETURN_HR_IF(E_UNEXPECTED, _itemId == VSITEMID_NIL);
			RETURN_HR_IF(E_UNEXPECTED, !_hier);
		}

		auto hr = root->QueryInterface(&_hier); RETURN_IF_FAILED(hr);
		_itemId = id;
		return S_OK;
	}

    virtual HRESULT STDMETHODCALLTYPE GetMkDocument (BSTR *pbstrMkDocument) override
	{
		return E_NOTIMPL; // We'll never support this
		/*
		if (!_name || !_name[0])
			RETURN_HR(E_UNEXPECTED);

		com_ptr<IChildNode> parentItem;
		auto hr = _parent->Resolve(&parentItem); RETURN_IF_FAILED(hr);
		unique_bstr parentMk;
		hr = parentItem->GetMkDocument(&parentMk); RETURN_IF_FAILED(hr);

		wil::unique_hlocal_string filePath;
		hr = PathAllocCombine (parentMk, _name, PathFlags, &filePath); RETURN_IF_FAILED(hr);
		*pbstrMkDocument = SysAllocString(filePath.get()); RETURN_IF_NULL_ALLOC(*pbstrMkDocument);
		return S_OK;
		*/
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
		if (propid == VSHPROPID_Parent) // -1000
			return InitVariantFromInt32 (_parentItemId, pvar);

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
			|| propid == VSHPROPID_ProvisionalViewingStatus // -2112
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

		if (propid == VSHPROPID_Parent)
		{
			RETURN_HR_IF(E_INVALIDARG, var.vt != VT_VSITEMID);
			if (V_VSITEMID(&var) != VSITEMID_NIL)
			{
				// setting it
				RETURN_HR_IF(E_UNEXPECTED, _parentItemId != VSITEMID_NIL);
			}
			else
			{
				// clearing it
				RETURN_HR_IF(E_UNEXPECTED, _parentItemId == VSITEMID_NIL);
			}

			_parentItemId = V_VSITEMID(&var);
			return S_OK;
		}

		if (propid == VSHPROPID_EditLabel)
		{
			RETURN_HR_IF(E_INVALIDARG, var.vt != VT_BSTR);
			wil::unique_bstr newName;
			hr = RenameDirOnFileSystem(var.bstrVal, &newName); RETURN_IF_FAILED_EXPECTED(hr);
			hr = RenameNode(newName.get()); RETURN_IF_FAILED(hr);
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

    virtual HRESULT STDMETHODCALLTYPE QueryStatus (const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText) override
	{
		return OLECMDERR_E_NOTSUPPORTED;
	}

    virtual HRESULT STDMETHODCALLTYPE Exec (const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
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

	virtual HRESULT STDMETHODCALLTYPE GetHierarchy (REFIID riid, void **ppvObject) override
	{
		return _hier->QueryInterface(riid, ppvObject);
	}
	#pragma endregion

	#pragma warning(push)
	#pragma warning(disable: 4995)
	HRESULT RenameDirOnFileSystem (const wchar_t* proposedName, BSTR* pbstrNewName)
	{
		HRESULT hr;

		com_ptr<IVsHierarchy> hier;
		hr = _hier->QueryInterface(&hier); RETURN_IF_FAILED(hr);

		hr = QueryEditProjectFile(hier); RETURN_IF_FAILED_EXPECTED(hr);

		auto newName = wil::make_bstr_nothrow(proposedName); RETURN_IF_NULL_ALLOC(newName);
		LUtilFixFilename(newName.get());

		wil::unique_process_heap_string oldFullPath;
		hr = GetPathOf(hier, _itemId, oldFullPath); RETURN_IF_FAILED(hr);

		wil::unique_process_heap_string newFullPath;
		hr = GetPathTo (hier, _parentItemId, newFullPath); RETURN_IF_FAILED(hr);
		hr = wil::str_concat_nothrow (newFullPath, L"\\", newName); RETURN_IF_FAILED(hr);

		BOOL bres = MoveFile(oldFullPath.get(), newFullPath.get());
		if (!bres)
			return HRESULT_FROM_WIN32(GetLastError());

		*pbstrNewName = newName.release();
		return S_OK;
	}

	HRESULT RenameNode (const wchar_t* newName)
	{
		HRESULT hr;

		_name = wil::make_bstr_nothrow(newName); RETURN_IF_NULL_ALLOC(_name);
		
		com_ptr<IEnumHierarchyEvents> enu;
		hr = _hier.try_query<IRootNode>()->EnumHierarchyEventSinks(&enu); RETURN_IF_FAILED(hr);
		com_ptr<IVsHierarchyEvents> sink;
		ULONG fetched;
		while (SUCCEEDED(enu->Next(1, &sink, &fetched)) && fetched)
		{
			sink->OnPropertyChanged (_itemId, VSHPROPID_SaveName, 0);
			sink->OnPropertyChanged (_itemId, VSHPROPID_Caption, 0);
			sink->OnPropertyChanged (_itemId, VSHPROPID_Name, 0);
			sink->OnPropertyChanged (_itemId, VSHPROPID_EditLabel, 0);
			sink->OnPropertyChanged (_itemId, VSHPROPID_DescriptiveName, 0);
		}

		// Make sure the property browser is updated.
		com_ptr<IVsUIShell> uiShell;
		hr = serviceProvider->QueryService (SID_SVsUIShell, &uiShell); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
			uiShell->RefreshPropertyBrowser(DISPID_UNKNOWN); // refresh all properties

		return S_OK;
	}
	#pragma warning(pop)
};

HRESULT MakeProjectFolder (IFolderNode** ppFolder)
{
	auto p = com_ptr(new (std::nothrow) ProjectFolder()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*ppFolder = p.detach();
	return S_OK;
}

