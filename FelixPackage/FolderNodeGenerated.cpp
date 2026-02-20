
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/inplace_function.h"
#include "Z80Xml.h"
#include "dispids.h"
#include "../FelixPackageUi/resource.h"

using namespace Microsoft::VisualStudio::Imaging;

struct FolderNodeGenerated : IFolderNode, IParentNode
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _parent;
	VSITEMID _itemId = VSITEMID_NIL;
	com_ptr<IChildNode> _next;
	com_ptr<IChildNode> _firstChild;
	WeakRefToThis _weakRefToThis;

public:
	HRESULT InitInstance()
	{
		HRESULT hr;
		hr = _weakRefToThis.InitInstance(static_cast<IFolderNode*>(this)); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IFolderNode*>(this), riid, ppvObject)
			|| TryQI<IFolderNode>(this, riid, ppvObject)
			|| TryQI<IChildNode>(this, riid, ppvObject)
			|| TryQI<IParentNode>(this, riid, ppvObject)
			|| TryQI<INode>(static_cast<IParentNode*>(this), riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IChildNode
	virtual VSITEMID STDMETHODCALLTYPE GetItemId() override
	{
		return _itemId;
	}

	virtual HRESULT STDMETHODCALLTYPE SetItemId (IProjectNode* root, IParentNode* parent) override
	{
		RETURN_HR_IF(E_INVALIDARG, !parent);
		RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL);
		RETURN_HR_IF(E_UNEXPECTED, _parent);
		auto hr = parent->QueryInterface(IID_PPV_ARGS(_parent.addressof())); RETURN_IF_FAILED(hr);
		_itemId = VSITEMID_GENFILES;
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

	virtual HRESULT STDMETHODCALLTYPE GetProperty (IProjectNode* proj, VSHPROPID propid, VARIANT *pvar) override
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
		)
			return InitVariantFromString(genFilesStr.get(), pvar);

		if (propid == VSHPROPID_Expandable) // -2006
			return InitVariantFromBoolean (_firstChild ? TRUE : FALSE, pvar);

		if (propid == VSHPROPID_ExpandByDefault) // -2011
			return InitVariantFromBoolean (FALSE, pvar);

		if (propid == VSHPROPID_SupportsIconMonikers) // -2159
			return InitVariantFromBoolean (TRUE, pvar);

		if (propid == VSHPROPID_IconMonikerId) // -2161
			return InitVariantFromInt32(KnownImageIds::FolderClosed, pvar);

		if (propid == VSHPROPID_OpenFolderIconMonikerId) // -2163
			return InitVariantFromInt32(KnownImageIds::FolderOpened, pvar);

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty (IProjectNode* proj, VSHPROPID propid, REFVARIANT var) override
	{
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

	virtual HRESULT STDMETHODCALLTYPE IsItemDirty (IUnknown *punkDocData, BOOL *pfDirty) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE QueryStatusCommand (IProjectNode* proj, const GUID *pguidCmdGroup, OLECMD* pCmd, OLECMDTEXT *pCmdText) override
	{
		return OLECMDERR_E_NOTSUPPORTED;
	}

	virtual HRESULT STDMETHODCALLTYPE ExecCommand (IProjectNode* proj, const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
	{
		HRESULT hr;

		if (*pguidCmdGroup == GUID_VsUIHierarchyWindowCmds)
		{
			if (nCmdID == UIHWCMDID_RightClick)
			{
				POINTS pts;
				memcpy (&pts, &pvaIn->uintVal, 4);
				hr = uiShell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_FOLDERNODE, pts, nullptr); RETURN_IF_FAILED(hr);
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

	virtual HRESULT STDMETHODCALLTYPE GetCanonicalName (IProjectNode* proj, BSTR* pbstrName) override
	{
		wil::unique_process_heap_string path;
		auto hr = GetPathOf(proj, this, path, true); RETURN_IF_FAILED(hr);
		*pbstrName = SysAllocString(path.get()); RETURN_IF_NULL_ALLOC(*pbstrName);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetMkDocument (IProjectNode* proj, BSTR* pbstrMkDocument) override
	{
		wil::unique_process_heap_string path;
		auto hr = GetPathOf(proj, this, path, false); RETURN_IF_FAILED(hr);
		*pbstrMkDocument = SysAllocString(path.get()); RETURN_IF_NULL_ALLOC(*pbstrMkDocument);
		return S_OK;
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

	#pragma region IFolderNode
	virtual IParentNode* AsParentNode() override { return this; }
	#pragma endregion
};

HRESULT MakeFolderNodeGenerated (IFolderNode** ppFolder)
{
	auto p = com_ptr(new (std::nothrow) FolderNodeGenerated()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*ppFolder = p.detach();
	return S_OK;
}

