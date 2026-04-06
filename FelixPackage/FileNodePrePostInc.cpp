
#include "pch.h"
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
#include "../FelixPackageUi/resource.h"
#include "dispids.h"

using namespace Microsoft::VisualStudio::Imaging;

class FileNodePrePostInc : public IFileNode
{
	ULONG _refCount = 0;
	VSITEMID _itemId = VSITEMID_NIL;
	wil::com_ptr_nothrow<IChildNode> _next;
	com_ptr<IWeakRef> _parent;
	VSCOOKIE _docCookie = VSDOCCOOKIE_NIL;
	bool _post;
	WeakRefToThis _weakRefToThis;

public:
	HRESULT InitInstance (bool post)
	{
		HRESULT hr;
		hr = _weakRefToThis.InitInstance(static_cast<IFileNode*>(this)); RETURN_IF_FAILED(hr);
		_post = post;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IFileNode*>(this), riid, ppvObject)
			|| TryQI<IFileNode>(this, riid, ppvObject)
			|| TryQI<IChildNode>(this, riid, ppvObject)
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
	virtual VSITEMID STDMETHODCALLTYPE GetItemId() override { return _itemId; }

	virtual HRESULT GetParent (IParentNode** ppParent) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !_parent);
		return _parent->QueryInterface(IID_PPV_ARGS(ppParent));
	}

	virtual HRESULT SetParent (IParentNode* parent) override
	{
		if (parent)
		{
			RETURN_HR_IF(E_UNEXPECTED, _parent);
			auto hr = parent->QueryInterface(IID_PPV_ARGS(_parent.addressof())); RETURN_IF_FAILED(hr);
		}
		else
		{
			RETURN_HR_IF(E_UNEXPECTED, !_parent);
			_parent = nullptr;
		}
		
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetItemId (IProjectNode* root) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL);
		_itemId = _post ? VSITEMID_POSTINCLUDE : VSITEMID_PREINCLUDE;
		return S_OK;
	}
		
	virtual HRESULT ClearItemId() override
	{
		RETURN_HR_IF(E_UNEXPECTED, _itemId == VSITEMID_NIL);
		_itemId = VSITEMID_NIL;
		return S_OK;
	}

	virtual IChildNode* STDMETHODCALLTYPE Next() override { return _next.get(); }

	virtual void STDMETHODCALLTYPE SetNext (IChildNode* next) override
	{
		//WI_ASSERT (!_next);
		_next = next;
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty (IProjectNode* proj, VSHPROPID propid, VARIANT* pvar) override
	{
		HRESULT hr;
		RETURN_HR_IF(E_UNEXPECTED, !_parent); // callable only while in a hierarchy
		switch (propid)
		{
			case VSHPROPID_Parent: // -1000
			{
				com_ptr<INode> parent;
				hr = _parent->QueryInterface(IID_PPV_ARGS(parent.addressof())); RETURN_IF_FAILED(hr);
				return InitVariantFromInt32 (parent->GetItemId(), pvar);
			}

			case VSHPROPID_FirstChild: // -1001
				return InitVariantFromInt32 (VSITEMID_NIL, pvar);

			case VSHPROPID_NextSibling: // -1002
				return InitVariantFromInt32 (_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

			case VSHPROPID_SaveName: // -2002
			case VSHPROPID_Caption: // -2003
			case VSHPROPID_Name: // -2012
				return InitVariantFromString (_post ? postincludeFilename.get() : preincludeFilename.get(), pvar);

			case VSHPROPID_ItemDocCookie: // -2034
				return InitVariantFromInt32 (_docCookie, pvar);

			case VSHPROPID_NextVisibleSibling: // -2042
				return InitVariantFromInt32 (_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

			case VSHPROPID_DescriptiveName: // -2108
			{
				// Tooltip when hovering the document tab with the mouse, maybe other things too.
				wil::unique_bstr path;
				hr = GetMkDocument(proj, &path); RETURN_IF_FAILED(hr);
				pvar->vt = VT_BSTR;
				pvar->bstrVal = path.release();
				return S_OK;
			}

			case VSHPROPID_SupportsIconMonikers: // -2159
				return InitVariantFromBoolean (TRUE, pvar);

			case VSHPROPID_IconMonikerId: // -2161
			{
				auto& _name = _post ? postincludeFilename : preincludeFilename;
				auto ext = PathFindExtension(_name.get());
				
				if (!_wcsicmp(ext, L".asm"))
					return InitVariantFromInt32(KnownImageIds::ASMFile, pvar);

				if (!_wcsicmp(ext, L".inc"))
					return InitVariantFromInt32(KnownImageIds::TextFile, pvar);

				return E_NOTIMPL;
			}

			default:
				return E_NOTIMPL;
		}
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty (IProjectNode* proj, VSHPROPID propid, const VARIANT& var) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !_parent); // callable only while in a hierarchy

		switch(propid)
		{
			case VSHPROPID_ItemDocCookie: // -2034
				RETURN_HR_IF(E_INVALIDARG, var.vt != VT_VSCOOKIE);
				_docCookie = var.lVal;// V_VSCOOKIE(&var);
				return S_OK;

			default:
				return E_NOTIMPL;
		}
	}

	virtual HRESULT STDMETHODCALLTYPE GetGuidProperty (VSHPROPID propid, GUID* pguid) override
	{
		if (propid == VSHPROPID_TypeGuid)
		{
			*pguid = GUID_ItemType_PhysicalFile;
			return S_OK;
		}

		if (propid == VSHPROPID_IconMonikerGuid) // -2160
			return (*pguid = KnownImageIds::ImageCatalogGuid), S_OK;

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE SetGuidProperty (VSHPROPID propid, REFGUID rguid) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsItemDirty (IUnknown *punkDocData, BOOL *pfDirty) override
	{
		wil::com_ptr_nothrow<IVsPersistDocData> docData;
		auto hr = punkDocData->QueryInterface(&docData); RETURN_IF_FAILED(hr);
		hr = docData->IsDocDataDirty(pfDirty); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryStatusCommand (IProjectNode* proj, const GUID *pguidCmdGroup, OLECMD* pCmd, OLECMDTEXT *pCmdText) override
	{
		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97)
		{
			// These are the cmdidXxxYyy constants from stdidcmd.h
			if (pCmd->cmdID >= 0xF000)
			{
				// debugger stuff, ignored for now
				pCmd->cmdf = OLECMDF_SUPPORTED;
			}
			else if (pCmd->cmdID >= 946 && pCmd->cmdID <= 957)
			{
				// refactoring stuff
				pCmd->cmdf = 0;
			}
			else if (pCmd->cmdID >= 122 && pCmd->cmdID <= cmdidSelectAllFields)
			{
				// database stuff
				pCmd->cmdf = 0;
			}
			else
			{
				switch (pCmd->cmdID)
				{
					case cmdidCopy: // 15
					case cmdidCut: // 16
					case cmdidMultiLevelRedo: // 30
					case cmdidMultiLevelUndo: // 44
					case cmdidNewProject: // 216
					case cmdidFileOpen: // 222
					case cmdidSaveSolution: // 224
					case cmdidGoto: // 231
					case cmdidOpen: // 261
					case cmdidFindInFiles: // 277
					case cmdidShellNavBackward: // 809
					case cmdidShellNavForward: // 810
						pCmd->cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
						break;

					case cmdidSaveProjectItem: // 331
					case cmdidSaveProjectItemAs: // 226
						// TODO: enable it only if it's in the running document list
						pCmd->cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
						break;

					// database stuff
					case cmdidVerifySQL: // 107
					case cmdidPrimaryKey: // 109
					case cmdidSortAscending: // 112
					case cmdidSortDescending: // 113
					case cmdidAppendQuery: // 114
					case cmdidDeleteQuery: // 116
					case cmdidMakeTableQuery: // 117
					case cmdidSelectQuery: // 118
					case cmdidUpdateQuery: // 119
					case cmdidTotals: // 121
					case cmdidRemoveFilter: // 164
					case cmdidJoinLeftAll: // 169
					case cmdidJoinRightAll: // 170
					case cmdidAddToOutput: // 171
					case cmdidGenerateChangeScript: // 173
					case cmdidRunQuery: // 201
					case cmdidClearQuery: // 202
					case cmdidPropertyPages: // 232
					case cmdidInsertValuesQuery: // 309
						pCmd->cmdf = 0;
						break;

					case cmdidOpenWith: // 199
					case cmdidViewForm: // 332
					case cmdidExceptions: // 339
					case cmdidViewCode: // 333
					case cmdidPreviewInBrowser: // 334
					case cmdidBrowseWith: // 336
					case cmdidPropSheetOrProperties: // 397
					case cmdidSolutionCfg:        // 684 - if we say "not supported" here, it seems VS will look in the project afterwards
					case cmdidSolutionCfgGetList: // 685 - same
						pCmd->cmdf = 0;
						break;

					default:
						pCmd->cmdf = 0;
				}
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet2K)
		{
			switch (pCmd->cmdID)
			{
				case ECMD_SLNREFRESH: // 222 - button in Solution Explorer
					pCmd->cmdf = OLECMDF_SUPPORTED;
					break;

				default:
					pCmd->cmdf = 0;
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet10)
		{
			switch (pCmd->cmdID)
			{
				case cmdidShellNavigate1First: // 1000
				case cmdidExtensionManager: // 3000
					pCmd->cmdf = 0;
					break;

				default:
					pCmd->cmdf = 0;
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet11)
		{
			switch (pCmd->cmdID)
			{
				case cmdidStartupProjectProperties: // 21
					pCmd->cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
					break;

				default:
					pCmd->cmdf = 0;
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet12)
			return OLECMDERR_E_UNKNOWNGROUP;

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet17)
			return OLECMDERR_E_UNKNOWNGROUP;

		if (*pguidCmdGroup == guidVSDebugCommand)
			// These are in VsDbgCmd.h
			return OLECMDERR_E_UNKNOWNGROUP;

		if (*pguidCmdGroup == CLSID_FelixPackageCmdSet)
			return OLECMDERR_E_NOTSUPPORTED;

		return OLECMDERR_E_NOTSUPPORTED;
	}

	virtual HRESULT STDMETHODCALLTYPE ExecCommand (IProjectNode* proj, const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
	{
		HRESULT hr;

		if (!pguidCmdGroup)
			return E_POINTER;

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97)
		{
			//	|| 
			//	|| nCmdID == cmdidSave   // ignore it here and VS will - if dirty - pass it to the project's IVsPersistHierarchyItem::SaveItem
			//	|| nCmdID == cmdidSaveAs // same
			//)
			//	return OLECMDERR_E_NOTSUPPORTED;

			if (nCmdID == cmdidOpen) // 261
			{
				wil::com_ptr_nothrow<IVsWindowFrame> windowFrame;
				hr = proj->AsVsProject()->OpenItem (_itemId, LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &windowFrame); RETURN_IF_FAILED_EXPECTED(hr);
				hr = windowFrame->Show(); RETURN_IF_FAILED_EXPECTED(hr);
				return S_OK;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == GUID_VsUIHierarchyWindowCmds)
		{
			if (nCmdID == UIHWCMDID_DoubleClick || nCmdID == UIHWCMDID_EnterKey)
			{
				wil::com_ptr_nothrow<IVsWindowFrame> windowFrame;
				hr = proj->AsVsProject()->OpenItem (_itemId, LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &windowFrame); RETURN_IF_FAILED_EXPECTED(hr);
				hr = windowFrame->Show();
				if (FAILED(hr))
					return hr;

				return S_OK;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		return OLECMDERR_E_UNKNOWNGROUP;
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
		wil::unique_process_heap_string mk;
		auto hr = GetPathOf (proj, this, mk, false); RETURN_IF_FAILED(hr);
		auto bstr = SysAllocString(mk.get()); RETURN_IF_NULL_ALLOC(bstr);
		*pbstrMkDocument = bstr;
		return S_OK;
	}
	#pragma endregion

	#pragma region IFileNode
	virtual HRESULT STDMETHODCALLTYPE GetPath (BSTR* pbstrPath) override
	{
		return GetBSTR(_post ? postincludeFilename.get() : preincludeFilename.get(), pbstrPath);
	}
	#pragma endregion
};

// ============================================================================

HRESULT MakeFileNodePrePostInc (bool post, IFileNode** file)
{
	auto p = com_ptr(new (std::nothrow) FileNodePrePostInc()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(post); RETURN_IF_FAILED(hr);
	*file = p.detach();
	return S_OK;
}
