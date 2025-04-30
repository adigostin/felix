
#include "pch.h"
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
#include "../FelixPackageUi/resource.h"
#include "dispids.h"
#include "guids.h"
#include <vsmanaged.h>
#include <KnownImageIds.h>
#include <variant>

using namespace Microsoft::VisualStudio::Imaging;

struct FileNode 
	: IFileNode
	, IFileNodeProperties
	, IVsPerPropertyBrowsing
	, IConnectionPointContainer
	, IPropertyNotifySink
{
	ULONG _refCount = 0;
	VSITEMID _itemId = VSITEMID_NIL;
	wil::com_ptr_nothrow<IChildNode> _next;
	com_ptr<IWeakRef> _parent;
	VSCOOKIE _docCookie = VSDOCCOOKIE_NIL;
	wil::unique_process_heap_string _path;
	BuildToolKind _buildTool = BuildToolKind::None;
	com_ptr<ICustomBuildToolProperties> _customBuildToolProps;
	com_ptr<ConnectionPointImpl<IID_IPropertyNotifySink>> _propNotifyCP;
	AdviseSinkToken _cbtPropNotifyToken;
	WeakRefToThis _weakRefToThis;

public:
	HRESULT InitInstance()
	{
		HRESULT hr;
		hr = _weakRefToThis.InitInstance(static_cast<IFileNode*>(this)); RETURN_IF_FAILED(hr);
		hr = ConnectionPointImpl<IID_IPropertyNotifySink>::CreateInstance(this, &_propNotifyCP); RETURN_IF_FAILED(hr);
		hr = MakeCustomBuildToolProperties(&_customBuildToolProps); RETURN_IF_FAILED(hr);
		hr = AdviseSink<IPropertyNotifySink>(_customBuildToolProps, _weakRefToThis, &_cbtPropNotifyToken); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	FileNode()
	{
	}

	~FileNode()
	{
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IFileNode>(this, riid, ppvObject)
			|| TryQI<IChildNode>(this, riid, ppvObject)
			|| TryQI<IFileNodeProperties>(this, riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

		#ifdef _DEBUG
		if (   riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IManagedObject // .Net stuff
			|| riid == IID_IInspectable // WinRT stuff
			|| riid == IID_IExtendedObject // not in MPF either
			|| riid == IID_IConvertible
			|| riid == IID_ICustomTypeDescriptor
			|| riid == IID_IComponent
			|| riid == IID_IPerPropertyBrowsing
			|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_ISpecifyPropertyPages
			|| riid == IID_ISupportErrorInfo
			|| riid == IID_IVsAggregatableProject // will never support
			|| riid == IID_IVsBuildPropertyStorage // something about MSBuild
			|| riid == IID_IVsProject // will never support
			|| riid == IID_IVsSolution // will never support
			|| riid == IID_IVsHasRelatedSaveItems // will never support for asm file
			|| riid == IID_IVsSupportItemHandoff
			|| riid == IID_IVsFilterAddProjectItemDlg
			|| riid == IID_IVsHierarchy
			|| riid == IID_IUseImmediateCommitPropertyPages
			|| riid == IID_IVsParentProject3
			|| riid == IID_IVsUIHierarchyEventsPrivate
			|| riid == IID_INoIdea3
			|| riid == IID_INoIdea4
			|| riid == IID_SolutionProperties
		)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IFileNodeProperties);

	#pragma region IChildNode
	virtual VSITEMID STDMETHODCALLTYPE GetItemId() override { return _itemId; }

	virtual HRESULT GetParent (IParentNode** ppParent) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !_parent);
		return _parent->QueryInterface(IID_PPV_ARGS(ppParent));
	}

	virtual HRESULT STDMETHODCALLTYPE SetItemId (IParentNode* parent, VSITEMID id) override
	{
		RETURN_HR_IF(E_INVALIDARG, id == VSITEMID_NIL);
		RETURN_HR_IF(E_INVALIDARG, !parent);
		RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL);
		RETURN_HR_IF(E_UNEXPECTED, _parent);

		com_ptr<IWeakRef> p;
		auto hr = parent->QueryInterface(IID_PPV_ARGS(p.addressof())); RETURN_IF_FAILED(hr);

		_parent = std::move(p);
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

	virtual IChildNode* STDMETHODCALLTYPE Next() override { return _next.get(); }

	virtual void STDMETHODCALLTYPE SetNext (IChildNode* next) override
	{
		//WI_ASSERT (!_next);
		_next = next;
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty (VSHPROPID propid, VARIANT* pvar) override
	{
		HRESULT hr;

		RETURN_HR_IF(E_UNEXPECTED, !_parent); // callable only while in a hierarchy
		/*
		OutputDebugStringA("ProjectFile::GetProperty propid=");
		char buffer[20];
		sprintf_s(buffer, "%d", propid);
		OutputDebugStringA(buffer);
		if (auto pn = PropIDToString(propid))
		{
			OutputDebugStringA(" (");
			OutputDebugStringA(pn);
			OutputDebugStringA(")");
		}
		OutputDebugStringA(".\r\n");
		*/
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
			case VSHPROPID_EditLabel: // -2026
				if (!_path || !_path.get()[0])
					return E_NOT_SET;
				return InitVariantFromString (PathFindFileName(_path.get()), pvar);

			case VSHPROPID_Expandable: // -2006
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_BrowseObject: // -2018
				return InitVariantFromDispatch (this, pvar);

			case VSHPROPID_ItemDocCookie: // -2034
				return InitVariantFromInt32 (_docCookie, pvar);

			case VSHPROPID_FirstVisibleChild: // -2041
				return InitVariantFromInt32 (VSITEMID_NIL, pvar);

			case VSHPROPID_NextVisibleSibling: // -2042
				return InitVariantFromInt32 (_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

			case VSHPROPID_IsHiddenItem: // -2043
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_IsNonMemberItem: // -2044
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_IsNonSearchable: // -2051
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_HasEnumerationSideEffects: // -2062
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_ExternalItem: // -2103
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_DescriptiveName: // -2108
			{
				// Tooltip when hovering the document tab with the mouse, maybe other things too.
				com_ptr<IVsHierarchy> hier;
				auto hr = FindHier(this, IID_PPV_ARGS(hier.addressof())); RETURN_IF_FAILED(hr);
				wil::unique_bstr path;
				hr = GetMkDocument(hier, &path); RETURN_IF_FAILED(hr);
				pvar->vt = VT_BSTR;
				pvar->bstrVal = path.release();
				return S_OK;
			}

			case VSHPROPID_ProvisionalViewingStatus: // -2112
				return InitVariantFromUInt32 (PVS_Disabled, pvar);

			case VSHPROPID_SupportsIconMonikers: // -2159
				return InitVariantFromBoolean (TRUE, pvar);

			case VSHPROPID_IconMonikerId: // -2161
			{
				auto ext = PathFindExtension(_path.get());
				
				if (!_wcsicmp(ext, L".asm"))
					return InitVariantFromInt32(KnownImageIds::ASMFile, pvar);

				if (!_wcsicmp(ext, L".inc"))
					return InitVariantFromInt32(KnownImageIds::TextFile, pvar);

				return E_NOTIMPL;
			}

			case VSHPROPID_OverlayIconIndex: // -2048
				// Since this code executes only while this file node in a hierarchy, a filename means the file
				// is under the project dir. Any directory component means the file is a link outside the project dir.
				return InitVariantFromUInt32(PathIsFileSpec(_path.get()) ? OVERLAYICON_NONE : OVERLAYICON_SHORTCUT, pvar);

			case VSHPROPID_IconIndex: // -2005
			case VSHPROPID_IconHandle: // -2013
			case VSHPROPID_OpenFolderIconHandle: // -2014
			case VSHPROPID_OpenFolderIconIndex: // -2015
			case VSHPROPID_AltHierarchy: // -2019
			case VSHPROPID_UserContext: // -2023
			case VSHPROPID_StateIconIndex: // -2029
			case VSHPROPID_IsNewUnsavedItem: // -2057,
			case VSHPROPID_ShowOnlyItemCaption: // -2058
			case VSHPROPID_KeepAliveDocument: // -2075
			case VSHPROPID_ProjectTreeCapabilities: // -2146
			case VSHPROPID_IsSharedItemsImportFile: // -2154
				return E_NOTIMPL;

			default:
				#ifdef _DEBUG
				RETURN_HR(E_NOTIMPL);
				#else
				return E_NOTIMPL;
				#endif
		}
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty (VSHPROPID propid, const VARIANT& var) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !_parent); // callable only while in a hierarchy

		switch(propid)
		{
			case VSHPROPID_Parent:
				RETURN_HR(E_UNEXPECTED); // Set this via SetItemId

			case VSHPROPID_EditLabel: // -2026
				RETURN_HR_IF(E_INVALIDARG, var.vt != VT_BSTR);
				return RenameFile (var.bstrVal);

			case VSHPROPID_ItemDocCookie: // -2034
				RETURN_HR_IF(E_INVALIDARG, var.vt != VT_VSCOOKIE);
				_docCookie = var.lVal;// V_VSCOOKIE(&var);
				return S_OK;

			default:
				#ifdef _DEBUG
				RETURN_HR(E_NOTIMPL);
				#else
				return E_NOTIMPL;
				#endif
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

		#ifdef _DEBUG
		if (propid == VSHPROPID_OpenFolderIconMonikerGuid) // -2162
			return E_NOTIMPL;

		RETURN_HR(E_NOTIMPL);
		#else
		return E_NOTIMPL;
		#endif
	}

	virtual HRESULT STDMETHODCALLTYPE SetGuidProperty (VSHPROPID propid, REFGUID rguid) override
	{
		#ifdef _DEBUG
		RETURN_HR(E_NOTIMPL);
		#else
		return E_NOTIMPL;
		#endif
	}

	virtual HRESULT STDMETHODCALLTYPE GetCanonicalName (BSTR* pbstrName) override
	{
		// Returns a unique, string name for an item in the hierarchy.
		// Used for workspace persistence, such as remembering window positions.
		return get_Path(pbstrName);
	}

	virtual HRESULT STDMETHODCALLTYPE IsItemDirty (IUnknown *punkDocData, BOOL *pfDirty) override
	{
		wil::com_ptr_nothrow<IVsPersistDocData> docData;
		auto hr = punkDocData->QueryInterface(&docData); RETURN_IF_FAILED(hr);
		hr = docData->IsDocDataDirty(pfDirty); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryStatusCommand (const GUID *pguidCmdGroup, OLECMD* pCmd, OLECMDTEXT *pCmdText) override
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

		#ifdef _DEBUG
		if (   *pguidCmdGroup == guidCommonIDEPackage
			|| *pguidCmdGroup == guidCmdSetTFS
			|| *pguidCmdGroup == DebugTargetTypeCommandGuid
			|| *pguidCmdGroup == GUID{ 0x25FD982B, 0x8CAE, 0x4CBD, { 0xA4, 0x40, 0xE0, 0x3F, 0xFC, 0xCD, 0xE1, 0x06 } } // NuGet
			|| *pguidCmdGroup == GUID{ 0x25113E5B, 0x9964, 0x4375, { 0x9D, 0xD1, 0x0A, 0x5E, 0x98, 0x40, 0x50, 0x7A } } // no idea
			|| *pguidCmdGroup == GUID{ 0x64A2F652, 0x0B0F, 0x4413, { 0x90, 0xC8, 0x3E, 0x6A, 0x25, 0x77, 0x52, 0x90 } }
			|| *pguidCmdGroup == GUID{ 0x319FF225, 0x2E9C, 0x4FA7, { 0x92, 0x67, 0x12, 0x8F, 0xE4, 0x24, 0x6B, 0xF1 } } // Web Tools something
			|| *pguidCmdGroup == GUID{ 0xB101F7CB, 0x4BB9, 0x46D0, { 0xA4, 0x89, 0x83, 0x0D, 0x45, 0x01, 0x16, 0x0A } } // something from Microsoft.VisualStudio.ProjectSystem.VS.Implementation.dll
			|| *pguidCmdGroup == GUID{ 0x665CC136, 0x6455, 0x491D, { 0xAB, 0x17, 0xEA, 0xF3, 0x84, 0x7A, 0x23, 0xBC } } // same
			|| *pguidCmdGroup == GUID{ 0x7F917E79, 0x7A75, 0x4DCA, { 0xAE, 0x0A, 0xE5, 0x7A, 0x4A, 0x1E, 0xC5, 0xCE } } // no idea
			|| *pguidCmdGroup == GUID{ 0xB39FEC88, 0x42C6, 0x4662, { 0x9F, 0x18, 0xFA, 0x56, 0xA4, 0x84, 0xB7, 0x65 } } // no idea
			|| *pguidCmdGroup == GUID{ 0xBDFA79D2, 0x2CD2, 0x474A, { 0xA8, 0x2A, 0xCE, 0x86, 0x94, 0x11, 0x68, 0x25 } } // no idea
			|| *pguidCmdGroup == GUID{ 0x9FAD7D7C, 0x0155, 0x46DE, { 0x93, 0x08, 0xBB, 0x3D, 0x97, 0xD7, 0xD3, 0xDE } } // no idea
			|| *pguidCmdGroup == GUID{ 0xD128E8B1, 0x03FA, 0x4495, { 0x8A, 0x97, 0xD6, 0xD2, 0x09, 0x66, 0xCE, 0x4C } } // no idea
			|| *pguidCmdGroup == GUID{ 0x001BD6E5, 0xE2CD, 0x4F47, { 0x98, 0xD4, 0x2D, 0x39, 0x21, 0x53, 0x59, 0xFC } } // no idea
			|| *pguidCmdGroup == GUID{ 0x6C1555D3, 0xB9B7, 0x4D39, { 0xB6, 0x57, 0x1A, 0x35, 0xA0, 0xF3, 0xC4, 0x61 } } // no idea
			|| *pguidCmdGroup == guidXamlCmdSet
			|| *pguidCmdGroup == guidVenusCmdId
			|| *pguidCmdGroup == guidSqlPkg
			|| *pguidCmdGroup == guidSourceExplorerPackage
			|| *pguidCmdGroup == guidWebProjectPackage
			|| *pguidCmdGroup == guidBrowserLinkCmdSet
			|| *pguidCmdGroup == guidUnknownMsenvDll
			|| *pguidCmdGroup == guidUnknownCmdGroup0
			|| *pguidCmdGroup == guidProjectandSolutionContextMenus
			|| *pguidCmdGroup == guidProjOverviewAppCapabilities
			|| *pguidCmdGroup == guidProjectAddTest
			|| *pguidCmdGroup == guidProjectAddWPF
			|| *pguidCmdGroup == guidUnknownCmdGroup5
			|| *pguidCmdGroup == guidUnknownCmdGroup6
			|| *pguidCmdGroup == guidUnknownCmdGroup7
			|| *pguidCmdGroup == guidUnknownCmdGroup8
			|| *pguidCmdGroup == CMDSETID_HtmEdGrp
			//|| *pguidCmdGroup == CMDSETID_WebForms
			|| *pguidCmdGroup == guidXmlPkg
			|| *pguidCmdGroup == guidSccProviderPackage
			|| *pguidCmdGroup == guidSccPkg
			|| *pguidCmdGroup == guidTrackProjectRetargetingCmdSet
			|| *pguidCmdGroup == guidUniversalProjectsCmdSet
			|| *pguidCmdGroup == guidProjectClassWizard
			|| *pguidCmdGroup == guidDataSources
			|| *pguidCmdGroup == guidSHLMainMenu
			|| *pguidCmdGroup == tfsCmdSet
			|| *pguidCmdGroup == tfsCmdSet1
			|| *pguidCmdGroup == WebAppCmdId
			|| *pguidCmdGroup == guidCmdGroupTestExplorer
			|| *pguidCmdGroup == guidCmdSetTaskRunnerExplorer
			|| *pguidCmdGroup == guidCmdSetBowerPackages
			|| *pguidCmdGroup == guidManagedProjectSystemOrderCommandSet
			|| *pguidCmdGroup == XsdDesignerPackage
			|| *pguidCmdGroup == guidCmdSetEdit
			|| *pguidCmdGroup == guidCmdGroupDatabase
			|| *pguidCmdGroup == guidCmdGroupTableQueryDesigner
			|| *pguidCmdGroup == guidVSEQTPackageCmdSet
			|| *pguidCmdGroup == guidSomethingResourcesCmdSet
			|| *pguidCmdGroup == guidCmdSetPerformance
			|| *pguidCmdGroup == guidCmdSetHelp
			|| *pguidCmdGroup == guidCmdSetCtxMenuPrjSolOtherClsWiz
			|| *pguidCmdGroup == guidCmdSetDiff
			|| *pguidCmdGroup == guidToolWindowTimestampButton
			|| *pguidCmdGroup == guidCrossProjectMultiItem
			|| *pguidCmdGroup == guidDataCmdId
			|| *pguidCmdGroup == guidCmdSetSomethingIntellisense
			|| *pguidCmdGroup == guidCmdSetSomethingAzure
			|| *pguidCmdGroup == guidCmdSetSomethingTable
			|| *pguidCmdGroup == guidCmdSetSomethingJson
			|| *pguidCmdGroup == guidCmdSetSomethingRefactor
			|| *pguidCmdGroup == guidCSharpGrpId
			|| *pguidCmdGroup == guidCmdSetCallHierarchy
			|| *pguidCmdGroup == guidCmdSetSomethingTerminal
			|| *pguidCmdGroup == CLSID_VsTaskListPackage
			|| *pguidCmdGroup == guidPropertyManagerPackage
		)
			return OLECMDERR_E_NOTSUPPORTED;
		
		return OLECMDERR_E_NOTSUPPORTED; // RETURN_HR(OLECMDERR_E_NOTSUPPORTED);
		#else
		return OLECMDERR_E_NOTSUPPORTED;
		#endif
	}

	virtual HRESULT STDMETHODCALLTYPE ExecCommand (const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
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
				com_ptr<IVsProject2> vsp;
				hr = FindHier(this, IID_PPV_ARGS(vsp.addressof())); RETURN_IF_FAILED(hr);
				wil::com_ptr_nothrow<IVsWindowFrame> windowFrame;
				hr = vsp->OpenItem (_itemId, LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &windowFrame); RETURN_IF_FAILED_EXPECTED(hr);
				hr = windowFrame->Show(); RETURN_IF_FAILED_EXPECTED(hr);
				return S_OK;
			}

			#ifdef _DEBUG
			if (   nCmdID == cmdidObjectVerbList0       // 137
				|| nCmdID == cmdidCloseSolution         // 219
				|| nCmdID == cmdidSaveSolution          // 224
				|| nCmdID == cmdidSaveProjectItemAs     // 226
				|| nCmdID == cmdidExit                  // 229
				|| nCmdID == cmdidNewFolder             // 245
				|| nCmdID == cmdidPaneActivateDocWindow // 289
				|| nCmdID == cmdidSaveProjectItem       // 331 - Ctrl + S on file in Solution Explorer, ignore it here and VS will somehow save the file
				|| nCmdID == cmdidPropSheetOrProperties // 397 - ignore it here and VS will ask for cmdidPropertyPages, then show the Properties Window
				|| nCmdID == cmdidSolutionCfg           // 684
				|| nCmdID == cmdidSolutionCfgGetList    // 685
			)
				return OLECMDERR_E_NOTSUPPORTED;

			RETURN_HR(OLECMDERR_E_NOTSUPPORTED);
			#else
			return OLECMDERR_E_NOTSUPPORTED;
			#endif
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet2K)
		{
			// https://docs.microsoft.com/en-us/previous-versions/visualstudio/dn604846(v=vs.111)
			if (   nCmdID == cmdidSolutionPlatform // 1990
				|| nCmdID == cmdidSolutionPlatformGetList // 1991
			)
			{
				return OLECMDERR_E_NOTSUPPORTED;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet11)
		{
			if (nCmdID == cmdidLocateFindTarget)
				return OLECMDERR_E_NOTSUPPORTED;

			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == GUID_VsUIHierarchyWindowCmds)
		{
			if (nCmdID == UIHWCMDID_DoubleClick || nCmdID == UIHWCMDID_EnterKey)
			{
				wil::com_ptr_nothrow<IVsProject2> vsp;
				hr = FindHier(this, IID_PPV_ARGS(vsp.addressof())); RETURN_IF_FAILED(hr);
				wil::com_ptr_nothrow<IVsWindowFrame> windowFrame;
				hr = vsp->OpenItem (_itemId, LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &windowFrame); RETURN_IF_FAILED_EXPECTED(hr);
				hr = windowFrame->Show();
				if (FAILED(hr))
					return hr;

				return S_OK;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		return OLECMDERR_E_UNKNOWNGROUP;
	}
	#pragma endregion

	#pragma region IFileNode
	virtual HRESULT STDMETHODCALLTYPE GetMkDocument (IVsHierarchy* hier, BSTR* pbstrMkDocument) override
	{
		HRESULT hr;

		if (PathIsFileSpec(_path.get()))
		{
			// (1)
			wil::unique_process_heap_string mk;
			hr = GetPathOf (this, mk, false); RETURN_IF_FAILED(hr);
			auto bstr = SysAllocString(mk.get()); RETURN_IF_NULL_ALLOC(bstr);
			*pbstrMkDocument = bstr;
			return S_OK;
		}
		else if (!wcsncmp(_path.get(), L"..\\", 3))
		{
			// (2)
			wil::unique_variant projectDir;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_UNEXPECTED, projectDir.vt != VT_BSTR);
			auto mk = wil::make_process_heap_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(mk);
			auto res = PathCombine(mk.get(), projectDir.bstrVal, _path.get()); RETURN_HR_IF_NULL(HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME), res);
			auto bstr = SysAllocString(mk.get()); RETURN_IF_NULL_ALLOC(bstr);
			*pbstrMkDocument = bstr;
			return S_OK;
		}
		else
		{
			// (3)
			auto bstr = SysAllocString(_path.get()); RETURN_IF_NULL_ALLOC(bstr);
			*pbstrMkDocument = bstr;
			return S_OK;
		}
	}
	#pragma endregion

	#pragma region IFileNodeProperties
	virtual HRESULT STDMETHODCALLTYPE get_Path (BSTR *pbstrPath) override
	{
		auto str = SysAllocString(_path.get()); RETURN_IF_NULL_ALLOC(str);
		*pbstrPath = str;
		return S_OK;
	}

	// The value of the Path property can have one of these formats:
	// 
	//  (1) If the file is located somewhere under the project dir, Path contains the file name and extension.
	//      Condition for this case: PathIsFileSpec returns TRUE.
	// 
	//  (2) If the file is located on the same drive as the project dir, but outside the project dir,
	//      Path begins with "..\", followed by zero or more directories, followed by file name and extension,
	//      Condition for this case: PathIsFileSpec returns FALSE, and the path begins with "..\".
	// 
	//  (3) If the file is located on a different drive than the project dir, Path contains a rooted path,
	//      starting for example with a drive number or a server share location.
	//      Condition for this case: PathIsFileSpec returns FALSE, and PathIsRelative returns FALSE.
	//
	// If none of the above conditions is satisfied, the path is incorrect.
	virtual HRESULT STDMETHODCALLTYPE put_Path (BSTR value) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL); // This function may not be called while in a hierarchy.

		RETURN_HR_IF(E_INVALIDARG, value == nullptr);

		// Verification
		if (PathIsFileSpec(value))
		{
			// Case (1)
		}
		else if (wcsncmp(value, L"..\\", 3) == 0)
		{
			// Case (2)
		}
		else if (!PathIsRelative(value))
		{
			// Case (3)
		}
		else
			RETURN_HR(E_INVALIDARG);

		size_t len = wcslen(value);
		auto newValue = wil::make_process_heap_string_nothrow(value, len); RETURN_IF_NULL_ALLOC(newValue);
		for (auto p = wcschr(newValue.get(), '/'); p; p = wcschr(p, '/'))
			*p = '\\';
		_path = std::move(newValue);
		// No need for notification since this is called only when loading from XML.
		// (Property is read-only in the Properties window - see IsPropertyReadOnly().)
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		// Shown by VS at the top of the Properties Window.
		auto name = PathFindFileName(_path.get());
		*value = SysAllocString(name); RETURN_IF_NULL_ALLOC(*value);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_BuildTool (enum BuildToolKind *value) override
	{
		*value = _buildTool;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_BuildTool (enum BuildToolKind value) override
	{
		if (_buildTool != value)
		{
			_buildTool = value;
			if (_parent)
			{
				com_ptr<IPropertyNotifySink> sink;
				auto hr = FindHier(this, IID_PPV_ARGS(sink.addressof())); RETURN_IF_FAILED(hr);
				sink->OnChanged(DISPID_UNKNOWN);
				_propNotifyCP->NotifyPropertyChanged(dispidBuildToolKind);
				_propNotifyCP->NotifyPropertyChanged(dispidCustomBuildToolProps);
			}
		}

		return S_OK;
	}

	virtual HRESULT get_CustomBuildToolProperties (ICustomBuildToolProperties** ppProps)
	{
		return wil::com_query_to_nothrow(_customBuildToolProps, ppProps);
	}
	#pragma endregion

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL *pfHide) override
	{
		if (dispid == dispidCustomBuildToolProps)
			return *pfHide = (_buildTool != BuildToolKind::CustomBuildTool), S_OK;

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override
	{
		if (dispid == dispidCustomBuildToolProps)
			return *pfDisplay = (_buildTool == BuildToolKind::CustomBuildTool), S_OK;

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidBuildToolKind || dispid == dispidPath)
		{
			*fDefault = FALSE;
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override
	{
		if (dispid == dispidPath || dispid == dispidCustomBuildToolProps)
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
		*pbstrClassName = SysAllocString(L"File");
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

	#pragma region IPropertyNotifySink
	virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
	{
		com_ptr<IPropertyNotifySink> sink;
		auto hr = FindHier(this, IID_PPV_ARGS(sink.addressof())); RETURN_IF_FAILED(hr);
		sink->OnChanged(DISPID_UNKNOWN);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	HRESULT RenameFile (BSTR newName)
	{
		HRESULT hr;

		RETURN_HR_IF(E_UNEXPECTED, _itemId == VSITEMID_NIL);
		
		RETURN_HR_IF(E_NOTIMPL, !PathIsFileSpec(_path.get()));

		com_ptr<IVsHierarchy> hier;
		hr = FindHier(this, IID_PPV_ARGS(hier.addressof())); RETURN_IF_FAILED(hr);
		com_ptr<IVsProject> project;
		hr = hier->QueryInterface(&project); RETURN_IF_FAILED(hr);

		wil::unique_process_heap_string pathToThis;
		hr = GetPathTo (this, pathToThis); RETURN_IF_FAILED(hr);

		wil::unique_process_heap_string oldFullPath;
		hr = wil::str_concat_nothrow(oldFullPath, pathToThis, L"\\", _path); RETURN_IF_FAILED(hr);

		hr = QueryEditProjectFile(hier); RETURN_IF_FAILED_EXPECTED(hr);

		// Check if the document is in the cache and rename document in the cache.
		com_ptr<IVsRunningDocumentTable> pRDT;
		hr = serviceProvider->QueryService(SID_SVsRunningDocumentTable, &pRDT);
		VSCOOKIE dwCookie = VSCOOKIE_NIL;
		if (SUCCEEDED(hr))
		{
			pRDT->FindAndLockDocument (RDT_NoLock, oldFullPath.get(), nullptr, nullptr, nullptr, &dwCookie); // ignore returned HRESULT as we're only interested in dwCookie
		}

		// TODO: repro and test this
		// if document is open and we are not owner of document, do not rename it
		//if (poHier != NULL)
		//{
		//	CComPtr<IVsHierarchy> pMyHier = (GetCVsHierarchy()->GetIVsHierarchy());
		//	CComPtr<IUnknown> punkRDTHier;
		//	CComPtr<IUnknown> punkMyHier;
		//	pMyHier->QueryInterface(IID_IUnknown, (void **)&punkMyHier);
		//	poHier->QueryInterface(IID_IUnknown, (void **)&punkRDTHier);
		//	if (punkRDTHier != punkMyHier)
		//		return S_OK;
		//}

		wil::unique_process_heap_string newFullPath;
		hr = wil::str_concat_nothrow(newFullPath, pathToThis, L"\\", newName); RETURN_IF_FAILED(hr);

		BOOL fRenameCanContinue = TRUE;
		com_ptr<IVsTrackProjectDocuments2> trackProjectDocs;
		hr = serviceProvider->QueryService (SID_SVsTrackProjectDocuments, &trackProjectDocs);
		if (SUCCEEDED(hr))
		{
			fRenameCanContinue = FALSE;
			hr = trackProjectDocs->OnQueryRenameFile (project, oldFullPath.get(), newFullPath.get(), VSRENAMEFILEFLAGS_NoFlags, &fRenameCanContinue);
			if (FAILED(hr) || !fRenameCanContinue)
				return OLE_E_PROMPTSAVECANCELLED;
		}

		auto otherName = wil::make_process_heap_string_nothrow(newName);

		if (!::MoveFile (oldFullPath.get(), newFullPath.get()))
			return HRESULT_FROM_WIN32(GetLastError());
		std::swap(_path, otherName);
		auto undoRename = wil::scope_exit([this, &otherName, &newFullPath, &oldFullPath]
			{
				std::swap (_path, otherName);
				::MoveFile (newFullPath.get(), oldFullPath.get());
			});


		if (dwCookie != VSCOOKIE_NIL)
		{
			hr = pRDT->RenameDocument (oldFullPath.get(), newFullPath.get(), HIERARCHY_DONTCHANGE, VSITEMID_NIL);
			if (FAILED(hr))
			{
				// We could rename the file on disk, but not in the Running Document Table.
				// Do we want the rename operation to fail? Not really. So no error checking.
			}
		}

		// TODO: ignore file change notifications
		//CSuspendFileChanges suspendFileChanges (strNewName, TRUE );

		// Tell packages that care that it happened. No error checking here.
		if (trackProjectDocs)
			trackProjectDocs->OnAfterRenameFile (project, oldFullPath.get(), newFullPath.get(), VSRENAMEFILEFLAGS_NoFlags);

		// This line was changing the build tool when the user renamed the file.
		// I eventually commented it out, to get behavior similar to that in VS projects
		// (build tool remains unchanged when renaming for example from .cpp to .h)
		//_buildTool = _wcsicmp(PathFindExtension(_pathRelativeToProjectDir.get()), L".asm") ? BuildToolKind::None : BuildToolKind::Assembler;

		com_ptr<IEnumHierarchyEvents> enu;
		hr = hier.try_query<IProjectNode>()->EnumHierarchyEventSinks(&enu); RETURN_IF_FAILED(hr);
		com_ptr<IVsHierarchyEvents> sink;
		ULONG fetched;
		while (SUCCEEDED(enu->Next(1, &sink, &fetched)) && fetched)
		{
			sink->OnPropertyChanged(_itemId, VSHPROPID_Caption, 0);
			sink->OnPropertyChanged(_itemId, VSHPROPID_Name, 0);
			sink->OnPropertyChanged(_itemId, VSHPROPID_SaveName, 0);
			sink->OnPropertyChanged(_itemId, VSHPROPID_DescriptiveName, 0);
			sink->OnPropertyChanged(_itemId, VSHPROPID_StateIconIndex, 0);
			sink->OnPropertyChanged(_itemId, VSHPROPID_IconMonikerId, 0);
		}

		// Make sure the property browser is updated.
		com_ptr<IVsUIShell> uiShell;
		hr = serviceProvider->QueryService (SID_SVsUIShell, &uiShell);
		if (SUCCEEDED(hr))
			uiShell->RefreshPropertyBrowser(DISPID_UNKNOWN); // refresh all properties

		// Mark project as dirty.
		com_ptr<IPropertyNotifySink> pns;
		hr = hier->QueryInterface(&pns); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
			pns->OnChanged(dispidItems);

		//~CSuspendFileChanges

		undoRename.release();
		return S_OK;
	}
};

// ============================================================================

FELIX_API HRESULT MakeFileNode (IFileNode** file)
{
	com_ptr<FileNode> p = new (std::nothrow) FileNode(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(); RETURN_IF_FAILED(hr);
	*file = p.detach();
	return S_OK;
}
