
#include "pch.h"
#include <ivstrackprojectdocuments2.h>
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
#include "../FelixPackageUi/resource.h"
#include "dispids.h"
#include "guids.h"

struct ProjectFile 
	: IProjectFile
	, IProvideClassInfo
	, IOleCommandTarget
	, IVsPerPropertyBrowsing
{
	VSITEMID _itemId;
	ULONG _refCount = 0;
	com_ptr<IVsUIHierarchy> _hier;
	wil::com_ptr_nothrow<IProjectItem> _next;
	VSITEMID _parentItemId = VSITEMID_NIL;
	VSCOOKIE _docCookie = VSDOCCOOKIE_NIL;
	wil::unique_hlocal_string _pathRelativeToProjectDir;
	BuildToolKind _buildTool = BuildToolKind::None;
	static inline HINSTANCE _uiLibrary;
	static inline wil::unique_hicon _iconAsmFile;
	static inline wil::unique_hicon _iconIncFile;
	
public:
	HRESULT InitInstance (VSITEMID itemId, IVsUIHierarchy* hier, VSITEMID parentItemId)
	{
		_itemId = itemId;
		_hier = hier;
		_parentItemId = parentItemId;
		if (!_uiLibrary)
		{
			wil::com_ptr_nothrow<IVsShell> shell;
			auto hr = serviceProvider->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);
			hr = shell->LoadUILibrary(CLSID_FelixPackage, 0, (DWORD_PTR*)&_uiLibrary); RETURN_IF_FAILED(hr);
			_iconAsmFile.reset(LoadIcon(_uiLibrary, MAKEINTRESOURCE(IDI_ASM_FILE))); WI_ASSERT(_iconAsmFile);
			_iconIncFile.reset(LoadIcon(_uiLibrary, MAKEINTRESOURCE(IDI_INC_FILE))); WI_ASSERT(_iconAsmFile);
		}
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IProjectFile>(this, riid, ppvObject)
			|| TryQI<IProjectItem>(this, riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProvideClassInfo>(this, riid, ppvObject)
			|| TryQI<IOleCommandTarget>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		if (riid == IID_IPerPropertyBrowsing)
			return E_NOINTERFACE;

		if (   riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IManagedObject // .Net stuff
			|| riid == IID_IInspectable // WinRT stuff
			|| riid == IID_IExtendedObject // not in MPF either
			|| riid == IID_IConvertible
			|| riid == IID_ICustomTypeDescriptor
			//|| riid == IID_IVSMDPerPropertyBrowsing
			|| riid == IID_IComponent
			)
			return E_NOINTERFACE;

		if (   riid == IID_ISpecifyPropertyPages
			|| riid == IID_ISupportErrorInfo
			|| riid == IID_IVsAggregatableProject // will never support
			|| riid == IID_IVsBuildPropertyStorage // something about MSBuild
			|| riid == IID_IVsProject // will never support
			|| riid == IID_IVsSolution // will never support
			|| riid == IID_IVsHasRelatedSaveItems // will never support for asm file
			|| riid == IID_IVsSupportItemHandoff
			|| riid == IID_IVsFilterAddProjectItemDlg
			|| riid == IID_IVsHierarchy
			//|| riid == IID_IUseImmediateCommitPropertyPages
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

	IMPLEMENT_IDISPATCH(IID_IProjectFile);

	#pragma region IProjectItem
	virtual VSITEMID STDMETHODCALLTYPE GetItemId() override { return _itemId; }

	virtual HRESULT STDMETHODCALLTYPE GetMkDocument (BSTR* pbstrMkDocument) override
	{
		wil::unique_variant location;
		auto hr = _hier->GetProperty (VSITEMID_ROOT, VSHPROPID_ProjectDir, &location); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, location.vt != VT_BSTR);

		wil::unique_hlocal_string filePath;
		hr = PathAllocCombine (location.bstrVal, _pathRelativeToProjectDir.get(), 
			PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS, &filePath); RETURN_IF_FAILED(hr);
		*pbstrMkDocument = SysAllocString(filePath.get()); RETURN_IF_NULL_ALLOC(*pbstrMkDocument);
		return S_OK;
	}
	
	virtual IProjectItem* STDMETHODCALLTYPE Next() override { return _next.get(); }

	virtual void STDMETHODCALLTYPE SetNext (IProjectItem* next) override
	{
		//WI_ASSERT (!_next);
		_next = next;
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty (VSHPROPID propid, VARIANT* pvar) override
	{
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
				return InitVariantFromInt32 (_parentItemId, pvar);

			case VSHPROPID_FirstChild: // -1001
				return InitVariantFromInt32 (VSITEMID_NIL, pvar);

			case VSHPROPID_NextSibling: // -1002
				return InitVariantFromInt32 (_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

			case VSHPROPID_SaveName: // -2002
			case VSHPROPID_Caption: // -2003
			case VSHPROPID_Name: // -2012
			case VSHPROPID_EditLabel: // -2026
			case VSHPROPID_DescriptiveName: // -2108
				return InitVariantFromString (PathFindFileName(_pathRelativeToProjectDir.get()), pvar);

			case VSHPROPID_Expandable: // -2006
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_IconHandle: // -2013
			{
				auto ext = PathFindExtension(_pathRelativeToProjectDir.get());

				if (!_wcsicmp(ext, L".asm"))
				{
					V_VT(pvar) = VT_VS_INT_PTR;
					V_VS_INT_PTR(pvar) = (INT_PTR)_iconAsmFile.get();
					return S_OK;
				}

				if (!_wcsicmp(ext, L".inc"))
				{
					V_VT(pvar) = VT_VS_INT_PTR;
					V_VS_INT_PTR(pvar) = (INT_PTR)_iconIncFile.get();
					return S_OK;
				}

				return E_NOTIMPL;
			}

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

			case VSHPROPID_IsNewUnsavedItem: // -2057
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_ShowOnlyItemCaption: // -2058
				return InitVariantFromBoolean (TRUE, pvar);

			case VSHPROPID_HasEnumerationSideEffects: // -2062
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_ExternalItem: // -2103
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_ProvisionalViewingStatus: // -2112
				return InitVariantFromUInt32 (PVS_Disabled, pvar);

			case VSHPROPID_IconIndex: // -2005
			case VSHPROPID_OpenFolderIconHandle: // -2014
			case VSHPROPID_OpenFolderIconIndex: // -2015
			case VSHPROPID_AltHierarchy: // -2019
			case VSHPROPID_ExtObject: // -2027
			case VSHPROPID_StateIconIndex: // -2029
			case VSHPROPID_OverlayIconIndex: // -2048
			case VSHPROPID_KeepAliveDocument: // -2075
			case VSHPROPID_ProjectTreeCapabilities: // -2146
			case VSHPROPID_IsSharedItemsImportFile: // -2154
			case VSHPROPID_SupportsIconMonikers: // -2159
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
		switch(propid)
		{
			case VSHPROPID_Parent:
				RETURN_HR_IF(E_INVALIDARG, var.vt != VT_VSITEMID);
				_parentItemId = V_VSITEMID(&var);
				return S_OK;

			case VSHPROPID_EditLabel	: // -2026
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

		#ifdef _DEBUG
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
		*pbstrName = SysAllocString(_pathRelativeToProjectDir.get());
		return *pbstrName ? S_OK : E_OUTOFMEMORY;
	}

	virtual HRESULT STDMETHODCALLTYPE IsItemDirty (IUnknown *punkDocData, BOOL *pfDirty) override
	{
		wil::com_ptr_nothrow<IVsPersistDocData> docData;
		auto hr = punkDocData->QueryInterface(&docData); RETURN_IF_FAILED(hr);
		hr = docData->IsDocDataDirty(pfDirty); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Close() override
	{
		_hier = nullptr;
		if (_next)
		{
			_next->Close();
			_next = nullptr;
		}
		return S_OK;
	}
	#pragma endregion

	#pragma region IProjectFile
	virtual HRESULT STDMETHODCALLTYPE get_Path (BSTR *pbstr) override
	{
		*pbstr = SysAllocString(_pathRelativeToProjectDir.get()); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Path (BSTR value) override
	{
		auto newValue = wil::make_hlocal_string_nothrow(value, SysStringLen(value)); RETURN_IF_NULL_ALLOC(newValue);
		_pathRelativeToProjectDir = std::move(newValue);
		// TODO: notification
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		// Shown by VS at the top of the Properties Window.
		auto name = PathFindFileName(_pathRelativeToProjectDir.get());
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
			if (auto sink = wil::try_com_query_nothrow<IPropertyNotifySink>(_hier))
				sink->OnChanged(DISPID_UNKNOWN);
			if (auto sink = wil::try_com_query_failfast<IVsHierarchyEvents>(_hier))
				sink->OnPropertyChanged(_itemId, VSHPROPID_IconHandle, 0);
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IProvideClassInfo
	virtual HRESULT STDMETHODCALLTYPE GetClassInfo (ITypeInfo** ppTI) override
	{
		return GetTypeInfo(0, 0, ppTI);
	}
	#pragma endregion

	#pragma region IOleCommandTarget
	virtual HRESULT STDMETHODCALLTYPE QueryStatus (const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText) override
	{
		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97)
		{
			// These are the cmdidXxxYyy constants from stdidcmd.h
			for (ULONG i = 0; i < cCmds; i++)
			{
				if (prgCmds[i].cmdID >= 0xF000)
				{
					// debugger stuff, ignored for now
					prgCmds[i].cmdf = OLECMDF_SUPPORTED;
				}
				else if (prgCmds[i].cmdID >= 946 && prgCmds[i].cmdID <= 957)
				{
					// refactoring stuff
					prgCmds[i].cmdf = 0;
				}
				else if (prgCmds[i].cmdID >= 122 && prgCmds[i].cmdID <= cmdidSelectAllFields)
				{
					// database stuff
					prgCmds[i].cmdf = 0;
				}
				else
				{
					switch (prgCmds[i].cmdID)
					{
						case cmdidCopy: // 15
						case cmdidCut: // 16
						case cmdidDelete: // 17
						case cmdidMultiLevelRedo: // 30
						case cmdidMultiLevelUndo: // 44
						case cmdidNewProject: // 216
						case cmdidAddNewItem: // 220
						case cmdidFileOpen: // 222
						case cmdidSaveSolution: // 224
						case cmdidGoto: // 231
						case cmdidOpen: // 261
						case cmdidFindInFiles: // 277
						case cmdidViewCode: // 333
						case cmdidPropSheetOrProperties: // 397
						case cmdidShellNavBackward: // 809
						case cmdidShellNavForward: // 810
						case cmdidBuildSln: // 882
						case cmdidRebuildSln: // 883
							prgCmds[i].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
							break;

						case cmdidSaveProjectItem: // 331
						case cmdidSaveProjectItemAs: // 226
							// TODO: enable it only if it's in the running document list
							prgCmds[i].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
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
							prgCmds[i].cmdf = 0;
							break;

						case cmdidOpenWith: // 199
						case cmdidViewForm: // 332
						case cmdidExceptions: // 339
						case cmdidPreviewInBrowser: // 334
						case cmdidBrowseWith: // 336
						case cmdidSolutionCfg:        // 684 - if we say "not supported" here, it seems VS will look in the project afterwards
						case cmdidSolutionCfgGetList: // 685 - same
							prgCmds[i].cmdf = 0;
							break;

						default:
							prgCmds[i].cmdf = 0;
					}
				}
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet2K)
		{
			// https://docs.microsoft.com/en-us/previous-versions/visualstudio/dn604846(v=vs.111)
			for (ULONG i = 0; i < cCmds; i++)
			{
				switch (prgCmds[i].cmdID)
				{
					case ECMD_COMPILE: // 350
					case cmdidCopyFullPathName: // 1610
					case cmdidOpenInCodeView: // 1634
					case cmdidBrowseToFileInExplorer: // 1642
					case cmdidSolutionPlatform: // 1990
					case cmdidSolutionPlatformGetList: // 1991
						prgCmds[i].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
						break;

					case ECMD_SLNREFRESH: // 222 - button in Solution Explorer
					case ECMD_INCLUDEINPROJECT: // 1109
					case ECMD_EXCLUDEFROMPROJECT: // 1110
						prgCmds[i].cmdf = OLECMDF_SUPPORTED;
						break;

					case ECMD_SUPERSCRIPT: // 298
					case ECMD_SUBSCRIPT: // 299
					case ECMD_UPDATEMGDRES: // 358 - update managed resources
					case ECMD_ADDRESOURCE: // 362
					case ECMD_ADDHTMLPAGE: // 400
					case ECMD_ADDMODULE: // 402
					case ECMD_ADDWFCFORM: // 406
					case ECMD_ADDWEBFORM: // 410
					case ECMD_ADDMASTERPAGE: // 411
					case ECMD_ADDUSERCONTROL: // 412
					case ECMD_ADDCONTENTPAGE: // 413
					case ECMD_ADDWEBUSERCONTROL: // 438
					case ECMD_ADDTBXCOMPONENT: // 442
					case ECMD_ADDWEBSERVICE: // 444
					case ECMD_ADDSTYLESHEET: // 445
					case ECMD_REFRESHFOLDER: // 447
					case ECMD_VIEWMARKUP: // 449
					case ECMD_VIEWCOMPONENTDESIGNER: // 457
					case ECMD_SHOWALLFILES: // 600
					case ECMD_SETASSTARTPAGE: // 1100
					case ECMD_RUNCUSTOMTOOL: // 1117
					case ECMD_DETACHLOCALDATAFILECTX: // 1128
					case ECMD_NESTRELATEDFILES: // 1401
					case cmdidViewInClassDiagram: // 1931
						prgCmds[i].cmdf = 0;
						break;

					default:
						prgCmds[i].cmdf = 0;
				}
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet10)
		{
			for (ULONG i = 0; i < cCmds; i++)
			{
				switch (prgCmds[i].cmdID)
				{
					case cmdidShellNavigate1First: // 1000
					case cmdidExtensionManager: // 3000
						prgCmds[i].cmdf = 0;
						break;

					default:
						prgCmds[i].cmdf = 0;
				}
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet11)
		{
			for (ULONG i = 0; i < cCmds; i++)
			{
				switch (prgCmds[i].cmdID)
				{
					case cmdidStartupProjectProperties: // 21
						prgCmds[i].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
						break;

					default:
						prgCmds[i].cmdf = 0;
				}
			}

			return S_OK;
		}

		if (*pguidCmdGroup == guidVSDebugCommand)
			// These are in VsDbgCmd.h
			return OLECMDERR_E_UNKNOWNGROUP;

		if (   *pguidCmdGroup == guidCommonIDEPackage
			|| *pguidCmdGroup == guidCmdSetTFS
			|| *pguidCmdGroup == DebugTargetTypeCommandGuid
		)
			return OLECMDERR_E_UNKNOWNGROUP; // not yet there

		if (*pguidCmdGroup == GUID{ 0x25FD982B, 0x8CAE, 0x4CBD, { 0xA4, 0x40, 0xE0, 0x3F, 0xFC, 0xCD, 0xE1, 0x06 } }) // NuGet
			return OLECMDERR_E_UNKNOWNGROUP;

			return OLECMDERR_E_UNKNOWNGROUP;

		if (   *pguidCmdGroup == GUID{ 0x25113E5B, 0x9964, 0x4375, { 0x9D, 0xD1, 0x0A, 0x5E, 0x98, 0x40, 0x50, 0x7A } } // no idea
			|| *pguidCmdGroup == GUID{ 0x64A2F652, 0x0B0F, 0x4413, { 0x90, 0xC8, 0x3E, 0x6A, 0x25, 0x77, 0x52, 0x90 } }
			|| *pguidCmdGroup == GUID{ 0x319FF225, 0x2E9C, 0x4FA7, { 0x92, 0x67, 0x12, 0x8F, 0xE4, 0x24, 0x6B, 0xF1 } } // Web Tools something
			|| *pguidCmdGroup == GUID{ 0xB101F7CB, 0x4BB9, 0x46D0, { 0xA4, 0x89, 0x83, 0x0D, 0x45, 0x01, 0x16, 0x0A } } // something from Microsoft.VisualStudio.ProjectSystem.VS.Implementation.dll
			|| *pguidCmdGroup == GUID{ 0x665CC136, 0x6455, 0x491D, { 0xAB, 0x17, 0xEA, 0xF3, 0x84, 0x7A, 0x23, 0xBC } } // same
			|| *pguidCmdGroup == GUID{ 0x7F917E79, 0x7A75, 0x4DCA, { 0xAE, 0x0A, 0xE5, 0x7A, 0x4A, 0x1E, 0xC5, 0xCE } } // no idea
		)
			return OLECMDERR_E_UNKNOWNGROUP;

		if (   *pguidCmdGroup == guidXamlCmdSet
			|| *pguidCmdGroup == guidVenusCmdId
			|| *pguidCmdGroup == guidSqlPkg
			|| *pguidCmdGroup == guidSourceExplorerPackage
			|| *pguidCmdGroup == guidWebProjectPackage
			|| *pguidCmdGroup == guidBrowserLinkCmdSet
			|| *pguidCmdGroup == guidUnknownMsenvDll
			|| *pguidCmdGroup == guidUnknownCmdGroup0
			|| *pguidCmdGroup == guidUnknownCmdGroup1
			|| *pguidCmdGroup == guidProjOverviewAppCapabilities
			|| *pguidCmdGroup == guidProjectAddTest
			|| *pguidCmdGroup == guidProjectAddWPF
			|| *pguidCmdGroup == guidUnknownCmdGroup5
			|| *pguidCmdGroup == guidUnknownCmdGroup6
			|| *pguidCmdGroup == guidUnknownCmdGroup7
			//|| *pguidCmdGroup == CMDSETID_HtmEdGrp
			//|| *pguidCmdGroup == CMDSETID_WebForms
			|| *pguidCmdGroup == guidXmlPkg
			|| *pguidCmdGroup == guidSccProviderPackage
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
		)
			return OLECMDERR_E_UNKNOWNGROUP;

		return OLECMDERR_E_UNKNOWNGROUP;
	}

	virtual HRESULT STDMETHODCALLTYPE Exec (const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
	{
		HRESULT hr;

		if (!pguidCmdGroup)
			return E_POINTER;

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97)
		{
			if (   nCmdID == cmdidSolutionCfg
				|| nCmdID == cmdidSolutionCfgGetList
				|| nCmdID == cmdidExit
				|| nCmdID == cmdidCloseSolution
				|| nCmdID == cmdidObjectVerbList0
				|| nCmdID == cmdidPropSheetOrProperties
				|| nCmdID == cmdidSave   // ignore it here and VS will - if dirty - pass it to the project's IVsPersistHierarchyItem::SaveItem
				|| nCmdID == cmdidSaveAs // same
			)
				return OLECMDERR_E_NOTSUPPORTED;

			if (nCmdID == cmdidBuildSel)
			{
				return OLECMDERR_E_NOTSUPPORTED;
			}
			
			if (nCmdID == cmdidOpen) // 261
			{
				com_ptr<IVsProject2> vsp;
				auto hr = _hier->QueryInterface(&vsp); RETURN_IF_FAILED(hr);
				wil::com_ptr_nothrow<IVsWindowFrame> windowFrame;
				hr = vsp->OpenItem (_itemId, LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &windowFrame); RETURN_IF_FAILED(hr);
				hr = windowFrame->Show(); RETURN_IF_FAILED(hr);
				return S_OK;
			}

			return OLECMDERR_E_NOTSUPPORTED;
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
			if (nCmdID == UIHWCMDID_RightClick)
			{
				POINTS pts;
				memcpy (&pts, &pvaIn->uintVal, 4);

				wil::com_ptr_nothrow<IVsUIShell> shell;
				hr = serviceProvider->QueryService(SID_SVsUIShell, &shell);
				if (FAILED(hr))
					return hr;

				return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_ITEMNODE, pts, this);
			}

			if (nCmdID == UIHWCMDID_DoubleClick || nCmdID == UIHWCMDID_EnterKey)
			{
				wil::com_ptr_nothrow<IVsProject2> vsp;
				hr = _hier->QueryInterface(&vsp); RETURN_IF_FAILED(hr);
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

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL *pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

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
		if (dispid == dispidPath)
		{
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

	HRESULT RenameFile (BSTR newName)
	{
		HRESULT hr;

		ULONG flags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;

		wil::unique_variant projectDir;
		hr = _hier->GetProperty (VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);

		wil::unique_hlocal_string oldFullPath;
		hr = PathAllocCombine (projectDir.bstrVal, _pathRelativeToProjectDir.get(), flags, oldFullPath.addressof()); RETURN_IF_FAILED(hr);

		VSQueryEditResult fEditVerdict;
		com_ptr<IVsQueryEditQuerySave2> queryEdit;
		hr = serviceProvider->QueryService (SID_SVsQueryEditQuerySave, &queryEdit); RETURN_IF_FAILED(hr);
		hr = queryEdit->QueryEditFiles (QEF_DisallowInMemoryEdits, 1, oldFullPath.addressof(), nullptr, nullptr, &fEditVerdict, nullptr); LOG_IF_FAILED(hr);
		if (FAILED(hr) || (fEditVerdict != QER_EditOK))
			return OLE_E_PROMPTSAVECANCELLED;

		// Check if the document is in the cache and rename document in the cache.
		com_ptr<IVsRunningDocumentTable> pRDT;
		hr = serviceProvider->QueryService(SID_SVsRunningDocumentTable, &pRDT); LOG_IF_FAILED(hr);
		VSCOOKIE dwCookie = VSCOOKIE_NIL;
		pRDT->FindAndLockDocument (RDT_NoLock, oldFullPath.get(), nullptr, nullptr, nullptr, &dwCookie); // ignore returned HRESULT as we're only interested in dwCookie

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

		com_ptr<IVsProject> project;
		hr = _hier->QueryInterface(&project); RETURN_IF_FAILED(hr);

		auto relativeDir = wil::make_hlocal_string_nothrow(_pathRelativeToProjectDir.get());
		hr = PathCchRemoveFileSpec (relativeDir.get(), wcslen(_pathRelativeToProjectDir.get()) + 1); RETURN_IF_FAILED(hr);
		wil::unique_hlocal_string newFullPath;
		hr = PathAllocCombine (projectDir.bstrVal, relativeDir.get(), flags, newFullPath.addressof()); RETURN_IF_FAILED(hr);
		hr = PathAllocCombine (newFullPath.get(), newName, flags, newFullPath.addressof()); RETURN_IF_FAILED(hr);

		com_ptr<IVsTrackProjectDocuments2> trackProjectDocs;
		hr = serviceProvider->QueryService (SID_SVsTrackProjectDocuments, &trackProjectDocs); RETURN_IF_FAILED(hr);
		BOOL fRenameCanContinue = FALSE;
		hr = trackProjectDocs->OnQueryRenameFile (project, oldFullPath.get(), newFullPath.get(), VSRENAMEFILEFLAGS_NoFlags, &fRenameCanContinue);
		if (FAILED(hr) || !fRenameCanContinue)
			return OLE_E_PROMPTSAVECANCELLED;

		wil::unique_hlocal_string otherPathRelativeToProjectDir;
		hr = PathAllocCombine (relativeDir.get(), newName, flags, otherPathRelativeToProjectDir.addressof()); RETURN_IF_FAILED(hr);

		if (!::MoveFile (oldFullPath.get(), newFullPath.get()))
			return HRESULT_FROM_WIN32(GetLastError());
		std::swap (_pathRelativeToProjectDir, otherPathRelativeToProjectDir);
		auto undoRename = wil::scope_exit([this, &otherPathRelativeToProjectDir, &newFullPath, &oldFullPath]
			{
				std::swap (_pathRelativeToProjectDir, otherPathRelativeToProjectDir);
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
		trackProjectDocs->OnAfterRenameFile (project, oldFullPath.get(), newFullPath.get(), VSRENAMEFILEFLAGS_NoFlags);

		// This line was changing the build tool when the user renamed the file.
		// I eventually commented it out, to get behavior similar to that in VS projects
		// (build tool remains unchanged when renaming for example from .cpp to .h)
		//_buildTool = _wcsicmp(PathFindExtension(_pathRelativeToProjectDir.get()), L".asm") ? BuildToolKind::None : BuildToolKind::Assembler;

		com_ptr<IVsHierarchyEvents> events;
		hr = _hier->QueryInterface(&events); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
		{
			events->OnPropertyChanged(_itemId, VSHPROPID_Caption, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_Name, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_SaveName, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_DescriptiveName, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_StateIconIndex, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_IconHandle, 0);
		}

		// Make sure the property browser is updated.
		com_ptr<IVsUIShell> uiShell;
		hr = serviceProvider->QueryService (SID_SVsUIShell, &uiShell); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
			uiShell->RefreshPropertyBrowser(DISPID_UNKNOWN); // refresh all properties

		// Mark project as dirty.
		com_ptr<IPropertyNotifySink> pns;
		hr = _hier->QueryInterface(&pns); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
			pns->OnChanged(dispidItems);

		//~CSuspendFileChanges

		undoRename.release();
		return S_OK;
	}
};

// ============================================================================

HRESULT MakeProjectFile (VSITEMID itemId, IVsUIHierarchy* hier, VSITEMID parentItemId, IProjectFile** file)
{
	com_ptr<ProjectFile> p = new (std::nothrow) ProjectFile(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(itemId, hier, parentItemId); RETURN_IF_FAILED(hr);
	*file = p.detach();
	return S_OK;
}
