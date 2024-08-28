
#include "pch.h"
#include <ivstrackprojectdocuments2.h>
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/com.h"
#include "../FelixPackageUi/resource.h"
#include "dispids.h"

struct Z80AsmFile : IZ80AsmFile, IProvideClassInfo, IOleCommandTarget, IVsGetCfgProvider
{
	static inline wil::com_ptr_nothrow<ITypeInfo> _typeInfo;
	VSITEMID _itemId;
	ULONG _refCount = 0;
	com_ptr<IVsUIHierarchy> _hier;
	wil::com_ptr_nothrow<IZ80ProjectItem> _next;
	VSITEMID _parentItemId = VSITEMID_NIL;
	VSCOOKIE _docCookie = VSDOCCOOKIE_NIL;
	wil::unique_hlocal_string _pathRelativeToProjectDir;
	
public:
	HRESULT InitInstance (VSITEMID itemId, IVsUIHierarchy* hier, VSITEMID parentItemId, ITypeLib* typeLib)
	{
		HRESULT hr;

		if (!_typeInfo)
		{
			hr = typeLib->GetTypeInfoOfGuid(__uuidof(IZ80AsmFile), &_typeInfo); RETURN_IF_FAILED(hr);
		}

		_itemId = itemId;
		_hier = hier;
		_parentItemId = parentItemId;
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDispatch*>(this), riid, ppvObject)
			|| TryQI<IZ80AsmFile>(this, riid, ppvObject)
			|| TryQI<IZ80SourceFile>(this, riid, ppvObject)
			|| TryQI<IZ80ProjectItem>(this, riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProvideClassInfo>(this, riid, ppvObject)
			|| TryQI<IOleCommandTarget>(this, riid, ppvObject)
			|| TryQI<IVsGetCfgProvider>(this, riid, ppvObject)
			)
			return S_OK;

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

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDispatch
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override
	{
		*pctinfo = 1;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override
	{
		_typeInfo.copy_to(ppTInfo);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override
	{
		if (cNames == 1 && !wcscmp(rgszNames[0], L"ExtenderCATID"))
			return DISP_E_UNKNOWNNAME; // For this one name we don't want any error logging

		auto hr = DispGetIDsOfNames (_typeInfo.get(), rgszNames, cNames, rgDispId); RETURN_IF_FAILED(hr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override
	{
		auto hr = DispInvoke (static_cast<IDispatch*>(this), _typeInfo.get(), dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	#pragma endregion

	#pragma region IZ80ProjectItem
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
	
	virtual IZ80ProjectItem* STDMETHODCALLTYPE Next() override { return _next.get(); }

	virtual void STDMETHODCALLTYPE SetNext (IZ80ProjectItem* next) override
	{
		//WI_ASSERT (!_next);
		_next = next;
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty (VSHPROPID propid, VARIANT* pvar) override
	{
		/*
		OutputDebugStringA("Z80AsmFile::GetProperty propid=");
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
			case VSHPROPID_IsHiddenItem: // -2043
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_NextVisibleSibling:
				return InitVariantFromInt32 (_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

			case VSHPROPID_NextSibling:
				return InitVariantFromInt32 (_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

			case VSHPROPID_Expandable: // -2006
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_SaveName: // -2002
			case VSHPROPID_Caption: // -2003
			case VSHPROPID_Name: // -2012
			case VSHPROPID_EditLabel: // -2026
			case VSHPROPID_DescriptiveName: // -2108
				return InitVariantFromString (PathFindFileName(_pathRelativeToProjectDir.get()), pvar);

			case VSHPROPID_ProvisionalViewingStatus: // -2112
				return InitVariantFromUInt32 (PVS_Disabled, pvar);

			case VSHPROPID_BrowseObject: // -2018
			{
				//wil::com_ptr_nothrow<IDispatch> props;
				//auto hr = Z80AsmFileProperties::CreateInstance (this, &props); RETURN_IF_FAILED(hr);
				//hr = InitVariantFromDispatch (props.get(), pvar); RETURN_IF_FAILED(hr);
				//return S_OK;
				return E_NOTIMPL;
			}

			case VSHPROPID_Parent: // -1000
				return InitVariantFromInt32 (_parentItemId, pvar);

			case VSHPROPID_ItemDocCookie: // -2034
				return InitVariantFromInt32 (_docCookie, pvar);

			case VSHPROPID_HasEnumerationSideEffects: // -2062
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_KeepAliveDocument: // -2075
				return E_NOTIMPL;

			case VSHPROPID_FirstChild:
				return InitVariantFromInt32 (VSITEMID_NIL, pvar);

			case VSHPROPID_FirstVisibleChild:
				return InitVariantFromInt32 (VSITEMID_NIL, pvar);

			case VSHPROPID_IsNonSearchable: // -2051
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_IsNonMemberItem: // -2044
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_ExtObject: // -2027
				return E_NOTIMPL;

			case VSHPROPID_ExternalItem: // -2103
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_IsNewUnsavedItem: // -2057
				return InitVariantFromBoolean (FALSE, pvar);

			case VSHPROPID_ShowOnlyItemCaption: // -2058
				return InitVariantFromBoolean (TRUE, pvar);

			case VSHPROPID_AltHierarchy: // -2019
			case VSHPROPID_StateIconIndex: // -2029
			case VSHPROPID_OverlayIconIndex: // -2048
			case VSHPROPID_ProjectTreeCapabilities: // -2146
			case VSHPROPID_IconIndex:
			case VSHPROPID_IconHandle:
			case VSHPROPID_OpenFolderIconHandle:
			case VSHPROPID_OpenFolderIconIndex:
			case VSHPROPID_IsSharedItemsImportFile: // -2154
			case VSHPROPID_SupportsIconMonikers: // -2159
				return E_NOTIMPL;

			default:
				return E_NOTIMPL;
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

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SetGuidProperty (VSHPROPID propid, REFGUID rguid) override
	{
		RETURN_HR(E_NOTIMPL);
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

	#pragma endregion

	#pragma region IZ80SourceFile
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
	#pragma endregion


	#pragma region IProvideClassInfo
	virtual HRESULT STDMETHODCALLTYPE GetClassInfo (ITypeInfo** ppTI) override
	{
		return E_NOTIMPL;
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
						case cmdidPropertyPages: // 232
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

				com_ptr<IServiceProvider> site;
				hr = _hier->GetSite(&site); RETURN_IF_FAILED(hr);
				wil::com_ptr_nothrow<IVsUIShell> shell;
				hr = site->QueryService(SID_SVsUIShell, &shell);
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

	#pragma region IVsGetCfgProvider
	virtual HRESULT STDMETHODCALLTYPE GetCfgProvider (IVsCfgProvider** ppCfgProvider) override
	{
		return _hier->QueryInterface(ppCfgProvider);
	}
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

		com_ptr<IVsHierarchyEvents> events;
		hr = _hier->QueryInterface(&events); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
		{
			events->OnPropertyChanged(_itemId, VSHPROPID_Caption, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_Name, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_SaveName, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_DescriptiveName, 0);
			events->OnPropertyChanged(_itemId, VSHPROPID_StateIconIndex, 0);
		}

		// TODO: repro this
		// Make sure the property browser is updated
		//com_ptr<IVsUIShell> uiShell;
		//hr = serviceProvider->QueryService (SID_SVsUIShell, &uiShell); LOG_IF_FAILED(hr);
		//if (SUCCEEDED(hr))
		//	uiShell->RefreshPropertyBrowser(DISPID_VALUE); // return value ignored on purpose

		// mark project as dirty
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

HRESULT MakeZ80AsmFile (VSITEMID itemId, IVsUIHierarchy* hier, VSITEMID parentItemId, ITypeLib* typeLib, IZ80AsmFile** file)
{
	com_ptr<Z80AsmFile> p = new (std::nothrow) Z80AsmFile(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(itemId, hier, parentItemId, typeLib); RETURN_IF_FAILED(hr);
	*file = p.detach();
	return S_OK;
}
