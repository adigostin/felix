
#include "pch.h"
#include "FelixPackage.h"
#include "shared/OtherGuids.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/inplace_function.h"
#include "shared/com.h"
#include "Z80Xml.h"
#include "dispids.h"
#include "guids.h"
#include "../FelixPackageUi/resource.h"

#pragma comment (lib, "Pathcch.lib")
#pragma comment (lib, "xmllite.lib")

static constexpr wchar_t ProjectElementName[] = L"Z80Project";
static constexpr wchar_t ConfigurationElementName[] = L"Configuration";
static constexpr wchar_t FileElementName[] = L"File";

// What MPF implements: https://docs.microsoft.com/en-us/visualstudio/extensibility/internals/project-model-core-components?view=vs-2022
class Z80Project
	: IZ80ProjectProperties    // includes IDispatch
	, IVsProject2              // includes IVsProject
	, IVsUIHierarchy           // includes IVsHierarchy
	, IPersistFileFormat       // includes IPersist
	, IVsPersistHierarchyItem
	, IVsGetCfgProvider
	, IVsCfgProvider2          // includes IVsCfgProvider
	, IProvideClassInfo
	, IOleCommandTarget
	, IPreferPropertyPagesWithTreeControl
	//, IConnectionPointContainer
	, IVsHierarchyDeleteHandler3
	, IXmlParent
	//, IVsProjectBuildSystem
	//, IVsBuildPropertyStorage
	//, IVsBuildPropertyStorage2
	, IProjectItemParent
	, IPropertyNotifySink
	, IVsHierarchyEvents // this implementation only forwards calls to subscribed sinks
	, IVsPerPropertyBrowsing
{
	wil::com_ptr_nothrow<IServiceProvider> _sp;
	ULONG _refCount = 0;
	GUID _projectInstanceGuid;
	wil::unique_hlocal_string _location;
	wil::unique_hlocal_string _filename;
	wil::unique_hlocal_string _caption;
	unordered_map_nothrow<VSCOOKIE, wil::com_ptr_nothrow<IVsHierarchyEvents>> _hierarchyEventSinks;
	VSCOOKIE _nextHierarchyEventSinkCookie = 1;
	bool _isDirty = false;
	bool _noScribble = false;
	wil::com_ptr_nothrow<IVsUIHierarchy> _parentHierarchy;
	VSITEMID _parentHierarchyItemId = VSITEMID_NIL; // item id of this project in the parent hierarchy
	com_ptr<IProjectItem> _firstChild;
	static inline VSITEMID _nextFileItemId = 1000; // VS uses 1..2..3 etc for the Solution item. I'll deal with this later.
	unordered_map_nothrow<VSCOOKIE, wil::com_ptr_nothrow<IVsCfgProviderEvents>> _cfgProviderEventSinks;
	VSCOOKIE _nextCfgProviderEventCookie = 1;
	VSCOOKIE _itemDocCookie = VSCOOKIE_NIL;
	vector_nothrow<com_ptr<IProjectConfig>> _configs;

	// I introduced this because VS sometimes retains a project for a long time after the user closes it.
	// For example in VS 17.11.2, when doing Close Solution while a project file was open, and then
	// opening another solution and another project file, VS calls GetCanonicalName on the old project.
	// This happens because VS retains a reference to the old project in
	// Microsoft.VisualStudio.PlatformUI.Packages.FileColor.ProjectFileGroupProvider.projectMap.
	// (Looks like a bug in OnBeforeCloseProject() in that class.)
	bool _closed = false;

	static HRESULT CreateProjectFilesFromTemplate (IServiceProvider* sp, const wchar_t* fromProjFilePath, const wchar_t* location, const wchar_t* filename)
	{
		// Prepare the file operations.

		wil::com_ptr_nothrow<IFileOperation> pfo;
		auto hr = CoCreateInstance(__uuidof(FileOperation), NULL, CLSCTX_ALL, IID_PPV_ARGS(&pfo)); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IVsUIShell> shell;
		hr = sp->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);

		HWND ownerHwnd;
		hr = shell->GetDialogOwnerHwnd(&ownerHwnd); RETURN_IF_FAILED(hr);
		hr = pfo->SetOwnerWindow(ownerHwnd); RETURN_IF_FAILED(hr);

		// -----------------------------------------
		// Create shell items for the source project file, source folder, destination folder.

		BOOL dirExists = PathFileExistsW(location);
		if (!dirExists)
		{
			DWORD gle = GetLastError();
			if (gle != ERROR_FILE_NOT_FOUND)
				RETURN_HR(HRESULT_FROM_WIN32(gle));

			BOOL bres = CreateDirectoryW(location, nullptr); RETURN_LAST_ERROR_IF(!bres);
		}

		wil::com_ptr_nothrow<IShellItem> destinationFolder;
		hr = SHCreateItemFromParsingName(location, nullptr, IID_PPV_ARGS(&destinationFolder)); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IShellItem> projectTemplateFile;
		hr = SHCreateItemFromParsingName(fromProjFilePath, nullptr, IID_PPV_ARGS(&projectTemplateFile)); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IPersistIDList> projectTemplateFileIDList;
		hr = projectTemplateFile->QueryInterface (&projectTemplateFileIDList); RETURN_IF_FAILED(hr);

		using itemidlist_t = wil::unique_any<ITEMIDLIST*, decltype(&::CoTaskMemFree), ::CoTaskMemFree>;
		itemidlist_t sourceProjectFileIDL;
		hr = projectTemplateFileIDList->GetIDList (&sourceProjectFileIDL); RETURN_IF_FAILED(hr);
		auto sourceProjectFileLastID = ILFindLastID(sourceProjectFileIDL.get());

		wil::com_ptr_nothrow<IShellItem> ptfp;
		hr = projectTemplateFile->GetParent(&ptfp); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IShellFolder> projectTemplateDir;
		hr = ptfp->BindToHandler(NULL, BHID_SFObject, IID_PPV_ARGS(&projectTemplateDir)); RETURN_IF_FAILED(hr);

		// ---------------------------------------
		// Copy only the project file for now.

		hr = pfo->CopyItem (projectTemplateFile.get(), destinationFolder.get(), filename, nullptr); RETURN_IF_FAILED(hr);

		// ---------------------------------------
		// Make a list of all the other files to be copied.

		wil::com_ptr_nothrow<IEnumIDList> enumIDList;
		hr = projectTemplateDir->EnumObjects (nullptr, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN, &enumIDList); RETURN_IF_FAILED(hr);

		itemidlist_t itemIDList[64];
		ULONG itemIDListCount = 0;
		ULONG fetched;
		while (true)
		{
			hr = enumIDList->Next(1, &itemIDList[itemIDListCount], &fetched); RETURN_IF_FAILED(hr);

			if (hr != S_OK)
				break;

			wil::unique_cotaskmem_string name;
			hr = SHGetNameFromIDList (itemIDList[itemIDListCount].get(), SIGDN_NORMALDISPLAY, &name);
			if (SUCCEEDED(hr) && !wcscmp(PathFindExtension(name.get()), L".vsdir"))
					continue;

			if (!ILIsEqual(sourceProjectFileLastID, itemIDList[itemIDListCount].get()))
			{
				itemIDListCount++;
				if (itemIDListCount == sizeof(itemIDList) / sizeof(itemIDList[0]))
					break;
			}
		}

		wil::com_ptr_nothrow<IShellItemArray> sourceFiles;
		hr = SHCreateShellItemArray (nullptr, projectTemplateDir.get(), (UINT)itemIDListCount, const_cast<LPCITEMIDLIST*>(itemIDList[0].addressof()), &sourceFiles); RETURN_IF_FAILED(hr);

		hr = pfo->CopyItems (sourceFiles.get(), destinationFolder.get()); RETURN_IF_FAILED(hr);

		// ----------------------------------------------------------------------

		hr = pfo->PerformOperations(); RETURN_IF_FAILED(hr);

		return S_OK;
	}

public:
	HRESULT InitInstance (IServiceProvider* sp, LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags)
	{
		HRESULT hr;

		_sp = sp;

		if (grfCreateFlags & CPF_CLONEFILE)
		{
			_location = wil::make_hlocal_string_nothrow(pszLocation); RETURN_IF_NULL_ALLOC(_location);
			_filename = wil::make_hlocal_string_nothrow(pszName); RETURN_IF_NULL_ALLOC(_filename);
			const wchar_t* ext = ::PathFindExtension(pszName);
			_caption = wil::make_hlocal_string_nothrow(pszName, ext - pszName); RETURN_IF_NULL_ALLOC(_caption);

			hr = CreateProjectFilesFromTemplate (sp, pszFilename, pszLocation, pszName); RETURN_IF_FAILED(hr);

			wil::unique_hlocal_string projFilePath;
			hr = PathAllocCombine (pszLocation, pszName, PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS, &projFilePath); RETURN_IF_FAILED(hr);

			com_ptr<IStream> stream;
			auto hr = SHCreateStreamOnFileEx(projFilePath.get(), STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, 0, nullptr, &stream); RETURN_IF_FAILED(hr);
			hr = LoadFromXml(this, ProjectElementName, stream.get()); RETURN_IF_FAILED(hr);

			hr = CoCreateGuid (&_projectInstanceGuid); RETURN_IF_FAILED(hr);

			hr = SHCreateStreamOnFileEx (projFilePath.get(), STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, 0, nullptr, &stream); RETURN_IF_FAILED(hr);
			hr = SaveToXml(this, ProjectElementName, stream.get()); RETURN_IF_FAILED(hr);

			_isDirty = false;
			return S_OK;
		}
		else if (grfCreateFlags & CPF_OPENFILE)
		{
			// pszFilename is the full path of the file to open, the others are NULL.
			_location = wil::make_hlocal_string_nothrow(pszFilename); RETURN_IF_NULL_ALLOC(_location);
			PathRemoveFileSpec(_location.get());

			const wchar_t* fn = PathFindFileName(pszFilename);
			_filename = wil::make_hlocal_string_nothrow(fn); RETURN_IF_NULL_ALLOC(_filename);

			const wchar_t* ext = PathFindExtension(pszFilename);
			_caption = wil::make_hlocal_string_nothrow(fn, ext - fn); RETURN_IF_NULL_ALLOC(_caption);

			com_ptr<IStream> stream;
			auto hr = SHCreateStreamOnFileEx(pszFilename, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, 0, nullptr, &stream); RETURN_IF_FAILED(hr);
			hr = LoadFromXml (this, ProjectElementName, stream.get()); RETURN_IF_FAILED(hr);

			_isDirty = false;
			return S_OK;
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	~Z80Project()
	{
		WI_ASSERT (_cfgProviderEventSinks.empty());
	}

	// If found, returns S_OK and ppItem is non-null.
	// If not found, return S_FALSE and ppItem is null.
	// If error during enumeration, returns an error code.
	// The filter function must return S_OK for a match (enumeration stops), S_FALSE for non-match (enumeration continues), or an error code (enumeration stops).
	HRESULT FindDescendant (const stdext::inplace_function<HRESULT(IProjectItem*), 32>& filter, IProjectItem** ppItem, IProjectItem** ppPrevSibling = nullptr, IProjectItemParent** ppParent = nullptr)
	{
		*ppItem = nullptr;
		if (ppPrevSibling)
			*ppPrevSibling = nullptr;
		if (ppParent)
			*ppParent = nullptr;

		IProjectItem* prev = nullptr;
		for (auto c = _firstChild.get(); c; c = c->Next())
		{
			wil::unique_variant fc;
			auto hr = c->GetProperty(VSHPROPID_FirstChild, &fc); RETURN_IF_FAILED(hr);
			if ((fc.vt == VT_VSITEMID) && (V_VSITEMID(fc.addressof()) != VSITEMID_NIL))
			{
				// TODO: search more than one level
				RETURN_HR(E_NOTIMPL);
			}

			hr = filter(c); RETURN_IF_FAILED(hr);
			if (hr == S_OK)
			{
				*ppItem = c;
				c->AddRef();
				
				if (ppPrevSibling && prev)
				{
					*ppPrevSibling = prev;
					prev->AddRef();
				}
				if (ppParent)
				{
					*ppParent = this;
					this->AddRef();
				}

				return S_OK;
			}

			prev = c;
		}

		return S_FALSE;
	}
	
	IProjectItem* FindDescendant (VSITEMID itemid, IProjectItem** ppPrevSibling = nullptr, IProjectItemParent** ppParent = nullptr)
	{
		auto sameId = [itemid](IProjectItem* item) { return (item->GetItemId() == itemid) ? S_OK : S_FALSE; };
		wil::com_ptr_nothrow<IProjectItem> item;
		auto hr = FindDescendant(sameId, &item, ppPrevSibling, ppParent); WI_ASSERT(SUCCEEDED(hr));
		return item.get();
	}

	HRESULT QueryStatusInternal (const GUID* pguidCmdGroup, ULONG cmdID, __RPC__out DWORD* cmdf, __RPC__out OLECMDTEXT *pCmdText)
	{
		if (*pguidCmdGroup == CLSID_FelixPackageCmdSet)
			return OLECMDERR_E_UNKNOWNGROUP;

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97)
		{
			// These are the cmdidXxxYyy constants from stdidcmd.h
			switch (cmdID)
			{
				case cmdidCopy: // 15
				case cmdidPaste: // 26
				case cmdidCut: // 16
				//case cmdidSave: // 110
				//case cmdidSaveAs: // 111
				case cmdidRename: // 150
				case cmdidAddNewItem: // 220
				case cmdidSaveProjectItemAs: // 226
				case cmdidExit: // 229
				case cmdidPropertyPages: // 232
				case cmdidAddExistingItem: // 244
				case cmdidStepInto: // 248
				case cmdidSaveProjectItem: // 331
				case cmdidPropSheetOrProperties: // 397
				case cmdidCloseDocument: // 658
				case cmdidBuildSln: // 882
				case cmdidRebuildSln: // 883
				case cmdidCleanSln: // 885
				//case cmdidBuildSel: // 886
				//case cmdidRebuildSel: // 887
				//case cmdidCleanSel: // 889
					*cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
					break;
					
				case cmdidToggleBreakpoint: // 255
				case cmdidInsertBreakpoint: // 375
				case cmdidEnableBreakpoint: // 376
					*cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
					break;
					
				default:
					*cmdf = 0; // not supported
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet2K)
		{
			switch (cmdID)
			{
				case cmdidBuildOnlyProject:   // 1603
				case cmdidRebuildOnlyProject: // 1604
				case cmdidCleanOnlyProject:   // 1605
					*cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
					break;

				default:
					*cmdf = 0; // not supported
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet10)
		{
			if (cmdID >= cmdidDynamicToolBarListFirst && cmdID <= cmdidDynamicToolBarListLast)
			{
				*cmdf = 0; // not supported
			}
			else
			{
				switch (cmdID)
				{
					case cmdidShellNavigate1First: // 1000
					case cmdidExtensionManager: // 3000
						*cmdf = 0; // not supported
						break;

					default:
						*cmdf = 0; // not supported
				}
			}

			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet11)
		{
			if (cmdID == cmdidStartupProjectProperties) // 21
			{
				*cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
				return S_ALLTHRESHOLD;
			}

			return E_NOTIMPL;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet12)
			return OLECMDERR_E_UNKNOWNGROUP;

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet16)
			return OLECMDERR_E_UNKNOWNGROUP;

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet17)
			return OLECMDERR_E_UNKNOWNGROUP;

		if (*pguidCmdGroup == guidVSDebugCommand)
		{
			// These are in VsDbgCmd.h
			if (cmdID == cmdidBreakpointsWindowShow)
			{
				*cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
				return S_OK;
			}

			return E_NOTIMPL;
		}

		#ifdef _DEBUG
		if (   *pguidCmdGroup == guidUnknownCmdGroup0
			|| *pguidCmdGroup == guidUnknownCmdGroup1
			|| *pguidCmdGroup == guidUnknownCmdGroup5
			|| *pguidCmdGroup == guidUnknownCmdGroup6
			|| *pguidCmdGroup == guidUnknownCmdGroup7
			|| *pguidCmdGroup == guidUnknownCmdGroup8
			|| *pguidCmdGroup == guidUnknownCmdGroup9
			|| *pguidCmdGroup == guidSourceExplorerPackage
			|| *pguidCmdGroup == WebAppCmdId
			|| *pguidCmdGroup == guidWebProjectPackage
			|| *pguidCmdGroup == guidBrowserLinkCmdSet
			|| *pguidCmdGroup == guidUnknownMsenvDll
			|| *pguidCmdGroup == CMDSETID_HtmEdGrp
			|| *pguidCmdGroup == guidXmlPkg
			|| *pguidCmdGroup == XsdDesignerPackage
			|| *pguidCmdGroup == guidSccProviderPackage
			|| *pguidCmdGroup == GUID{ 0xB101F7CB, 0x4BB9, 0x46D0, { 0xA4, 0x89, 0x83, 0x0D, 0x45, 0x01, 0x16, 0x0A } } // something from Microsoft.VisualStudio.ProjectSystem.VS.Implementation.dll
			|| *pguidCmdGroup == GUID{ 0x665CC136, 0x6455, 0x491D, { 0xAB, 0x17, 0xEA, 0xF3, 0x84, 0x7A, 0x23, 0xBC } } // same
			|| *pguidCmdGroup == guidProjOverviewAppCapabilities
			|| *pguidCmdGroup == guidTrackProjectRetargetingCmdSet
			|| *pguidCmdGroup == guidProjectAddTest
			|| *pguidCmdGroup == guidProjectAddWPF
			|| *pguidCmdGroup == guidProjectClassWizard
			|| *pguidCmdGroup == guidDataSources
			|| *pguidCmdGroup == tfsCmdSet
			|| *pguidCmdGroup == tfsCmdSet1
			|| *pguidCmdGroup == guidCmdGroupClassDiagram
			|| *pguidCmdGroup == guidCmdGroupDatabase
			|| *pguidCmdGroup == guidCmdGroupTableQueryDesigner
			|| *pguidCmdGroup == guidXamlCmdSet
			|| *pguidCmdGroup == guidCmdGroupTestExplorer
			|| *pguidCmdGroup == guidVSEQTPackageCmdSet
			|| *pguidCmdGroup == guidSomethingResourcesCmdSet
			|| *pguidCmdGroup == guidCmdSetPerformance
			|| *pguidCmdGroup == guidCmdSetEdit
			|| *pguidCmdGroup == guidCmdSetHelp
			|| *pguidCmdGroup == guidCmdSetTaskRunnerExplorer
			|| *pguidCmdGroup == guidCmdSetCtxMenuPrjSolOtherClsWiz
			|| *pguidCmdGroup == guidCmdSetCodeAnalysis
			|| *pguidCmdGroup == guidCmdSetCodeMetrics
			|| *pguidCmdGroup == guidCmdSetClassViewProject
			|| *pguidCmdGroup == guidCmdSetSolExplPivotStartList
			|| *pguidCmdGroup == guidCmdSetWebToolsScaffolding
			|| *pguidCmdGroup == guidCmdSetWebToolsScaffolding2
			|| *pguidCmdGroup == guidUniversalProjectsCmdSet
			|| *pguidCmdGroup == guidCmdSetNetAspire
			|| *pguidCmdGroup == guidCmdSetBowerPackages
			|| *pguidCmdGroup == guidDotNetCoreWebCmdId
			|| *pguidCmdGroup == guidCmdSetCSharpInteractive
			|| *pguidCmdGroup == guidCmdSetTFS
			|| *pguidCmdGroup == guidSccPkg
			|| *pguidCmdGroup == guidToolWindowTimestampButton
			|| *pguidCmdGroup == guidCmdSetSomethingIntellisense
			|| *pguidCmdGroup == guidCmdSetSomethingTerminal
			|| *pguidCmdGroup == guidSHLMainMenu
			|| *pguidCmdGroup == guidCommonIDEPackage
		)
			return OLECMDERR_E_UNKNOWNGROUP;
		#endif

		if (*pguidCmdGroup == guidNuGetDialogCmdSet || *pguidCmdGroup == guidNuGetSomethingCmdSet)
		{
			*cmdf = 0;
			return S_OK;
		}

		if (*pguidCmdGroup == DebugTargetTypeCommandGuid)
		{
			if (cmdID == 0x100)
			{
				// DebugTargetTypeCommandId
				return E_NOTIMPL;
			}

			return E_NOTIMPL;
		}

		//BreakIntoDebugger();
		return OLECMDERR_E_UNKNOWNGROUP;
	}

	HRESULT ExecInternal(const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn, VARIANT* pvaOut)
	{
		if (!pguidCmdGroup)
			return E_POINTER;

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97)
		{
			if (   nCmdID == cmdidObjectVerbList0 // 137
				|| nCmdID == cmdidCloseSolution // 219
				|| nCmdID == cmdidExit // 229
				|| nCmdID == cmdidOutputWindow // 237
				|| nCmdID == cmdidPropSheetOrProperties // 397
				|| nCmdID == cmdidCloseDocument // 658
				|| nCmdID == cmdidBuildSel // 886
				|| nCmdID == cmdidStart          // 295 - ignore it here, and the IDE will pass it to our IVsBuildableProjectCfg
				|| nCmdID == cmdidStop           // 179 - same
				|| nCmdID == cmdidDebugProcesses // 213 - same
				|| nCmdID == cmdidSaveProjectItem // 331 - ignore it here and VS will - if dirty - pass it to IPersistFileFormat::Save
				|| nCmdID == cmdidSaveProjectItemAs // 226 - same
			)
				return OLECMDERR_E_NOTSUPPORTED;
			
			if (nCmdID == cmdidAddNewItem) // 220
				return ProcessCommandAddNewItem(TRUE);
			if (nCmdID == cmdidAddExistingItem) // 244
				return ProcessCommandAddNewItem(FALSE);

			//BreakIntoDebugger();
			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet11)
		{
			if (nCmdID == cmdidLocateFindTarget)
				return OLECMDERR_E_NOTSUPPORTED;

			//BreakIntoDebugger();
			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet2K)
		{
			if (nCmdID == ECMD_COMPILE)
				return OLECMDERR_E_NOTSUPPORTED;

			//BreakIntoDebugger();
			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == GUID_VsUIHierarchyWindowCmds)
		{
			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == guidNuGetDialogCmdSet || *pguidCmdGroup == guidNuGetSomethingCmdSet)
		{
			wil::com_ptr_nothrow<IVsUIShell> shell;
			auto hr = _sp->QueryService (SID_SVsUIShell, &shell);
			if (SUCCEEDED(hr))
			{
				LONG result;
				const wchar_t text[] = L"Sorry, Visual Studio won't allow me to remove the Nuget nonsense from the menus";
				shell->ShowMessageBox (0, GUID_NULL, nullptr, const_cast<LPOLESTR>(text),
					nullptr, 0, OLEMSGBUTTON_OK, OLEMSGDEFBUTTON_FIRST, OLEMSGICON_INFO, FALSE, &result);
			}

			return S_OK;
		}

		return OLECMDERR_E_UNKNOWNGROUP;
	}

	HRESULT ProcessCommandAddNewItem (BOOL fAddNewItem)
	{
		wil::com_ptr_nothrow<IVsAddProjectItemDlg> dlg;
		auto hr = _sp->QueryService(SID_SVsAddProjectItemDlg, &dlg); RETURN_IF_FAILED(hr);
		VSADDITEMFLAGS flags;
		if (fAddNewItem)
			flags = VSADDITEM_AddNewItems | VSADDITEM_SuggestTemplateName | VSADDITEM_NoOnlineTemplates; /* | VSADDITEM_ShowLocationField */
		else
			flags = VSADDITEM_AddExistingItems | VSADDITEM_AllowMultiSelect | VSADDITEM_AllowStickyFilter;

		// TODO: To specify a sticky behavior for the location field, which is the recommended behavior, remember the last location field value and pass it back in when you open the dialog box again.
		// TODO: To specify sticky behavior for the filter field, which is the recommended behavior, remember the last filter field value and pass it back in when you open the dialog box again.
		hr = dlg->AddProjectItemDlg (VSITEMID_ROOT, __uuidof(IZ80ProjectProperties), this, flags, nullptr, nullptr, nullptr, nullptr, nullptr);
		if (FAILED(hr) && (hr != OLE_E_PROMPTSAVECANCELLED))
			return hr;

		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		RETURN_HR_IF_EXPECTED(E_NOINTERFACE, riid == IID_ICustomCast); // VS abuses this, let's check it first

		if (   TryQI<IZ80ProjectProperties>(this, riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IUnknown>(static_cast<IVsUIHierarchy*>(this), riid, ppvObject)
			|| TryQI<IVsHierarchy>(this, riid, ppvObject)
			|| TryQI<IVsUIHierarchy>(this, riid, ppvObject)
			|| TryQI<IPersist>(this, riid, ppvObject)
			|| TryQI<IPersistFileFormat>(this, riid, ppvObject)
			|| TryQI<IVsProject>(this, riid, ppvObject)
			|| TryQI<IVsProject2>(this, riid, ppvObject)
			|| TryQI<IOleCommandTarget>(this, riid, ppvObject)
			|| TryQI<IProvideClassInfo>(this, riid, ppvObject)
			|| TryQI<IVsGetCfgProvider>(this, riid, ppvObject)
			|| TryQI<IVsCfgProvider>(this, riid, ppvObject)
			|| TryQI<IVsCfgProvider2>(this, riid, ppvObject)
			//|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
			|| TryQI<IVsHierarchyDeleteHandler3>(this, riid, ppvObject)
			|| TryQI<IVsPersistHierarchyItem>(this, riid, ppvObject)
			|| TryQI<IPreferPropertyPagesWithTreeControl>(this, riid, ppvObject)
			|| TryQI<IXmlParent>(this, riid, ppvObject)
			//|| TryQI<IVsProjectBuildSystem>(this, riid, ppvObject)
			//|| TryQI<IVsBuildPropertyStorage>(this, riid, ppvObject)
			//|| TryQI<IVsBuildPropertyStorage2>(this, riid, ppvObject)
			|| TryQI<IProjectItemParent>(this, riid, ppvObject)
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
			|| TryQI<IVsHierarchyEvents>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		// These will never be implemented.
		if (   riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IManagedObject
			|| riid == IID_IEnumerable
			|| riid == IID_IConvertible
			|| riid == IID_IInspectable
			|| riid == IID_IWeakReferenceSource
			|| riid == IID_ICustomTypeDescriptor
			|| riid == IID_IComponent
			//|| riid == IID_VSProject  // C# and VB stuff - When removing these, remove also the #include <vslangprojX.h> lines at the top of this file.
			//|| riid == IID_VSProject2 // same
			//|| riid == IID_VSProject3 // same
			//|| riid == IID_VSProject4 // same
			|| riid == IID_IUseImmediateCommitPropertyPages // Some VS internal thing
			|| riid == GUID{ 0xDBF177F2, 0x06DB, 0x4A47, { 0x8A, 0xAD, 0xC8, 0xE1, 0x2B, 0xFD, 0x6C, 0x86 } } // VCProject
			|| riid == GUID{ 0xF70D3A58, 0x8522, 0x4D8B, { 0x81, 0x89, 0xFC, 0x0D, 0x25, 0x35, 0x46, 0x44 } } // no idea
			|| riid == GUID{ 0xFC1E6B3F, 0x5123, 0x3042, { 0xBA, 0xEB, 0xEF, 0xBF, 0x10, 0x45, 0x54, 0x0A } } // something from Microsoft.VisualStudio.ProjectSystem.VS.Implementation.dll
			|| riid == GUID{ 0x19893E06, 0x22DA, 0x4A7D, { 0xAD, 0x25, 0xDF, 0xED, 0x10, 0x40, 0x90, 0x3B } } // same
			|| riid == GUID{ 0x779DC4EF, 0x2E1A, 0x4F3A, { 0xBA, 0x0B, 0x52, 0x96, 0xD2, 0x93, 0xE2, 0xB9 } } // something from Microsoft.VisualStudio.Project.VisualC.VCProjectEngine.dll
			|| riid == GUID{ 0xE952CF02, 0xF120, 0x4119, { 0x8F, 0x94, 0x5D, 0x66, 0xF0, 0xC3, 0x8B, 0xF0 } } // no idea
			|| riid == __uuidof(IVsParentProject)
			|| riid == __uuidof(IVsAggregatableProject)
			|| riid == IID_IVsParentProject3
			|| riid == IID_IVsUIHierarchyEventsPrivate
			|| riid == IID_INoIdea3
			|| riid == IID_INoIdea4
			|| riid == GUID{ 0x266e178e, 0x46b9, 0x4e95, { 0x91, 0x0c, 0x43, 0xaf, 0x54, 0xe3, 0x02, 0x3c } } // IVsStubWindowPrivate
			|| riid == IID_IProvideMultipleClassInfo
			|| riid == IID_IExtendedObject
			|| riid == IID_IVsSccProjectProviderBinding
			|| riid == IID_IVsSccProject2
			|| riid == GUID{ 0x101d210d, 0x5b28, 0x4e02, { 0xb2, 0x20, 0x19, 0x94, 0x9f, 0xf4, 0x02, 0x3b } } // IID_IVsProjectAsyncOpen
			|| riid == IID_SolutionProperties
		)
			return E_NOINTERFACE;

		if (   riid == __uuidof(IVsProjectBuildSystem)
			|| riid == __uuidof(IVsBuildPropertyStorage)
			|| riid == __uuidof(IVsBuildPropertyStorage2)
			|| riid == IID_IVsBooleanSymbolPresenceChecker
		)
			return E_NOINTERFACE;

		if (riid == IID_IVsFilterAddProjectItemDlg || riid == IID_IVsSolution)
			return E_NOINTERFACE;

		if (riid == IID_IPersistFile)
			return E_NOINTERFACE;
		if (riid == IID_IVsPersistDocData)
			return E_NOINTERFACE;
		if (riid == IID_IVsPersistDocData2)
			return E_NOINTERFACE;

		if (riid == __uuidof(IVsSolution))
			return E_NOINTERFACE;

		if (riid == __uuidof(IVsFileBackup))
			return E_NOINTERFACE;

		if (riid == __uuidof(IVsSaveOptionsDlg))
			return E_NOINTERFACE;

		if (riid == __uuidof(ISupportErrorInfo))
			return E_NOINTERFACE;

		if (   riid == __uuidof(IVsProject4)
			|| riid == __uuidof(IVsProject5)
			|| riid == __uuidof(IVsHasRelatedSaveItems)
		)
			return E_NOINTERFACE;

		if (riid == __uuidof(ICategorizeProperties))
			return E_NOINTERFACE;

		if (   riid == IID_IPerPropertyBrowsing
//			|| riid == IID_IVSMDPerPropertyBrowsing
			//|| riid == IID_IProvidePropertyBuilder
			|| riid == IID_IVsFilterAddProjectItemDlg
			|| riid == IID_IVsWindowFrame
			|| riid == IID_IVsTextBuffer
			|| riid == IID_IVsProjectUpgrade
			|| riid == IID_IVsTextLines
			|| riid == IID_IVsTextBufferProvider
			|| riid == IID_IVsDependencyProvider
			|| riid == IID_IVsSupportItemHandoff
			|| riid == GUID{ 0xB773DCB5, 0x5690, 0x342A, { 0x92, 0x32, 0xE9, 0x20, 0xD7, 0x5C, 0xA4, 0x53 } } // some weird IID that VS looks for when closing the VS application after some debugging was done
			|| riid == IID_IVsCurrentCfg
		)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IZ80ProjectProperties);

	#pragma region IVsHierarchy
	virtual HRESULT STDMETHODCALLTYPE SetSite(IServiceProvider* pSP) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetSite(IServiceProvider** ppSP) override
	{
		*ppSP = _sp.get();
		_sp->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryClose(BOOL* pfCanClose) override
	{
		*pfCanClose = TRUE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Close() override
	{
		_parentHierarchy = nullptr;
		_parentHierarchyItemId = VSITEMID_NIL;
		if (_firstChild)
		{
			_firstChild->Close();
			_firstChild = nullptr;
		}
		_configs.clear();
		_closed = true;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetGuidProperty(VSITEMID itemid, VSHPROPID propid, GUID* pguid) override
	{
		if (_closed)
			return E_UNEXPECTED;

		if (itemid == VSITEMID_ROOT)
		{
			if (propid == VSHPROPID_TypeGuid) // -1004
			{
				*pguid = GUID_ItemType_PhysicalFile;
				return S_OK;
			}

			if (propid == VSHPROPID_CmdUIGuid) // -2016
			{
				// TODO: https://microsoft.public.vstudio.extensibility.narkive.com/bCIjzuRK/adding-a-command-to-project-node-context-menu
				*pguid = __uuidof(IZ80ProjectProperties);
				return S_OK;
			}

			if (propid == VSHPROPID_PreferredLanguageSID) // -2054
				return E_NOTIMPL;

			if (propid == VSHPROPID_ProjectIDGuid) // -2059
			{
				*pguid = _projectInstanceGuid;
				return S_OK;
			}

			if (propid == VSHPROPID_BrowseObjectCATID) // -2068
			{
				// have a look in HierarchyNode.cs, "VSHPROPID_BrowseObjectCATID"
				RETURN_HR(E_NOTIMPL);
			}

			if (propid == VSHPROPID_AddItemTemplatesGuid) // -2070
				return E_NOTIMPL;

			// see the related VSHPROPID_SupportsProjectDesigner
			if (propid == VSHPROPID_ProjectDesignerEditor) // -2088
				return E_NOTIMPL;

			RETURN_HR(E_NOTIMPL);
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->GetGuidProperty(propid, pguid);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE SetGuidProperty(VSITEMID itemid, VSHPROPID propid, REFGUID rguid) override
	{
		if (_closed)
			return E_UNEXPECTED;

		// TODO: property change notifications

		if (itemid == VSITEMID_ROOT)
		{
			if (propid == VSHPROPID_ProjectIDGuid)
			{
				_projectInstanceGuid = rguid;
				// TODO: property change notification
				return S_OK;
			}

			RETURN_HR(E_NOTIMPL);
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->SetGuidProperty(propid, rguid);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty (VSITEMID itemid, VSHPROPID propid, VARIANT* pvar) override
	{
		// https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.__vshpropid?view=visualstudiosdk-2017

		if (_closed)
			return E_UNEXPECTED;

		#pragma region Some properties return the same thing for any node of our hierarchy
		if (propid == VSHPROPID_ParentHierarchy) // -2032
			return InitVariantFromUnknown (_parentHierarchy.get(), pvar);

		if (propid == VSHPROPID_ParentHierarchyItemid) // -2033
			return InitVariantFromInt32 (_parentHierarchyItemId, pvar);
		#pragma endregion

		if (itemid == VSITEMID_ROOT)
		{
			if (propid == VSHPROPID_Parent) // -1000
				// The VSITEMID of the node’s parent item. For example, in projects a folder node is the parent of
				// the files put into that folder. The property uses a System.Int32 value. If the node is a direct
				// parent of the root node, VSITEMID_ROOT is retrieved. If the node is at the root level and it
				// does not have a parent, VSITEMID_NIL is returned.
				return InitVariantFromInt32 (VSITEMID_NIL, pvar);

			if (propid == VSHPROPID_FirstChild) // -1001
				return InitVariantFromInt32 (_firstChild ? _firstChild->GetItemId() : VSITEMID_NIL, pvar);

			if (propid == VSHPROPID_SaveName) // -2002
				return InitVariantFromString (_filename.get(), pvar);

			if (propid == VSHPROPID_Caption) // -2003
				return InitVariantFromString (_caption.get(), pvar);

			if (propid == VSHPROPID_Expandable) // -2006
				return InitVariantFromBoolean (_firstChild ? TRUE : FALSE, pvar);

			if (propid == VSHPROPID_ExpandByDefault) // -2011
				return InitVariantFromBoolean (TRUE, pvar);

			if (propid == VSHPROPID_Name) // -2012
				return InitVariantFromString (_filename.get(), pvar);

			if (propid == VSHPROPID_BrowseObject) // -2018
			{
				// Be aware that VS2019 will QueryInterface for IVsHierarchy on the object we return here,
				// then call again this GetProperty function to get another browse object. It does this
				// recursively and will stop when (1) we return an error HRESULT, or (2) we return the same
				// browse object as in the previous call, or (3) it overflows the stack (that exception
				// seems to be caught and turned into a freeze).
				auto hr = InitVariantFromDispatch (this, pvar); RETURN_IF_FAILED(hr);
				return S_OK;
			}

			if (propid == VSHPROPID_ProjectDir) //-2021
				return InitVariantFromString (_location.get(), pvar);

			if (propid == VSHPROPID_EditLabel) // -2026,
			{
				const wchar_t* ext = ::PathFindExtension(_filename.get());
				BSTR bs = SysAllocStringLen(_filename.get(), (UINT)(ext - _filename.get())); RETURN_IF_NULL_ALLOC(bs);
				pvar->vt = VT_BSTR;
				pvar->bstrVal = bs;
				return S_OK;
			}

			if (propid == VSHPROPID_TypeName) // -2030
				return InitVariantFromString (L"Z80", pvar); // Called by 17.11.5 from Help -> About.

			if (propid == VSHPROPID_HandlesOwnReload) // -2031
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_ItemDocCookie) // -2034
				return InitVariantFromInt32 (_itemDocCookie, pvar);

			if (propid == VSHPROPID_IsHiddenItem) // -2043
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_IsNonLocalStorage) // -2045
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_IsNonSearchable) // -2051
				return InitVariantFromBoolean (TRUE, pvar);

			if (propid == VSHPROPID_CanBuildFromMemory) // -2053
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_IsNewUnsavedItem) // -2057
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_ShowOnlyItemCaption) // -2058
				return InitVariantFromBoolean (TRUE, pvar);

			if (propid == VSHPROPID_HasEnumerationSideEffects) // -2062
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_DefaultEnableBuildProjectCfg) // -2063
				return InitVariantFromBoolean (TRUE, pvar);

			if (propid == VSHPROPID_DefaultEnableDeployProjectCfg) // -2064
				return InitVariantFromBoolean (FALSE, pvar);

			// See the related VSHPROPID_ProjectDesignerEditor.
			if (propid == VSHPROPID_SupportsProjectDesigner) // -2076
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_EnableDataSourceWindow) // -2083
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_TargetPlatformIdentifier) // -2114
				return InitVariantFromString (L"ZX Spectrum", pvar);

			if (propid == VSHPROPID_TargetPlatformVersion) // -2115
				return InitVariantFromString (L"100K", pvar);

			if (propid == VSHPROPID_IsFaulted) // -2122
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_MonikerSameAsPersistFile) // -2130
				return InitVariantFromBoolean (TRUE, pvar);

			#ifdef _DEBUG
			if (   propid == VSHPROPID_IconIndex                  // -2005
				|| propid == VSHPROPID_IconHandle                 // -2013
				|| propid == VSHPROPID_OpenFolderIconHandle       // -2014
				|| propid == VSHPROPID_OpenFolderIconIndex        // -2015
				|| propid == VSHPROPID_AltHierarchy               // -2019
				|| propid == VSHPROPID_SortPriority               // -2022 - requested when reverting in Git an open project file
				|| propid == VSHPROPID_ExtObject                  // -2027
				|| propid == VSHPROPID_StateIconIndex             // -2029
				|| propid == VSHPROPID_ConfigurationProvider      // -2036
				|| propid == VSHPROPID_ImplantHierarchy           // -2037 - "This property is optional."
				|| propid == VSHPROPID_FirstVisibleChild          // -2041 - We show all children, so no need to handle this.
				|| propid == VSHPROPID_OverlayIconIndex           // -2048
				|| propid == VSHPROPID_ShowProjInSolutionPage     // -2055
				|| propid == VSHPROPID_StatusBarClientText        // -2072
				|| propid == VSHPROPID_DebuggeeProcessId          // -2073
				|| propid == VSHPROPID_DebuggerSourcePaths        // -2085
				|| propid == VSHPROPID_NoDefaultNestedHierSorting // -2090 - Default should be good enough
				|| propid == VSHPROPID_ProductBrandName           // -2099 - Default should be good enough
				|| propid == VSHPROPID_TargetFrameworkMoniker     // -2102 - We'll never support this
				|| propid == VSHPROPID_ExternalItem               // -2103 - MPF doesn't implement this
				|| propid == VSHPROPID_DescriptiveName            // -2108
				|| propid == VSHPROPID_ProvisionalViewingStatus   // -2112
				|| propid == VSHPROPID_TargetRuntime              // -2116 - We'll never support this as it has values only for JS, CLR, Native
				|| propid == VSHPROPID_AppContainer               // -2117 - Something about .Net
				|| propid == VSHPROPID_OutputType                 // -2118 - "This property is optional."
				|| propid == VSHPROPID_ProjectUnloadStatus        // -2120
				|| propid == VSHPROPID_DemandLoadDependencies     // -2121 - Much later, if ever
				|| propid == VSHPROPID_ProjectCapabilities        // -2124 - MPF doesn't implement this
				|| propid == VSHPROPID_ProjectRetargeting         // -2134 - We won't be supporting IVsRetargetProject
				|| propid == VSHPROPID_Subcaption                 // -2136 - This is shown on the project node between parentheses after the project name.
				|| propid == VSHPROPID_SharedItemsImportFullPaths // -2145
				|| propid == VSHPROPID_ProjectTreeCapabilities    // -2146 - MPF doesn't implement this, so we'll probably not implement it either
				|| propid == VSHPROPID_OneAppCapabilities         // -2149 - Virtually no info about this one. We probably don't need it.
				|| propid == VSHPROPID_SharedAssetsProject        // -2153
				|| propid == VSHPROPID_CanBuildQuickCheck         // -2156 - Much later, if ever
				|| propid == VSHPROPID_SupportsIconMonikers       // -2159
				|| propid == VSHPROPID_ProjectCapabilitiesChecker // -2173
				|| propid == -9089 // VSHPROPID_SlowEnumeration   // -9089
			)
				return E_NOTIMPL;

			RETURN_HR(E_NOTIMPL);
			#else
			return E_NOTIMPL;
			#endif
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->GetProperty(propid, pvar);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty (VSITEMID itemid, VSHPROPID propid, VARIANT var) override
	{
		if (_closed)
			return E_UNEXPECTED;

		#pragma region Some properties return the same thing for any node of our hierarchy
		if (propid == VSHPROPID_ParentHierarchy)
		{
			// TODO: need to advise hierarchy events also for the parent?
			RETURN_HR_IF(E_INVALIDARG, var.vt != VT_UNKNOWN);
			return var.punkVal->QueryInterface(&_parentHierarchy);
		}
		
		if (propid == VSHPROPID_ParentHierarchyItemid)
		{
			RETURN_HR_IF(E_INVALIDARG, var.vt != VT_VSITEMID);
			_parentHierarchyItemId = V_VSITEMID(&var);
			return S_OK;
		}
		#pragma endregion

		// https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.__vshpropid?view=visualstudiosdk-2017

		if (itemid == VSITEMID_ROOT)
		{
			HRESULT hr;
			if (propid == VSHPROPID_Caption)
			{
				RETURN_HR_IF(E_INVALIDARG, var.vt != VT_BSTR);
				_caption = wil::make_hlocal_string_nothrow(var.bstrVal); RETURN_IF_NULL_ALLOC(_caption);
				hr = S_OK;
			}
			else if (propid == VSHPROPID_ItemDocCookie) // -2034
			{
				RETURN_HR_IF(E_INVALIDARG, var.vt != VT_VSCOOKIE);
				_itemDocCookie = V_VSCOOKIE(&var);
				hr = S_OK;
			}
			else if (propid == VSHPROPID_EditLabel)
			{
				RETURN_HR_IF(E_INVALIDARG, var.vt != VT_BSTR);
				hr = RenameProject(var.bstrVal); RETURN_IF_FAILED_EXPECTED(hr);
			}
			else
			{
				#ifdef _DEBUG
				RETURN_HR(E_NOTIMPL);
				#else
				return E_NOTIMPL;
				#endif
			}

			if (SUCCEEDED(hr))
			{
				for (auto& s : _hierarchyEventSinks)
					s.second->OnPropertyChanged(itemid, propid, 0);
			}

			return hr;
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->SetProperty(propid, var);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE GetNestedHierarchy(VSITEMID itemid, REFIID iidHierarchyNested, void** ppHierarchyNested, VSITEMID* pitemidNested) override
	{
		// "If itemid is not a nested hierarchy, this method returns E_FAIL."
		return E_FAIL;
	}

	// A hierarchy is an object that contains many items, but does not necessarily contain an object
	// for each of those items. Thus, to get information about any of the hierarchy items, you need 
	// to query the hierarchy object for that information. The item identifier (itemid) is used to 
	// identify the requested item in that query. Using the GetCanonicalName method, you pass in the 
	// itemid and the canonical name is returned. The canonical name is a unique name used to distinguish 
	// a particular item in the hierarchy from every other item in the hierarchy.
	virtual HRESULT STDMETHODCALLTYPE GetCanonicalName(VSITEMID itemid, BSTR* pbstrName) override
	{
		if (_closed)
			return E_UNEXPECTED;

		// Returns a unique, string name for an item in the hierarchy.
		// Used for workspace persistence, such as remembering window positions.

		if (itemid == VSITEMID_ROOT)
		{
			wil::unique_cotaskmem_string buffer;
			DWORD unused;
			auto hr = GetCurFile (&buffer, &unused); RETURN_IF_FAILED(hr);
			*pbstrName = SysAllocString (buffer.get()); RETURN_IF_NULL_ALLOC(*pbstrName);
			return S_OK;
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->GetCanonicalName(pbstrName);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE ParseCanonicalName(LPCOLESTR pszName, VSITEMID* pitemid) override
	{
		// Note AGO: VS calls this function with full paths, with the call stack to this function showing
		// parameters with "mk" in their names. This happens for example while renaming a project, during the
		// call to OnAfterRenameProject.
		wil::unique_bstr projCN;
		auto hr = this->GetCanonicalName(VSITEMID_ROOT, &projCN); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr) && !_wcsicmp(projCN.get(), pszName))
		{
			*pitemid = VSITEMID_ROOT;
			return S_OK;
		}

		com_ptr<IProjectItem> c;
		hr = FindDescendant([pszName](IProjectItem* c)
			{
				wil::unique_bstr childCN;
				auto hr = c->GetCanonicalName(&childCN); RETURN_IF_FAILED(hr);
				return _wcsicmp(childCN.get(), pszName) ? S_FALSE : S_OK;
			}, &c);
		RETURN_IF_FAILED(hr);
		if (hr == S_OK)
		{
			*pitemid = c->GetItemId();
			return S_OK;
		}

		return E_FAIL;
	}

	virtual HRESULT STDMETHODCALLTYPE Unused0() override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseHierarchyEvents(IVsHierarchyEvents* pEventSink, VSCOOKIE* pdwCookie) override
	{
		bool inserted = _hierarchyEventSinks.try_insert ({ _nextHierarchyEventSinkCookie, pEventSink }); RETURN_HR_IF(E_OUTOFMEMORY, !inserted);
		*pdwCookie = _nextHierarchyEventSinkCookie;
		_nextHierarchyEventSinkCookie++;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseHierarchyEvents(VSCOOKIE dwCookie) override
	{
		auto it = _hierarchyEventSinks.find(dwCookie);
		if (it == _hierarchyEventSinks.end())
		{
			RETURN_HR(E_INVALIDARG);
		}

		_hierarchyEventSinks.erase(it);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Unused1() override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE Unused2() override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE Unused3() override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE Unused4() override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	HRESULT RenameProject (BSTR newName)
	{
		auto newNameLen = SysStringLen(newName); RETURN_HR_IF(E_INVALIDARG, newNameLen==0);

		// See if new name has an extension and that it is the right one.
		wil::unique_bstr defaultExt;
		auto hr = GetDefaultProjectFileExtension (&defaultExt); RETURN_IF_FAILED(hr);
		auto* newNameExt = wcsrchr(newName, '.');

		wil::unique_hlocal_string newFilename;
		wil::unique_hlocal_string newCaption;
		if(!newNameExt || _wcsicmp(newNameExt + 1, defaultExt.get()))
		{
			newCaption = wil::make_hlocal_string_nothrow(newName); RETURN_IF_NULL_ALLOC(newCaption);
			size_t newFilenameLen = newNameLen + 1 + SysStringLen(defaultExt.get());
			newFilename = wil::make_hlocal_string_nothrow(newName, newFilenameLen); RETURN_IF_NULL_ALLOC(newFilename);
			hr = StringCchCat(newFilename.get(), newFilenameLen + 1, L"."); RETURN_IF_FAILED(hr);
			hr = StringCchCat(newFilename.get(), newFilenameLen + 1, defaultExt.get()); RETURN_IF_FAILED(hr);
		}
		else
		{
			newCaption = wil::make_hlocal_string_nothrow(newName, newNameExt - newName); RETURN_IF_NULL_ALLOC(newCaption);
			newFilename = wil::make_hlocal_string_nothrow(newName, newNameLen); RETURN_IF_NULL_ALLOC(newFilename);
		}

		wil::unique_hlocal_string newFullPath;
		ULONG flags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
		hr = PathAllocCombine (_location.get(), newFilename.get(), flags, &newFullPath); RETURN_IF_FAILED(hr);

		// Check if it exists, but only if the user changed more than character casing.
		if (_wcsicmp(_filename.get(), newFilename.get()))
		{
			HANDLE hFile = ::CreateFile (newFullPath.get(), GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
			if (hFile != INVALID_HANDLE_VALUE)
			{
				CloseHandle(hFile);
				SetErrorInfo1 (HRESULT_FROM_WIN32(ERROR_FILE_EXISTS), IDS_RENAMEFILEALREADYEXISTS, newFullPath.get());
				return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
			}
		}

		BOOL renameCanContinue = FALSE;
		com_ptr<IVsSolution> solution;
		hr = serviceProvider->QueryService(SID_SVsSolution, &solution); RETURN_IF_FAILED(hr);
		wil::unique_bstr oldFullPath;
		hr = this->GetMkDocument(VSITEMID_ROOT, &oldFullPath); RETURN_IF_FAILED(hr);
		hr = solution->QueryRenameProject (this, oldFullPath.get(), newFullPath.get(), 0, &renameCanContinue);
		if(FAILED(hr) || !renameCanContinue)
			return OLE_E_PROMPTSAVECANCELLED;

		if (!::MoveFile (oldFullPath.get(), newFullPath.get()))
			RETURN_LAST_ERROR();

		// Bookkeeping time. We need to update our project name, our title, 
		// and force our rootnode to update its caption.
		_filename = std::move(newFilename);
		_caption = std::move(newCaption);

		// Make sure the property browser is updated
		com_ptr<IVsUIShell> uiShell;
		hr = serviceProvider->QueryService (SID_SVsUIShell, &uiShell); RETURN_IF_FAILED(hr);
		uiShell->RefreshPropertyBrowser(DISPID_VALUE); // return value ignored on purpose

		// Let the world know that the project is renamed. Solution needs to be told of rename. 
		hr = solution->OnAfterRenameProject (this, oldFullPath.get(), newFullPath.get(), 0); LOG_IF_FAILED(hr);

		for (auto& s : _hierarchyEventSinks)
		{
			s.second->OnPropertyChanged(VSITEMID_ROOT, VSHPROPID_Caption, 0);
			s.second->OnPropertyChanged(VSITEMID_ROOT, VSHPROPID_Name, 0);
			s.second->OnPropertyChanged(VSITEMID_ROOT, VSHPROPID_SaveName, 0);
			s.second->OnPropertyChanged(VSITEMID_ROOT, VSHPROPID_StateIconIndex, 0);
		}

		return S_OK;
	}

	#pragma region IVsUIHierarchy
	virtual HRESULT STDMETHODCALLTYPE QueryStatusCommand (VSITEMID itemid, const GUID* pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT* pCmdText) override
	{
		if (itemid == VSITEMID_ROOT)
			return QueryStatus (pguidCmdGroup, cCmds, prgCmds, pCmdText);

		RETURN_HR_IF_EXPECTED(OLECMDERR_E_NOTSUPPORTED, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(OLECMDERR_E_NOTSUPPORTED, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->QueryStatus (pguidCmdGroup, cCmds, prgCmds, pCmdText);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE ExecCommand (VSITEMID itemid, const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn, VARIANT* pvaOut) override
	{
		// Some commands such as UIHWCMDID_RightClick apply to items of all kinds.
		if (*pguidCmdGroup == GUID_VsUIHierarchyWindowCmds && nCmdID == UIHWCMDID_RightClick)
			return ShowContextMenu(itemid, pvaIn);

		if (itemid == VSITEMID_ROOT)
			return ExecInternal (pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

		// We need to return OLECMDERR_E_NOTSUPPORTED. If we return something else, VS won't like it.
		// For example, if the user selects two project items in the Solution Explorer and hits Escape,
		// and if we return E_NOTIMPL for VSITEMID_SELECTION, VS will say "operation cannot be completed".
		RETURN_HR_IF_EXPECTED(OLECMDERR_E_NOTSUPPORTED, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(OLECMDERR_E_NOTSUPPORTED, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->Exec (pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	HRESULT ShowContextMenu (VSITEMID itemid, const VARIANT* pvaIn)
	{
		POINTS pts;
		memcpy (&pts, &pvaIn->uintVal, 4);

		com_ptr<IVsUIShell> shell;
		auto hr = _sp->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);

		if (itemid == VSITEMID_ROOT)
		{
			return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_PROJNODE, pts, nullptr);
		}
		else if (itemid == VSITEMID_SELECTION)
		{
			// It seems if we have any kind of selection across multiple projects/hierachies in a solution,
			// the right-click is handled somewhere else, so we get here only with selections from 
			// a single project/hierarchy. This means GetCurrentSelection returns non-null in first parameter.
			com_ptr<IVsMonitorSelection> monsel;
			auto hr = serviceProvider->QueryService(SID_SVsShellMonitorSelection, &monsel); RETURN_IF_FAILED(hr);
			com_ptr<IVsHierarchy> hier;
			VSITEMID selitemid;
			com_ptr<IVsMultiItemSelect> mis;
			com_ptr<ISelectionContainer> sc;
			hr = monsel->GetCurrentSelection (&hier, &selitemid, &mis, &sc); RETURN_IF_FAILED(hr);
			RETURN_HR_IF_NULL_EXPECTED(E_NOTIMPL, hier);
			RETURN_HR_IF_EXPECTED(E_NOTIMPL, selitemid != VSITEMID_SELECTION);
			ULONG cItems;
			BOOL singleHier;
			hr = mis->GetSelectionInfo(&cItems, &singleHier); RETURN_IF_FAILED(hr);
			RETURN_HR_IF_EXPECTED(E_NOTIMPL, !singleHier);
			auto items = wil::make_unique_nothrow<VSITEMSELECTION[]>(cItems); RETURN_IF_NULL_ALLOC(items);
			hr = mis->GetSelectedItems (GSI_fOmitHierPtrs, cItems, items.get()); RETURN_IF_FAILED(hr);
			bool projectNodeIncluded = false;
			ULONG fileNodesIncluded = 0;
			for (ULONG i = 0; i < cItems; i++)
			{
				if (items[i].itemid == VSITEMID_ROOT)
					projectNodeIncluded = true;
				else
				{
					if (auto d = FindDescendant(items[i].itemid))
					{
						if (wil::try_com_query_nothrow<IProjectFile>(d))
							fileNodesIncluded++;
					}
				}
			}

			if (projectNodeIncluded && !fileNodesIncluded)
				return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_PROJNODE, pts, nullptr);
			else if (projectNodeIncluded && fileNodesIncluded)
				return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_XPROJ_PROJITEM, pts, nullptr);
			else if (!projectNodeIncluded && fileNodesIncluded)
				return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_XPROJ_MULTIITEM, pts, nullptr);
			return E_NOTIMPL;
		}
		else
		{
			if (auto d = com_ptr(FindDescendant(itemid)))
			{
				if (auto file = wil::try_com_query_nothrow<IProjectFile>(d))
					return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_ITEMNODE, pts, nullptr);

				return E_NOTIMPL;
			}
			else
				return E_INVALIDARG;
		}
	}
	#pragma endregion

	#pragma region IPersistFileFormat
	virtual HRESULT STDMETHODCALLTYPE GetClassID(CLSID* pClassID) override
	{
		*pClassID = __uuidof(IZ80ProjectProperties);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE IsDirty(BOOL* pfIsDirty) override
	{
		*pfIsDirty = _isDirty ? TRUE : FALSE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE InitNew(DWORD nFormatIndex) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Load(LPCOLESTR pszFilename, DWORD grfMode, BOOL fReadOnly) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Save(LPCOLESTR pszFilename, BOOL fRemember, DWORD nFormatIndex) override
	{
		HRESULT hr;

		RETURN_HR_IF(E_NOTIMPL, nFormatIndex != DEF_FORMAT_INDEX);

		if (!pszFilename)
		{
			// "Save" operation (save under existing name)
			RETURN_HR_IF(E_INVALIDARG, !_filename || !_filename.get()[0]);
			RETURN_HR_IF(E_INVALIDARG, !_location || !_location.get()[0]);
			wil::unique_hlocal_string filePath;
			ULONG flags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
			hr = PathAllocCombine(_location.get(), _filename.get(), flags, &filePath); RETURN_IF_FAILED(hr);
			_noScribble = true;

			// TODO: pause monitoring file change notifications
			//CSuspendFileChanges suspendFileChanges(CString(pszFileName), TRUE);
			com_ptr<IStream> stream;
			hr = SHCreateStreamOnFile(filePath.get(), STGM_WRITE | STGM_SHARE_DENY_WRITE, &stream); RETURN_IF_FAILED(hr);
			hr = SaveToXml(this, ProjectElementName, stream); RETURN_IF_FAILED(hr);
			// TODO: resume monitoring file change notifications

			_isDirty = false; // TODO: notify property changes
		}
		else
		{
			// "Save As" or "Save A Copy As"
			if (PathIsRelative(pszFilename))
				RETURN_HR(E_NOTIMPL);
			_noScribble = true;
			com_ptr<IStream> stream;
			hr = SHCreateStreamOnFile(pszFilename, STGM_WRITE | STGM_SHARE_DENY_WRITE, &stream); RETURN_IF_FAILED(hr);
			hr = SaveToXml(this, ProjectElementName, stream.get()); RETURN_IF_FAILED(hr);

			// If pszFileName is null, the implementation ignores the fRemember flag.
			if (fRemember)
			{
				// "Save As"
				auto filePart = PathFindFileName(pszFilename);
				_location = wil::make_hlocal_string_nothrow(pszFilename, filePart - pszFilename); RETURN_IF_NULL_ALLOC(_location);
				_filename = wil::make_hlocal_string_nothrow(filePart); RETURN_IF_NULL_ALLOC(_filename);
				_isDirty = false;
				// TODO: notify property changes for filename and isDirty.
			}
			else
			{
				// "Save A Copy As"
			}
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SaveCompleted(LPCOLESTR pszFilename) override
	{
		WI_ASSERT(_noScribble);
		_noScribble = false;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetCurFile(LPOLESTR* ppszFilename, DWORD* pnFormatIndex) override
	{
		// Note that we return true for VSHPROPID_MonikerSameAsPersistFile, meaning that
		// IVsProject::GetMkDocument returns the same as GetCurFile for VSITEMID_ROOT.

		size_t reservedLen = wcslen(_location.get()) + wcslen(_filename.get()) + 10;

		auto buffer = wil::make_cotaskmem_string_nothrow(nullptr, reservedLen); RETURN_IF_NULL_ALLOC(buffer);
		ULONG flags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
		auto hr = PathCchCombineEx (buffer.get(), reservedLen, _location.get(), _filename.get(), flags); RETURN_IF_FAILED(hr);

		*ppszFilename = buffer.release();
		*pnFormatIndex = 0;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetFormatList(LPOLESTR* ppszFormatList) override
	{
		// Keep this in sync with DisplayProjectFileExtensions from the .pkgdef file.
		static constexpr wchar_t list[] = L"Z80 Project Files (*.flx)\n*.flx\n";
		*ppszFormatList = (wchar_t*)CoTaskMemAlloc(sizeof(list));
		if (!*ppszFormatList)
			return E_OUTOFMEMORY;
		wcscpy_s (*ppszFormatList, _countof(list), list);
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsProject3
	virtual HRESULT STDMETHODCALLTYPE IsDocumentInProject(LPCOLESTR pszMkDocument, BOOL* pfFound, VSDOCUMENTPRIORITY* pdwPriority, VSITEMID* pitemid) override
	{
		RETURN_HR_IF(E_POINTER, !pszMkDocument || !pfFound || !pdwPriority || !pitemid);

		bool isRelative = PathIsRelative(pszMkDocument);

		wil::com_ptr_nothrow<IProjectItem> c;
		auto hr = FindDescendant(
			[pszMkDocument, isRelative](IProjectItem* c)
			{
				wil::unique_bstr path;
				if (isRelative)
				{
					auto hr = c->GetCanonicalName(&path); RETURN_IF_FAILED(hr);
				}
				else
				{
					auto hr = c->GetMkDocument(&path); RETURN_IF_FAILED(hr);
				}

				return wcscmp(pszMkDocument, path.get()) ? S_FALSE : S_OK;
			}, &c); RETURN_IF_FAILED(hr);

		if (hr == S_FALSE)
		{
			*pfFound = FALSE;
			return hr;
		}

		*pfFound = TRUE;
		*pdwPriority = DP_Standard;
		*pitemid = c->GetItemId();
		return S_OK;
	}

	// File-based project types must return the path from this method.
	virtual HRESULT STDMETHODCALLTYPE GetMkDocument (VSITEMID itemid, BSTR* pbstrMkDocument) override
	{
		if (itemid == VSITEMID_ROOT)
		{
			// Note that we return true for VSHPROPID_MonikerSameAsPersistFile, meaning that
			// IVsProject::GetMkDocument returns the same as GetCurFile for VSITEMID_ROOT.

			wil::unique_cotaskmem_string buffer;
			DWORD unused;
			auto hr = GetCurFile (&buffer, &unused); RETURN_IF_FAILED(hr);
			*pbstrMkDocument = SysAllocString (buffer.get()); RETURN_IF_NULL_ALLOC(*pbstrMkDocument);
			return S_OK;
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->GetMkDocument(pbstrMkDocument);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE OpenItem(VSITEMID itemid, REFGUID rguidLogicalView, IUnknown* punkDocDataExisting, IVsWindowFrame** ppWindowFrame) override
	{
		auto d = FindDescendant(itemid); RETURN_HR_IF(E_INVALIDARG, d == nullptr);

		wil::unique_bstr mkDocument;
		auto hr = d->GetMkDocument(&mkDocument); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IVsUIShellOpenDocument> uiShellOpenDocument;
		hr = _sp->QueryService(SID_SVsUIShellOpenDocument, &uiShellOpenDocument); RETURN_IF_FAILED(hr);

		hr = uiShellOpenDocument->OpenStandardEditor(OSE_ChooseBestStdEditor, mkDocument.get(), rguidLogicalView,
			L"%3", this, itemid, punkDocDataExisting, _sp.get(), ppWindowFrame); RETURN_IF_FAILED_EXPECTED(hr);

		wil::unique_variant var;
		hr = (*ppWindowFrame)->GetProperty(VSFPROPID_DocCookie, &var); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr) && (var.vt == VT_VSCOOKIE) && (V_VSCOOKIE(&var) != VSDOCCOOKIE_NIL))
		{
			hr = d->SetProperty (VSHPROPID_ItemDocCookie, var); LOG_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetItemContext(VSITEMID itemid, IServiceProvider** ppSP) override
	{
		*ppSP = _sp.get();
		_sp->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GenerateUniqueItemName(VSITEMID itemidLoc, LPCOLESTR pszExt, LPCOLESTR pszSuggestedRoot, BSTR* pbstrItemName) override
	{
		HRESULT hr;

		if (itemidLoc != VSITEMID_ROOT)
			RETURN_HR(E_NOTIMPL);

		for (uint32_t i = 0; i < 1000; i++)
		{
			wchar_t buffer[50];
			swprintf_s(buffer, L"%s%u%s", pszSuggestedRoot, i, pszExt);
			wil::unique_hlocal_string dest;
			hr = PathAllocCombine (_location.get(), buffer, PATHCCH_ALLOW_LONG_PATHS, &dest); RETURN_IF_FAILED(hr);
			if (!PathFileExists(dest.get()))
			{
				*pbstrItemName = SysAllocString(buffer); RETURN_IF_NULL_ALLOC(*pbstrItemName);
				return S_OK;
			}
		}

		RETURN_HR(E_FAIL);
	}

	HRESULT AddNewFile (VSITEMID itemidLoc, LPCTSTR pszFullPathSource, LPCTSTR pszNewFileName, IProjectItem** ppNewNode)
	{
		RETURN_HR_IF_NULL(E_INVALIDARG, pszFullPathSource);
		RETURN_HR_IF_NULL(E_INVALIDARG, pszNewFileName);
		RETURN_HR_IF_NULL(E_POINTER, ppNewNode);

		wil::unique_hlocal_string dest;
		DWORD flags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
		auto hr = PathAllocCombine (_location.get(), pszNewFileName, flags, &dest); RETURN_IF_FAILED(hr);
		COPYFILE2_EXTENDED_PARAMETERS params = { 
			.dwSize = (DWORD)sizeof(params), 
			.dwCopyFlags = COPY_FILE_ALLOW_DECRYPTED_DESTINATION | COPY_FILE_FAIL_IF_EXISTS
		};
		hr = CopyFile2 (pszFullPathSource, dest.get(), &params); RETURN_IF_FAILED_EXPECTED(hr);
		// TODO: if the file exists, ask user whether to overwrite

		// template was read-only, but our file should not be
		SetFileAttributes(dest.get(), FILE_ATTRIBUTE_ARCHIVE);

		return AddExistingFile (itemidLoc, dest.get(), ppNewNode);
	}

	HRESULT AddExistingFile (VSITEMID itemidLoc, LPCTSTR pszFullPathSource, IProjectItem** ppNewFile, BOOL fSilent = FALSE, BOOL fLoad = FALSE)
	{
		HRESULT hr;

		// Check if the item exists in the project already.
		wchar_t relativeUgly[MAX_PATH];
		BOOL bRes = PathRelativePathTo (relativeUgly, _location.get(), FILE_ATTRIBUTE_DIRECTORY, pszFullPathSource, 0);
		if (!bRes)
			return SetErrorInfo(E_INVALIDARG, L"Can't make a relative path from '%s' relative to '%s'.", pszFullPathSource, _location.get());
		wil::unique_hlocal_string relative;
		hr = PathAllocCanonicalize (relativeUgly, 0, relative.addressof());
		if (FAILED(hr))
			return SetErrorInfo(hr, L"Can't make a relative path from '%s' relative to '%s'.", pszFullPathSource, _location.get());

		com_ptr<IProjectItem> existing;
		hr = FindDescendant([relative=relative.get()](IProjectItem* c)
			{
				if (auto sf = wil::try_com_query_nothrow<IProjectFile>(c))
				{
					wil::unique_bstr rel;
					auto hr = sf->get_Path(&rel); RETURN_IF_FAILED(hr);
					if (!_wcsicmp(relative, rel.get()))
						return S_OK;
					else
						return S_FALSE;
				}

				return S_FALSE;
			}, &existing); RETURN_IF_FAILED(hr);
		if (hr == S_OK)
			return SetErrorInfo(HRESULT_FROM_WIN32(ERROR_FILE_EXISTS), L"File already in project:\r\n\r\n%s", pszFullPathSource);

		com_ptr<IProjectFile> file;
		VSITEMID itemId = _nextFileItemId++;
		hr = MakeProjectFile (itemId, this, itemidLoc, &file); RETURN_IF_FAILED(hr);
		auto path = wil::make_bstr_nothrow(pszFullPathSource); RETURN_IF_NULL_ALLOC(path);
		hr = file->put_Path(relative.get()); RETURN_IF_FAILED(hr);
		auto buildTool = _wcsicmp(PathFindExtension(path.get()), L".asm") ? BuildToolKind::None : BuildToolKind::Assembler;
		hr = file->put_BuildTool(buildTool); RETURN_IF_FAILED(hr);

		if (!_firstChild)
			_firstChild = file;
		else
		{
			IProjectItem* last = _firstChild;
			while(last->Next())
				last = last->Next();
			last->SetNext(file);
		}

		_isDirty = true;

		for (auto& sink : _hierarchyEventSinks)
			sink.second->OnItemsAppended(itemidLoc);

		com_ptr<IVsWindowFrame> frame;
		hr = this->OpenItem(itemId, LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &frame);
		if (SUCCEEDED(hr))
			frame->Show();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE AddItem(VSITEMID itemidLoc, VSADDITEMOPERATION dwAddItemOperation, LPCOLESTR pszItemName, ULONG cFilesToOpen, LPCOLESTR rgpszFilesToOpen[], HWND hwndDlgOwner, VSADDRESULT* pResult) override
	{
		HRESULT hr;

		if (itemidLoc != VSITEMID_ROOT)
			RETURN_HR(E_NOTIMPL);

		switch(dwAddItemOperation)
		{
			case VSADDITEMOP_CLONEFILE:
			{
				// Add New File
				RETURN_HR_IF(E_INVALIDARG, cFilesToOpen != 1);
				com_ptr<IProjectItem> pNewNode;
				hr = AddNewFile (itemidLoc, rgpszFilesToOpen[0], pszItemName, &pNewNode); RETURN_IF_FAILED_EXPECTED(hr);
				*pResult = ADDRESULT_Success;
				return S_OK;
			}

			case VSADDITEMOP_LINKTOFILE:
				// Because we are a reference-based project system our handling for LINKTOFILE is the same as OPENFILE.
				// A storage-based project system which handles OPENFILE by copying the file into the project directory
				// would have distinct handling for LINKTOFILE vs. OPENFILE.
			case VSADDITEMOP_OPENFILE:
			{
				// Add Existing File
				for (DWORD i = 0; i < cFilesToOpen; i++)
				{
					com_ptr<IProjectItem> pNewNode;
					hr = AddExistingFile(itemidLoc, rgpszFilesToOpen[i], &pNewNode); RETURN_IF_FAILED_EXPECTED(hr);
				}

				*pResult = ADDRESULT_Success;
				return hr;
			}

			default:
				RETURN_HR(E_NOTIMPL);
		}
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveItem(DWORD dwReserved, VSITEMID itemid, BOOL* pfResult) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE ReopenItem(VSITEMID itemid, REFGUID rguidEditorType, LPCOLESTR pszPhysicalView, REFGUID rguidLogicalView, IUnknown* punkDocDataExisting, IVsWindowFrame** ppWindowFrame) override
	{
		auto d = FindDescendant(itemid); RETURN_HR_IF(E_INVALIDARG, d == nullptr);

		wil::unique_bstr mkDocument;
		auto hr = d->GetMkDocument(&mkDocument); RETURN_IF_FAILED(hr);

		com_ptr<IVsUIShellOpenDocument> uiShellOpenDocument;
		hr = _sp->QueryService(SID_SVsUIShellOpenDocument, &uiShellOpenDocument); RETURN_IF_FAILED(hr);

		hr = uiShellOpenDocument->OpenSpecificEditor(0, mkDocument.get(), rguidEditorType, pszPhysicalView, rguidLogicalView,
			L"%3", this, itemid, punkDocDataExisting, _sp.get(), ppWindowFrame); RETURN_IF_FAILED_EXPECTED(hr);

		wil::unique_variant var;
		hr = (*ppWindowFrame)->GetProperty(VSFPROPID_DocCookie, &var); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr) && (var.vt == VT_VSCOOKIE) && (V_VSCOOKIE(&var) != VSDOCCOOKIE_NIL))
		{
			hr = d->SetProperty (VSHPROPID_ItemDocCookie, var); LOG_IF_FAILED(hr);
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IVsGetCfgProvider
	virtual HRESULT STDMETHODCALLTYPE GetCfgProvider (IVsCfgProvider** ppCfgProvider) override
	{
		*ppCfgProvider = this;
		AddRef();
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsCfgProvider
	virtual HRESULT STDMETHODCALLTYPE GetCfgs (ULONG celt, IVsCfg *rgpcfg[], ULONG* pcActual, VSCFGFLAGS *prgfFlags) override
	{
		if (celt == 0)
		{
			*pcActual = _configs.size();
			return S_OK;
		}

		ULONG& i = *pcActual;
		for (i = 0; i < std::min(celt, (ULONG)_configs.size()); i++)
		{
			auto hr = _configs[i]->QueryInterface(&rgpcfg[i]); RETURN_IF_FAILED(hr);
			if (prgfFlags)
				prgfFlags[i] = 0;
		}

		return (i < celt) ? S_FALSE : S_OK;
	}
	#pragma endregion

	#pragma region IProvideClassInfo
	virtual HRESULT STDMETHODCALLTYPE GetClassInfo (ITypeInfo **ppTI) override
	{
		return GetTypeInfo(0, 0, ppTI);
	}
	#pragma endregion

	#pragma region IOleCommandTarget
	virtual HRESULT STDMETHODCALLTYPE QueryStatus (const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[],OLECMDTEXT *pCmdText) override
	{
		if (!pguidCmdGroup)
			return E_POINTER;

		for (ULONG i = 0; i < cCmds; i++)
		{
			auto hr = this->QueryStatusInternal (pguidCmdGroup, prgCmds[i].cmdID, &prgCmds[i].cmdf, pCmdText ? &pCmdText[i] : nullptr);
			if (FAILED(hr))
				return hr;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Exec (const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) override
	{
		if (!pguidCmdGroup)
			return E_POINTER;

		return ExecInternal (pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
	}
	#pragma endregion
	/*
	#pragma region IConnectionPointContainer
	virtual HRESULT STDMETHODCALLTYPE EnumConnectionPoints (IEnumConnectionPoints **ppEnum) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE FindConnectionPoint (REFIID riid, IConnectionPoint **ppCP) override
	{
		if (riid == IID_IPropertyNotifySink)
		{
			BreakIntoDebugger();
			return E_NOTIMPL;
		}

		BreakIntoDebugger();
		return E_NOTIMPL;
	}
	#pragma endregion
	*/

	#pragma region IVsCfgProvider2
	virtual HRESULT STDMETHODCALLTYPE GetCfgNames(ULONG celt, BSTR rgbstr[], ULONG* pcActual) override
	{
		vector_nothrow<wil::unique_bstr> names;
		for (auto& c : _configs)
		{
			wil::unique_bstr n;
			auto hr = c->get_ConfigName(&n); RETURN_IF_FAILED(hr);
			auto it = names.find_if([n=n.get()](auto& en) { return !wcscmp(en.get(), n); });
			if (it == names.end())
			{
				bool pushed = names.try_push_back(std::move(n)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
			}
		}

		if (celt == 0)
		{
			*pcActual = names.size();
			return S_OK;
		}

		ULONG& i = *pcActual;
		for (i = 0; i < std::min(celt, (ULONG)names.size()); i++)
			rgbstr[i] = names[i].release();

		return (i < celt) ? S_FALSE : S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetPlatformNames(ULONG celt, BSTR rgbstr[], ULONG* pcActual) override
	{
		vector_nothrow<wil::unique_bstr> names;
		for (auto& c : _configs)
		{
			wil::unique_bstr n;
			auto hr = c->get_PlatformName(&n); RETURN_IF_FAILED(hr);
			auto it = names.find_if([n=n.get()](auto& en) { return !wcscmp(en.get(), n); });
			if (it == names.end())
			{
				bool pushed = names.try_push_back(std::move(n)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
			}
		}
		if (celt == 0)
		{
			*pcActual = names.size();
			return S_OK;
		}

		ULONG& i = *pcActual;
		for (i = 0; i < std::min(celt, (ULONG)names.size()); i++)
			rgbstr[i] = names[i].release();

		return (i < celt) ? S_FALSE : S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetCfgOfName(LPCOLESTR pszCfgName, LPCOLESTR pszPlatformName, IVsCfg** ppCfg) override
	{
		for (auto& cfg : _configs)
		{
			wil::unique_bstr cn, pn;
			auto hr = cfg->get_ConfigName(&cn); RETURN_IF_FAILED(hr);
			hr = cfg->get_PlatformName(&pn); RETURN_IF_FAILED(hr);
			if (!wcscmp(pszCfgName, cn.get()) && !wcscmp(pszPlatformName, pn.get()))
			{
				auto hr = cfg->QueryInterface(ppCfg); RETURN_IF_FAILED(hr);
				return S_OK;
			}
		}

		RETURN_HR(E_INVALIDARG);
	}

	virtual HRESULT STDMETHODCALLTYPE AddCfgsOfCfgName(LPCOLESTR pszCfgName, LPCOLESTR pszCloneCfgName, BOOL fPrivate) override
	{
		com_ptr<IProjectConfig> newConfig;
		auto hr = ProjectConfig_CreateInstance (this, &newConfig); RETURN_IF_FAILED_EXPECTED(hr);

		if (pszCloneCfgName)
		{
			com_ptr<IProjectConfig> existing;
			for (auto& c : _configs)
			{
				wil::unique_bstr name;
				hr = c->get_ConfigName(&name); RETURN_IF_FAILED(hr);
				if (wcscmp(name ? name.get() : L"", pszCloneCfgName) == 0)
				{
					existing = c;
					break;
				}
			}

			RETURN_HR_IF(E_INVALIDARG, !existing);

			auto stream = com_ptr(SHCreateMemStream(nullptr, 0)); RETURN_IF_NULL_ALLOC(stream);
		
			hr = SaveToXml(existing, L"Temp", stream); RETURN_IF_FAILED_EXPECTED(hr);

			hr = stream->Seek({ 0 }, STREAM_SEEK_SET, nullptr); RETURN_IF_FAILED(hr);

			hr = LoadFromXml(newConfig, L"Temp", stream); RETURN_IF_FAILED_EXPECTED(hr);
		}

		auto newConfigNameBstr = wil::make_bstr_nothrow(pszCfgName); RETURN_IF_NULL_ALLOC(newConfigNameBstr);
		hr = newConfig->put_ConfigName(newConfigNameBstr.get()); RETURN_IF_FAILED(hr);

		bool pushed = _configs.try_push_back(std::move(newConfig)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		auto removeBack = wil::scope_exit([this] { _configs.remove_back(); });

		_isDirty = true;

		for (auto& sink : _cfgProviderEventSinks)
			sink.second->OnCfgNameAdded(pszCfgName);

		removeBack.release();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE DeleteCfgsOfCfgName(LPCOLESTR pszCfgName) override
	{
		HRESULT hr;

		auto it = _configs.begin();
		while (it != _configs.end())
		{
			wil::unique_bstr name;
			hr = it->get()->get_ConfigName(&name); RETURN_IF_FAILED(hr);
			if (wcscmp(name ? name.get() : L"", pszCfgName) == 0)
				break;
			it++;
		}

		if (it == _configs.end())
		{
			// When the user deletes a solution configuration, VS asks us to delete
			// all project configs with that name, even if we have no such project config.
			return E_INVALIDARG;
		}

		_configs.erase(it);

		_isDirty = true;

		for (auto& sink : _cfgProviderEventSinks)
			sink.second->OnCfgNameDeleted(pszCfgName);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE RenameCfgsOfCfgName(LPCOLESTR pszOldName, LPCOLESTR pszNewName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE AddCfgsOfPlatformName(LPCOLESTR pszPlatformName, LPCOLESTR pszClonePlatformName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE DeleteCfgsOfPlatformName(LPCOLESTR pszPlatformName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetSupportedPlatformNames(ULONG celt, BSTR rgbstr[], ULONG* pcActual) override
	{
		return GetPlatformNames(celt, rgbstr, pcActual);
	}

	virtual HRESULT STDMETHODCALLTYPE GetCfgProviderProperty(VSCFGPROPID propid, VARIANT* pvar) override
	{
		if (propid == VSCFGPROPID_SupportsCfgAdd)
			return InitVariantFromBoolean(TRUE, pvar);
		if (propid == VSCFGPROPID_SupportsCfgDelete)
			return InitVariantFromBoolean(TRUE, pvar);
		if (propid == VSCFGPROPID_SupportsCfgRename)
			return InitVariantFromBoolean(TRUE, pvar);
		if (propid == VSCFGPROPID_SupportsPlatformAdd)
			return InitVariantFromBoolean(FALSE, pvar);
		if (propid == VSCFGPROPID_SupportsPlatformDelete)
			return InitVariantFromBoolean(FALSE, pvar);

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseCfgProviderEvents(IVsCfgProviderEvents* pCPE, VSCOOKIE* pdwCookie) override
	{
		bool inserted = _cfgProviderEventSinks.try_insert({ _nextCfgProviderEventCookie, pCPE }); RETURN_HR_IF(E_OUTOFMEMORY, !inserted);
		*pdwCookie = _nextCfgProviderEventCookie;
		_nextCfgProviderEventCookie++;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseCfgProviderEvents(VSCOOKIE dwCookie) override
	{
		auto it = _cfgProviderEventSinks.find(dwCookie);
		if (it == _cfgProviderEventSinks.end())
			RETURN_HR(E_INVALIDARG);

		_cfgProviderEventSinks.erase(it);
		return S_OK;
	}
	#pragma endregion
	/*
	#pragma region IVsProjectBuildSystem
	virtual HRESULT STDMETHODCALLTYPE SetHostObject(LPCOLESTR pszTargetName, LPCOLESTR pszTaskName, IUnknown* punkHostObject) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE StartBatchEdit(void) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE EndBatchEdit(void) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE CancelBatchEdit(void) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE BuildTarget(LPCOLESTR pszTargetName, VARIANT_BOOL* pbSuccess) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetBuildSystemKind(BuildSystemKindFlags* pBuildSystemKind) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IVsBuildPropertyStorage
	virtual HRESULT STDMETHODCALLTYPE GetPropertyValue(LPCOLESTR pszPropName, LPCOLESTR pszConfigName, PersistStorageType storage, BSTR* pbstrPropValue) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE SetPropertyValue(LPCOLESTR pszPropName, LPCOLESTR pszConfigName, PersistStorageType storage, LPCOLESTR pszPropValue) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveProperty(LPCOLESTR pszPropName, LPCOLESTR pszConfigName, PersistStorageType storage) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetItemAttribute(VSITEMID item, LPCOLESTR pszAttributeName, BSTR* pbstrAttributeValue) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE SetItemAttribute(VSITEMID item, LPCOLESTR pszAttributeName, LPCOLESTR pszAttributeValue) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IVsBuildPropertyStorage2
	virtual HRESULT STDMETHODCALLTYPE SetPropertyValueEx(LPCOLESTR pszPropName, LPCOLESTR pszPropertyGroupCondition, PersistStorageType storage, LPCOLESTR pszPropValue) override
	{
		BreakIntoDebugger();
		return E_NOTIMPL;
	}
	#pragma endregion
	*/
	#pragma region IVsHierarchyDeleteHandler3
	virtual HRESULT STDMETHODCALLTYPE QueryDeleteItems (
		ULONG cItems,
		VSDELETEITEMOPERATION dwDelItemOp,
		__RPC__in_ecount_full(cItems) VSITEMID itemid[  ],
		__RPC__inout_ecount_full(cItems) VARIANT_BOOL pfCanDelete[  ]) override
	{
		for (ULONG i = 0; i < cItems; i++)
			pfCanDelete[i] = VARIANT_TRUE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE DeleteItems (
		ULONG cItems,
		VSDELETEITEMOPERATION dwDelItemOp,
		__RPC__in_ecount_full(cItems) VSITEMID itemid[  ],
		VSDELETEHANDLEROPTIONS dwFlags) override
	{
		if ((dwDelItemOp != DELITEMOP_RemoveFromProject) && (dwDelItemOp != DELITEMOP_DeleteFromStorage))
			RETURN_HR(E_ABORT);

		for (ULONG i = 0; i < cItems; i++)
			RETURN_HR_IF(E_NOTIMPL, itemid[i] == VSITEMID_ROOT); // see example in PrjHier.cpp

		com_ptr<IVsUIShell> shell;
		auto hr = _sp->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);

		com_ptr<IVsSolution> solution;
		hr = _sp->QueryService(SID_SVsSolution, &solution); RETURN_IF_FAILED(hr);

		com_ptr<IVsRunningDocumentTable> rdt;
		hr = _sp->QueryService(SID_SVsRunningDocumentTable, &rdt); RETURN_IF_FAILED(hr);

		HWND ownerHwnd;
		hr = shell->GetDialogOwnerHwnd(&ownerHwnd); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IFileOperation> pfo;
		hr = CoCreateInstance(__uuidof(FileOperation), NULL, CLSCTX_ALL, IID_PPV_ARGS(&pfo)); RETURN_IF_FAILED(hr);
		hr = pfo->SetOwnerWindow(ownerHwnd); RETURN_IF_FAILED(hr);

		for (ULONG i = 0; i < cItems; i++)
		{
			com_ptr<IProjectItem> prevSibling;
			com_ptr<IProjectItemParent> parent;
			auto d = FindDescendant(itemid[i], &prevSibling, &parent); RETURN_HR_IF_NULL(E_INVALIDARG, d);

			wil::unique_bstr mk;
			hr = d->GetMkDocument(&mk); RETURN_IF_FAILED(hr);

			com_ptr<IVsHierarchy> docHier;
			VSITEMID docItemId;
			com_ptr<IUnknown> docData;
			VSCOOKIE docCookie;
			hr = rdt->FindAndLockDocument (RDT_NoLock, (LPCOLESTR)mk.get(), &docHier, &docItemId, &docData, &docCookie); LOG_IF_FAILED(hr);
			if (hr == S_OK && docHier.get() == this && docItemId == itemid[i])
			{
				auto slnCloseOpts = (dwDelItemOp == DELITEMOP_DeleteFromStorage) ? SLNSAVEOPT_NoSave : SLNSAVEOPT_PromptSave;
				hr = solution->CloseSolutionElement (slnCloseOpts, nullptr, docCookie); RETURN_IF_FAILED_EXPECTED(hr);
			}

			wil::unique_variant firstChild;
			hr = d->GetProperty(VSHPROPID_FirstChild, &firstChild); RETURN_IF_FAILED(hr);
			if (V_VSITEMID(&firstChild) != VSITEMID_NIL)
			{
				if (!(dwFlags & DHO_SUPPRESS_UI))
				{
					LONG res;
					hr = shell->ShowMessageBox(0, CLSID_NULL, (LPOLESTR)L"Not Implemented", (LPOLESTR)L"Deleting parent nodes not yet implemented",
						nullptr, 0, OLEMSGBUTTON_OK, OLEMSGDEFBUTTON_FIRST, OLEMSGICON_WARNING, FALSE, &res); RETURN_IF_FAILED(hr);
				}
				return S_FALSE;
			}
			auto next = com_ptr(d->Next());
			d->SetNext(nullptr);
			hr = d->SetProperty(VSHPROPID_Parent, MakeVariantFromVSITEMID(VSITEMID_NIL)); RETURN_IF_FAILED(hr);

			if (dwDelItemOp == DELITEMOP_DeleteFromStorage)
			{
				com_ptr<IShellItem> si;
				hr = SHCreateItemFromParsingName(mk.get(), nullptr, IID_PPV_ARGS(&si));
				if (SUCCEEDED(hr))
				{
					hr = pfo->DeleteItem(si, nullptr);
					if (SUCCEEDED(hr))
						hr = pfo->PerformOperations();
				}
			}

			if (!prevSibling)
				// removing the first item in a list
				parent->SetFirstChild(next);
			else
				// removing item that's not first
				prevSibling->SetNext(next);

			_isDirty = true;

			for (auto& sink : _hierarchyEventSinks)
				sink.second->OnItemDeleted(d->GetItemId());
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IVsPersistHierarchyItem
	virtual HRESULT STDMETHODCALLTYPE IsItemDirty (VSITEMID itemid, IUnknown *punkDocData, BOOL *pfDirty) override
	{
		if (itemid == VSITEMID_ROOT)
		{
			wil::com_ptr_nothrow<IVsPersistDocData> docData;
			auto hr = punkDocData->QueryInterface(&docData); RETURN_IF_FAILED(hr);
			hr = docData->IsDocDataDirty(pfDirty); RETURN_IF_FAILED(hr);
			return S_OK;
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		if (auto d = FindDescendant(itemid))
			return d->IsItemDirty(punkDocData, pfDirty);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE SaveItem (VSSAVEFLAGS dwSave, LPCOLESTR pszSilentSaveAsName, VSITEMID itemid, IUnknown *punkDocData, BOOL *pfCanceled) override
	{
		// https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.ivspersisthierarchyitem.saveitem?view=visualstudiosdk-2022

		//if (itemid == VSITEMID_ROOT)
		//{
			if (dwSave == VSSAVE_Save || dwSave == VSSAVE_SaveAs)
			{
				wil::com_ptr_nothrow<IVsPersistDocData> docData;
				auto hr = punkDocData->QueryInterface(&docData);
				if (FAILED(hr))
					return hr;
				wil::unique_bstr mkNew;
				return docData->SaveDocData(dwSave, &mkNew, pfCanceled);
			}

			if (dwSave == VSSAVE_SilentSave)
			{
				wil::com_ptr_nothrow<IPersistFileFormat> ff;
				auto hr = punkDocData->QueryInterface(&ff);
				if (FAILED(hr))
					return hr;

				wil::com_ptr_nothrow<IVsUIShell> shell;
				hr = _sp->QueryService(SID_SVsUIShell, &shell);
				if (FAILED(hr))
					return hr;

				wil::unique_bstr docNew;
				return shell->SaveDocDataToFile (VSSAVE_SilentSave, ff.get(), pszSilentSaveAsName, &docNew, pfCanceled);
			}

			RETURN_HR(E_NOTIMPL);
		//}

		//RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IZ80ProjectProperties
	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		auto ext = PathFindExtension(_filename.get());
		*value = SysAllocStringLen(_filename.get(), (UINT)(ext - _filename.get())); RETURN_IF_NULL_ALLOC(*value);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Configurations (SAFEARRAY** configs) override
	{
		*configs = nullptr;
		SAFEARRAYBOUND bound;
		bound.cElements = (ULONG)_configs.size();
		bound.lLbound = 0;
		auto sa = unique_safearray(SafeArrayCreate(VT_UNKNOWN, 1, &bound)); RETURN_HR_IF(E_OUTOFMEMORY, !sa);
		for (LONG i = 0; i < (LONG)bound.cElements; i++)
		{
			auto hr = SafeArrayPutElement(sa.get(), &i, _configs[i].get()); RETURN_IF_FAILED(hr);
		}
		*configs = sa.release();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Configurations (SAFEARRAY* sa) override
	{
		VARTYPE vt;
		auto hr = SafeArrayGetVartype(sa, &vt); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, vt != VT_UNKNOWN);
		UINT dim = SafeArrayGetDim(sa);
		RETURN_HR_IF(E_NOTIMPL, dim != 1);
		LONG lbound;
		hr = SafeArrayGetLBound(sa, 1, &lbound); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, lbound != 0);
		LONG ubound;
		hr = SafeArrayGetUBound(sa, 1, &ubound); RETURN_IF_FAILED(hr);

		vector_nothrow<com_ptr<IProjectConfig>> newConfigs;
		for (LONG i = 0; i <= ubound; i++)
		{
			wil::com_ptr_nothrow<IUnknown> child;
			hr = SafeArrayGetElement (sa, &i, child.addressof()); RETURN_IF_FAILED(hr);
			com_ptr<IProjectConfig> config;
			hr = child->QueryInterface(&config); RETURN_IF_FAILED(hr);
			bool pushed = newConfigs.try_push_back(std::move(config)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}

		_configs = std::move(newConfigs);

		_isDirty = true;

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Items (SAFEARRAY** items) override
	{
		*items = nullptr;

		SAFEARRAYBOUND bound;
		bound.cElements = 0;
		bound.lLbound = 0;
		for (auto c = _firstChild.get(); c; c = c->Next())
			bound.cElements++;

		auto sa = unique_safearray(SafeArrayCreate(VT_UNKNOWN, 1, &bound)); RETURN_HR_IF(E_OUTOFMEMORY, !sa);
		LONG i = 0;
		for (auto c = _firstChild.get(); c; c = c->Next())
		{
			auto hr = SafeArrayPutElement(sa.get(), &i, static_cast<IUnknown*>(c)); RETURN_IF_FAILED(hr);
			i++;
		}
		*items = sa.release();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Items (SAFEARRAY* sa) override
	{
		VARTYPE vt;
		auto hr = SafeArrayGetVartype(sa, &vt); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, vt != VT_UNKNOWN);
		UINT dim = SafeArrayGetDim(sa);
		RETURN_HR_IF(E_NOTIMPL, dim != 1);
		LONG lbound;
		hr = SafeArrayGetLBound(sa, 1, &lbound); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, lbound != 0);
		LONG ubound;
		hr = SafeArrayGetUBound(sa, 1, &ubound); RETURN_IF_FAILED(hr);

		_firstChild = nullptr;
		if (ubound >= 0)
		{
			vector_nothrow<com_ptr<IProjectItem>> items;
			bool resized = items.try_resize(ubound + 1); RETURN_HR_IF(E_OUTOFMEMORY, !resized);

			for (LONG i = 0; i <= ubound; i++)
			{
				wil::com_ptr_nothrow<IUnknown> child;
				hr = SafeArrayGetElement (sa, &i, child.addressof()); RETURN_IF_FAILED(hr);
				hr = child->QueryInterface(&items[i]); RETURN_IF_FAILED(hr);
			}

			_firstChild = items[0];
			auto _lastChild = _firstChild.get();
			for (uint32_t i = 1; i < items.size(); i++)
			{
				_lastChild->SetNext(items[i].get());
				_lastChild = items[i].get();
			}
		}

		// This is called only from LoadXml, no need to set dirty flag or send notifications.

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Guid (BSTR* value) override
	{
		const size_t bufferSize = 40;
		wchar_t buffer[bufferSize];
		int ires = StringFromGUID2(_projectInstanceGuid, buffer, (int)bufferSize); RETURN_HR_IF(E_FAIL, !ires);
		*value = SysAllocStringLen(buffer, (UINT)ires - 1); RETURN_IF_NULL_ALLOC(*value);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Guid (BSTR value) override
	{
		GUID guid;
		auto hr = CLSIDFromString(value, &guid); RETURN_IF_FAILED(hr);
		if (_projectInstanceGuid != guid)
		{
			_projectInstanceGuid = guid;
			// This is called only from LoadXml, no need to set dirty flag or send notifications.
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IXmlParent
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IUnknown* child, BSTR* xmlElementNameOut) override
	{
		if (dispidProperty == dispidConfigurations)
		{
			*xmlElementNameOut = SysAllocString(ConfigurationElementName); RETURN_IF_NULL_ALLOC(*xmlElementNameOut);
			return S_OK;
		}

		if (dispidProperty == dispidItems)
		{
			wil::com_ptr_nothrow<IProjectFile> sourceFile;
			if (SUCCEEDED(child->QueryInterface(&sourceFile)))
			{
				*xmlElementNameOut = SysAllocString(FileElementName); RETURN_IF_NULL_ALLOC(*xmlElementNameOut);
				return S_OK;
			}

			RETURN_HR(E_NOTIMPL);
		}

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE CreateChild (DISPID dispidProperty, PCWSTR xmlElementName, IDispatch** childOut) override
	{
		if (dispidProperty == dispidConfigurations)
		{
			wil::com_ptr_nothrow<IProjectConfig> config;
			auto hr = ProjectConfig_CreateInstance(this, &config); RETURN_IF_FAILED(hr);
			*childOut = config.detach();
			return S_OK;
		}

		if (dispidProperty == dispidItems)
		{
			// Files saved by newer versions use "File" for the name of the XML element.
			// For backward compatibility, we also recognize the name "AsmFile".
			// When we'll have directories, they'll use a different XML element name.

			if (!wcscmp(xmlElementName, FileElementName) || !wcscmp(xmlElementName, L"AsmFile"))
			{
				wil::com_ptr_nothrow<IProjectFile> file;
				VSITEMID itemId = _nextFileItemId++;
				auto hr = MakeProjectFile (itemId, this, VSITEMID_ROOT, &file); RETURN_IF_FAILED(hr);
				if (!wcscmp(xmlElementName, L"AsmFile"))
					file->put_BuildTool(BuildToolKind::Assembler);
				*childOut = file.detach();
				return S_OK;
			}

			RETURN_HR(E_NOTIMPL);
		}

		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetIDOfName (ITypeInfo* typeInfo, LPCWSTR name, MEMBERID* pMemId) override
	{
		return typeInfo->GetIDsOfNames(&const_cast<LPOLESTR&>(name), 1, pMemId);
	}
	#pragma endregion

	#pragma region IProjectItemParent
	virtual IProjectItem* STDMETHODCALLTYPE FirstChild() override
	{
		return _firstChild;
	}

	virtual void STDMETHODCALLTYPE SetFirstChild (IProjectItem *next) override
	{
		_firstChild = next;
	}
	#pragma endregion

	#pragma region IPropertyNotifySink
	virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
	{
		_isDirty = true;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IVsHierarchyEvents
	virtual HRESULT STDMETHODCALLTYPE OnItemAdded (VSITEMID itemidParent, VSITEMID itemidSiblingPrev, VSITEMID itemidAdded) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemsAppended (VSITEMID itemidParent) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemDeleted (VSITEMID itemid) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged (VSITEMID itemid, VSHPROPID propid, DWORD flags) override
	{
		for (auto& s : _hierarchyEventSinks)
			s.second->OnPropertyChanged(itemid, propid, flags);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnInvalidateItems (VSITEMID itemidParent) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE OnInvalidateIcon (HICON hicon) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL *pfHide) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override
	{
		if (dispid == dispidProjectGuid)
		{
			*fReadOnly = TRUE;
			return S_OK;
		}

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetClassName (BSTR *pbstrClassName) override
	{
		// Shown by VS at the top of the Properties Window.
		*pbstrClassName = SysAllocString(L"Project");
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE CanResetPropertyValue (DISPID dispid, BOOL* pfCanReset) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE ResetPropertyValue (DISPID dispid) override { return E_NOTIMPL; }
	#pragma endregion

};

HRESULT MakeFelixProject (IServiceProvider* sp, LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags, REFIID iidProject, void** ppvProject)
{
	wil::com_ptr_nothrow<Z80Project> p = new (std::nothrow) Z80Project(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(sp, pszFilename, pszLocation, pszName, grfCreateFlags); RETURN_IF_FAILED_EXPECTED(hr);
	return p->QueryInterface (iidProject, ppvProject);
}
