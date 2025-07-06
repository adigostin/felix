
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

// What MPF implements: https://docs.microsoft.com/en-us/visualstudio/extensibility/internals/project-model-core-components?view=vs-2022
class ProjectNode
	: IProjectNodeProperties    // includes IDispatch
	, IVsProject2              // includes IVsProject
	, IVsUIHierarchy           // includes IVsHierarchy
	, IPersistFileFormat       // includes IPersist
	, IVsPersistHierarchyItem
	, IVsGetCfgProvider
	, IVsCfgProvider2          // includes IVsCfgProvider
	, IProvideClassInfo
	//, IConnectionPointContainer
	, IVsHierarchyDeleteHandler3
	, IXmlParent
	//, IVsProjectBuildSystem
	//, IVsBuildPropertyStorage
	//, IVsBuildPropertyStorage2
	, IProjectNode
	, IParentNode
	, IPropertyNotifySink // this implementation only used to mark the project as dirty
	, IVsPerPropertyBrowsing
	, IVsUpdateSolutionEvents
	, IVsHierarchyEvents // this implementation forwards to sinks hierarchy events generated in descendants
{
	ULONG _refCount = 0;
	GUID _projectInstanceGuid;
	wil::unique_hlocal_string _projectDir;
	wil::unique_hlocal_string _filename;
	wil::unique_hlocal_string _caption;
	unordered_map_nothrow<VSCOOKIE, wil::com_ptr_nothrow<IVsHierarchyEvents>> _hierarchyEventSinks;
	VSCOOKIE _nextHierarchyEventSinkCookie = 1;
	bool _isDirty = false;
	wil::com_ptr_nothrow<IVsUIHierarchy> _parentHierarchy;
	VSITEMID _parentHierarchyItemId = VSITEMID_NIL; // item id of this project in the parent hierarchy
	com_ptr<IChildNode> _firstChild;
	VSITEMID _nextItemId = 1000;
	unordered_map_nothrow<VSCOOKIE, wil::com_ptr_nothrow<IVsCfgProviderEvents>> _cfgProviderEventSinks;
	VSCOOKIE _nextCfgProviderEventCookie = 1;
	VSCOOKIE _itemDocCookie = VSDOCCOOKIE_NIL;
	vector_nothrow<com_ptr<IProjectConfig>> _configs;

	// I introduced this because VS sometimes retains a project for a long time after the user closes it.
	// For example in VS 17.11.2, when doing Close Solution while a project file was open, and then
	// opening another solution and another project file, VS calls GetCanonicalName on the old project.
	// This happens because VS retains a reference to the old project in
	// Microsoft.VisualStudio.PlatformUI.Packages.FileColor.ProjectFileGroupProvider.projectMap.
	// (Looks like a bug in OnBeforeCloseProject() in that class.)
	bool _closed = false;

	WeakRefToThis _weakRefToThis;
	wil::unique_bstr _autoOpenFiles;
	VSCOOKIE _updateBuildSolutionEventsCookie = VSCOOKIE_NIL;

	static HRESULT CreateProjectFilesFromTemplate (const wchar_t* fromProjFilePath, const wchar_t* location, const wchar_t* filename)
	{
		// Prepare the file operations.

		wil::com_ptr_nothrow<IFileOperation> pfo;
		auto hr = CoCreateInstance(__uuidof(FileOperation), NULL, CLSCTX_ALL, IID_PPV_ARGS(&pfo)); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IVsUIShell> shell;
		hr = serviceProvider->QueryService(SID_SVsUIShell, &shell);
		if (SUCCEEDED(hr))
		{
			HWND ownerHwnd;
			hr = shell->GetDialogOwnerHwnd(&ownerHwnd); RETURN_IF_FAILED(hr);
			hr = pfo->SetOwnerWindow(ownerHwnd); RETURN_IF_FAILED(hr);
		}

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

		wil::unique_process_heap_string destinationProjFullPath;
		hr = wil::str_concat_nothrow (destinationProjFullPath, location, L"\\", filename); RETURN_IF_FAILED(hr);
		BOOL bres = CopyFileW (fromProjFilePath, destinationProjFullPath.get(), FALSE); RETURN_LAST_ERROR_IF(!bres);

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

		if (itemIDListCount)
		{
			wil::com_ptr_nothrow<IShellItemArray> sourceFiles;
			hr = SHCreateShellItemArray (nullptr, projectTemplateDir.get(), (UINT)itemIDListCount, const_cast<LPCITEMIDLIST*>(itemIDList[0].addressof()), &sourceFiles); RETURN_IF_FAILED(hr);

			hr = pfo->CopyItems (sourceFiles.get(), destinationFolder.get()); RETURN_IF_FAILED(hr);

			hr = pfo->PerformOperations(); RETURN_IF_FAILED(hr);
		}

		// ----------------------------------------------------------------------

		return S_OK;
	}

public:
	HRESULT InitInstance (LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags)
	{
		HRESULT hr;

		hr = _weakRefToThis.InitInstance(static_cast<IVsHierarchy*>(this));

		if (grfCreateFlags & CPF_CLONEFILE)
		{
			hr = EnsureDirHasBackslash (pszLocation, _projectDir); RETURN_IF_FAILED(hr);

			_filename = wil::make_hlocal_string_nothrow(pszName); RETURN_IF_NULL_ALLOC(_filename);
			const wchar_t* ext = ::PathFindExtension(pszName);
			_caption = wil::make_hlocal_string_nothrow(pszName, ext - pszName); RETURN_IF_NULL_ALLOC(_caption);

			hr = CreateProjectFilesFromTemplate (pszFilename, pszLocation, pszName); RETURN_IF_FAILED(hr);

			wchar_t projFilePath[MAX_PATH];
			PathCombine (projFilePath, pszLocation, pszName);

			com_ptr<IStream> stream;
			auto hr = SHCreateStreamOnFileEx(projFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, 0, nullptr, &stream); RETURN_IF_FAILED(hr);
			hr = LoadFromXml(this, ProjectElementName, stream.get()); RETURN_IF_FAILED(hr);

			hr = CoCreateGuid (&_projectInstanceGuid); RETURN_IF_FAILED(hr);

			hr = SHCreateStreamOnFileEx (projFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, 0, nullptr, &stream); RETURN_IF_FAILED(hr);
			hr = SaveToXml(this, ProjectElementName, 0, stream.get()); RETURN_IF_FAILED(hr);
		}
		else if (grfCreateFlags & CPF_OPENFILE)
		{
			// pszFilename is the full path of the file to open, the others are NULL.
			const wchar_t* fn = PathFindFileName(pszFilename);
			
			_projectDir = wil::make_hlocal_string_nothrow(pszFilename, fn - pszFilename); RETURN_IF_NULL_ALLOC(_projectDir);
			
			_filename = wil::make_hlocal_string_nothrow(fn); RETURN_IF_NULL_ALLOC(_filename);
			const wchar_t* ext = PathFindExtension(pszFilename);
			_caption = wil::make_hlocal_string_nothrow(fn, ext - fn); RETURN_IF_NULL_ALLOC(_caption);

			com_ptr<IStream> stream;
			hr = SHCreateStreamOnFileEx(pszFilename, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, 0, nullptr, &stream); RETURN_IF_FAILED_EXPECTED(hr);
			hr = LoadFromXml (this, ProjectElementName, stream.get()); RETURN_IF_FAILED(hr);
		}
		else
		{
			// Neither clone nor open. So far I haven't found a scenario where VS calls us like this.
			// It's useful for tests, where we need a blank project.
			RETURN_HR_IF(E_UNEXPECTED, !!pszFilename);
			RETURN_HR_IF(E_UNEXPECTED, !pszLocation);
			hr = EnsureDirHasBackslash (pszLocation, _projectDir); RETURN_IF_FAILED(hr);
		}

		for (auto& c : _configs)
			c->SetSite(this);

		com_ptr<IVsSolutionBuildManager> buildManager;
		if (SUCCEEDED(serviceProvider->QueryService(SID_SVsSolutionBuildManager, IID_PPV_ARGS(&buildManager))))
			buildManager->AdviseUpdateSolutionEvents(this, &_updateBuildSolutionEventsCookie);

		_isDirty = false;
		return S_OK;
	}

	~ProjectNode()
	{
		WI_ASSERT (_updateBuildSolutionEventsCookie == VSCOOKIE_NIL);
		WI_ASSERT (_cfgProviderEventSinks.empty());
	}

	static HRESULT EnsureDirHasBackslash (LPCOLESTR pszLocation, wil::unique_hlocal_string& dir)
	{
		RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME), !pszLocation);
		RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME), !pszLocation[0]);
		RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME), PathIsRelative(pszLocation));

		size_t len = wcslen(pszLocation);
		if (pszLocation[len - 1] == '\\')
		{
			dir = wil::make_hlocal_string_nothrow(pszLocation, len); RETURN_IF_NULL_ALLOC(dir);
		}
		else if (pszLocation[len - 1] == '/')
		{
			dir = wil::make_hlocal_string_nothrow(pszLocation, len); RETURN_IF_NULL_ALLOC(dir);
			dir.get()[len - 1] = '\\';
		}
		else
		{
			dir = wil::make_hlocal_string_failfast(pszLocation, len + 1); RETURN_IF_NULL_ALLOC(dir);
			dir.get()[len] = '\\';
			dir.get()[len + 1] = 0;
		}

		return S_OK;
	}

	// If found, returns S_OK and ppItem is non-null.
	// If not found, return S_FALSE and ppItem is null.
	// If error during enumeration, returns an error code.
	// The filter function must return S_OK for a match (enumeration stops), S_FALSE for non-match (enumeration continues), or an error code (enumeration stops).
	HRESULT FindDescendantIf (const stdext::inplace_function<HRESULT(IChildNode*), 32>& predicate, IChildNode** ppItem)
	{
		if (ppItem)
			*ppItem = nullptr;

		stdext::inplace_function<HRESULT(IParentNode* parent)> enumChildren;

		enumChildren = [&predicate, ppItem, &enumChildren](IParentNode* parent) -> HRESULT
		{
			for (auto c = parent->FirstChild(); c; c = c->Next())
			{
				auto hr = predicate(c); RETURN_IF_FAILED_EXPECTED(hr);
				if (hr == S_OK)
				{
					if (ppItem)
					{
						*ppItem = c;
						(*ppItem)->AddRef();
					}
					return S_OK;
				}

				com_ptr<IParentNode> cp;
				if (SUCCEEDED(c->QueryInterface(&cp)))
				{
					auto hr = enumChildren(cp); RETURN_IF_FAILED_EXPECTED(hr);
					if (hr == S_OK)
						return S_OK;
				}
			}

			return S_FALSE;
		};

		return enumChildren(this);
	}
	
	// Returns S_OK when found, S_FALSE when not found.
	HRESULT FindDescendant (VSITEMID itemid, IChildNode** ppItem)
	{
		if (ppItem)
			*ppItem = nullptr;

		stdext::inplace_function<HRESULT(IParentNode* parent)> enumChildren;

		enumChildren = [itemid, ppItem, &enumChildren](IParentNode* parent) -> HRESULT
		{
			for (auto c = parent->FirstChild(); c; c = c->Next())
			{
				if (c->GetItemId() == itemid)
				{
					if (ppItem)
					{
						*ppItem = c;
						(*ppItem)->AddRef();
					}
					return S_OK;
				}

				com_ptr<IParentNode> cp;
				if (SUCCEEDED(c->QueryInterface(&cp)))
				{
					auto hr = enumChildren(cp); RETURN_IF_FAILED_EXPECTED(hr);
					if (hr == S_OK)
						return S_OK;
				}
			}

			return S_FALSE;
		};

		return enumChildren(this);
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		RETURN_HR_IF_EXPECTED(E_NOINTERFACE, riid == IID_ICustomCast); // VS abuses this, let's check it first

		if (   TryQI<IProjectNodeProperties>(this, riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IUnknown>(static_cast<IVsUIHierarchy*>(this), riid, ppvObject)
			|| TryQI<IVsHierarchy>(this, riid, ppvObject)
			|| TryQI<IVsUIHierarchy>(this, riid, ppvObject)
			|| TryQI<IPersist>(this, riid, ppvObject)
			|| TryQI<IPersistFileFormat>(this, riid, ppvObject)
			|| TryQI<IVsProject>(this, riid, ppvObject)
			|| TryQI<IVsProject2>(this, riid, ppvObject)
			|| TryQI<IProvideClassInfo>(this, riid, ppvObject)
			|| TryQI<IVsGetCfgProvider>(this, riid, ppvObject)
			|| TryQI<IVsCfgProvider>(this, riid, ppvObject)
			|| TryQI<IVsCfgProvider2>(this, riid, ppvObject)
			//|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
			|| TryQI<IVsHierarchyDeleteHandler3>(this, riid, ppvObject)
			|| TryQI<IVsPersistHierarchyItem>(this, riid, ppvObject)
			|| TryQI<IXmlParent>(this, riid, ppvObject)
			//|| TryQI<IVsProjectBuildSystem>(this, riid, ppvObject)
			//|| TryQI<IVsBuildPropertyStorage>(this, riid, ppvObject)
			//|| TryQI<IVsBuildPropertyStorage2>(this, riid, ppvObject)
			|| TryQI<IParentNode>(this, riid, ppvObject)
			|| TryQI<IProjectNode>(this, riid, ppvObject)
			|| TryQI<INode>(this, riid, ppvObject)
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
			|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)
			|| TryQI<IVsUpdateSolutionEvents>(this, riid, ppvObject)
			|| TryQI<IVsHierarchyEvents>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

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
			|| riid == IID_TypeDescriptor_IUnimplemented
			|| riid == IID_PropertyGrid_IUnimplemented
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

	IMPLEMENT_IDISPATCH(IID_IProjectNodeProperties);

	#pragma region IVsHierarchy
	virtual HRESULT STDMETHODCALLTYPE SetSite(IServiceProvider* pSP) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetSite(IServiceProvider** ppSP) override
	{
		*ppSP = serviceProvider.get();
		(*ppSP)->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryClose(BOOL* pfCanClose) override
	{
		*pfCanClose = TRUE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Close() override
	{
		if (_updateBuildSolutionEventsCookie)
		{
			com_ptr<IVsSolutionBuildManager> buildManager;
			if (SUCCEEDED(serviceProvider->QueryService(SID_SVsSolutionBuildManager, IID_PPV_ARGS(&buildManager))))
				buildManager->UnadviseUpdateSolutionEvents(_updateBuildSolutionEventsCookie);
			_updateBuildSolutionEventsCookie = VSCOOKIE_NIL;
		}

		_parentHierarchy = nullptr;
		_parentHierarchyItemId = VSITEMID_NIL;
		_firstChild = nullptr;
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
				*pguid = __uuidof(IProjectNodeProperties);
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

			#ifdef _DEBUG
			RETURN_HR(E_NOTIMPL);
			#else
			return E_NOTIMPL;
			#endif
		}

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
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

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
			return d->SetGuidProperty(propid, rguid);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty (VSITEMID itemid, VSHPROPID propid, VARIANT* pvar) override
	{
		if (_closed)
			return E_UNEXPECTED;

		#pragma region Some properties apply to the hierarchy as a whole
		if (propid == VSHPROPID_ExtObject) // -2027
			return E_NOTIMPL;

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

			if (propid == VSHPROPID_Name || propid == VSHPROPID_EditLabel) // -2012, -2026
			{
				const wchar_t* ext = ::PathFindExtension(_filename.get());
				BSTR bs = SysAllocStringLen(_filename.get(), (UINT)(ext - _filename.get())); RETURN_IF_NULL_ALLOC(bs);
				pvar->vt = VT_BSTR;
				pvar->bstrVal = bs;
				return S_OK;
			}

			if (propid == VSHPROPID_BrowseObject) // -2018
			{
				// Be aware that VS2019 will QueryInterface for IVsHierarchy on the object we return here,
				// then call again this GetProperty function to get another browse object. It does this
				// recursively and will stop when (1) we return an error HRESULT, or (2) we return the same
				// browse object as in the previous call, or (3) it overflows the stack (that exception
				// seems to be caught and turned into a freeze).
				return InitVariantFromDispatch (this, pvar);
			}

			if (propid == VSHPROPID_ProjectDir) //-2021
				return InitVariantFromString (_projectDir.get(), pvar);

			if (propid == VSHPROPID_TypeName) // -2030
				return InitVariantFromString (L"Z80", pvar); // Called by 17.11.5 from Help -> About.

			if (propid == VSHPROPID_HandlesOwnReload) // -2031
				return InitVariantFromBoolean (FALSE, pvar);

			if (propid == VSHPROPID_ItemDocCookie) // -2034
			{
				if (_itemDocCookie != VSDOCCOOKIE_NIL)
					return InitVariantFromInt32 (_itemDocCookie, pvar);
				else
					return E_FAIL;
			}

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
			if (   propid == VSHPROPID_NextSibling                // -1002 - VS asks for this when opening an empty project
				|| propid == VSHPROPID_IconIndex                  // -2005
				|| propid == VSHPROPID_IconHandle                 // -2013
				|| propid == VSHPROPID_OpenFolderIconHandle       // -2014
				|| propid == VSHPROPID_OpenFolderIconIndex        // -2015
				|| propid == VSHPROPID_AltHierarchy               // -2019
				|| propid == VSHPROPID_SortPriority               // -2022 - requested when reverting in Git an open project file
				|| propid == VSHPROPID_UserContext                // -2023
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
				|| propid == VSHPROPID_AlwaysBuildOnDebugLaunch   // -2109
				|| propid == VSHPROPID_ProvisionalViewingStatus   // -2112
				|| propid == VSHPROPID_TargetRuntime              // -2116 - We'll never support this as it has values only for JS, CLR, Native
				|| propid == VSHPROPID_AppContainer               // -2117 - Something about .Net
				|| propid == VSHPROPID_OutputType                 // -2118 - "This property is optional."
				|| propid == VSHPROPID_ProjectUnloadStatus        // -2120
				|| propid == VSHPROPID_DemandLoadDependencies     // -2121 - Much later, if ever
				|| propid == VSHPROPID_ProjectCapabilities        // -2124 - MPF doesn't implement this
				|| propid == VSHPROPID_RequiresReloadForExternalFileChange	// -2125,
				|| propid == VSHPROPID_ProjectRetargeting         // -2134 - We won't be supporting IVsRetargetProject
				|| propid == VSHPROPID_Subcaption                 // -2136 - This is shown on the project node between parentheses after the project name.
				|| propid == VSHPROPID_SharedItemsImportFullPaths // -2145
				|| propid == VSHPROPID_ProjectTreeCapabilities    // -2146 - MPF doesn't implement this, so we'll probably not implement it either
				|| propid == VSHPROPID_OneAppCapabilities         // -2149 - Virtually no info about this one. We probably don't need it.
				|| propid == VSHPROPID_SharedAssetsProject        // -2153
				|| propid == VSHPROPID_CanBuildQuickCheck         // -2156 - Much later, if ever
				|| propid == VSHPROPID_SupportsIconMonikers       // -2159
				|| propid == VSHPROPID_IconMonikerImageList       // -2164
				|| propid == VSHPROPID_ProjectCapabilitiesChecker // -2173
				|| propid == -2176 // VSHPROPID_HasRunningOperation
				|| propid == -2177 // VSHPROPID_PreserveExpandCollapseState
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

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
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

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
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

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
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

		com_ptr<IChildNode> c;
		hr = FindDescendantIf([pszName](IChildNode* c)
			{
				wil::unique_bstr childCN;
				auto hr = c->GetCanonicalName(&childCN);
				if (hr == E_NOTIMPL)
					return S_FALSE;
				RETURN_IF_FAILED_EXPECTED(hr);
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

		auto newFullPath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(newFullPath);
		PathCombine (newFullPath.get(), _projectDir.get(), newFilename.get());

		// Check if it exists, but only if the user changed more than character casing.
		if (_wcsicmp(_filename.get(), newFilename.get()))
		{
			HANDLE hFile = ::CreateFile (newFullPath.get(), GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
			if (hFile != INVALID_HANDLE_VALUE)
			{
				CloseHandle(hFile);
				return SetErrorInfo1 (HRESULT_FROM_WIN32(ERROR_FILE_EXISTS), IDS_RENAMEFILEALREADYEXISTS, newFullPath.get());
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


	HRESULT QueryStatusCommandOnProjectNode (const GUID* pguidCmdGroup, ULONG cmdID, __RPC__out DWORD* cmdf, __RPC__out OLECMDTEXT *pCmdText)
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
			case cmdidExit: // 229
			case cmdidPropertyPages: // 232
			case cmdidAddExistingItem: // 244
			case cmdidStepInto: // 248
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

			case cmdidSaveProjectItem: // 331
				*cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
				break;
			case cmdidSaveProjectItemAs: // 226
			{
				// This command is actually "Save Selection As", which we want to support for files,
				// but not for projects (it's too complicated for projects and Visual Studio doesn't
				// support it either for Visual C++ projects, so it's probably not a big deal).

				com_ptr<IVsMonitorSelection> monsel;
				auto hr = serviceProvider->QueryService(SID_SVsShellMonitorSelection, &monsel); RETURN_IF_FAILED(hr);
				com_ptr<IVsHierarchy> hier;
				VSITEMID selitemid;
				com_ptr<IVsMultiItemSelect> mis;
				com_ptr<ISelectionContainer> sc;
				hr = monsel->GetCurrentSelection (&hier, &selitemid, &mis, &sc); RETURN_IF_FAILED(hr);

				// Enable it only for single-item selection (Save As doesn't make sense for
				// multiple selections), and only if it is a file node (not a project node).
				bool enable = (selitemid != VSITEMID_SELECTION) && (selitemid != VSITEMID_ROOT);
				*cmdf = OLECMDF_SUPPORTED | (enable ? OLECMDF_ENABLED : 0);

				// Note that we still have a scenario where this command is enabled but doesn't work:
				// Right after a project is opened, if no files were auto-opened, if a file node is selected
				// in the Solution Explorer, if the focus is in the Solution Explorer window.
				// We'll deal with this some other time.

				break;
			}

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
				return S_OK;
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
			*cmdf = OLECMDF_SUPPORTED | OLECMDF_INVISIBLE;
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

	HRESULT ExecCommandOnProjectNode (const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn, VARIANT* pvaOut)
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

		return OLECMDERR_E_UNKNOWNGROUP;
	}

	HRESULT ProcessCommandAddItem (VSITEMID location, BOOL fAddNewItem)
	{
		wil::com_ptr_nothrow<IVsAddProjectItemDlg> dlg;
		auto hr = serviceProvider->QueryService(SID_SVsAddProjectItemDlg, &dlg); RETURN_IF_FAILED(hr);
		VSADDITEMFLAGS flags;
		if (fAddNewItem)
			flags = VSADDITEM_AddNewItems | VSADDITEM_SuggestTemplateName | VSADDITEM_NoOnlineTemplates; /* | VSADDITEM_ShowLocationField */
		else
			flags = VSADDITEM_AddExistingItems | VSADDITEM_AllowMultiSelect | VSADDITEM_AllowStickyFilter;

		// TODO: To specify a sticky behavior for the location field, which is the recommended behavior, remember the last location field value and pass it back in when you open the dialog box again.
		// TODO: To specify sticky behavior for the filter field, which is the recommended behavior, remember the last filter field value and pass it back in when you open the dialog box again.
		hr = dlg->AddProjectItemDlg (location, __uuidof(IProjectNodeProperties), this, flags, nullptr, nullptr, nullptr, nullptr, nullptr);
		if (FAILED(hr) && (hr != OLE_E_PROMPTSAVECANCELLED))
			return hr;

		return S_OK;
	}

	#pragma region IVsUIHierarchy
	virtual HRESULT STDMETHODCALLTYPE QueryStatusCommand (VSITEMID itemid, const GUID* pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT* pCmdText) override
	{
		if (*pguidCmdGroup == CMDSETID_StandardCommandSet2K && cCmds == 1 && prgCmds[0].cmdID == ECMD_SLNREFRESH)
		{
			// We treat this here no matter which node is selected.
			prgCmds[0].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
			return S_OK;
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97 && cCmds == 1 && prgCmds[0].cmdID == cmdidNewFolder) // 245
		{
			com_ptr<IChildNode> node;
			if (itemid != VSITEMID_SELECTION
				&& itemid != VSITEMID_NIL
				&& (itemid == VSITEMID_ROOT
					|| (FindDescendant(itemid, &node) == S_OK && node.try_query<IFolderNode>())))
			{
				prgCmds[0].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED;
				return S_OK;
			}
		}

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97 && cCmds == 1
			&& (prgCmds->cmdID == cmdidAddNewItem || prgCmds->cmdID == cmdidAddExistingItem))
		{
			if (itemid == VSITEMID_NIL || itemid == VSITEMID_ROOT
				|| (itemid != VSITEMID_SELECTION && FindDescendant(itemid, nullptr) == S_OK))
				return (prgCmds->cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED), S_OK;
		}

		if (itemid == VSITEMID_ROOT)
		{
			if (cCmds == 1)
				return QueryStatusCommandOnProjectNode (pguidCmdGroup, prgCmds->cmdID, &prgCmds->cmdf, pCmdText);
			RETURN_HR(E_NOTIMPL);
		}

		if (itemid == VSITEMID_SELECTION)
			return OLECMDERR_E_NOTSUPPORTED;
		if (itemid == VSITEMID_NIL)
			return OLECMDERR_E_NOTSUPPORTED;

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
		{
			if (cCmds == 1)
				return d->QueryStatusCommand (pguidCmdGroup, prgCmds, pCmdText);
			RETURN_HR(E_NOTIMPL);
		}

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE ExecCommand (VSITEMID itemid, const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn, VARIANT* pvaOut) override
	{
		// Some commands such as UIHWCMDID_RightClick apply to items of all kinds.
		if (*pguidCmdGroup == GUID_VsUIHierarchyWindowCmds && nCmdID == UIHWCMDID_RightClick)
			return ShowContextMenu(itemid, pvaIn);

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet2K && nCmdID == ECMD_SLNREFRESH)
			return RefreshHierarchy();

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97 && nCmdID == cmdidNewFolder)
			return ProcessCommandAddNewFolder(itemid, pvaOut);

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97 && nCmdID == cmdidAddNewItem)
			return ProcessCommandAddItem(itemid, TRUE);

		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97 && nCmdID == cmdidAddExistingItem)
			return ProcessCommandAddItem(itemid, FALSE);

		if (itemid == VSITEMID_ROOT)
			return ExecCommandOnProjectNode (pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

		// We need to return OLECMDERR_E_NOTSUPPORTED. If we return something else, VS won't like it.
		// For example, if the user selects two project items in the Solution Explorer and hits Escape,
		// and if we return E_NOTIMPL for VSITEMID_SELECTION, VS will say "operation cannot be completed".
		RETURN_HR_IF_EXPECTED(OLECMDERR_E_NOTSUPPORTED, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(OLECMDERR_E_NOTSUPPORTED, itemid == VSITEMID_NIL);

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
			return d->ExecCommand (pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	HRESULT ShowContextMenu (VSITEMID itemid, const VARIANT* pvaIn)
	{
		POINTS pts;
		memcpy (&pts, &pvaIn->uintVal, 4);

		com_ptr<IVsUIShell> shell;
		auto hr = serviceProvider->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);

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
					com_ptr<IChildNode> d;
					if (FindDescendant(items[i].itemid, &d) == S_OK)
					{
						if (wil::try_com_query_nothrow<IFileNode>(d))
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
			com_ptr<IChildNode> d;
			if (FindDescendant(itemid, &d) == S_OK)
			{
				if (wil::try_com_query_nothrow<IFileNode>(d))
					return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_ITEMNODE, pts, nullptr);
				else if (wil::try_com_query_nothrow<IFolderNode>(d))
					return shell->ShowContextMenu (0, guidSHLMainMenu, IDM_VS_CTXT_FOLDERNODE, pts, nullptr);
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
		*pClassID = __uuidof(IProjectNodeProperties);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE IsDirty(BOOL* pfIsDirty) override
	{
		// VS calls this one to find out whether the project file is dirty.
		// (See related IsItemDirty in this class.)
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

		// "If the object is in the untitled state and null is passed as the pszFilename, the object returns E_INVALIDARG."
		bool isTitled = _filename && _filename.get()[0] && _projectDir && _projectDir.get()[0];
		RETURN_HR_IF_EXPECTED(E_INVALIDARG, !isTitled && !pszFilename);

		wil::unique_hlocal_string currentFilePath;
		if (isTitled)
		{
			currentFilePath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(currentFilePath);
			PathCombine(currentFilePath.get(), _projectDir.get(), _filename.get());
		}

		if (!pszFilename || (currentFilePath && !wcscmp(pszFilename, currentFilePath.get())))
		{
			// "Save" operation (save under existing name)

			// TODO: pause monitoring file change notifications
			//CSuspendFileChanges suspendFileChanges(CString(pszFileName), TRUE);

			// Save XML to temporary stream first. During debugging, the SaveToXml file crashes often and we lose the file content.
			auto memStreamRaw = SHCreateMemStream(nullptr, 0); RETURN_IF_NULL_ALLOC(memStreamRaw);
			com_ptr<IStream> memStream;
			memStream.attach (memStreamRaw);
			hr = SaveToXml(this, ProjectElementName, 0, memStream); RETURN_IF_FAILED(hr);
			
			STATSTG stat;
			hr = memStream->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
			RETURN_HR_IF(ERROR_FILE_TOO_LARGE, !!stat.cbSize.HighPart);
			hr = memStream->Seek({ 0 }, STREAM_SEEK_SET, nullptr); RETURN_IF_FAILED(hr);

			com_ptr<IStream> stream;
			hr = SHCreateStreamOnFile(currentFilePath.get(), STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &stream); RETURN_IF_FAILED(hr);
			hr = IStream_Copy(memStream, stream, stat.cbSize.LowPart); RETURN_IF_FAILED(hr);
			stream.reset();

			// TODO: resume monitoring file change notifications

			_isDirty = false; // TODO: notify property changes
		}
		else
		{
			if (fRemember)
			{
				// "Save As"
				// We don't support this. We disabled cmdidSaveProjectItemAs, so we shouldn't get here.
				// If/when we'll support it, we should go through RenameProject().
				RETURN_HR(E_NOTIMPL);
			}
			else
			{
				// "Save A Copy As"
				RETURN_HR(E_NOTIMPL);
			}
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SaveCompleted(LPCOLESTR pszFilename) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetCurFile(LPOLESTR* ppszFilename, DWORD* pnFormatIndex) override
	{
		// Note that we return true for VSHPROPID_MonikerSameAsPersistFile, meaning that
		// IVsProject::GetMkDocument returns the same as GetCurFile for VSITEMID_ROOT.

		auto buffer = wil::make_cotaskmem_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(buffer);
		PathCombine (buffer.get(), _projectDir.get(), _filename.get());

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

		wil::com_ptr_nothrow<IChildNode> c;
		auto hr = FindDescendantIf(
			[pszMkDocument, isRelative](IChildNode* c)
			{
				wil::unique_bstr path;
				if (isRelative)
				{
					auto hr = c->GetCanonicalName(&path);
					if (hr == E_NOTIMPL)
						return S_FALSE;
					RETURN_IF_FAILED(hr);
					return wcscmp(pszMkDocument, path.get()) ? S_FALSE : S_OK;
				}
				else
				{
					wil::unique_process_heap_string path;
					auto hr = GetPathOf (c, path); RETURN_IF_FAILED(hr);
					return wcscmp(pszMkDocument, path.get()) ? S_FALSE : S_OK;
				}
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

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
		{
			if (auto file = d.try_query<IFileNode>())
				return file->GetMkDocument(this, pbstrMkDocument);

			return E_INVALIDARG;
		}

		RETURN_HR_MSG(E_INVALIDARG, "itemid=%u", itemid);
	}

	virtual HRESULT STDMETHODCALLTYPE OpenItem(VSITEMID itemid, REFGUID rguidLogicalView, IUnknown* punkDocDataExisting, IVsWindowFrame** ppWindowFrame) override
	{
		HRESULT hr;

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) != S_OK)
			return E_INVALIDARG;

		com_ptr<IFileNode> file;
		hr = d->QueryInterface(IID_PPV_ARGS(&file)); RETURN_IF_FAILED(hr);
		wil::unique_bstr mkDocument;
		hr = file->GetMkDocument(this, &mkDocument); RETURN_IF_FAILED(hr);

		com_ptr<IVsUIShellOpenDocument> uiShellOpenDocument;
		hr = serviceProvider->QueryService(SID_SVsUIShellOpenDocument, &uiShellOpenDocument); RETURN_IF_FAILED_EXPECTED(hr);

		hr = uiShellOpenDocument->OpenStandardEditor(OSE_ChooseBestStdEditor, mkDocument.get(), rguidLogicalView,
			L"%3", this, itemid, punkDocDataExisting, nullptr, ppWindowFrame);
		if (FAILED(hr))
		{
			if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
				return SetErrorInfo(hr, L"File not found.\r\n\r\n%s", mkDocument.get());
			if (hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
				return SetErrorInfo(hr, L"Path to the file was not found.\r\n\r\n%s", mkDocument.get());
			return hr;
		}

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
		*ppSP = serviceProvider.get();
		(*ppSP)->AddRef();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GenerateUniqueItemName(VSITEMID itemidLoc, LPCOLESTR pszExt, LPCOLESTR pszSuggestedRoot, BSTR* pbstrItemName) override
	{
		HRESULT hr;

		wil::unique_process_heap_string path;
		if (itemidLoc == VSITEMID_ROOT)
		{
			path = wil::make_process_heap_string_nothrow(_projectDir.get()); RETURN_IF_NULL_ALLOC(path);
		}
		else
		{
			com_ptr<IChildNode> n;
			hr = FindDescendant(itemidLoc, &n); RETURN_IF_FAILED_EXPECTED(hr); RETURN_HR_IF_EXPECTED(E_INVALIDARG, hr != S_OK);
			hr = GetPathOf(n, path); RETURN_IF_FAILED_EXPECTED(hr);
		}

		for (uint32_t i = 0; i < 1000; i++)
		{
			wchar_t buffer[50];
			swprintf_s(buffer, L"%s%u%s", pszSuggestedRoot, i, pszExt);
			wchar_t dest[MAX_PATH];
			PathCombine (dest, path.get(), buffer);
			if (!PathFileExists(dest))
			{
				*pbstrItemName = SysAllocString(buffer); RETURN_IF_NULL_ALLOC(*pbstrItemName);
				return S_OK;
			}
		}

		RETURN_HR(E_FAIL);
	}

	HRESULT AddNewFile (IParentNode* location, LPCTSTR pszFullPathSource, LPCTSTR pszNewFileName, IChildNode** ppNewNode)
	{
		HRESULT hr;

		RETURN_HR_IF_NULL(E_INVALIDARG, pszFullPathSource);
		RETURN_HR_IF_NULL(E_INVALIDARG, pszNewFileName);
		RETURN_HR_IF(E_INVALIDARG, !!wcspbrk(pszNewFileName, L":/\\"));
		RETURN_HR_IF_NULL(E_POINTER, ppNewNode);

		// Do we already have an item with the same name under the selected location? If yes, we give an
		// error message and return. Let's not bother asking the user whether to replace the existing item;
		// that existing item might be a folder, and our replacing code would have to be really complicated.
		for (auto c = location->FirstChild(); c; c = c->Next())
		{
			wil::unique_variant saveName;
			hr = c->GetProperty(VSHPROPID_SaveName, &saveName); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_UNEXPECTED, saveName.vt != VT_BSTR);
			if (saveName.bstrVal && !_wcsicmp(saveName.bstrVal, pszNewFileName))
			{
				com_ptr<IVsShell> shell;
				hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);
				wil::unique_bstr text;
				hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_ITEM_ALREADY_EXISTS_IN_LOCATION, &text); RETURN_IF_FAILED(hr);

				return SetErrorInfo(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), text.get());
			}
		}

		// The user chose a hierarchy node as location, but that node may or may not have
		// a corresponding directory in the file system. Let's make sure the directory exists.
		wil::unique_process_heap_string locationDir;
		hr = CreatePathOfNode(location, locationDir); RETURN_IF_FAILED_EXPECTED(hr);
		size_t locationDirLen = wcslen(locationDir.get());
		WI_ASSERT(locationDirLen && locationDir.get()[locationDirLen - 1] == L'\\');

		wil::unique_hlocal_string dest;
		hr = wil::str_concat_nothrow (dest, locationDir, pszNewFileName); RETURN_IF_FAILED(hr);

		BOOL bres = CopyFile(pszFullPathSource, dest.get(), TRUE);
		if (!bres)
			return HRESULT_FROM_WIN32(GetLastError());

		// template was read-only, but our file should not be
		SetFileAttributes(dest.get(), FILE_ATTRIBUTE_ARCHIVE);

		return AddExistingFile (location, dest.get(), ppNewNode);
	}

	HRESULT EnsureFilePathUniqueInProject (LPCWSTR path)
	{
		auto hr = FindDescendantIf([path](IChildNode* node)
			{
				if (auto fn = wil::try_com_query_nothrow<IFileNodeProperties>(node))
				{
					wil::unique_bstr existingPath;
					if (SUCCEEDED(fn->get_Path(&existingPath)) && !_wcsicmp(path, existingPath.get()))
						return S_OK;
				}
				return S_FALSE;
			}, nullptr);
		if (hr == S_OK)
			return SetErrorInfo1(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), IDS_FILE_ALREADY_IN_PROJECT, path);
		return S_OK;
	}

	HRESULT AddExistingFile (IParentNode* location, LPCTSTR pszFullPathSource, IChildNode** ppNewFile, BOOL fSilent = FALSE, BOOL fLoad = FALSE)
	{
		HRESULT hr;

		hr = QueryEditProjectFile(this); RETURN_IF_FAILED_EXPECTED(hr);

		// TODO: there's a lot that needs to be called in IVsTrackProjectDocumentsEvents2

		com_ptr<IFileNode> file;
		if (PathIsSameRoot(pszFullPathSource, _projectDir.get()))
		{
			// File is on same drive as the project. We'll keep its path in a form relative to the project dir.
			wchar_t relativeUgly[MAX_PATH];
			BOOL bRes = PathRelativePathTo (relativeUgly, _projectDir.get(), FILE_ATTRIBUTE_DIRECTORY, pszFullPathSource, 0);
			if (!bRes)
				return SetErrorInfo(E_INVALIDARG, L"Can't make a relative path from '%s' relative to '%s'.", pszFullPathSource, _projectDir.get());

			for (auto* p = wcschr(relativeUgly, '/'); p; p = wcschr(p, '/'))
				*p = '\\';

			if (wcsncmp(relativeUgly, L"..\\", 3))
			{
				// File is under the project dir.
				const wchar_t* relative = relativeUgly;
				while (relative[0] == '.' && relative[1] == '\\')
					relative += 2;

				// Find or create path of directories where we need to add our file node.
				IParentNode* parent = this;
				auto ptrComponent = relative;
				while (auto nextComp = wcschr(ptrComponent, L'\\'))
				{
					auto dir = wil::make_process_heap_string_nothrow (ptrComponent, nextComp - ptrComponent); RETURN_IF_NULL_ALLOC(dir);
					ptrComponent = nextComp + 1;
					com_ptr<IFolderNode> ch;
					hr = GetOrCreateChildFolder(parent, dir.get(), false, &ch); RETURN_IF_FAILED(hr);
					parent = ch->AsParentNode(); 
				}
				if (FindChildFileByName(parent, ptrComponent))
					return SetErrorInfo0(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), IDS_FILE_ALREADY_IN_PROJECT);

				hr = MakeFileNodeForExistingFile (ptrComponent, &file); RETURN_IF_FAILED(hr);
				hr = AddFileToParent(file, parent); RETURN_IF_FAILED(hr);
			}
			else
			{
				// File is outside the project dir, but on the same drive.
				// Add as link to the folder selected by the user when starting the UI command.
				hr = EnsureFilePathUniqueInProject(relativeUgly); RETURN_IF_FAILED_EXPECTED(hr);
				hr = MakeFileNodeForExistingFile (relativeUgly, &file); RETURN_IF_FAILED(hr);
				hr = AddFileToParent(file, location); RETURN_IF_FAILED(hr);
			}
		}
		else
		{
			// File is on different drive, not on the project's drive. We keep it absolute.
			// Add as link to the folder selected by the user when starting the UI command.
			hr = EnsureFilePathUniqueInProject(pszFullPathSource); RETURN_IF_FAILED_EXPECTED(hr);
			hr = MakeFileNodeForExistingFile (pszFullPathSource, &file); RETURN_IF_FAILED(hr);
			hr = AddFileToParent(file, location); RETURN_IF_FAILED(hr);
		}

		_isDirty = true;

		*ppNewFile = file.detach();
		return S_OK;
	}

	/// <summary>
	/// 
	/// </summary>
	/// <param name="itemidLoc"></param>
	/// <param name="dwAddItemOperation">VSADDITEMOP_CLONEFILE / VSADDITEMOP_OPENFILE </param>
	/// <param name="pszItemName">For VSADDITEMOP_CLONEFILE it's the name of the new item. For VSADDITEMOP_OPENFILE must be NULL.</param>
	/// <param name="cFilesToOpen">For VSADDITEMOP_CLONEFILE must be 1. For VSADDITEMOP_OPENFILE it's the number of files to open.</param>
	/// <param name="rgpszFilesToOpen">For VSADDITEMOP_CLONEFILE it's the full path of the template file. For VSADDITEMOP_OPENFILE it's an array of full paths to the files to open.</param>
	/// <param name="hwndDlgOwner"></param>
	/// <param name="pResult"></param>
	/// <returns></returns>
	virtual HRESULT STDMETHODCALLTYPE AddItem(VSITEMID itemidLoc, VSADDITEMOPERATION dwAddItemOperation,
		LPCOLESTR pszItemName, ULONG cFilesToOpen, LPCOLESTR rgpszFilesToOpen[], HWND hwndDlgOwner,
		VSADDRESULT* pResult) override
	{
		HRESULT hr;

		for (ULONG i = 0; i < cFilesToOpen; i++)
			RETURN_HR_IF(CO_E_BAD_PATH, PathIsRelative(rgpszFilesToOpen[i]));

		com_ptr<IParentNode> location;
		if (itemidLoc == VSITEMID_ROOT)
			location = this;
		else
		{
			com_ptr<IChildNode> cn;
			if (FindDescendant(itemidLoc, &cn) != S_OK)
				RETURN_HR(E_INVALIDARG);
			hr = cn->QueryInterface(IID_PPV_ARGS(&location)); RETURN_IF_FAILED(hr);
		}

		switch(dwAddItemOperation)
		{
			case VSADDITEMOP_CLONEFILE:
			{
				// Add New File
				RETURN_HR_IF(E_INVALIDARG, cFilesToOpen != 1);
				com_ptr<IChildNode> pNewNode;
				hr = AddNewFile (location, rgpszFilesToOpen[0], pszItemName, &pNewNode); RETURN_IF_FAILED_EXPECTED(hr);
				if (pResult)
					*pResult = ADDRESULT_Success;

				com_ptr<IVsWindowFrame> frame;
				hr = this->OpenItem (pNewNode->GetItemId(), LOGVIEWID_Primary, DOCDATAEXISTING_UNKNOWN, &frame);
				if (SUCCEEDED(hr))
					frame->Show();

				return S_OK;
			}

			case VSADDITEMOP_LINKTOFILE:
				// Because we are a reference-based project system our handling for LINKTOFILE is the same as OPENFILE.
				// A storage-based project system which handles OPENFILE by copying the file into the project directory
				// would have distinct handling for LINKTOFILE vs. OPENFILE.
			case VSADDITEMOP_OPENFILE:
			{
				// Add Existing File
				RETURN_HR_IF(E_INVALIDARG, pszItemName != nullptr);
				for (DWORD i = 0; i < cFilesToOpen; i++)
				{
					com_ptr<IChildNode> pNewNode;
					hr = AddExistingFile(location, rgpszFilesToOpen[i], &pNewNode); RETURN_IF_FAILED_EXPECTED(hr);
				}

				if (pResult)
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
		HRESULT hr;

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) != S_OK)
			return E_INVALIDARG;

		com_ptr<IFileNode> file;
		hr = d->QueryInterface(IID_PPV_ARGS(&file)); RETURN_IF_FAILED_EXPECTED(hr);
		wil::unique_bstr mkDocument;
		hr = file->GetMkDocument(this, &mkDocument); RETURN_IF_FAILED_EXPECTED(hr);

		com_ptr<IVsUIShellOpenDocument> uiShellOpenDocument;
		hr = serviceProvider->QueryService(SID_SVsUIShellOpenDocument, &uiShellOpenDocument); RETURN_IF_FAILED(hr);

		hr = uiShellOpenDocument->OpenSpecificEditor(0, mkDocument.get(), rguidEditorType, pszPhysicalView, rguidLogicalView,
			L"%3", this, itemid, punkDocDataExisting, serviceProvider, ppWindowFrame); RETURN_IF_FAILED_EXPECTED(hr);

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
			auto hr = c->AsProjectConfigProperties()->get_ConfigName(&n); RETURN_IF_FAILED(hr);
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
			auto hr = c->AsProjectConfigProperties()->get_PlatformName(&n); RETURN_IF_FAILED(hr);
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
			auto hr = cfg->AsProjectConfigProperties()->get_ConfigName(&cn); RETURN_IF_FAILED(hr);
			hr = cfg->AsProjectConfigProperties()->get_PlatformName(&pn); RETURN_IF_FAILED(hr);
			if (!wcscmp(pszCfgName, cn.get()) && !wcscmp(pszPlatformName, pn.get()))
			{
				auto hr = cfg->QueryInterface(ppCfg); RETURN_IF_FAILED(hr);
				return S_OK;
			}
		}

		return E_INVALIDARG;
	}

	virtual HRESULT STDMETHODCALLTYPE AddCfgsOfCfgName(LPCOLESTR pszCfgName, LPCOLESTR pszCloneCfgName, BOOL fPrivate) override
	{
		com_ptr<IProjectConfig> newConfig;
		auto hr = MakeProjectConfig (&newConfig); RETURN_IF_FAILED_EXPECTED(hr);

		if (pszCloneCfgName)
		{
			com_ptr<IProjectConfigProperties> existing;
			for (auto& c : _configs)
			{
				wil::unique_bstr name;
				hr = c->AsProjectConfigProperties()->get_ConfigName(&name); RETURN_IF_FAILED(hr);
				if (wcscmp(name ? name.get() : L"", pszCloneCfgName) == 0)
				{
					existing = c->AsProjectConfigProperties();
					break;
				}
			}

			RETURN_HR_IF(E_INVALIDARG, !existing);

			auto stream = com_ptr(SHCreateMemStream(nullptr, 0)); RETURN_IF_NULL_ALLOC(stream);
		
			hr = SaveToXml(existing, L"Temp", 0, stream); RETURN_IF_FAILED_EXPECTED(hr);

			hr = stream->Seek({ 0 }, STREAM_SEEK_SET, nullptr); RETURN_IF_FAILED(hr);

			hr = LoadFromXml(newConfig->AsProjectConfigProperties(), L"Temp", stream); RETURN_IF_FAILED_EXPECTED(hr);
		}

		auto newConfigNameBstr = wil::make_bstr_nothrow(pszCfgName); RETURN_IF_NULL_ALLOC(newConfigNameBstr);
		hr = newConfig->AsProjectConfigProperties()->put_ConfigName(newConfigNameBstr.get()); RETURN_IF_FAILED(hr);

		bool pushed = _configs.try_push_back(std::move(newConfig)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		hr = _configs.back()->SetSite(this);
		if (FAILED(hr))
		{
			_configs.remove_back();
			RETURN_HR(hr);
		}

		_isDirty = true;

		for (auto& sink : _cfgProviderEventSinks)
			sink.second->OnCfgNameAdded(pszCfgName);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE DeleteCfgsOfCfgName(LPCOLESTR pszCfgName) override
	{
		HRESULT hr;

		auto it = _configs.begin();
		while (it != _configs.end())
		{
			wil::unique_bstr name;
			hr = it->get()->AsProjectConfigProperties()->get_ConfigName(&name); RETURN_IF_FAILED(hr);
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
		HRESULT hr;

		RETURN_HR_IF(E_INVALIDARG, !pszOldName);
		RETURN_HR_IF(E_INVALIDARG, !pszNewName);
		RETURN_HR_IF(E_INVALIDARG, !pszNewName[0]);

		auto newNameBstr = wil::make_bstr_nothrow(pszNewName); RETURN_IF_NULL_ALLOC(newNameBstr);

		bool any = false;
		for (auto& c : _configs)
		{
			IProjectConfigProperties* props = c->AsProjectConfigProperties();
			wil::unique_bstr n;
			hr = props->get_ConfigName(&n); RETURN_IF_FAILED(hr);
			if (!wcscmp(n.get(), pszOldName))
			{
				hr = props->put_ConfigName(newNameBstr.get()); RETURN_IF_FAILED(hr);
				any = true;
			}
		}

		RETURN_HR_IF(E_INVALIDARG, !any);

		for (auto& sink : _cfgProviderEventSinks)
			sink.second->OnCfgNameRenamed(pszOldName, pszNewName);

		return S_OK;
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
		HRESULT hr;

		for (ULONG i = 0; i < cItems; i++)
			RETURN_HR_IF(E_NOTIMPL, itemid[i] == VSITEMID_ROOT); // see example in PrjHier.cpp

		com_ptr<IVsSolution> solution;
		hr = serviceProvider->QueryService(SID_SVsSolution, &solution); RETURN_IF_FAILED(hr);

		com_ptr<IVsRunningDocumentTable> rdt;
		hr = serviceProvider->QueryService(SID_SVsRunningDocumentTable, &rdt); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IFileOperation> pfo;
		hr = CoCreateInstance(__uuidof(FileOperation), NULL, CLSCTX_ALL, IID_PPV_ARGS(&pfo)); RETURN_IF_FAILED(hr);
		DWORD fof;
		if (dwFlags & DHO_SUPPRESS_UI)
		{
			hr = pfo->SetOperationFlags(FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT); RETURN_IF_FAILED(hr);
		}
		else
		{
			com_ptr<IVsUIShell> shell;
			auto hr = serviceProvider->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);
			HWND ownerHwnd;
			hr = shell->GetDialogOwnerHwnd(&ownerHwnd); RETURN_IF_FAILED(hr);
			hr = pfo->SetOwnerWindow(ownerHwnd); RETURN_IF_FAILED(hr);
			hr = pfo->SetOperationFlags(FOF_NOCONFIRMATION); RETURN_IF_FAILED(hr);
		}

		for (ULONG i = 0; i < cItems; i++)
		{
			com_ptr<IChildNode> d;
			hr = FindDescendant(itemid[i], &d); 
			RETURN_HR_IF(E_INVALIDARG, hr != S_OK);

			wil::unique_process_heap_string mk;
			hr = GetPathOf(d, mk); RETURN_IF_FAILED(hr);

			com_ptr<IVsHierarchy> docHier;
			VSITEMID docItemId;
			com_ptr<IUnknown> docData;
			VSCOOKIE docCookie;
			hr = rdt->FindAndLockDocument (RDT_NoLock, (LPCOLESTR)mk.get(), &docHier, &docItemId, &docData, &docCookie);
			if (hr == S_OK && docHier.get() == this && docItemId == itemid[i])
			{
				auto slnCloseOpts = (dwDelItemOp == DELITEMOP_DeleteFromStorage) ? SLNSAVEOPT_NoSave : SLNSAVEOPT_PromptSave;
				hr = solution->CloseSolutionElement (slnCloseOpts, nullptr, docCookie); RETURN_IF_FAILED_EXPECTED(hr);
			}

			hr = RemoveChildFromParent(this, d); RETURN_IF_FAILED(hr);

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

			_isDirty = true;
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IVsPersistHierarchyItem
	virtual HRESULT STDMETHODCALLTYPE IsItemDirty (VSITEMID itemid, IUnknown *punkDocData, BOOL *pfDirty) override
	{
		// VS calls this to find out if files open in the editor are dirty.
		// It doesn't seem to ever call this for VSITEMID_ROOT.
		RETURN_HR_IF(E_NOTIMPL, itemid == VSITEMID_ROOT);

		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_SELECTION);
		RETURN_HR_IF_EXPECTED(E_NOTIMPL, itemid == VSITEMID_NIL);

		com_ptr<IChildNode> d;
		if (FindDescendant(itemid, &d) == S_OK)
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
				hr = serviceProvider->QueryService(SID_SVsUIShell, &shell);
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

	#pragma region IProjectNodeProperties
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
		auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound)); RETURN_HR_IF(E_OUTOFMEMORY, !sa);
		for (LONG i = 0; i < (LONG)bound.cElements; i++)
		{
			com_ptr<IDispatch> pdisp;
			auto hr = _configs[i]->QueryInterface(&pdisp); RETURN_IF_FAILED(hr);
			hr = SafeArrayPutElement(sa.get(), &i, pdisp.get()); RETURN_IF_FAILED(hr);
		}
		*configs = sa.release();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Configurations (SAFEARRAY* sa) override
	{
		VARTYPE vt;
		auto hr = SafeArrayGetVartype(sa, &vt); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, vt != VT_DISPATCH);
		UINT dim = SafeArrayGetDim(sa);
		RETURN_HR_IF(E_NOTIMPL, dim != 1);
		LONG lbound;
		hr = SafeArrayGetLBound(sa, 1, &lbound); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_NOTIMPL, lbound != 0);
		LONG ubound;
		hr = SafeArrayGetUBound(sa, 1, &ubound); RETURN_IF_FAILED(hr);

		for (LONG i = 0; i <= ubound; i++)
		{
			com_ptr<IDispatch> child;
			hr = SafeArrayGetElement (sa, &i, child.addressof()); RETURN_IF_FAILED(hr);
			com_ptr<IProjectConfig> config;
			hr = child->QueryInterface(&config); RETURN_IF_FAILED(hr);
			bool pushed = _configs.try_push_back(std::move(config)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}

		// This is meant to be called only from LoadXml, no need to set dirty flag or send notifications.

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

	virtual HRESULT STDMETHODCALLTYPE get_AutoOpenFiles (BSTR *pbstrFilenames) override
	{
		// Never save this to XML. It's supposed to be added by hand in template files only.
		*pbstrFilenames = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_AutoOpenFiles (BSTR bstrFilenames) override
	{
		if (!bstrFilenames || !bstrFilenames[0])
			return (_autoOpenFiles.reset()), S_OK;

		auto f = wil::make_bstr_nothrow(bstrFilenames); RETURN_IF_NULL_ALLOC(f);
		_autoOpenFiles = std::move(f);
		return S_OK;
	}
	#pragma endregion

	#pragma region IXmlParent
	virtual HRESULT STDMETHODCALLTYPE GetChildXmlElementName (DISPID dispidProperty, IDispatch* child, BSTR* xmlElementNameOut) override
	{
		if (dispidProperty == dispidConfigurations)
		{
			*xmlElementNameOut = SysAllocString(ConfigurationElementName); RETURN_IF_NULL_ALLOC(*xmlElementNameOut);
			return S_OK;
		}

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

		if (dispidProperty == dispidConfigurations)
		{
			com_ptr<IProjectConfig> config;
			hr = MakeProjectConfig(&config); RETURN_IF_FAILED(hr);
			com_ptr<IProjectConfigProperties> props;
			hr = config->QueryInterface(&props);
			*childOut = props.detach();
			return S_OK;
		}

		if (dispidProperty == dispidItems)
		{
			// Files saved by newer versions use "File" for the name of the XML element.
			// For backward compatibility, we also recognize the name "AsmFile".
			if (!wcscmp(xmlElementName, FileElementName) || !wcscmp(xmlElementName, L"AsmFile"))
			{
				wil::com_ptr_nothrow<IFileNode> file;
				hr = MakeFileNode(&file); RETURN_IF_FAILED(hr);
				com_ptr<IFileNodeProperties> fileProps;
				hr = file->QueryInterface(&fileProps); RETURN_IF_FAILED(hr);
				if (!wcscmp(xmlElementName, L"AsmFile"))
					fileProps->put_BuildTool(BuildToolKind::Assembler);
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

	#pragma region IParentNode
	virtual VSITEMID STDMETHODCALLTYPE GetItemId() override { return VSITEMID_ROOT; }

	virtual IChildNode* STDMETHODCALLTYPE FirstChild() override
	{
		return _firstChild;
	}

	virtual void STDMETHODCALLTYPE SetFirstChild (IChildNode *child) override
	{
		_firstChild = child;
	}
	#pragma endregion

	#pragma region IProjectNode
	virtual VSITEMID MakeItemId() override
	{
		return _nextItemId++;
	}

	virtual HRESULT GetAutoOpenFiles (BSTR* pbstrFilenames) override
	{
		if (!_autoOpenFiles || !_autoOpenFiles.get()[0])
			return E_NOT_SET;

		auto f = SysAllocString(_autoOpenFiles.get()); RETURN_IF_NULL_ALLOC(f);
		*pbstrFilenames = f;
		return S_OK;
	}

	virtual IParentNode* AsParentNode() override { return this; }

	virtual IVsUIHierarchy* AsHierarchy() override { return this; }

	virtual IVsProject* AsVsProject() override { return this; }
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

	#pragma region IVsPerPropertyBrowsing
	virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL *pfHide) override
	{
		if (dispid == dispidAutoOpenFiles)
			return (*pfHide = TRUE), S_OK;

		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { return E_NOTIMPL; }

	virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
	{
		if (dispid == dispidAutoOpenFiles)
			// Never save this to XML. It's supposed to be added by hand in template files only.
			return (*fDefault = TRUE), S_OK;

		return E_NOTIMPL;
	}

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

	#pragma region IVsUpdateSolutionEvents
	virtual HRESULT STDMETHODCALLTYPE UpdateSolution_Begin (BOOL *pfCancelUpdate) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UpdateSolution_Done (BOOL fSucceeded, BOOL fModified, BOOL fCancelCommand) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UpdateSolution_StartUpdate (BOOL *pfCancelUpdate) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UpdateSolution_Cancel() override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnActiveProjectCfgChange (IVsHierarchy *pIVsHierarchy) override
	{
		HRESULT hr;

		if (pIVsHierarchy == this)
		{
			hr = GeneratePrePostIncludeFiles(this, nullptr); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IVsHierarchyEvents
	virtual HRESULT STDMETHODCALLTYPE OnItemAdded (VSITEMID itemidParent, VSITEMID itemidSiblingPrev, VSITEMID itemidAdded) override
	{
		for (auto& sink : _hierarchyEventSinks)
			sink.second->OnItemAdded (itemidParent, itemidSiblingPrev, itemidAdded);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemsAppended (VSITEMID itemidParent) override
	{
		for (auto& sink : _hierarchyEventSinks)
			sink.second->OnItemsAppended(itemidParent);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemDeleted (VSITEMID itemid) override
	{
		for (auto& sink : _hierarchyEventSinks)
			sink.second->OnItemDeleted(itemid);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged (VSITEMID itemid, VSHPROPID propid, DWORD flags) override
	{
		for (auto& sink : _hierarchyEventSinks)
			sink.second->OnPropertyChanged(itemid, propid, flags);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnInvalidateItems (VSITEMID itemidParent) override
	{
		for (auto& sink : _hierarchyEventSinks)
			sink.second->OnInvalidateItems(itemidParent);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnInvalidateIcon (HICON hicon) override
	{
		for (auto& sink : _hierarchyEventSinks)
			sink.second->OnInvalidateIcon(hicon);
		return S_OK;
	}
	#pragma endregion

	HRESULT RefreshHierarchy()
	{
		com_ptr<IVsUIShell> shell;
		auto hr = serviceProvider->QueryService(SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);
		LONG result;
		hr = shell->ShowMessageBox (0, GUID_NULL, L"title", L"Refresh not yet implemented", 0, 0,
			OLEMSGBUTTON_OK, OLEMSGDEFBUTTON_FIRST, OLEMSGICON_INFO, FALSE, &result);
		return S_OK;
	}

	HRESULT ProcessCommandAddNewFolder (VSITEMID parentItemId, VARIANT* pvaOut)
	{
		HRESULT hr;

		com_ptr<IParentNode> parent;
		wil::unique_process_heap_string parentPath;
		if (parentItemId == VSITEMID_ROOT)
		{
			parent = this;
			parentPath = wil::make_process_heap_string_nothrow(_projectDir.get()); RETURN_IF_NULL_ALLOC(parentPath);
		}
		else
		{
			com_ptr<IChildNode> node;
			hr = FindDescendant(parentItemId, &node); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_INVALIDARG, hr == S_FALSE);
			hr = node->QueryInterface(IID_PPV_ARGS(&parent)); RETURN_IF_FAILED(hr);
			hr = GetPathOf (node, parentPath); RETURN_IF_FAILED(hr);
		}

		com_ptr<IVsShell> shell;
		hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);
		wil::unique_bstr newFolderName;
		hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_NEW_FOLDER_NAME, &newFolderName); RETURN_IF_FAILED(hr);

		wchar_t dirName[20];
		for (uint32_t i = 1;;)
		{
			swprintf_s(dirName, newFolderName.get(), i);

			// Search for a node with the same name in the same location (not deeper).
			bool nameExists = false;
			for (auto c = parent->FirstChild(); !!c; c = c->Next())
			{
				wil::unique_variant nameVar;
				hr = c->GetProperty(VSHPROPID_Name, &nameVar); RETURN_IF_FAILED(hr);
				if (!_wcsicmp(nameVar.bstrVal, dirName))
				{
					nameExists = true;
					break;
				}
			}
			
			if (!nameExists)
			{
				wil::unique_hlocal_string dirPath;
				hr = wil::str_concat_nothrow(dirPath, parentPath, L"\\", dirName); RETURN_IF_FAILED(hr);
				int ires = SHCreateDirectoryExW(nullptr, dirPath.get(), nullptr);
				if (ires == 0 || ires == ERROR_ALREADY_EXISTS)
					break;
				RETURN_WIN32(ires);
			}

			i++;
			if (i == 100)
				RETURN_HR(E_UNEXPECTED);
		}

		com_ptr<IFolderNode> newFolder;
		hr = GetOrCreateChildFolder (parent, dirName, false, &newFolder); RETURN_IF_FAILED(hr);

		_isDirty = true;

		com_ptr<IVsUIHierarchyWindow> uiWindow;
		if (SUCCEEDED(GetHierarchyWindow(uiWindow.addressof())))
		{
			// we need to get into label edit mode now...
			// so first select the new guy...
			if (SUCCEEDED(uiWindow->ExpandItem(this, newFolder->GetItemId(), EXPF_SelectItem)))
			{
				// them post the rename command to the shell. Folder verification and creation will
				// happen in the setlabel code...
				com_ptr<IVsUIShell> shell;
				if (SUCCEEDED(serviceProvider->QueryService(SID_SVsUIShell, IID_PPV_ARGS(shell.addressof()))))
				{
					wil::unique_variant dummy;
					shell->PostExecCommand (&CMDSETID_StandardCommandSet97, cmdidRename, 0, &dummy);
				}
			}
		}

		// We return the item ID just for testing purposes.
		if (pvaOut)
			InitVariantFromVSITEMID(newFolder->GetItemId(), pvaOut);

		return S_OK;
	}
};

FELIX_API HRESULT MakeProjectNode (LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags, REFIID iidProject, void** ppvProject)
{
	wil::com_ptr_nothrow<ProjectNode> p = new (std::nothrow) ProjectNode(); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance (pszFilename, pszLocation, pszName, grfCreateFlags); RETURN_IF_FAILED_EXPECTED(hr);
	return p->QueryInterface (iidProject, ppvProject);
}
