
#pragma once
#include "FelixPackage_h.h"
#include "Simulator.h"

// We use DLL exports only to be able to call various functions from unit test projects.
#ifdef FELIX_EXPORTS
#define FELIX_API __declspec(dllexport)
#else
#define FELIX_API __declspec(dllimport)
#endif

using unique_safearray = wil::unique_any<SAFEARRAY*, decltype(SafeArrayDestroy), &SafeArrayDestroy>;

__interface IProjectConfigBuilderCallback : IUnknown
{
	HRESULT OnBuildComplete (bool success);
};

// Purpose of this interface is to allow separation of "project config"
// functionality (mostly a VS GUI thing) from the "build" functionality (a "felix" thing).
__interface IProjectConfigBuilder : IUnknown
{
	// The build runs synchronously if possible, in which case this function
	// calls the callback and then returns S_OK.
	// 
	// If the build cannot run synchronously, this function returns S_FALSE and the build
	// runs in the background. The message loop must be pumped for the build to run in the background.
	// In this case the implementation calls the callback at a later time (provided the message
	// loop is pumped). If this function is called and afterwards the last reference
	// to the object is released while the build is running (i.e., before the callback is called
	// and before CancelBuild is called), then the callback will not be called.
	//
	// If this function fails, it returns an error code and doesn't call the callback.
	HRESULT STDMETHODCALLTYPE StartBuild (IProjectConfigBuilderCallback* callback);

	// This function can be called while a build is running in the background - that is,
	// after StartBuild returned S_FALSE and before the callback was called.
	// The implementation calls the callback before this function returns, either with
	// fSuccess=TRUE if the build completed successfully before calling this function,
	// or with fSuccess=FALSE otherwise.
	HRESULT STDMETHODCALLTYPE CancelBuild();
};

struct IChildNode;

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("29DDAB6E-5A2E-4BD8-A617-E1EAA90E30DA")
INode : IUnknown
{
	virtual VSITEMID GetItemId() = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("D930CCA1-E515-4569-8DC6-959CD6367654")
IParentNode : IUnknown
{
	virtual VSITEMID GetItemId() = 0;
	virtual IChildNode* FirstChild() = 0;
	virtual void SetFirstChild (IChildNode* child) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("F36A3A6C-01AF-423B-86FD-DB071AA47E97")
IProjectNode : INode
{
	virtual VSITEMID MakeItemId() = 0;
	virtual HRESULT GetAutoOpenFiles (BSTR* pbstrFilenames) = 0; // TODO: AsProjectNodeProperties()->get_AutoOpenFile
	virtual IParentNode* AsParentNode() = 0;
	virtual IVsUIHierarchy* AsHierarchy() = 0;
	virtual IVsProject* AsVsProject() = 0;
};

// This is the base interface for every node except IProjectNode.
struct DECLSPEC_NOVTABLE DECLSPEC_UUID("A2EE7852-34B1-49A9-A3DB-36232AC6680C")
IChildNode : INode
{
	virtual HRESULT GetParent (IParentNode** ppParent) = 0;
	virtual HRESULT SetItemId (IParentNode* parent, VSITEMID itemId) = 0;
	virtual HRESULT ClearItemId() = 0;
	virtual IChildNode* Next() = 0; // TODO: keep an unordered_map with itemid/itemptr, then get rid of Next, SetNext, FindDescendant
	virtual void SetNext (IChildNode* next) = 0;
	virtual HRESULT GetProperty (VSHPROPID propid, VARIANT* pvar) = 0;
	virtual HRESULT SetProperty (VSHPROPID propid, REFVARIANT var) = 0;
	virtual HRESULT GetGuidProperty (VSHPROPID propid, GUID* pguid) = 0;
	virtual HRESULT SetGuidProperty (VSHPROPID propid, REFGUID rguid) = 0;
	virtual HRESULT GetCanonicalName (BSTR* pbstrName) = 0; // returns the path relative to project if possible, otherwise the full path -- all lowercase
	virtual HRESULT IsItemDirty (IUnknown *punkDocData, BOOL *pfDirty) = 0;
	virtual HRESULT QueryStatusCommand (const GUID* pguidCmdGroup, OLECMD* pCmd, OLECMDTEXT *pCmdText) = 0;
	virtual HRESULT ExecCommand (const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn, VARIANT* pvaOut) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("5F6EA158-4DA8-469A-8FD8-E8C04F31244E")
IFileNode : IChildNode
{
	virtual HRESULT STDMETHODCALLTYPE GetMkDocument (IVsHierarchy* hier, BSTR* pbstrMkDocument) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("E5498B79-7C01-4F49-B5EC-8D1C98FF935D")
IFolderNode : IChildNode
{
	virtual IParentNode* AsParentNode() = 0;
	virtual IFolderNodeProperties* AsFolderNodeProperties() = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("56831DCD-0782-48BE-BF8A-57827FC0D6CA")
IProjectConfig : IUnknown
{
	virtual HRESULT SetSite (IProjectNode* proj) = 0;
	virtual HRESULT GetSite (REFIID riid, void** ppvObject) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetOutputDirectory (BSTR* pbstr) = 0;
	virtual IProjectConfigAssemblerProperties* AsmProps() = 0;
	virtual IProjectConfigProperties* AsProjectConfigProperties() = 0;
	virtual IVsProjectCfg* AsVsProjectConfig() = 0;
};

enum SymbolKind : DWORD
{
	SK_None = 0,
	SK_Code = 1,
	SK_Data = 2,
	SK_Both = 3
};

#define E_UNRECOGNIZED_SLD_VERSION           MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x201)
#define E_INVALID_SLD_LINE                   MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x202)
#define E_UNRECOGNIZED_Z80SYM_VERSION        MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x203)
#define E_INVALID_Z80SYM_LINE                MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x204)
#define E_SYMBOL_NOT_IN_SYMBOL_FILE          MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x205)

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("32BEBBF2-86DF-4D79-88DF-1123548C4D8E") IFelixSymbols : IUnknown
{
	// Returns S_OK if the symbol file contains mapping between source code lines and instruction addresses;
	// in this case GetSourceLocationFromAddress and GetAddressFromSourceLocation can be called to attempt
	// mapping between lines and addresses.
	// 
	// Returns S_FALSE if the symbol file doesn't contain this mapping;
	// in this case GetSourceLocationFromAddress and GetAddressFromSourceLocation return E_NOTIMPL.
	// 
	// TODO: implement this as an optional interface on the same object.
	virtual HRESULT STDMETHODCALLTYPE HasSourceLocationInformation() = 0;

	virtual HRESULT STDMETHODCALLTYPE GetSourceLocationFromAddress(
		__RPC__in uint16_t address,
		__RPC__deref_out BSTR* srcFilename,
		__RPC__out UINT32* srcLineIndex) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSourceLocation(
		__RPC__in LPCWSTR src_filename,
		__RPC__in uint32_t line_index,
		__RPC__out UINT16* address_out) = 0;

	// Finds the symbol (code and/or data) located at a given address.
	// If "offset" is NULL, looks only for an exact address match.
	virtual HRESULT STDMETHODCALLTYPE GetSymbolAtAddress(
		__RPC__in uint16_t address,
		__RPC__in SymbolKind searchKind,
		__RPC__deref_out_opt SymbolKind* foundKind,
		__RPC__deref_out_opt BSTR* foundSymbol,
		__RPC__deref_out_opt UINT16* foundOffset) = 0;

	// Returns S_OK or an error if not found.
	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSymbol (LPCWSTR symbolName, UINT16* address) = 0;
};

FELIX_API extern wil::com_ptr_nothrow<IServiceProvider> serviceProvider;
extern wil::com_ptr_nothrow<ISimulator> simulator;
extern wil::com_ptr_nothrow<ISimulator> _simulator;

extern const wchar_t Z80AsmLanguageName[];
extern const wchar_t SingleDebugPortName[];
extern const GUID Z80AsmLanguageGuid;
extern const wchar_t SettingsCollection[];
extern const wchar_t SettingLoadSavePath[];
extern const LCID InvariantLCID;
FELIX_API extern const wchar_t ProjectElementName[];
extern const wchar_t ConfigurationElementName[];
extern const wchar_t FileElementName[];
extern const wchar_t FolderElementName[];

const char* PropIDToString (VSHPROPID propid);
void PrintProperty (const char* prefix, VSHPROPID propid, const VARIANT* pvar);
HRESULT MakeBstrFromString (const char* name, BSTR* bstr);
HRESULT MakeBstrFromString (const char* name, size_t len, BSTR* bstr);
HRESULT MakeBstrFromString (const char* sl_name_from, const char* sl_name_to, BSTR* to);
FELIX_API HRESULT MakeBstrFromStreamOnHGlobal (IStream* stream, BSTR* pBstr);

template<typename from_string_type, typename to_string_type>
HRESULT UTF8ToUTF16 (const from_string_type& fromStringUTF8, to_string_type& toStringUTF16)
{
	const char* fromRaw = fromStringUTF8.get();
	size_t fromLen = strlen(fromRaw);
	if (fromLen == 0)
	{
		auto toString = wil::make_unique_string_nothrow<to_string_type>(L""); RETURN_IF_NULL_ALLOC(toString);
		toStringUTF16 = std::move(toString);
		return S_OK;
	}

	// TODO: start with a buffer with the same number of characters, which will almost always be enough.
	int wcharCount = MultiByteToWideChar (CP_UTF8, 0, fromRaw, (int)fromLen, nullptr, 0); RETURN_LAST_ERROR_IF(wcharCount == 0);
	auto toString = wil::make_unique_string_nothrow<to_string_type>(nullptr, wcharCount); RETURN_IF_NULL_ALLOC(toString);
	int ires = MultiByteToWideChar (CP_UTF8, 0, fromRaw, (int)fromLen, toString.get(), wcharCount); RETURN_LAST_ERROR_IF(ires == 0);
	toString.get()[wcharCount] = 0;
	toStringUTF16 = std::move(toString);
	return S_OK;
}

template<typename from_string_type, typename to_string_type>
HRESULT UTF16ToUTF8 (const from_string_type& fromStringUTF16, to_string_type& toStringUTF8, size_t* toStringLen = nullptr)
{
	const wchar_t* fromRaw = wil::str_raw_ptr(fromStringUTF16);
	size_t fromLen = wcslen(fromRaw);
	if (fromLen == 0)
	{
		auto toString = wil::make_unique_ansistring_nothrow<to_string_type>(""); RETURN_IF_NULL_ALLOC(toString);
		toStringUTF8 = std::move(toString);
		return S_OK;
	}

	// TODO: start with a buffer with the same number of characters, which will almost always be enough.
	int mblen = WideCharToMultiByte (CP_UTF8, 0, fromRaw, (int)fromLen, nullptr, 0, nullptr, nullptr); RETURN_LAST_ERROR_IF(mblen == 0);
	auto mb = wil::make_unique_ansistring_nothrow<to_string_type>(nullptr, mblen); RETURN_IF_NULL_ALLOC(mb);
	int ires = WideCharToMultiByte (CP_UTF8, 0, fromRaw, (int)fromLen, mb.get(), mblen, nullptr, nullptr); RETURN_LAST_ERROR_IF(ires == 0);
	mb.get()[mblen] = 0;
	toStringUTF8 = std::move(mb);
	if (toStringLen)
		*toStringLen = mblen;
	return S_OK;
}

VARIANT MakeVariantFromVSITEMID (VSITEMID itemid);
inline HRESULT InitVariantFromVSITEMID (VSITEMID itemid, VARIANT* pvar)
{
	pvar->vt = VT_VSITEMID;
	V_VSITEMID(pvar) = itemid;
	return S_OK;
}

FELIX_API HRESULT MakeProjectNode (LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags, REFIID iidProject, void** ppvProject);
FELIX_API HRESULT MakeFileNode (IFileNode** file);
HRESULT MakeProjectConfig (IProjectConfig** to);
HRESULT MakeProjectFactory (IVsProjectFactory** to);
HRESULT MakePGPropertyPage (UINT titleStringResId, REFGUID pageGuid, DISPID dispidChildObj, IPropertyPage** to);
HRESULT MakeAsmPropertyPage (IPropertyPage** to);
HRESULT SimulatorWindowPane_CreateInstance (IVsWindowPane** to);
HRESULT Z80AsmLanguageInfo_CreateInstance (IVsLanguageInfo** to);
HRESULT MakeDebugPortSupplier (IDebugPortSupplier2** to);
HRESULT MakeDebugEngine (IDebugEngine2** to);
HRESULT MakeLaunchOptions (IFelixLaunchOptions** ppOptions);
HRESULT GetDefaultProjectFileExtension (BSTR* ppExt);
HRESULT SetErrorInfo0 (HRESULT errorHR, ULONG packageStringResId);
HRESULT SetErrorInfo1 (HRESULT errorHR, ULONG packageStringResId, LPCWSTR arg1);
FELIX_API HRESULT MakeCustomBuildToolProperties (ICustomBuildToolProperties** to);
FELIX_API HRESULT MakeProjectConfigBuilder (IProjectNode* project, IProjectConfig* config,
	IVsOutputWindowPane2* outputWindowPane, IProjectConfigBuilder** to);
HRESULT ShowCommandLinePropertyBuilder (HWND hwndParent, BSTR valueBefore, BSTR* valueAfter);
HRESULT GeneratePrePostIncludeFiles (IProjectNode* project, IProjectConfig* configOrNullForActive);
FELIX_API HRESULT MakeSjasmCommandLine (IVsHierarchy* hier, IProjectConfig* config, IProjectConfigAssemblerProperties* asmPropsOverride, BSTR* ppCmdLine);
HRESULT MakeFolderNode (IFolderNode** ppFolder);
BOOL LUtilFixFilename (wchar_t* strName);
HRESULT QueryEditProjectFile (IVsHierarchy* hier);
HRESULT GetHierarchyWindow (IVsUIHierarchyWindow** ppHierWindow);
HRESULT GetPathTo (IChildNode* node, wil::unique_process_heap_string& dir, bool relativeToProjectDir = false);
HRESULT GetPathOf (IChildNode* node, wil::unique_process_heap_string& path, bool relativeToProjectDir = false);
HRESULT FindHier (IChildNode* from, REFIID riid, void** ppvHier);
HRESULT FindHier (IParentNode* from, REFIID riid, void** ppvHier);
HRESULT AddFileToParent (IFileNode* child, IParentNode* addTo);
FELIX_API HRESULT GetOrCreateChildFolder (IParentNode* parent, const wchar_t* folderName, bool createDirectoryOnFileSystem, IFolderNode** ppFolder);
HRESULT RemoveChildFromParent (IProjectNode* root, IChildNode* child);
HRESULT CreatePathOfNode (IParentNode* node, wil::unique_process_heap_string& pathOut);
HRESULT GetItems (IParentNode* itemsIn, SAFEARRAY** itemsOut);
HRESULT PutItems (SAFEARRAY* sa, IParentNode* items);
HRESULT CreateFileFromTemplate (LPCWSTR fromPath, LPCWSTR toPath, IProjectConfig* config);
IFileNode* FindChildFileByName (IParentNode* parent, const wchar_t* fileName);
HRESULT MakeFileNodeForExistingFile (LPCWSTR path, IFileNode** ppFile);
HRESULT ParseNumber (LPCWSTR str, DWORD* value); // returns S_OK or S_FALSE
HRESULT MakeSldSymbols (const wchar_t* symbolsFullPath, IFelixSymbols** to);
HRESULT MakeZ80SymSymbols (const wchar_t* symbolsFullPath, IFelixSymbols** to);
