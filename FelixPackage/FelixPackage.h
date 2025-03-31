
#pragma once
#include "FelixPackage_h.h"
#include "Simulator.h"

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

extern wil::com_ptr_nothrow<IServiceProvider> serviceProvider;

constexpr DWORD PathFlags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;

extern GUID SID_Simulator;
extern const wchar_t Z80AsmLanguageName[];
extern const wchar_t SingleDebugPortName[];
extern const GUID Z80AsmLanguageGuid;
extern const wchar_t SettingsCollection[];
extern const wchar_t SettingLoadSavePath[];

const char* PropIDToString (VSHPROPID propid);
void PrintProperty (const char* prefix, VSHPROPID propid, const VARIANT* pvar);
HRESULT MakeBstrFromString (const char* name, BSTR* bstr);
HRESULT MakeBstrFromString (const char* name, size_t len, BSTR* bstr);
HRESULT MakeBstrFromString (const char* sl_name_from, const char* sl_name_to, BSTR* to);
HRESULT MakeBstrFromStreamOnHGlobal (IStream* stream, BSTR* pBstr);

inline HRESULT InitVariantFromVSITEMID (VSITEMID itemid, VARIANT* pvar)
{
	pvar->vt = VT_VSITEMID;
	V_VSITEMID(pvar) = itemid;
	return S_OK;
}

HRESULT MakeFelixProject (IServiceProvider* sp, LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags, REFIID iidProject, void** ppvProject);
HRESULT MakeProjectFile (VSITEMID itemId, IVsUIHierarchy* hier, VSITEMID parentItemId, IProjectFile** file);
HRESULT ProjectConfig_CreateInstance (IVsHierarchy* hier, IProjectConfig** to);
HRESULT Z80ProjectFactory_CreateInstance (IServiceProvider* sp, IVsProjectFactory** to);
HRESULT MakePGPropertyPage (UINT titleStringResId, REFGUID pageGuid, DISPID dispidChildObj, IPropertyPage** to);
HRESULT MakeAsmPropertyPage (IPropertyPage** to);
HRESULT SimulatorWindowPane_CreateInstance (IVsWindowPane** to);
HRESULT Z80AsmLanguageInfo_CreateInstance (IVsLanguageInfo** to);
HRESULT MakeDebugPortSupplier (IDebugPortSupplier2** to);
HRESULT MakeDebugEngine (IDebugEngine2** to);
HRESULT MakeLaunchOptions (IFelixLaunchOptions** ppOptions);
HRESULT GetDefaultProjectFileExtension (BSTR* ppExt);
HRESULT SetErrorInfo1 (HRESULT errorHR, ULONG packageStringResId, LPCWSTR arg1);
HRESULT AssemblerPageProperties_CreateInstance (IProjectConfig* config, IProjectConfigAssemblerProperties** to);
HRESULT DebuggingPageProperties_CreateInstance (IProjectConfigDebugProperties** to);
HRESULT MakeCustomBuildToolProperties (ICustomBuildToolProperties** to);
HRESULT MakeProjectConfigBuilder (IVsHierarchy* hier, IProjectConfig* config,
	IVsOutputWindowPane2* outputWindowPane, IProjectConfigBuilder** to);
HRESULT PrePostBuildPageProperties_CreateInstance (bool post, IProjectConfigPrePostBuildProperties** to);
HRESULT ShowCommandLinePropertyBuilder (HWND hwndParent, BSTR valueBefore, BSTR* valueAfter);
HRESULT MakeSjasmCommandLine (IVsHierarchy* hier, IProjectConfig* config, IProjectConfigAssemblerProperties* asmProps, BSTR* ppCmdLine);
