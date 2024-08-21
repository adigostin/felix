
#pragma once
#include "FelixPackage_h.h"
#include "Simulator/Simulator.h"

using unique_safearray = wil::unique_any<SAFEARRAY*, decltype(SafeArrayDestroy), &SafeArrayDestroy>;

extern wil::com_ptr_nothrow<IServiceProvider> serviceProvider;

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

inline VARIANT MakeVariantFromVSITEMID (VSITEMID itemid)
{
	VARIANT variant;
	variant.vt = VT_VSITEMID;
	V_VSITEMID(&variant) = itemid;
	return variant;
}

HRESULT MakeFelixProject (IServiceProvider* sp, LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags, REFIID iidProject, void** ppvProject);
HRESULT MakeZ80AsmFile (VSITEMID itemId, IVsUIHierarchy* hier, VSITEMID parentItemId, ITypeLib* typeLib, IZ80AsmFile** file);
HRESULT Z80ProjectConfig_CreateInstance (IVsUIHierarchy* hier, ITypeLib* typeLib, IZ80ProjectConfig** to);
HRESULT Z80ProjectFactory_CreateInstance (IServiceProvider* sp, IVsProjectFactory** to);
HRESULT MakePGPropertyPage (UINT titleStringResId, REFGUID pageGuid, DISPID dispidChildObj, IPropertyPage** to);
HRESULT SimulatorWindowPane_CreateInstance (IVsWindowPane** to);
HRESULT Z80AsmLanguageInfo_CreateInstance (IVsLanguageInfo** to);
HRESULT MakeDebugPortSupplier (IDebugPortSupplier2** to);
HRESULT MakeDebugEngine (IDebugEngine2** to);
HRESULT MakeLaunchOptions (IFelixLaunchOptions** ppOptions);
HRESULT GetDefaultProjectFileExtension (BSTR* ppExt);
