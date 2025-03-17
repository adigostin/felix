
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "Guids.h"
#include "../FelixPackageUi/Resource.h"
#include "dispids.h"

HRESULT FelixPackage_CreateInstance (IVsPackage** out);

BOOL APIENTRY DllMain(HMODULE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	wil::DLLMain (hinstDLL, fdwReason, lpvReserved);

	switch (fdwReason)
	{
		case DLL_PROCESS_ATTACH:
		{
			wchar_t buffer[MAX_PATH];
			GetProcessImageFileName(GetCurrentProcess(), buffer, MAX_PATH);
			wil::g_fBreakOnFailure = IsDebuggerPresent() && _wcsicmp(PathFindFileName(buffer), L"testhost.exe");
			break;
		}
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
	}

	return TRUE;
}

extern "C" HRESULT __stdcall DllGetClassObject (REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	*ppv = nullptr;

	wil::com_ptr_nothrow<IUnknown> p;
	if (rclsid == CLSID_FelixPackage)
	{
		p = new (std::nothrow) ClassObjectImpl<IVsPackage>(FelixPackage_CreateInstance); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == AssemblerPropertyPage_CLSID)
	{
		static const auto make = [](IPropertyPage** to)
			{ return MakePGPropertyPage(IDS_ASSEMBLER_PROP_PAGE_TITLE, AssemblerPropertyPage_CLSID, dispidAssemblerProperties, to); };
		p = new (std::nothrow) ClassObjectImpl<IPropertyPage>(make); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == DebugPropertyPage_CLSID)
	{
		static const auto make = [](IPropertyPage** to)
			{ return MakePGPropertyPage(IDS_DEBUGGING_PROP_PAGE_TITLE, DebugPropertyPage_CLSID, dispidDebuggingProperties, to); };
		p = new (std::nothrow) ClassObjectImpl<IPropertyPage>(make); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == PreBuildPropertyPage_CLSID)
	{
		static const auto make = [](IPropertyPage** to)
			{ return MakePGPropertyPage(IDS_PRE_BUILD_PROP_PAGE_TITLE, PreBuildPropertyPage_CLSID, dispidPreBuildProperties, to); };
		p = new (std::nothrow) ClassObjectImpl<IPropertyPage>(make); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == PostBuildPropertyPage_CLSID)
	{
		static const auto make = [](IPropertyPage** to)
			{ return MakePGPropertyPage(IDS_POST_BUILD_PROP_PAGE_TITLE, PostBuildPropertyPage_CLSID, dispidPostBuildProperties, to); };
		p = new (std::nothrow) ClassObjectImpl<IPropertyPage>(make); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == PortSupplier_CLSID)
	{
		p = new (std::nothrow) ClassObjectImpl(MakeDebugPortSupplier); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == Engine_CLSID)
	{
		p = new (std::nothrow) ClassObjectImpl(MakeDebugEngine);
	}

	else
		RETURN_HR(E_INVALIDARG);

	return p->QueryInterface(riid, ppv);
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
	return S_FALSE;
}
