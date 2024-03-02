
#include "pch.h"
#include "FelixPackage.h"
#include "shared/ClassFactory.h"
#include "Guids.h"
#include "../FelixPackageUi/Resource.h"

HRESULT FelixPackage_CreateInstance (IVsPackage** out);

BOOL APIENTRY DllMain(HMODULE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	wil::DLLMain (hinstDLL, fdwReason, lpvReserved);

	switch (fdwReason)
	{
		case DLL_PROCESS_ATTACH:
			wil::g_fBreakOnFailure = IsDebuggerPresent();
			break;
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
		p = new (std::nothrow) Z80ClassFactory<IVsPackage>(FelixPackage_CreateInstance); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == GeneralPropertyPage_CLSID)
	{
		static const auto make = [](IPropertyPage** to)
			{ return MakePGPropertyPage(IDS_GENERAL_PROP_PAGE_TITLE, GeneralPropertyPage_CLSID, to); };
		p = new (std::nothrow) Z80ClassFactory<IPropertyPage>(make); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == DebugPropertyPage_CLSID)
	{
		static const auto make = [](IPropertyPage** to)
			{ return MakePGPropertyPage(IDS_DEBUGGING_PROP_PAGE_TITLE, DebugPropertyPage_CLSID, to); };
		p = new (std::nothrow) Z80ClassFactory<IPropertyPage>(make); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == PortSupplier_CLSID)
	{
		p = new (std::nothrow) Z80ClassFactory(MakeDebugPortSupplier); RETURN_IF_NULL_ALLOC(p);
	}
	else if (rclsid == Engine_CLSID)
	{
		p = new (std::nothrow) Z80ClassFactory(MakeDebugEngine);
	}

	else
		RETURN_HR(E_INVALIDARG);

	return p->QueryInterface(riid, ppv);
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
	return S_FALSE;
}
