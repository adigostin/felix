
#pragma once
#include "..\TestsCommon.h"
#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>
namespace VxDTE
{
	#include <dte80.h>
	#include <dte90.h>
}

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace UITests
{
	wil::com_ptr_failfast<VxDTE::DTE2> GetDefaultVSInstance();
	wil::com_ptr_failfast<VxDTE::DTE2> LaunchVS (const wchar_t* envVar = nullptr);
	void CloseVS (VxDTE::DTE2* dte, bool hard = false);

	inline VSITEMID GetProperty_VSITEMID (IVsHierarchy* hier, VSITEMID itemid, VSHPROPID propid)
	{
		wil::unique_variant var;
		auto hr = hier->GetProperty(itemid, propid, &var);
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::IsTrue(SUCCEEDED(hr));
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::AreEqual<VARTYPE>(VT_VSITEMID, var.vt);
		return V_VSITEMID(&var);
	}

	inline wil::com_ptr_failfast<IDispatch> GetProperty_Dispatch (IVsHierarchy* hier, VSITEMID itemid, VSHPROPID propid)
	{
		wil::unique_variant var;
		auto hr = hier->GetProperty(itemid, propid, &var);
		Assert::IsTrue(SUCCEEDED(hr));
		Assert::AreEqual<VARTYPE>(VT_DISPATCH, var.vt);
		wil::com_ptr_failfast<IDispatch> res;
		res.attach(var.release().pdispVal);
		return res;
	}

	inline wil::unique_bstr GetProperty_String (IVsHierarchy* hier, VSITEMID itemid, VSHPROPID propid)
	{
		wil::unique_variant var;
		auto hr = hier->GetProperty(itemid, propid, &var);
		Assert::IsTrue(SUCCEEDED(hr));
		Assert::AreEqual<VARTYPE>(VT_BSTR, var.vt);
		return wil::unique_bstr(var.release().bstrVal);
	}
}


