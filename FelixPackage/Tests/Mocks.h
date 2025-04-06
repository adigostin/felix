
#pragma once
#include "shared/com.h"
#include "FelixPackage.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

extern wchar_t tempPath[MAX_PATH + 1];

extern com_ptr<IProjectConfig> MakeMockProjectConfig (IVsHierarchy* hier);
extern com_ptr<IVsHierarchy> MakeMockVsHierarchy (const wchar_t* projectDir);
extern com_ptr<IProjectFile> MakeMockSourceFile (BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir);
extern com_ptr<IVsOutputWindowPane2> MakeMockOutputWindowPane (IStream* outputStreamUTF16);
extern com_ptr<IServiceProvider> MakeMockServiceProvider();
com_ptr<IVsHierarchyEvents> MakeMockHierarchyEventSink();

extern void WriteFileOnDisk(const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir, const char* fileContent);
extern void DeleteFileOnDisk(const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir);

inline VSITEMID GetProperty_VSITEMID (IVsHierarchy* hier, VSITEMID itemid, VSHPROPID propid)
{
	wil::unique_variant var;
	auto hr = hier->GetProperty(itemid, propid, &var);
	Microsoft::VisualStudio::CppUnitTestFramework::Assert::IsTrue(SUCCEEDED(hr));
	Microsoft::VisualStudio::CppUnitTestFramework::Assert::AreEqual<VARTYPE>(VT_VSITEMID, var.vt);
	return V_VSITEMID(&var);
}

inline com_ptr<IDispatch> GetProperty_Dispatch (IVsHierarchy* hier, VSITEMID itemid, VSHPROPID propid)
{
	wil::unique_variant var;
	auto hr = hier->GetProperty(itemid, propid, &var);
	Assert::IsTrue(SUCCEEDED(hr));
	Assert::AreEqual<VARTYPE>(VT_DISPATCH, var.vt);
	com_ptr<IDispatch> res;
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

