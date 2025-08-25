
#pragma once
#include "shared/com.h"
#include "shared/inplace_function.h"
#include "FelixPackage.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

extern wil::unique_process_heap_string templateFullPath;
extern wil::unique_process_heap_string TemplatePath_EmptyProject;
extern wil::unique_process_heap_string TemplatePath_EmptyFile;
extern wchar_t tempPath[MAX_PATH + 1];
extern const GUID FelixProjectType;

namespace FelixTests
{
	struct DECLSPEC_NOVTABLE DECLSPEC_UUID("B82AFB90-FF61-4B05-AD2F-760CD09443E2") IMockPropertyNotifySink : IUnknown
	{
		virtual bool IsChanged (DISPID dispid) const = 0;
	};

	com_ptr<IFileNode> MakeFileNode (const wchar_t* pathRelativeToProjectDir);
	VSITEMID AddFolderNode (IVsUIHierarchy* hier, VSITEMID addTo, const wchar_t* name);
	void WriteFileOnDisk(const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir, const char* fileContent = nullptr);
	void DeleteFileOnDisk(const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir);

	com_ptr<IProjectConfig> AddDebugProjectConfig (IVsHierarchy* hier);
	com_ptr<IVsOutputWindowPane2> MakeMockOutputWindowPane (IStream* outputStreamUTF16);
	com_ptr<IServiceProvider> MakeMockServiceProvider();
	com_ptr<IMockPropertyNotifySink> MakeMockPropertyNotifySink();
	com_ptr<IVsDebugger> MakeMockDebugger();
}

using PropChangedCallback = stdext::inplace_function<void(VSITEMID itemid, VSHPROPID propid, DWORD flags)>;
com_ptr<IVsHierarchyEvents> MakeMockHierarchyEventSink (PropChangedCallback propChanged);

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

