
#pragma once
#include "../FelixPackage/FelixPackage_h.h"
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
	extern wchar_t tempPath[MAX_PATH + 1];
	extern wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
	extern wil::unique_process_heap_string TemplatePath_OneConfigOneCustomBuildTool;
	extern wil::unique_process_heap_string TemplatePath_EmptyProject;
	extern wil::unique_process_heap_string TemplatePath_EmptyFile;
	extern LPCOLESTR TemplateEmptyFile[1];

	wil::com_ptr_failfast<VxDTE::DTE2> GetDefaultVSInstance();
	wil::com_ptr_failfast<VxDTE::DTE2> LaunchVS (const wchar_t* envVar = nullptr);
	void CloseVS (VxDTE::DTE2* dte, bool hard = false);

	inline VSITEMID GetProperty_VSITEMID (IVsHierarchy* hier, VSITEMID itemid, VSHPROPID propid)
	{
		wil::unique_variant var;
		auto hr = hier->GetProperty(itemid, propid, &var);
		Assert::IsTrue(SUCCEEDED(hr));
		Assert::AreEqual<VARTYPE>(VT_VSITEMID, var.vt);
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

	template <typename... arguments>
	wil::unique_process_heap_string str_concat (arguments&&... args)
	{
		wil::unique_process_heap_string result{};
		FAIL_FAST_IF_FAILED(wil::str_concat_nothrow(result, wistd::forward<arguments>(args)...));
		return result;
	}

	// Returns true if the condition was met before the timeout expired.
	template<typename condition_t> requires std::is_invocable_r_v<bool, condition_t>
	bool WaitWithMessageLoop (const condition_t& condition, DWORD timeoutMilliseconds)
	{
		DWORD tickStart = GetTickCount();
		while (timeoutMilliseconds == INFINITE || GetTickCount() - tickStart < timeoutMilliseconds)
		{
			if (condition())
				return true;

			MSG msg;
			while(PeekMessage(&msg,0,0,0,PM_NOREMOVE))
			{
				if (::GetMessage(&msg, NULL, 0, 0) > 0)
					::DispatchMessage(&msg);
			}

			Sleep(20);
		}

		return false;
	}

	void DisableGeneratedFiles (VxDTE::Project* proj);

	struct DECLSPEC_NOVTABLE DECLSPEC_UUID("EBD70B25-9BE7-4C0E-B562-2FE4CDE6F14A") ITestHierarchyEventSink : IVsHierarchyEvents
	{
		virtual bool PropertyChanged (VSITEMID itemid, VSHPROPID propid) const = 0;
		virtual bool ItemAdded (VSITEMID itemidParent, VSITEMID itemidAdded) const = 0;
		virtual bool ItemRemoved (VSITEMID itemid) const = 0;
		virtual bool ChildItemsInvalidated(VSITEMID itemidParent) const = 0;
	};
	wil::com_ptr_failfast<ITestHierarchyEventSink> MakeTestHierarchyEventSink();

	struct DECLSPEC_NOVTABLE DECLSPEC_UUID("B358F183-4FE2-43CB-A5C4-FC0DFC37FB84") ITestPropertyChangeSink : IPropertyChangeSink
	{
		virtual bool Called (IDispatch* obj, std::initializer_list<DISPID> dispIDs) const = 0;
	};
	wil::com_ptr_failfast<ITestPropertyChangeSink> MakeTestPropertyChangeSink();

	struct DECLSPEC_NOVTABLE DECLSPEC_UUID("5F94F823-F39F-4311-8067-4F851D1DDAAC") ITestPropertyNotifySink : IPropertyNotifySink
	{
		virtual bool Called (std::initializer_list<DISPID> dispIDs) const = 0;
	};
	wil::com_ptr_failfast<ITestPropertyNotifySink> MakeTestPropertyNotifySink();

	template<typename string_type>
	void WriteFileCreateDirs (const string_type& path_str, const char* fileContent)
	{
		const wchar_t* path = wil::str_raw_ptr(path_str);
		Assert::IsFalse(PathIsRelativeW(path));
		auto fn = PathFindFileName(path);
		Assert::IsTrue(fn > path);
		fn[-1] = L'\0'; // temporarily terminate to get the directory path
		auto hr = wil::CreateDirectoryDeepNoThrow(path);
		Assert::IsTrue(SUCCEEDED(hr));

		fn[-1] = L'\\'; // restore full path
		wil::unique_hfile handle (CreateFile(path, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
		Assert::IsTrue(handle.is_valid());
		if (fileContent)
		{
			BOOL bres = WriteFile(handle.get(), fileContent, (DWORD)strlen(fileContent), NULL, NULL);
			Assert::IsTrue(bres);
		}
	}

	template<typename string_type>
	void RemoveDirectoryTree (const string_type& dir_str)
	{
		const wchar_t* dir = wil::str_raw_ptr(dir_str);
		auto buffer = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s%c", dir, L'\0');
		SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = buffer.get(), .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
		int ires = SHFileOperation(&file_op);
		Microsoft::VisualStudio::CppUnitTestFramework::Assert::AreEqual(0, ires);
	}
}


