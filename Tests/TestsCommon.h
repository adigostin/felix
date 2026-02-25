
#pragma once
#include "../FelixPackage/FelixPackage_h.h"
#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>

extern wchar_t tempPath[MAX_PATH + 1];
extern wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
extern wil::unique_process_heap_string TemplatePath_EmptyProject;
extern wil::unique_process_heap_string TemplatePath_EmptyFile;

void MakeTemplates (const wchar_t* tempPath);
wil::unique_process_heap_string CombinePath (const wchar_t* dir, const wchar_t* pathRelativeToDir);
void WriteFileOnDisk (wchar_t* path, const char* fileContent);
void RemoveDirectoryTree (const wchar_t* dir);

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("EBD70B25-9BE7-4C0E-B562-2FE4CDE6F14A") ITestHierarchyEventSink : IVsHierarchyEvents
{
	virtual bool PropertyChanged (VSITEMID itemid, VSHPROPID propid) const = 0;
	virtual bool ItemAdded (VSITEMID itemidParent, VSITEMID itemidAdded) const = 0;
	virtual bool ItemRemoved (VSITEMID itemid) const = 0;
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
