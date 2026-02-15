
#pragma once
#include "../FelixPackage/FelixPackage_h.h"
#include "shared/inplace_function.h"

extern wchar_t tempPath[MAX_PATH + 1];
extern wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
extern wil::unique_process_heap_string TemplatePath_EmptyProject;
extern wil::unique_process_heap_string TemplatePath_EmptyFile;

void MakeTemplates (const wchar_t* tempPath);
wil::unique_hlocal_string CombinePath (const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir);
void WriteFileOnDisk (wchar_t* path, const char* fileContent);
void RemoveDirectoryTree (const wchar_t* dir);

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("EBD70B25-9BE7-4C0E-B562-2FE4CDE6F14A") IMockHierarchyEventSink : IVsHierarchyEvents
{
	virtual bool PropertyChanged (VSITEMID itemid, VSHPROPID propid) const = 0;
	virtual bool ItemAdded (VSITEMID itemidParent, VSITEMID itemidAdded) const = 0;
	virtual bool ItemRemoved (VSITEMID itemid) const = 0;
};
wil::com_ptr_failfast<IMockHierarchyEventSink> MakeMockHierarchyEventSink();

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

bool WaitWithMessageLoop (const stdext::inplace_function<bool()>& condition, DWORD timeoutMilliseconds);
