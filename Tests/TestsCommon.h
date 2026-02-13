
#pragma once

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
};

wil::com_ptr_failfast<IMockHierarchyEventSink> MakeMockHierarchyEventSink();
