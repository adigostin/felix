
#pragma once
#include "../FelixPackage/FelixPackage_h.h"
#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>

extern wchar_t tempPath[MAX_PATH + 1];
extern wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
extern wil::unique_process_heap_string TemplatePath_OneConfigOneCustomBuildTool;
extern wil::unique_process_heap_string TemplatePath_EmptyProject;
extern wil::unique_process_heap_string TemplatePath_EmptyFile;
extern LPCOLESTR TemplateEmptyFile[1];

void MakeTemplates (const wchar_t* tempPath);
wil::unique_process_heap_string CombinePath (const wchar_t* dir, const wchar_t* pathRelativeToDir);
void WriteFileOnDisk (wchar_t* path, const char* fileContent);
void RemoveDirectoryTree (const wchar_t* dir);
