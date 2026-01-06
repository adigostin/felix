
#pragma once

extern wchar_t tempPath[MAX_PATH + 1];
extern wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
extern wil::unique_process_heap_string TemplatePath_EmptyProject;
extern wil::unique_process_heap_string TemplatePath_EmptyFile;

void MakeTemplates (const wchar_t* tempPath);
wil::unique_hlocal_string CombinePath (const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir);
void WriteFileOnDisk (wchar_t* path, const char* fileContent);
