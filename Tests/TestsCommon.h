
#pragma once
#include "../FelixPackage/FelixPackage_h.h"
#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>


wil::unique_process_heap_string CombinePath (const wchar_t* dir, const wchar_t* pathRelativeToDir);
void RemoveDirectoryTree (const wchar_t* dir);
