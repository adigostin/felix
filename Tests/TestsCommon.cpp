
#include "pch.h"
#include "TestsCommon.h"
#include "shared/com.h"

void RemoveDirectoryTree (const wchar_t* dir)
{
	auto buffer = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s%c", dir, L'\0');
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = buffer.get(), .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	int ires = SHFileOperation(&file_op);
	Microsoft::VisualStudio::CppUnitTestFramework::Assert::AreEqual(0, ires);
}
