
#include "pch.h"
#include "TestsCommon.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wil::unique_process_heap_string CombinePath (const wchar_t* dir, const wchar_t* pathRelativeToDir)
{
	Assert::IsTrue(!PathIsRelative(dir));
	Assert::IsTrue(PathIsRelative(pathRelativeToDir));
	wil::unique_hlocal_string path;
	auto hr = PathAllocCombine (dir, pathRelativeToDir, PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS, path.addressof());
	Assert::IsTrue(SUCCEEDED(hr));
	for (wchar_t* p = path.get(); *p; ++p)
	{
		if (*p == L'/')
			*p = L'\\';
	}
	return wil::make_process_heap_string_failfast(path.get());
}

void RemoveDirectoryTree (const wchar_t* dir)
{
	auto buffer = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s%c", dir, L'\0');
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = buffer.get(), .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	int ires = SHFileOperation(&file_op);
	Microsoft::VisualStudio::CppUnitTestFramework::Assert::AreEqual(0, ires);
}
