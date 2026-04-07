
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

void WriteFileOnDisk (wchar_t* path, const char* fileContent)
{
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

void RemoveDirectoryTree (const wchar_t* dir)
{
	auto buffer = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s%c", dir, L'\0');
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = buffer.get(), .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	int ires = SHFileOperation(&file_op);
	Microsoft::VisualStudio::CppUnitTestFramework::Assert::AreEqual(0, ires);
}
