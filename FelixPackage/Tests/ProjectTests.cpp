
#include "pch.h"
#include "CppUnitTest.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wchar_t tempPath[MAX_PATH + 1];

#pragma warning (push)
#pragma warning (disable: 4995)

TEST_MODULE_INITIALIZE(ModuleInit)
{
	serviceProvider = MakeMockServiceProvider();

	GetTempPathW (MAX_PATH + 1, tempPath);
	wchar_t* end;
	StringCchCatExW (tempPath, _countof(tempPath), L"FelixTest", &end, NULL, 0);
	Assert::IsTrue (end < tempPath + _countof(tempPath) - 1);
	end[1] = 0;
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = tempPath, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	SHFileOperation(&file_op);

	BOOL bres = CreateDirectory(tempPath, 0);
	Assert::IsTrue(bres);
}

TEST_MODULE_CLEANUP(ModuleCleanup)
{
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = tempPath, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	SHFileOperation(&file_op);

	serviceProvider = nullptr;
}


#pragma warning (pop)
