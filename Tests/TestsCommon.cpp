
#include "pch.h"
#include "TestsCommon.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wchar_t tempPath[MAX_PATH + 1];
wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
wil::unique_process_heap_string TemplatePath_OneConfigOneCustomBuildTool;
wil::unique_process_heap_string TemplatePath_EmptyProject;
wil::unique_process_heap_string TemplatePath_EmptyFile;
LPCOLESTR TemplateEmptyFile[1];

static const char TemplateTwoConfigsOneFileXML[] = ""
	"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
	"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
	"  <Configurations>\r\n"
	"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />\r\n"
	"    <Configuration ConfigName=\"Release\" PlatformName=\"ZX Spectrum 48K\" />\r\n"
	"  </Configurations>\r\n"
	"  <Items>\r\n"
	"    <File Path=\"file.asm\" BuildTool=\"Assembler\" />\r\n"
	"  </Items>\r\n"
	"</Z80Project>\r\n";

static const char TemplateXML_OneConfigOneCustomBuildTool[] = ""
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
"  <Configurations>\r\n"
"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />\r\n"
"  </Configurations>\r\n"
"  <Items>\r\n"
"    <File Path=\"file.asm\" BuildTool=\"CustomBuildTool\" >\r\n"
"      <CustomBuildToolProperties />\r\n"
"    </File>\r\n"
"  </Items>\r\n"
"</Z80Project>\r\n";

static const char TemplateXML_EmptyProject[] = ""
	"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
	"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
	"  <Configurations>\r\n"
	"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\">\r\n"
	"      <AssemblerProperties GeneratePrePostIncludeFiles=\"False\" />\r\n"
	"    </Configuration>\r\n"
	"  </Configurations>\r\n"
	"</Z80Project>\r\n";

void MakeTemplates (const wchar_t* tempDirName)
{
	GetTempPathW (MAX_PATH + 1, tempPath);
	swprintf_s (tempPath, L"%s%s\\", tempPath, tempDirName);
	if (PathFileExists(tempPath))
	{
		auto buffer = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s*%c", tempPath, L'\0');
		SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = buffer.get(), .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
		int ires = SHFileOperation(&file_op);
		Assert::AreEqual(0, ires);
	}
	else
		Assert::IsTrue(CreateDirectory(tempPath, 0));

	auto templateDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"TemplateTwoConfigsOneFile\\");
	wil::str_concat_nothrow(TemplatePath_TwoConfigsOneFile, templateDir, L"proj.flx");
	WriteFileOnDisk(TemplatePath_TwoConfigsOneFile.get(), TemplateTwoConfigsOneFileXML);
	auto file = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"file.asm");
	WriteFileOnDisk(file.get(), "start:\tret");

	templateDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"TemplateOneConfigOneCustomBuildTool\\");
	wil::str_concat_nothrow(TemplatePath_OneConfigOneCustomBuildTool, templateDir, L"proj.flx");
	WriteFileOnDisk(TemplatePath_OneConfigOneCustomBuildTool.get(), TemplateXML_OneConfigOneCustomBuildTool);
	file = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"file.asm");
	WriteFileOnDisk(file.get(), "start:\tret");

	wil::str_concat_nothrow(TemplatePath_EmptyProject, tempPath, L"TemplateEmpty\\proj.flx");
	WriteFileOnDisk(TemplatePath_EmptyProject.get(), TemplateXML_EmptyProject);

	wil::str_concat_nothrow(TemplatePath_EmptyFile, tempPath, L"template.asm");
	wil::unique_hfile (CreateFile(TemplatePath_EmptyFile.get(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
	TemplateEmptyFile[0] = TemplatePath_EmptyFile.get();
}

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
