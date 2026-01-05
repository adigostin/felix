
#include "pch.h"
#include "TestsCommon.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wchar_t tempPath[MAX_PATH + 1];
wil::unique_process_heap_string templateFullPath;
wil::unique_process_heap_string TemplatePath_EmptyProject;
wil::unique_process_heap_string TemplatePath_EmptyFile;

static const char TemplateOneConfigOneFileXML[] = ""
	"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
	"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">"
	"  <Configurations>"
	"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />"
	"  </Configurations>"
	"  <Items>"
	"    <File Path=\"file.asm\" BuildTool=\"Assembler\" />"
	"  </Items>"
	"</Z80Project>";

static const char TemplateXML_EmptyProject[] = ""
	"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
	"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\" />";

void MakeTemplates (const wchar_t* tempDirName)
{
	GetTempPathW (MAX_PATH + 1, tempPath);
	swprintf_s (tempPath, L"%s%s\\%c", tempPath, tempDirName, '\0');
	if (PathFileExists(tempPath))
		std::filesystem::remove_all(tempPath);
	Assert::IsTrue(CreateDirectory(tempPath, 0));

	auto templateDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"TemplateOneConfigOneFile\\");
	//CreateDirectory(templateDir.get(), nullptr);
	wil::str_concat_nothrow(templateFullPath, templateDir, L"proj.flx");
	//wil::unique_hfile th (CreateFile(templateFullPath.get(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
	//Assert::IsTrue(th.is_valid());
	//DWORD bytesWritten;
	//bres = WriteFile(th.get(), TemplateOneConfigOneFileXML, sizeof(TemplateOneConfigOneFileXML) - 1, &bytesWritten, nullptr);
	//Assert::IsTrue(bres);
	WriteFileOnDisk(templateFullPath.get(), TemplateOneConfigOneFileXML);

	auto file = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"file.asm");
	//th.reset(CreateFile(file.get(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
	//Assert::IsTrue(th.is_valid());
	//bres = WriteFile(th.get(), "\tend", 4, &bytesWritten, nullptr);
	//Assert::IsTrue(bres);
	//th.reset();
	WriteFileOnDisk(file.get(), "\tend");

	wil::str_concat_nothrow(TemplatePath_EmptyProject, tempPath, L"TemplateEmpty\\proj.flx");
	//CreateDirectory(TemplatePath_EmptyProject.get(), nullptr);
	//wil::str_concat_nothrow(TemplatePath_EmptyProject, L"proj.flx");
	//th.reset(CreateFile(TemplatePath_EmptyProject.get(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
	//Assert::IsTrue(th.is_valid());
	//bres = WriteFile(th.get(), TemplateXML_EmptyProject, sizeof(TemplateXML_EmptyProject) - 1, &bytesWritten, nullptr);
	//Assert::IsTrue(bres);
	//th.reset();
	WriteFileOnDisk(TemplatePath_EmptyProject.get(), TemplateXML_EmptyProject);

	wil::str_concat_nothrow(TemplatePath_EmptyFile, tempPath, L"template.asm");
	wil::unique_hfile (CreateFile(TemplatePath_EmptyFile.get(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
}

wil::unique_hlocal_string CombinePath (const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir)
{
	Assert::IsTrue(!PathIsRelative(projectDir));
	Assert::IsTrue(PathIsRelative(pathRelativeToProjectDir));
	wil::unique_hlocal_string path;
	auto hr = PathAllocCombine (projectDir, pathRelativeToProjectDir, PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS, path.addressof());
	Assert::IsTrue(SUCCEEDED(hr));
	for (wchar_t* p = path.get(); *p; ++p)
	{
		if (*p == L'/')
			*p = L'\\';
	}
	return path;
}

void WriteFileOnDisk (wchar_t* path, const char* fileContent)
{
	wchar_t* lastsep = wcsrchr(path, L'\\');
	Assert::IsNotNull(lastsep);
	*lastsep = L'\0'; // temporarily terminate to get the directory path
	auto hr = wil::CreateDirectoryDeepNoThrow(path);
	Assert::IsTrue(SUCCEEDED(hr));

	*lastsep = L'\\'; // restore full path
	wil::unique_hfile handle (CreateFile(path, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
	Assert::IsTrue(handle.is_valid());
	if (fileContent)
	{
		BOOL bres = WriteFile(handle.get(), fileContent, (DWORD)strlen(fileContent), NULL, NULL);
		Assert::IsTrue(bres);
	}
}
