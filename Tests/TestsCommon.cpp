
#include "pch.h"
#include "TestsCommon.h"
#include "shared/com.h"
#include <unordered_map>
#include <set>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wchar_t tempPath[MAX_PATH + 1];
wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
wil::unique_process_heap_string TemplatePath_EmptyProject;
wil::unique_process_heap_string TemplatePath_EmptyFile;

static const char TemplateTwoConfigsOneFileXML[] = ""
	"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
	"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">"
	"  <Configurations>"
	"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />"
	"    <Configuration ConfigName=\"Release\" PlatformName=\"ZX Spectrum 48K\" />"
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
		RemoveDirectoryTree(tempPath);
	Assert::IsTrue(CreateDirectory(tempPath, 0));

	auto templateDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"TemplateTwoConfigsOneFile\\");
	wil::str_concat_nothrow(TemplatePath_TwoConfigsOneFile, templateDir, L"proj.flx");
	WriteFileOnDisk(TemplatePath_TwoConfigsOneFile.get(), TemplateTwoConfigsOneFileXML);

	auto file = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"file.asm");
	WriteFileOnDisk(file.get(), "start:\tret");

	wil::str_concat_nothrow(TemplatePath_EmptyProject, tempPath, L"TemplateEmpty\\proj.flx");
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

struct MockHierarchyEventSink : IMockHierarchyEventSink
{
	ULONG _refCount = 0;
	std::unordered_map<VSITEMID, std::set<VSHPROPID>> _changedProps;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IVsHierarchyEvents>(this, riid, ppvObject)
			|| TryQI<IMockHierarchyEventSink>(this, riid, ppvObject)
		)
			return S_OK;

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsHierarchyEvents
	virtual HRESULT STDMETHODCALLTYPE OnItemAdded (VSITEMID itemidParent, VSITEMID itemidSiblingPrev, VSITEMID itemidAdded) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemsAppended (VSITEMID itemidParent) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemDeleted (VSITEMID itemid) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged (VSITEMID itemid, VSHPROPID propid, DWORD flags) override
	{
		_changedProps[itemid].insert(propid);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnInvalidateItems (VSITEMID itemidParent) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnInvalidateIcon (HICON hicon) override
	{
		return S_OK;
	}
	#pragma endregion

	#pragma region IMockHierarchyEventSink
	virtual bool PropertyChanged (VSITEMID itemid, VSHPROPID propid) const override
	{
		auto it = _changedProps.find(itemid);
		if (it == _changedProps.end())
			return false;
		return it->second.contains(propid);
	}
	#pragma endregion
};

wil::com_ptr_failfast<IMockHierarchyEventSink> MakeMockHierarchyEventSink()
{
	return wil::com_ptr_failfast(new (std::nothrow) MockHierarchyEventSink());
}
