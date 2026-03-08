
#include "pch.h"
#include "TestsCommon.h"
#include "shared/com.h"
#include <unordered_map>
#include <set>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wchar_t tempPath[MAX_PATH + 1];
wil::unique_process_heap_string TemplatePath_TwoConfigsOneFile;
wil::unique_process_heap_string TemplatePath_OneConfigOneCustomBuildTool;
wil::unique_process_heap_string TemplatePath_EmptyProject;
wil::unique_process_heap_string TemplatePath_EmptyFile;

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
	swprintf_s (tempPath, L"%s%s\\%c", tempPath, tempDirName, '\0');
	if (PathFileExists(tempPath))
		RemoveDirectoryTree(tempPath);
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

struct TestHierarchyEventSink : ITestHierarchyEventSink
{
	ULONG _refCount = 0;
	std::unordered_map<VSITEMID, std::set<VSHPROPID>> _changedProps;

	struct Added { VSITEMID itemidParent; VSITEMID itemidSiblingPrev; VSITEMID itemidAdded; };
	std::vector<Added> _added;

	std::set<VSITEMID> _removed;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IVsHierarchyEvents>(this, riid, ppvObject)
			|| TryQI<ITestHierarchyEventSink>(this, riid, ppvObject)
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
		_added.push_back({ itemidParent, itemidSiblingPrev, itemidAdded });
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemsAppended (VSITEMID itemidParent) override
	{
		Assert::Fail();
	}

	virtual HRESULT STDMETHODCALLTYPE OnItemDeleted (VSITEMID itemid) override
	{
		_removed.insert(itemid);
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

	#pragma region ITestHierarchyEventSink
	virtual bool PropertyChanged (VSITEMID itemid, VSHPROPID propid) const override
	{
		auto it = _changedProps.find(itemid);
		if (it == _changedProps.end())
			return false;
		return it->second.contains(propid);
	}

	virtual bool ItemAdded (VSITEMID itemidParent, VSITEMID itemidAdded) const override
	{
		for (auto& a : _added)
		{
			if (a.itemidParent == itemidParent && a.itemidAdded == itemidAdded)
				return true;
		}

		return false;
	}

	virtual bool ItemRemoved (VSITEMID itemid) const override
	{
		return _removed.contains(itemid);
	}
	#pragma endregion
};

wil::com_ptr_failfast<ITestHierarchyEventSink> MakeTestHierarchyEventSink()
{
	return wil::com_ptr_failfast(new (std::nothrow) TestHierarchyEventSink());
}

struct TestPropertyChangeSink : ITestPropertyChangeSink
{
	ULONG _refCount = 0;
	std::unordered_map<wil::com_ptr_failfast<IDispatch>, std::set<DISPID>, std::hash<IDispatch*>> _changing;
	std::unordered_map<wil::com_ptr_failfast<IDispatch>, std::set<DISPID>, std::hash<IDispatch*>> _changed;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IPropertyChangeSink>(this, riid, ppvObject)
			|| TryQI<ITestPropertyChangeSink>(this, riid, ppvObject)
		)
			return S_OK;

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IPropertyChangeSink
	virtual HRESULT STDMETHODCALLTYPE OnPropertyChanging (IDispatch* pObject, DISPID dispID, PropertyChangeArgs args) override
	{
		_changing[pObject].insert(dispID);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged (IDispatch* pObject, DISPID dispID, PropertyChangeArgs args) override
	{
		_changed[pObject].insert(dispID);
		return S_OK;
	}
	#pragma endregion

	#pragma region ITestPropertyChangeSink
	virtual bool Called (IDispatch* obj, std::initializer_list<DISPID> dispIDs) const override
	{
		for (DISPID dispID : dispIDs)
		{
			auto it = _changing.find(obj);
			if (it == _changing.end() || !it->second.contains(dispID))
				return false;
			it = _changed.find(obj);
			if (it == _changed.end() || !it->second.contains(dispID))
				return false;
		}

		return true;
	}

	#pragma endregion
};

wil::com_ptr_failfast<ITestPropertyChangeSink> MakeTestPropertyChangeSink()
{
	return new TestPropertyChangeSink();
}

struct TestPropertyNotifySink : ITestPropertyNotifySink
{
	ULONG _refCount = 0;
	vector_nothrow<DISPID> _changed;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IPropertyNotifySink*>(this), riid, ppvObject)
			|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
			|| TryQI<ITestPropertyNotifySink>(this, riid, ppvObject)
			)
			return S_OK;

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IPropertyNotifySink
	virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
	{
		auto it = _changed.find(dispID);
		if (it == _changed.end())
			_changed.try_push_back(dispID);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region ITestPropertyNotifySink
	virtual bool Called (std::initializer_list<DISPID> dispIDs) const override
	{
		for (auto dispID : dispIDs)
		{
			if (_changed.find(dispID) == _changed.end())
				return false;
		}

		return true;
	}
	#pragma endregion
};

wil::com_ptr_failfast<ITestPropertyNotifySink> MakeTestPropertyNotifySink()
{
	return new (std::nothrow) TestPropertyNotifySink();
}

void DisableGeneratedFiles (VxDTE::Project* proj)
{
	wil::com_ptr_failfast<IVsCfg> cfg;
	ULONG actual;
	VSCFGFLAGS flags;
	wil::com_query_failfast<IVsCfgProvider>(proj)->GetCfgs(1, &cfg, &actual, &flags);
	com_ptr<IProjectConfigAssemblerProperties> asmProps;
	cfg.query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
	auto hr = asmProps->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);
	Assert::AreEqual(S_OK, hr);
}
