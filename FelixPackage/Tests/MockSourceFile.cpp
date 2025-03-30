
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockSourceFile : IProjectFile
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _hier;
	VSITEMID _itemId;
	BuildToolKind _buildTool;
	wil::unique_hlocal_string _pathRelativeToProjectDir;
	com_ptr<ICustomBuildToolProperties> _customBuildToolProps;

	HRESULT InitInstance (IVsHierarchy* hier, VSITEMID itemId,
		BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir, std::string_view fileContent)
	{
		auto hr = hier->QueryInterface(IID_PPV_ARGS(_hier.addressof())); RETURN_IF_FAILED(hr);
		_itemId = itemId;
		_buildTool = buildTool;
		_pathRelativeToProjectDir = wil::make_hlocal_string_nothrow(pathRelativeToProjectDir); RETURN_IF_NULL_ALLOC_EXPECTED(_pathRelativeToProjectDir);
		hr = MakeCustomBuildToolProperties(&_customBuildToolProps); RETURN_IF_FAILED_EXPECTED(hr);

		wil::unique_bstr mkDoc;
		hr = GetMkDocument(&mkDoc);
		Assert::IsTrue(SUCCEEDED(hr));
		if (!fileContent.empty())
		{
			wil::unique_hfile handle (CreateFile(mkDoc.get(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			Assert::IsTrue(handle.is_valid());
			BOOL bres = WriteFile(handle.get(), fileContent.data(), (DWORD)fileContent.size(), NULL, NULL);
			Assert::IsTrue(bres);
		}
		else
		{
			if (PathFileExists(mkDoc.get()))
			{
				BOOL bres = DeleteFile(mkDoc.get());
				Assert::IsTrue(bres);
			}
		}

		return S_OK;
	}

	~MockSourceFile()
	{
		// TODO: delete file from disk
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IProjectItem*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IProjectItem>(this, riid, ppvObject)
			|| TryQI<IProjectFile>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IProjectItem);

	#pragma region IProjectItem
	virtual VSITEMID STDMETHODCALLTYPE GetItemId(void) override { return _itemId; }

	HRESULT STDMETHODCALLTYPE GetMkDocument(BSTR* pbstrMkDocument) override
	{
		com_ptr<IVsHierarchy> hier;
		auto hr = _hier->QueryInterface(&hier);
		Assert::IsTrue(SUCCEEDED(hr));
		Assert::IsNotNull(hier.get());
		wil::unique_variant projectDir;
		hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir);
		Assert::IsTrue(SUCCEEDED(hr));
		Assert::IsTrue(projectDir.vt == VT_BSTR);
		wil::unique_hlocal_string filePath;
		hr = PathAllocCombine(projectDir.bstrVal, _pathRelativeToProjectDir.get(), PathFlags, &filePath);
		Assert::IsTrue(SUCCEEDED(hr));
		*pbstrMkDocument = SysAllocString(filePath.get()); RETURN_IF_NULL_ALLOC(*pbstrMkDocument);
		return S_OK;
	}

	IProjectItem* STDMETHODCALLTYPE Next(void) override
	{
		return nullptr;
	}
	void STDMETHODCALLTYPE SetNext(IProjectItem* next) override
	{
	}
	HRESULT STDMETHODCALLTYPE GetProperty(VSHPROPID propid, VARIANT* pvar) override
	{
		if (propid == VSHPROPID_BrowseObject)
			return InitVariantFromDispatch(this, pvar);

		if (propid == VSHPROPID_NextSibling)
			return InitVariantFromVSITEMID(VSITEMID_NIL, pvar);

		return E_NOTIMPL;
	}

	HRESULT STDMETHODCALLTYPE SetProperty(VSHPROPID propid, REFVARIANT var) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetGuidProperty(VSHPROPID propid, GUID* pguid) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE SetGuidProperty(VSHPROPID propid, REFGUID rguid) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetCanonicalName(BSTR* pbstrName) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE IsItemDirty(IUnknown* punkDocData, BOOL* pfDirty) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE QueryStatus(const GUID* pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT* pCmdText) override
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE Exec(const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn, VARIANT* pvaOut) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IProjectFile
	virtual HRESULT STDMETHODCALLTYPE get___id (BSTR *value) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_Path (BSTR *value) override
	{
		*value = SysAllocString(_pathRelativeToProjectDir.get());
		return *value ? S_OK : E_OUTOFMEMORY;
	}

	virtual HRESULT STDMETHODCALLTYPE put_Path (BSTR value) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_BuildTool (enum BuildToolKind *value) override
	{
		*value = _buildTool;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE put_BuildTool (enum BuildToolKind value) override
	{
		_buildTool = value;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE get_CustomBuildToolProperties (ICustomBuildToolProperties **ppProps) override
	{
		return wil::com_query_to_nothrow(_customBuildToolProps, ppProps);
	}
	#pragma endregion
};

com_ptr<IProjectFile> MakeMockSourceFile (IVsHierarchy* hier, VSITEMID itemId,
	BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir, std::string_view fileContent)
{
	auto p = com_ptr (new (std::nothrow) MockSourceFile());
	Assert::IsNotNull(p.get());
	auto hr = p->InitInstance (hier, itemId, buildTool, pathRelativeToProjectDir, fileContent);
	Assert::IsTrue(SUCCEEDED(hr));
	return p;
}
