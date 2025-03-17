
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockSourceFile : IProjectFile
{
	ULONG _refCount = 0;
	VSITEMID _itemId;
	BuildToolKind _buildTool;
	wil::unique_hlocal_string _pathRelativeToProjectDir;
	com_ptr<ICustomBuildToolProperties> _customBuildToolProps;

	MockSourceFile()
	{
	}

	HRESULT InitInstance (VSITEMID itemId, BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir)
	{
		_itemId = itemId;
		_buildTool = buildTool;
		_pathRelativeToProjectDir = wil::make_hlocal_string_nothrow(pathRelativeToProjectDir); RETURN_IF_NULL_ALLOC_EXPECTED(_pathRelativeToProjectDir);
		auto hr = MakeCustomBuildToolProperties(&_customBuildToolProps); RETURN_IF_FAILED_EXPECTED(hr);
		return S_OK;
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
		return E_NOTIMPL;
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
	HRESULT STDMETHODCALLTYPE Close(void) override
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

com_ptr<IProjectFile> MakeMockSourceFile (VSITEMID itemId, BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir)
{
	auto p = com_ptr (new (std::nothrow) MockSourceFile());
	Assert::IsNotNull(p.get());
	auto hr = p->InitInstance(itemId, buildTool, pathRelativeToProjectDir);
	Assert::IsTrue(SUCCEEDED(hr));
	return p;
}
