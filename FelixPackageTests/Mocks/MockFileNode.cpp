
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockFileNode : IFileNode, IFileNodeProperties
{
	ULONG _refCount = 0;
	com_ptr<IWeakRef> _parent;
	VSITEMID _itemId = VSITEMID_NIL;
	com_ptr<IChildNode> _next;
	BuildToolKind _buildTool;
	wil::unique_hlocal_string _pathRelativeToProjectDir;
	com_ptr<ICustomBuildToolProperties> _customBuildToolProps;

	HRESULT InitInstance (BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir)
	{
		HRESULT hr;
		_buildTool = buildTool;
		_pathRelativeToProjectDir = wil::make_hlocal_string_nothrow(pathRelativeToProjectDir); RETURN_IF_NULL_ALLOC_EXPECTED(_pathRelativeToProjectDir);
		hr = MakeCustomBuildToolProperties(&_customBuildToolProps); RETURN_IF_FAILED_EXPECTED(hr);
		return S_OK;
	}

	~MockFileNode()
	{
		// TODO: delete file from disk
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IChildNode*>(this), riid, ppvObject)
			|| TryQI<IDispatch>(this, riid, ppvObject)
			|| TryQI<IChildNode>(this, riid, ppvObject)
			|| TryQI<IFileNode>(this, riid, ppvObject)
			|| TryQI<IFileNodeProperties>(this, riid, ppvObject)
		)
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH_(__uuidof(IFileNodeProperties), L"FelixPackage.dll");

	#pragma region IChildNode
	virtual HRESULT GetParent (IParentNode** ppParent) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !_parent);
		return _parent->QueryInterface(IID_PPV_ARGS(ppParent));
	}

	virtual VSITEMID STDMETHODCALLTYPE GetItemId(void) override { return _itemId; }

	virtual HRESULT STDMETHODCALLTYPE SetItemId (IParentNode* parent, VSITEMID id) override
	{
		RETURN_HR_IF(E_INVALIDARG, id == VSITEMID_NIL);
		RETURN_HR_IF(E_INVALIDARG, !parent);
		RETURN_HR_IF(E_UNEXPECTED, _itemId != VSITEMID_NIL);
		RETURN_HR_IF(E_UNEXPECTED, _parent);
		auto hr = parent->QueryInterface(IID_PPV_ARGS(_parent.addressof())); RETURN_IF_FAILED(hr);
		_itemId = id;
		return S_OK;
	}

	virtual HRESULT ClearItemId() override
	{
		RETURN_HR_IF(E_UNEXPECTED, _itemId == VSITEMID_NIL);
		RETURN_HR_IF(E_UNEXPECTED, !_parent);
		_itemId = VSITEMID_NIL;
		_parent = nullptr;
		return S_OK;
	}

	IChildNode* STDMETHODCALLTYPE Next(void) override
	{
		return _next;
	}

	void STDMETHODCALLTYPE SetNext(IChildNode* next) override
	{
		_next = next;
	}

	HRESULT STDMETHODCALLTYPE GetProperty(VSHPROPID propid, VARIANT* pvar) override
	{
		switch (propid)
		{
			case VSHPROPID_SaveName: // -2002
			case VSHPROPID_Caption: // -2003
			case VSHPROPID_Name: // -2012
			case VSHPROPID_EditLabel: // -2026
				if (!_pathRelativeToProjectDir || !_pathRelativeToProjectDir.get()[0])
					return E_NOT_SET;
				return InitVariantFromString (PathFindFileName(_pathRelativeToProjectDir.get()), pvar);

			case VSHPROPID_BrowseObject:
				return InitVariantFromDispatch(this, pvar);

			case VSHPROPID_NextSibling:
				return InitVariantFromVSITEMID(_next ? _next->GetItemId() : VSITEMID_NIL, pvar);

			case VSHPROPID_Parent:
			{
				com_ptr<INode> parent;
				auto hr = _parent->QueryInterface(IID_PPV_ARGS(parent.addressof())); RETURN_IF_FAILED(hr);
				return InitVariantFromInt32 (parent->GetItemId(), pvar);
			}
		}

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

	#pragma region IFileNode
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

com_ptr<IFileNode> MakeMockFileNode (BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir)
{
	auto p = com_ptr (new (std::nothrow) MockFileNode());
	Assert::IsNotNull(p.get());
	auto hr = p->InitInstance (buildTool, pathRelativeToProjectDir);
	Assert::IsTrue(SUCCEEDED(hr));
	return p;
}

void WriteFileOnDisk (const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir, const char* fileContent)
{
	wchar_t mkDoc[MAX_PATH];
	auto res = PathCombineW (mkDoc, projectDir, pathRelativeToProjectDir);
	Assert::IsNotNull(res);

	wil::unique_hfile handle (CreateFile(mkDoc, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
	Assert::IsTrue(handle.is_valid());
	if (fileContent)
	{
		BOOL bres = WriteFile(handle.get(), fileContent, (DWORD)strlen(fileContent), NULL, NULL);
		Assert::IsTrue(bres);
	}
}

void DeleteFileOnDisk (const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir)
{
	wchar_t mkDoc[MAX_PATH];
	auto res = PathCombineW (mkDoc, projectDir, pathRelativeToProjectDir);
	Assert::IsNotNull(res);

	// If the caller specified no file content, we must make sure there's no leftover file.
	if (PathFileExists(mkDoc))
	{
		BOOL bres = DeleteFile(mkDoc);
		Assert::IsTrue(bres);
	}
}
