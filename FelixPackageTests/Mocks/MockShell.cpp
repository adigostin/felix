
#include "pch.h"
#include "FelixPackageTests.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockShell : IVsShell, IVsUIShell
{
	ULONG _refCount = 0;
	wil::unique_hmodule _ui;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (TryQI<IUnknown>(static_cast<IVsShell*>(this), riid, ppvObject)
			|| TryQI<IVsShell>(this, riid, ppvObject)
			|| TryQI<IVsUIShell>(this, riid, ppvObject)
		)
			return S_OK;

		Assert::Fail();
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsShell
	virtual HRESULT STDMETHODCALLTYPE GetPackageEnum( 
		/* [out] */ __RPC__deref_out_opt IEnumPackages **ppEnum) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty( 
		/* [in] */ VSSPROPID propid,
		/* [out] */ __RPC__out VARIANT *pvar) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty( 
		/* [in] */ VSSPROPID propid,
		/* [in] */ VARIANT var) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseBroadcastMessages( 
		/* [in] */ __RPC__in_opt IVsBroadcastMessageEvents *pSink,
		/* [out] */ __RPC__out VSCOOKIE *pdwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseBroadcastMessages( 
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseShellPropertyChanges( 
		/* [in] */ __RPC__in_opt IVsShellPropertyEvents *pSink,
		/* [out] */ __RPC__out VSCOOKIE *pdwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseShellPropertyChanges( 
		/* [in] */ VSCOOKIE dwCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE LoadPackage( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [retval][out] */ __RPC__deref_out_opt IVsPackage **ppPackage) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE LoadPackageString( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [in] */ ULONG resid,
		/* [retval][out] */ __RPC__deref_out_opt BSTR *pbstrOut) override
	{
		Assert::IsTrue(guidPackage == CLSID_FelixPackage);
		if (!_ui)
		{
			_ui.reset(LoadLibrary(L"FelixPackageUi.dll"));
			Assert::IsNotNull(_ui.get());
		}

		const wchar_t* str;
		int ires = LoadString(_ui.get(), resid, (LPWSTR)&str, 0);
		if (ires == 0)
		{
			DWORD gle = GetLastError();
			Assert::Fail();
		}

		BSTR s = SysAllocStringLen(str, (UINT)ires);
		Assert::IsNotNull(s);

		*pbstrOut = s;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE LoadUILibrary( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [in] */ DWORD dwExFlags,
		/* [retval][out] */ __RPC__out DWORD_PTR *phinstOut) override
	{
		Assert::IsTrue(guidPackage == GUID{ 0x768BC57B, 0x42A8, 0x42AB, { 0xB3, 0x89, 0x45, 0x79, 0x46, 0xC4, 0xFC, 0x6A } });
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE IsPackageInstalled( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [retval][out] */ __RPC__out BOOL *pfInstalled) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE IsPackageLoaded( 
		/* [in] */ __RPC__in REFGUID guidPackage,
		/* [retval][out] */ __RPC__deref_out_opt IVsPackage **ppPackage) override
	{
		Assert::Fail(L"Not Implemented");
	}
	#pragma endregion

	#pragma region IVsUIShell
	virtual HRESULT STDMETHODCALLTYPE GetToolWindowEnum(IEnumWindowFrames** ppEnum) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetDocumentWindowEnum(IEnumWindowFrames** ppEnum) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE FindToolWindow(VSFINDTOOLWIN grfFTW, REFGUID rguidPersistenceSlot, IVsWindowFrame** ppWindowFrame) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE CreateToolWindow(VSCREATETOOLWIN grfCTW, DWORD dwToolWindowId, IUnknown* punkTool, REFCLSID rclsidTool, REFGUID rguidPersistenceSlot, REFGUID rguidAutoActivate, IServiceProvider* pSP, LPCOLESTR pszCaption, BOOL* pfDefaultPosition, IVsWindowFrame** ppWindowFrame) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE CreateDocumentWindow(VSCREATEDOCWIN grfCDW, LPCOLESTR pszMkDocument, IVsUIHierarchy* pUIH, VSITEMID itemid, IUnknown* punkDocView, IUnknown* punkDocData, REFGUID rguidEditorType, LPCOLESTR pszPhysicalView, REFGUID rguidCmdUI, IServiceProvider* pSP, LPCOLESTR pszOwnerCaption, LPCOLESTR pszEditorCaption, BOOL* pfDefaultPosition, IVsWindowFrame** ppWindowFrame) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetErrorInfo(HRESULT hr, LPCOLESTR pszDescription, DWORD dwReserved, LPCOLESTR pszHelpKeyword, LPCOLESTR pszSource) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE ReportErrorInfo(HRESULT hr) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetDialogOwnerHwnd(HWND* phwnd) override
	{
		*phwnd = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE EnableModeless(BOOL fEnable) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SaveDocDataToFile(VSSAVEFLAGS grfSave, IUnknown* pPersistFile, LPCOLESTR pszUntitledPath, BSTR* pbstrDocumentNew, BOOL* pfCanceled) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetupToolbar(HWND hwnd, IVsToolWindowToolbar* ptwt, IVsToolWindowToolbarHost **pptwth) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetForegroundWindow() override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE TranslateAcceleratorAsACmd(MSG* pMsg) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UpdateCommandUI(BOOL fImmediateUpdate) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE UpdateDocDataIsDirtyFeedback(VSCOOKIE docCookie, BOOL fDirty) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE RefreshPropertyBrowser(DISPID dispid) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetWaitCursor() override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE PostExecCommand(const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE ShowContextMenu(DWORD dwCompRole, REFCLSID rclsidActive, LONG nMenuId, REFPOINTS pos, IOleCommandTarget* pCmdTrgtActive) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE ShowMessageBox(DWORD dwCompRole, REFCLSID rclsidComp, LPOLESTR pszTitle, LPOLESTR pszText, LPOLESTR pszHelpFile, DWORD dwHelpContextID, OLEMSGBUTTON msgbtn, OLEMSGDEFBUTTON msgdefbtn, OLEMSGICON msgicon, BOOL fSysAlert, LONG* pnResult) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetMRUComboText(const GUID* pguidCmdGroup, DWORD dwCmdID, LPSTR lpszText, BOOL fAddToList) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetToolbarVisibleInFullScreen(const GUID* pguidCmdGroup, DWORD dwToolbarId, BOOL fVisibleInFullScreen) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE FindToolWindowEx(VSFINDTOOLWIN grfFTW, REFGUID rguidPersistenceSlot, DWORD dwToolWinId, IVsWindowFrame** ppWindowFrame) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetAppName(BSTR* pbstrAppName) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetVSSysColor(VSSYSCOLOR dwSysColIndex, DWORD* pdwRGBval) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetMRUComboTextW(const GUID* pguidCmdGroup, DWORD dwCmdID, LPWSTR pwszText, BOOL fAddToList) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE PostSetFocusMenuCommand(const GUID* pguidCmdGroup, DWORD nCmdID) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetCurrentBFNavigationItem(IVsWindowFrame** ppWindowFrame, BSTR* pbstrData, IUnknown** ppunk) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AddNewBFNavigationItem(IVsWindowFrame* pWindowFrame, BSTR bstrData, IUnknown* punk, BOOL fReplaceCurrent) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE OnModeChange(DBGMODE dbgmodeNew) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetErrorInfo (BSTR *pbstrErrText) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetOpenFileNameViaDlg(VSOPENFILENAMEW* pOpenFileName) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetSaveFileNameViaDlg(VSSAVEFILENAMEW* pSaveFileName) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetDirectoryViaBrowseDlg(VSBROWSEINFOW* pBrowse) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE CenterDialogOnWindow(HWND hwndDialog, HWND hwndParent) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetPreviousBFNavigationItem(IVsWindowFrame** ppWindowFrame, BSTR* pbstrData, IUnknown** ppunk) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetNextBFNavigationItem(IVsWindowFrame** ppWindowFrame, BSTR* pbstrData, IUnknown** ppunk) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetURLViaDlg(LPCOLESTR pszDlgTitle, LPCOLESTR pszStaticLabel, LPCOLESTR pszHelpTopic, BSTR* pbstrURL) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveAdjacentBFNavigationItem(RemoveBFDirection rdDir) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveCurrentNavigationDupes(RemoveBFDirection rdDir) override
	{
		Assert::Fail(L"Not Implemented");
	}
	#pragma endregion
};

com_ptr<IVsShell> MakeMockShell()
{
	return com_ptr(new MockShell());
}
