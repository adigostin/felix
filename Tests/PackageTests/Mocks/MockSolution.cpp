
#include "pch.h"
#include "CppUnitTest.h"
#include "shared/com.h"
#include "shared/unordered_map_nothrow.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockSolution : IVsSolution, IVsRegisterProjectTypes
{
	ULONG _refCount = 0;
	VSCOOKIE _nextProjectTypeCookie = 1;
	unordered_map_nothrow<VSCOOKIE, std::pair<GUID, com_ptr<IVsProjectFactory>>> _projectFactories;

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IVsSolution*>(this), riid, ppvObject)
			|| TryQI<IVsSolution>(this, riid, ppvObject)
			|| TryQI<IVsRegisterProjectTypes>(this, riid, ppvObject)
		)
			return S_OK;

		Assert::Fail();
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsSolution
	virtual HRESULT STDMETHODCALLTYPE GetProjectEnum(
		/* [in] */ VSENUMPROJFLAGS grfEnumFlags,
		/* [in] */ __RPC__in REFGUID rguidEnumOnlyThisType,
		/* [out] */ __RPC__deref_out_opt IEnumHierarchies** ppEnum) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE CreateProject(
		/* [in] */ __RPC__in REFGUID rguidProjectType,
		/* [in] */ __RPC__in LPCOLESTR lpszMoniker,
		/* [in] */ __RPC__in LPCOLESTR lpszLocation,
		/* [in] */ __RPC__in LPCOLESTR lpszName,
		/* [in] */ VSCREATEPROJFLAGS grfCreateFlags,
		/* [in] */ __RPC__in REFIID iidProject,
		/* [iid_is][out] */ __RPC__deref_out_opt void** ppProject) override
	{
		auto it = _projectFactories.find_if([&rguidProjectType](auto& p) { return p.second.first == rguidProjectType; });
		Assert::IsTrue(it != _projectFactories.end());
		BOOL canceled;
		auto hr = it->second.second->CreateProject(lpszMoniker, lpszLocation, lpszName, grfCreateFlags, iidProject, ppProject, &canceled);
		Assert::IsTrue(SUCCEEDED(hr));
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GenerateUniqueProjectName(
		/* [in] */ __RPC__in LPCOLESTR lpszRoot,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrProjectName) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjectOfGuid(
		/* [in] */ __RPC__in REFGUID rguidProjectID,
		/* [out] */ __RPC__deref_out_opt IVsHierarchy** ppHierarchy) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetGuidOfProject(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [out] */ __RPC__out GUID* pguidProjectID) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetSolutionInfo(
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrSolutionDirectory,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrSolutionFile,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrUserOptsFile) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseSolutionEvents(
		/* [in] */ __RPC__in_opt IVsSolutionEvents* pSink,
		/* [out] */ __RPC__out VSCOOKIE* pdwCookie) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseSolutionEvents(
		/* [in] */ VSCOOKIE dwCookie) override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SaveSolutionElement(
		/* [in] */ VSSLNSAVEOPTIONS grfSaveOpts,
		/* [in] */ __RPC__in_opt IVsHierarchy* pHier,
		/* [in] */ VSCOOKIE docCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE CloseSolutionElement(
		/* [in] */ VSSLNCLOSEOPTIONS grfCloseOpts,
		/* [in] */ __RPC__in_opt IVsHierarchy* pHier,
		/* [in] */ VSCOOKIE docCookie) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjectOfProjref(
		/* [in] */ __RPC__in LPCOLESTR pszProjref,
		/* [out] */ __RPC__deref_out_opt IVsHierarchy** ppHierarchy,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrUpdatedProjref,
		/* [out] */ __RPC__out VSUPDATEPROJREFREASON* puprUpdateReason) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjrefOfProject(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrProjref) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjectInfoOfProjref(
		/* [in] */ __RPC__in LPCOLESTR pszProjref,
		/* [in] */ VSHPROPID propid,
		/* [out] */ __RPC__out VARIANT* pvar) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AddVirtualProject(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [in] */ VSADDVPFLAGS grfAddVPFlags) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetItemOfProjref(
		/* [in] */ __RPC__in LPCOLESTR pszProjref,
		/* [out] */ __RPC__deref_out_opt IVsHierarchy** ppHierarchy,
		/* [out] */ __RPC__out VSITEMID* pitemid,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrUpdatedProjref,
		/* [out] */ __RPC__out VSUPDATEPROJREFREASON* puprUpdateReason) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjrefOfItem(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [in] */ VSITEMID itemid,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrProjref) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetItemInfoOfProjref(
		/* [in] */ __RPC__in LPCOLESTR pszProjref,
		/* [in] */ VSHPROPID propid,
		/* [out] */ __RPC__out VARIANT* pvar) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjectOfUniqueName(
		/* [in] */ __RPC__in LPCOLESTR pszUniqueName,
		/* [out] */ __RPC__deref_out_opt IVsHierarchy** ppHierarchy) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetUniqueNameOfProject(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrUniqueName) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProperty(
		/* [in] */ VSPROPID propid,
		/* [out] */ __RPC__out VARIANT* pvar) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE SetProperty(
		/* [in] */ VSPROPID propid,
		/* [in] */ VARIANT var) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE OpenSolutionFile(
		/* [in] */ VSSLNOPENOPTIONS grfOpenOpts,
		/* [in] */ __RPC__in LPCOLESTR pszFilename) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE QueryEditSolutionFile(
		/* [out] */ __RPC__out DWORD* pdwEditResult) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE CreateSolution(
		/* [unique][in] */ __RPC__in_opt LPCOLESTR lpszLocation,
		/* [unique][in] */ __RPC__in_opt LPCOLESTR lpszName,
		/* [in] */ VSCREATESOLUTIONFLAGS grfCreateFlags) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjectFactory(
		/* [in] */ DWORD dwReserved,
		/* [out][in] */ __RPC__inout GUID* pguidProjectType,
		/* [in] */ __RPC__in LPCOLESTR pszMkProject,
		/* [retval][out] */ __RPC__deref_out_opt IVsProjectFactory** ppProjectFactory) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjectTypeGuid(
		/* [in] */ DWORD dwReserved,
		/* [in] */ __RPC__in LPCOLESTR pszMkProject,
		/* [retval][out] */ __RPC__out GUID* pguidProjectType) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE OpenSolutionViaDlg(
		__RPC__in LPCOLESTR pszStartDirectory,
		BOOL fDefaultToAllProjectsFilter) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE AddVirtualProjectEx(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [in] */ VSADDVPFLAGS grfAddVPFlags,
		/* [in] */ __RPC__in REFGUID rguidProjectID) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE QueryRenameProject(
		/* [in] */ __RPC__in_opt IVsProject* pProject,
		/* [in] */ __RPC__in LPCOLESTR pszMkOldName,
		/* [in] */ __RPC__in LPCOLESTR pszMkNewName,
		/* [in] */ DWORD dwReserved,
		/* [out] */ __RPC__out BOOL* pfRenameCanContinue) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE OnAfterRenameProject(
		/* [in] */ __RPC__in_opt IVsProject* pProject,
		/* [in] */ __RPC__in LPCOLESTR pszMkOldName,
		/* [in] */ __RPC__in LPCOLESTR pszMkNewName,
		/* [in] */ DWORD dwReserved) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveVirtualProject(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [in] */ VSREMOVEVPFLAGS grfRemoveVPFlags) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE CreateNewProjectViaDlg(
		/* [in] */ __RPC__in LPCOLESTR pszExpand,
		/* [in] */ __RPC__in LPCOLESTR pszSelect,
		/* [in] */ DWORD dwReserved) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetVirtualProjectFlags(
		/* [in] */ __RPC__in_opt IVsHierarchy* pHierarchy,
		/* [out] */ __RPC__out VSADDVPFLAGS* pgrfAddVPFlags) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GenerateNextDefaultProjectName(
		/* [in] */ __RPC__in LPCOLESTR pszBaseName,
		/* [in] */ __RPC__in LPCOLESTR pszLocation,
		/* [out] */ __RPC__deref_out_opt BSTR* pbstrProjectName) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE GetProjectFilesInSolution(
		/* [in] */ VSGETPROJFILESFLAGS grfGetOpts,
		/* [in] */ ULONG cProjects,
		/* [length_is][size_is][out] */ __RPC__out_ecount_part(cProjects, *pcProjectsFetched) BSTR* rgbstrProjectNames,
		/* [out] */ __RPC__out ULONG* pcProjectsFetched) override
	{
		Assert::Fail(L"Not Implemented");
	}

	virtual HRESULT STDMETHODCALLTYPE CanCreateNewProjectAtLocation(
		/* [in] */ BOOL fCreateNewSolution,
		/* [in] */ __RPC__in LPCOLESTR pszFullProjectFilePath,
		/* [out] */ __RPC__out BOOL* pfCanCreate) override
	{
		Assert::Fail(L"Not Implemented");
	}
	#pragma endregion

	#pragma region IVsRegisterProjectTypes
	virtual HRESULT STDMETHODCALLTYPE RegisterProjectType( 
		REFGUID rguidProjType,
		IVsProjectFactory *pVsPF,
		VSCOOKIE *pdwCookie) override
	{
		VSCOOKIE cookie = _nextProjectTypeCookie++;
		(void)_projectFactories.try_insert({ cookie, { rguidProjType, pVsPF } });
		*pdwCookie = cookie;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnregisterProjectType (VSCOOKIE dwCookie) override
	{
		auto it = _projectFactories.find(dwCookie);
		Assert::IsTrue (it != _projectFactories.end());
		_projectFactories.erase(it);
		return S_OK;
	}
	#pragma endregion
};

com_ptr<IVsSolution> MakeMockSolution()
{
	auto p = com_ptr(new (std::nothrow) MockSolution());
	Assert::IsNotNull(p.get());
	return p;
}
