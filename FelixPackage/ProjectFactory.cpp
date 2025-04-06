
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"

class ProjectFactory : public IVsProjectFactory
{
	ULONG _refCount = 0;

	wil::com_ptr_nothrow<IServiceProvider> _sp;

public:
	static HRESULT CreateInstance (IServiceProvider* sp, IVsProjectFactory** to)
	{
		wil::com_ptr_nothrow<ProjectFactory> p = new (std::nothrow) ProjectFactory(); RETURN_IF_NULL_ALLOC(p);
		p->_sp = sp;
		*to = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IVsProjectFactory>(this, riid, ppvObject))
			return S_OK;

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IVsProjectFactory
	virtual HRESULT __stdcall CanCreateProject(LPCOLESTR pszFilename, VSCREATEPROJFLAGS grfCreateFlags, BOOL * pfCanCreate) override
	{
		*pfCanCreate = TRUE;
		return S_OK;
	}

	virtual HRESULT __stdcall CreateProject(LPCOLESTR pszFilename, LPCOLESTR pszLocation, LPCOLESTR pszName, VSCREATEPROJFLAGS grfCreateFlags, REFIID iidProject, void ** ppvProject, BOOL * pfCanceled) override
	{
		// Better doc here: https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.ivssolution.createproject?view=visualstudiosdk-2022

		RETURN_HR_IF(E_POINTER, !ppvProject || !pfCanceled);
		*ppvProject = nullptr;
		*pfCanceled = TRUE;

		auto hr = MakeProjectNode (pszFilename, pszLocation, pszName, grfCreateFlags, iidProject, ppvProject); RETURN_IF_FAILED_EXPECTED(hr);
	
		*pfCanceled = FALSE;
		return S_OK;
	}

	virtual HRESULT __stdcall SetSite(IServiceProvider* pSP) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall Close() override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeProjectFactory (IServiceProvider* sp, IVsProjectFactory** to)
{
	return ProjectFactory::CreateInstance(sp, to);
}

