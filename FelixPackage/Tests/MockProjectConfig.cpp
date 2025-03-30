
#include "pch.h"
#include "CppUnitTest.h"
#include "FelixPackage.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct MockProjectConfig : IProjectConfig
{
	ULONG _refCount = 0;
	com_ptr<IVsHierarchy> _hier;
	com_ptr<IProjectConfigPrePostBuildProperties> _preBuildProps;
	com_ptr<IProjectConfigPrePostBuildProperties> _postBuildProps;
	com_ptr<IProjectConfigAssemblerProperties> _assemblerProps;


	HRESULT InitInstance (IVsHierarchy* hier)
	{
		HRESULT hr;
		_hier = hier;
		hr = PrePostBuildPageProperties_CreateInstance(false, &_preBuildProps); RETURN_IF_FAILED_EXPECTED(hr);
		hr = PrePostBuildPageProperties_CreateInstance(true, &_postBuildProps); RETURN_IF_FAILED_EXPECTED(hr);
		hr = AssemblerPageProperties_CreateInstance(&_assemblerProps); RETURN_IF_FAILED_EXPECTED(hr);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { RETURN_HR(E_NOTIMPL); }

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	IMPLEMENT_IDISPATCH(IID_IProjectConfig);

	#pragma region IProjectConfig
	virtual HRESULT STDMETHODCALLTYPE get___id(BSTR* value) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_AssemblerProperties(IProjectConfigAssemblerProperties** ppProps) override
	{
		return wil::com_copy_to_nothrow(_assemblerProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE put_GeneralProperties(IProjectConfigAssemblerProperties* pDispatch) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_DebuggingProperties(IProjectConfigDebugProperties** ppProps) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_PreBuildProperties(IProjectConfigPrePostBuildProperties** ppProps) override
	{
		return wil::com_copy_to_nothrow(_preBuildProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE get_PostBuildProperties(IProjectConfigPrePostBuildProperties** ppProps) override
	{
		return wil::com_copy_to_nothrow(_postBuildProps, ppProps);
	}

	virtual HRESULT STDMETHODCALLTYPE get_ConfigName(BSTR* pbstr) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE put_ConfigName(BSTR pbstr) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE get_PlatformName(BSTR* pbstr) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE put_PlatformName(BSTR pbstr) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetOutputDirectory(BSTR* pbstr) override
	{
		wil::unique_variant project_dir;
		auto hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &project_dir); RETURN_IF_FAILED(hr);
		if (project_dir.vt != VT_BSTR)
			return E_FAIL;
		*pbstr = project_dir.release().bstrVal;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetOutputFileName(BSTR* pbstr) override
	{
		*pbstr = SysAllocString(L"output.bin"); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSldFileName(BSTR* pbstr) override
	{
		*pbstr = SysAllocString(L"output.sld"); RETURN_IF_NULL_ALLOC(*pbstr);
		return S_OK;
	}
	#pragma endregion
};

com_ptr<IProjectConfig> MakeMockProjectConfig (IVsHierarchy* hier)
{
	auto config = com_ptr(new (std::nothrow) MockProjectConfig());
	auto hr = config->InitInstance(hier);
	Assert::IsTrue(SUCCEEDED(hr));
	return config;
}
