
#include "pch.h"
#include "Mocks.h"

com_ptr<IFileNode> MakeFileNode (const wchar_t* pathRelativeToProjectDir)
{
	com_ptr<IFileNode> f;
	auto hr = MakeFileNode(&f);
	Assert::IsTrue(SUCCEEDED(hr));
	com_ptr<IFileNodeProperties> props;
	hr = f->QueryInterface(IID_PPV_ARGS(props.addressof()));
	Assert::IsTrue(SUCCEEDED(hr));
	auto p = wil::make_bstr_nothrow(pathRelativeToProjectDir);
	Assert::IsNotNull(p.get());
	hr = props->put_Path(p.get());
	Assert::IsTrue(SUCCEEDED(hr));
	return f;
}

com_ptr<IProjectConfig> AddDebugProjectConfig (IVsHierarchy* hier)
{
	HRESULT hr;

	com_ptr<IVsGetCfgProvider> getCfgProvider;
	hr = hier->QueryInterface(IID_PPV_ARGS(&getCfgProvider));
	Assert::IsTrue(SUCCEEDED(hr));

	com_ptr<IVsCfgProvider> cfgProvider;
	hr = getCfgProvider->GetCfgProvider(&cfgProvider);
	Assert::IsTrue(SUCCEEDED(hr));

	com_ptr<IVsCfgProvider2> cfgProvider2;
	hr = cfgProvider->QueryInterface(&cfgProvider2);
	Assert::IsTrue(SUCCEEDED(hr));

	hr = cfgProvider2->AddCfgsOfCfgName(L"Debug", nullptr, FALSE);
	Assert::IsTrue(SUCCEEDED(hr));

	com_ptr<IVsCfg> cfg;
	ULONG actual;
	hr = cfgProvider2->GetCfgs (1, cfg.addressof(), &actual, NULL);
	Assert::IsTrue(SUCCEEDED(hr));
	Assert::AreEqual<ULONG>(1, actual);

	com_ptr<IProjectConfig> config;
	hr = cfg->QueryInterface(&config);
	Assert::IsTrue(SUCCEEDED(hr));

	return config;
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
