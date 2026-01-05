
#include "pch.h"
#include "CppUnitTest.h"
#include "shared/com.h"
#include "FelixPackageTests.h"
#include "TestsCommon.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

extern com_ptr<IServiceProvider> MakeMockServiceProvider();


static HMODULE dll;
using GCO = HRESULT (STDMETHODCALLTYPE*)(REFCLSID rclsid, REFIID riid, LPVOID* ppv);
static GCO getClassObject;
com_ptr<IVsPackage> package;

const GUID CLSID_FelixPackage = { 0x768BC57B, 0x42A8, 0x42AB, { 0xB3, 0x89, 0x45, 0x79, 0x46, 0xC4, 0xFC, 0x6A } };

const GUID FelixProjectType = { 0xD438161C, 0xF032, 0x4014, { 0xBC, 0x5C, 0x20, 0xA8, 0x0E, 0xAF, 0xF5, 0x9B } };

namespace FelixTests
{
	TEST_MODULE_INITIALIZE(InitModule)
	{
		HRESULT hr;

		MakeTemplates (L"FelixTest");

		dll = LoadLibrary(L"FelixPackage.dll");
		Assert::IsNotNull(dll);
		getClassObject = (GCO)GetProcAddress(dll, "DllGetClassObject");
		Assert::IsNotNull((void*)getClassObject);
	
		com_ptr<IClassFactory> packageFactory;
		hr = getClassObject(CLSID_FelixPackage, IID_PPV_ARGS(&packageFactory));
		Assert::IsTrue(SUCCEEDED(hr));

		hr = packageFactory->CreateInstance(nullptr, IID_PPV_ARGS(&package));
		Assert::IsTrue(SUCCEEDED(hr));

		auto sp = MakeMockServiceProvider();
		hr = package->SetSite(sp);
		Assert::IsTrue(SUCCEEDED(hr));
	}

	TEST_MODULE_CLEANUP(CleanupModule)
	{
		if (package)
		{
			package->Close();
			ULONG refCount = package.detach()->Release();
			Assert::AreEqual((ULONG)0, refCount);
		}

		if (serviceProvider)
		{
			ULONG refCount = serviceProvider.detach()->Release();
			Assert::AreEqual((ULONG)0, refCount);
		}
	}

	TEST_CLASS(PackageTests)
	{
	public:
		
		TEST_METHOD(CloneProject)
		{
			HRESULT hr;

			com_ptr<IVsSolution> sol;
			hr = serviceProvider->QueryService(SID_SVsSolution, IID_PPV_ARGS(&sol));
			Assert::IsTrue(SUCCEEDED(hr));

			static const wchar_t ProjFileName[] = L"TestProject.flx";
			com_ptr<IVsHierarchy> hier;
			hr = sol->CreateProject(FelixProjectType, templateFullPath.get(), tempPath, ProjFileName, CPF_CLONEFILE, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_hlocal_string path;
			hr = wil::str_concat_nothrow(path, tempPath, L"\\", ProjFileName);
			Assert::IsTrue(SUCCEEDED(hr));
			DeleteFile(path.get());
		}
	};
}
