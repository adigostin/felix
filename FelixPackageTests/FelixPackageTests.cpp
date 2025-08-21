
#include "pch.h"
#include "CppUnitTest.h"
#include "shared/com.h"
#include "../FelixPackage/FelixPackage.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

extern com_ptr<IServiceProvider> MakeMockServiceProvider();


static HMODULE dll;
using GCO = HRESULT (STDMETHODCALLTYPE*)(REFCLSID rclsid, REFIID riid, LPVOID* ppv);
static GCO getClassObject;
com_ptr<IVsPackage> package;
wchar_t tempPath[MAX_PATH + 1];
wil::unique_process_heap_string templateFullPath;
wil::unique_process_heap_string TemplatePath_EmptyProject;
wil::unique_process_heap_string TemplatePath_EmptyFile;

static const GUID CLSID_FelixPackage = { 0x768BC57B, 0x42A8, 0x42AB, { 0xB3, 0x89, 0x45, 0x79, 0x46, 0xC4, 0xFC, 0x6A } };

const GUID FelixProjectType = { 0xD438161C, 0xF032, 0x4014, { 0xBC, 0x5C, 0x20, 0xA8, 0x0E, 0xAF, 0xF5, 0x9B } };

static const char TemplateXML[] = ""
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">"
"  <Configurations>"
"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />"
"  </Configurations>"
"  <Items>"
"    <File Path=\"file.asm\" BuildTool=\"Assembler\" />"
"  </Items>"
"</Z80Project>";

static const char TemplateXML_EmptyProject[] = ""
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\" />";

namespace FelixTests
{
	TEST_MODULE_INITIALIZE(InitModule)
	{
		HRESULT hr;

		GetTempPathW (MAX_PATH + 1, tempPath);
		wchar_t* end;
		StringCchCatExW (tempPath, _countof(tempPath), L"FelixTest\\", &end, NULL, 0);
		Assert::IsTrue (end < tempPath + _countof(tempPath) - 1);
		if (PathFileExists(tempPath))
		{
			end[1] = 0;
			SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = tempPath, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
			int ires = SHFileOperation(&file_op);
			Assert::AreEqual(0, ires);
		}
		BOOL bres = CreateDirectory(tempPath, 0);
		Assert::IsTrue(bres);

		wil::str_concat_nothrow(templateFullPath, tempPath, L"TemplateOneConfigOneFile\\");
		CreateDirectory(templateFullPath.get(), nullptr);
		wil::str_concat_nothrow(templateFullPath, L"proj.flx");
		wil::unique_hfile th (CreateFile(templateFullPath.get(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
		Assert::IsTrue(th.is_valid());
		DWORD bytesWritten;
		bres = WriteFile(th.get(), TemplateXML, sizeof(TemplateXML) - 1, &bytesWritten, nullptr);
		Assert::IsTrue(bres);
		th.reset();

		wil::str_concat_nothrow(TemplatePath_EmptyProject, tempPath, L"TemplateEmpty\\");
		CreateDirectory(TemplatePath_EmptyProject.get(), nullptr);
		wil::str_concat_nothrow(TemplatePath_EmptyProject, L"proj.flx");
		th.reset(CreateFile(TemplatePath_EmptyProject.get(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
		Assert::IsTrue(th.is_valid());
		bres = WriteFile(th.get(), TemplateXML_EmptyProject, sizeof(TemplateXML_EmptyProject) - 1, &bytesWritten, nullptr);
		Assert::IsTrue(bres);
		th.reset();

		wil::str_concat_nothrow(TemplatePath_EmptyFile, tempPath, L"template.asm");
		WriteFileOnDisk(tempPath, L"template.asm");

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
