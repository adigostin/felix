
#include "pch.h"
#include "CppUnitTest.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wchar_t tempPath[MAX_PATH + 1];

#pragma warning (push)
#pragma warning (disable: 4995)

TEST_MODULE_INITIALIZE(ModuleInit)
{
	serviceProvider = MakeMockServiceProvider();

	GetTempPathW (MAX_PATH + 1, tempPath);
	wchar_t* end;
	StringCchCatExW (tempPath, _countof(tempPath), L"FelixTest", &end, NULL, 0);
	Assert::IsTrue (end < tempPath + _countof(tempPath) - 1);
	end[1] = 0;
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = tempPath, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	SHFileOperation(&file_op);

	BOOL bres = CreateDirectory(tempPath, 0);
	Assert::IsTrue(bres);
}

TEST_MODULE_CLEANUP(ModuleCleanup)
{
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = tempPath, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	SHFileOperation(&file_op);

	serviceProvider = nullptr;
}

namespace FelixTests
{
	TEST_CLASS(ProjectTests)
	{
		TEST_METHOD(PutItemsOneFile)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeFelixProject (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file = MakeMockSourceFile(hier, 1000, VSITEMID_ROOT, BuildToolKind::Assembler, L"test.asm");
			//Assert::AreEqual<VSITEMID>(VSITEMID_NIL, file->GetItemId());

			{
				SAFEARRAYBOUND bound = { .cElements = 1, .lLbound = 0 };
				auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
				LONG i = 0;
				SafeArrayPutElement (sa.get(), &i, file.try_query<IDispatch>());
				hr = hier.try_query<IZ80ProjectProperties>()->put_Items(sa.get());
				Assert::IsTrue(SUCCEEDED(hr));

				Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file->GetItemId());

				auto pip = hier.try_query<IProjectItemParent>();
				IUnknown* firstChild = wil::try_com_query_nothrow<IUnknown>(pip->FirstChild());
				Assert::AreEqual<void*>(file.try_query<IUnknown>().get(), firstChild);

				wil::unique_variant parentID;
				hr = file->GetProperty(VSHPROPID_Parent, &parentID);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, V_VSITEMID(&parentID));
			}

			ULONG refCount = hier.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);

			refCount = file.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);
		}
	};
}

#pragma warning (pop)
