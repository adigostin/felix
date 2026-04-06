
#include "pch.h"
#include "PackageTests.h"
#include "../FelixPackage/Z80Xml.h"
#include "../FelixPackageUi/resource.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(ProjectTests)
	{


		TEST_METHOD(ProjectDirPropertyEndsWithBackslash)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant projDir;
			hr = hier->GetProperty (VSITEMID_ROOT, VSHPROPID_ProjectDir, &projDir);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_BSTR, projDir.vt);
			Assert::AreEqual(L'\\', projDir.bstrVal[SysStringLen(projDir.bstrVal) - 1]);
		}
	};
}
