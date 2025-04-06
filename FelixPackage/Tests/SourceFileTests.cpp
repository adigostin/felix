
#include "pch.h"
#include "CppUnitTest.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(SourceFileTests)
	{
		TEST_METHOD(SourceFileUnnamedTest)
		{
			auto hier = MakeMockVsHierarchy(tempPath);
			auto config = MakeMockProjectConfig(hier);
			auto pane = MakeMockOutputWindowPane(nullptr);

			com_ptr<IProjectFile> file;
			auto hr = MakeProjectFile (1000, hier, VSITEMID_ROOT, &file);
			Assert::IsTrue(SUCCEEDED(hr));
			hier.try_query<IProjectItemParent>()->SetFirstChild(file);

			wil::unique_variant value;
			hr = hier->GetProperty(1000, VSHPROPID_SaveName, &value);
			Assert::IsTrue(FAILED(hr));
			hr = hier->GetProperty(1000, VSHPROPID_Caption, &value);
			Assert::IsTrue(FAILED(hr));
		}
	};
}
