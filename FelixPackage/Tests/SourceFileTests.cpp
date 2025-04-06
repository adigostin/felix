
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
			auto hr = MakeProjectFile (&file);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = AddFileToParent(file, hier.try_query<IProjectItemParent>());
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant value;
			hr = hier->GetProperty(1000, VSHPROPID_SaveName, &value);
			Assert::IsTrue(FAILED(hr));
			hr = hier->GetProperty(1000, VSHPROPID_Caption, &value);
			Assert::IsTrue(FAILED(hr));
		}

		TEST_METHOD(put_Items_EmptyProject)
		{
		}

		TEST_METHOD(put_Items_NonEmptyProject)
		{
		}
	};
}
