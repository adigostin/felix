
#include "pch.h"
#include "PackageTests.h"

#pragma comment (lib, "synchronization.lib")

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct TestBuildCallback : IProjectConfigBuilderCallback
{
	ULONG _refCount = 0;
	bool _complete = false;
	bool _success = false;
	stdext::inplace_function<void(bool)> _buildComplete;

	TestBuildCallback (stdext::inplace_function<void(bool)> buildComplete = nullptr)
		: _buildComplete(std::move(buildComplete))
	{ }

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { return E_NOTIMPL; }
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	virtual HRESULT STDMETHODCALLTYPE OnBuildComplete (bool success) override
	{
		Assert::IsFalse(_complete);
		_complete = TRUE;
		_success = success;
		if (_buildComplete)
			_buildComplete(success);
		return S_OK;
	}
};

namespace FelixTests
{
	TEST_CLASS(BuilderTests)
	{
		struct TD
		{
			wil::unique_process_heap_string testDir;

			TD()
			{
				testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"FolderTest\\");
				Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));
			}

			~TD()
			{
				RemoveDirectoryTree(testDir.get());
			}
		};

		static std::pair<wil::com_ptr_failfast<IProjectNode>, wil::com_ptr_failfast<IProjectConfigBuilder>> MakeSjasmProjectBuilder (const wchar_t* testDir, const char* asmFileContent)
		{
			HRESULT hr;
			com_ptr<IProjectNode> project;
			hr = MakeProjectNode (nullptr, testDir, nullptr, 0, IID_PPV_ARGS(&project));
			Assert::IsTrue(SUCCEEDED(hr));
			auto config = AddDebugProjectConfig(project->AsHierarchy());
			config->AsmProps()->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);

			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"test.asm");
			if (asmFileContent)
			{
				wil::unique_hfile h (CreateFile(filePath.get(), GENERIC_WRITE, 0, 0, CREATE_NEW, 0, 0));
				Assert::IsTrue(h.is_valid());
				BOOL bres = WriteFile(h.get(), asmFileContent, strlen(asmFileContent), NULL, NULL);
				Assert::IsTrue(bres);
			}
			else
			{
				BOOL bres = DeleteFile(filePath.get());
				Assert::IsTrue(bres);
			}
			LPCOLESTR filesToOpen[] = { filePath.get() };
			hr = project->AsVsProject()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, filesToOpen, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			auto pane = MakeMockOutputWindowPane(nullptr);
			com_ptr<IProjectConfigBuilder> builder;
			hr = MakeProjectConfigBuilder (project, config, pane, &builder);
			Assert::IsTrue(SUCCEEDED(hr));
			return { std::move(project), std::move(builder) };
		}

		// sourceFileContent - empty string view to skip creating the file on disk
		static std::pair<wil::com_ptr_failfast<IProjectNode>, wil::com_ptr_failfast<IProjectConfigBuilder>> MakeProjectWithCustomBuildTool (
			const wchar_t* testDir,
			const wchar_t* sourceFileName, const char* sourceFileContent,
			const wchar_t* cbtDescription,
			const wchar_t* cbtCmdLine, IStream* outputStreamUTF16)
		{
			HRESULT hr;

			wil::com_ptr_failfast<IProjectNode> project;
			hr = MakeProjectNode (nullptr, testDir, nullptr, 0, IID_PPV_ARGS(&project));
			Assert::IsTrue(SUCCEEDED(hr));
			auto config = AddDebugProjectConfig(project->AsHierarchy());
			config->AsmProps()->put_GeneratePrePostIncludeFiles(VARIANT_FALSE);

			if (sourceFileContent)
				WriteFileOnDisk(CombinePath(testDir, sourceFileName).get(), sourceFileContent);
			else
				DeleteFileOnDisk(testDir, sourceFileName);
			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, sourceFileName);
			hr = project->AsVsProject()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)filePath.addressof(), nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			auto sourceFile = wil::com_query_failfast<IFileNodeProperties>(project->FirstChild());

			hr = sourceFile->put_BuildTool(BuildToolKind::CustomBuildTool);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<ICustomBuildToolProperties> cbtProps;
			hr = sourceFile->get_CustomBuildToolProperties(&cbtProps);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = cbtProps->put_Description(wil::make_bstr_nothrow(cbtDescription).get());
			Assert::IsTrue(SUCCEEDED(hr));
			hr = cbtProps->put_CommandLine(wil::make_bstr_nothrow(cbtCmdLine).get());
			Assert::IsTrue(SUCCEEDED(hr));

			auto pane = MakeMockOutputWindowPane(outputStreamUTF16);

			wil::com_ptr_failfast<IProjectConfigBuilder> builder;
			hr = MakeProjectConfigBuilder (project, config, pane, &builder);
			Assert::IsTrue(SUCCEEDED(hr));
			return { std::move(project), std::move(builder) };
		}

		TEST_METHOD(TestBuilderDestroyedWhenReleasedWithPendingBuild)
		{
			TD td;
			auto [proj, builder] = MakeProjectWithCustomBuildTool(td.testDir.get(), L"test.xxx", "content", nullptr, L"cmd /c pause", nullptr);
			auto close = wil::scope_exit([&proj] { proj->AsHierarchy()->Close(); });
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			WaitWithMessageLoop([callback] { return callback->_complete; }, 1000);
			Assert::IsFalse(callback->_complete);
			ULONG remainingRefCount = builder.detach()->Release();
			Assert::AreEqual((ULONG)0, remainingRefCount);
		}
	};
}
