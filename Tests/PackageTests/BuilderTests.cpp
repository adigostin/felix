
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

		TEST_METHOD(Test_SjasmMissingExe)
		{
			TD td;
			wil::unique_process_heap_string dllDir;
			wil::GetModuleFileNameW((HMODULE)&__ImageBase, dllDir);
			*PathFindFileName(dllDir.get()) = 0;

			auto sjasmOrigPath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine (sjasmOrigPath.get(), dllDir.get(), L"sjasmplus.exe");
			auto sjasmTempPath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine(sjasmTempPath.get(), dllDir.get(), L"sjasmplus.tmp");
			BOOL bres = MoveFileExW (sjasmOrigPath.get(), sjasmTempPath.get(), MOVEFILE_REPLACE_EXISTING);
			Assert::IsTrue(bres);

			auto [proj, builder] = MakeSjasmProjectBuilder(td.testDir.get(), "");
			auto close = wil::scope_exit([&proj] { proj->AsHierarchy()->Close(); });
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild (callback);

			bres = MoveFileExW(sjasmTempPath.get(), sjasmOrigPath.get(), MOVEFILE_REPLACE_EXISTING);
			Assert::IsTrue(bres);

			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), hr);
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

		/*
		On second thought, this scenario is not legal COM. The application is supposed to
		hold a reference to "builder" until _after_ the call to CancelBuild returns.
		TEST_METHOD(TestCustomBuildToolWaitingUserInput_CallbackReleasesBuilder)
		{
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", { }, L"cmd /c pause", nullptr);
			auto callback = com_ptr(new TestBuildCallback([&builder](bool success) { builder.reset(); }));
			auto hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			WaitCallbackWithMessageLoop(1000, callback);
			Assert::IsFalse(callback->_complete);
			hr = builder->CancelBuild();
			Assert::IsTrue(callback->_complete);
			Assert::IsFalse(callback->_success);
			Assert::IsTrue(SUCCEEDED(hr));
		}
		*/
		static void CancelAfterAsyncBuildProcessExited (const wchar_t* testDir, const wchar_t* command, BOOL* complete, BOOL* success)
		{
			auto [proj, builder] = MakeProjectWithCustomBuildTool (testDir, L"test.xxx", "content", nullptr, command, nullptr);
			auto close = wil::scope_exit([&proj] { proj->AsHierarchy()->Close(); });
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsFalse(callback->_complete);
			Sleep(1000); // give the process time to run and exit, but don't pump the message loop
			Assert::IsFalse(callback->_complete);

			hr = builder->CancelBuild();
			Assert::IsTrue(SUCCEEDED(hr));
			*complete = callback->_complete;
			*success = callback->_success;
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode0)
		{
			TD td;
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited (td.testDir.get(), L"cmd /c exit 0", &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode1)
		{
			TD td;
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited (td.testDir.get(), L"cmd /c exit 1", &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}

		static void CancelAfterAsyncBuildProcessExited_NotOnLastCmd (const wchar_t* testDir, DWORD firstCommandExitCode, BOOL* complete, BOOL* success)
		{
			wchar_t tempFilename[MAX_PATH];
			UINT uires = GetTempFileNameW (testDir, L"TST", 0, tempFilename);
			Assert::IsTrue(uires > 0);
			static const wchar_t Format[] = L"cmd /c exit %u\r\ncmd /c del \"%s\"";
			size_t allocLen = _countof(Format) + 10 + wcslen(tempFilename);
			auto cmd = wil::make_hlocal_string_nothrow(nullptr, allocLen);
			Assert::IsNotNull(cmd.get());
			swprintf_s (cmd.get(), allocLen, Format, firstCommandExitCode, tempFilename);
			CancelAfterAsyncBuildProcessExited (testDir, cmd.get(), complete, success);

			// Since we canceled the build right after the first command, the second command
			// (the one that deletes the temporary file), shouldn't have been executed.
			BOOL fileExists = PathFileExists(tempFilename);
			DeleteFileW(tempFilename);
			Assert::IsTrue(fileExists);
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode0_NotOnLastCmd)
		{
			TD td;
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited_NotOnLastCmd(td.testDir.get(), 0, &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode1_NotOnLastCmd)
		{
			TD td;
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited_NotOnLastCmd(td.testDir.get(), 1, &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}
	};
}
