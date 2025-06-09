
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/inplace_function.h"
#include "Mocks.h"

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
		TEST_METHOD(BuildFailsOnEmptyProject)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto config = AddDebugProjectConfig(hier);
			auto pane = MakeMockOutputWindowPane(nullptr);

			com_ptr<IProjectConfigBuilder> builder;
			hr = MakeProjectConfigBuilder (hier, config, pane, &builder);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = builder->StartBuild (nullptr);
			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES), hr);
		}

		TEST_METHOD(ProjectConfigHasPrePostBuildProps)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto config = AddDebugProjectConfig(hier);

			com_ptr<IProjectConfigPrePostBuildProperties> props;
			hr = config->AsProjectConfigProperties()->get_PreBuildProperties(&props);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(props.get());

			hr = config->AsProjectConfigProperties()->get_PostBuildProperties(&props);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(props.get());
		}

		static com_ptr<IProjectConfigBuilder> MakeSjasmProjectBuilder (const char* asmFileContent)
		{
			HRESULT hr;
			com_ptr<IVsHierarchy> hier;
			hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			if (asmFileContent)
				WriteFileOnDisk(tempPath, L"test.asm", asmFileContent);
			else
				DeleteFileOnDisk(tempPath, L"test.asm");
			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\test.asm");
			LPCOLESTR filesToOpen[] = { filePath.get() };
			hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, filesToOpen, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			auto config = AddDebugProjectConfig(hier);
			auto pane = MakeMockOutputWindowPane(nullptr);
			com_ptr<IProjectConfigBuilder> builder;
			hr = MakeProjectConfigBuilder (hier, config, pane, &builder);
			Assert::IsTrue(SUCCEEDED(hr));
			return builder;
		}

		static void WaitCallbackWithMessageLoop (DWORD milliseconds, TestBuildCallback* callback)
		{
			DWORD tickStart = GetTickCount();
			while (!callback->_complete && (milliseconds == INFINITE || GetTickCount() - tickStart < milliseconds))
			{
				MSG msg;
				while(PeekMessage(&msg,0,0,0,PM_NOREMOVE))
				{
					if (::GetMessage(&msg, NULL, 0, 0) > 0)
						::DispatchMessage(&msg);
				}

				Sleep(10);
			}
		}

		TEST_METHOD(Test_SjasmMissingExe)
		{
			wil::unique_hlocal_string moduleFilename;
			auto hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, moduleFilename);
			Assert::IsTrue(SUCCEEDED(hr));
			PathFindFileName(moduleFilename.get())[0] = 0;
			auto sjasmOrigPath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine (sjasmOrigPath.get(), moduleFilename.get(), L"sjasmplus.exe");
			Assert::IsTrue(SUCCEEDED(hr));
			auto sjasmTempPath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine(sjasmTempPath.get(), moduleFilename.get(), L"sjasmplus.tmp");
			Assert::IsTrue(SUCCEEDED(hr));
			BOOL bres = MoveFileExW (sjasmOrigPath.get(), sjasmTempPath.get(), MOVEFILE_REPLACE_EXISTING);
			Assert::IsTrue(bres);

			auto builder = MakeSjasmProjectBuilder({ });
			auto callback = com_ptr(new TestBuildCallback());
			hr = builder->StartBuild (callback);

			bres = MoveFileExW(sjasmTempPath.get(), sjasmOrigPath.get(), MOVEFILE_REPLACE_EXISTING);
			Assert::IsTrue(bres);

			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), hr);
		}

		TEST_METHOD(Test_SjasmCommandLine_ExitCodeZero)
		{
			auto builder = MakeSjasmProjectBuilder("\tend\r\n");
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild (callback);
			Assert::IsTrue(SUCCEEDED(hr));

			WaitCallbackWithMessageLoop(INFINITE, callback);

			Assert::IsTrue(callback->_complete);
			Assert::IsTrue(callback->_success);
		}

		TEST_METHOD(Test_SjasmCommandLine_ExitCodeNonzero)
		{
			auto builder = MakeSjasmProjectBuilder({ });
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild (callback);
			Assert::IsTrue(SUCCEEDED(hr));

			WaitCallbackWithMessageLoop(INFINITE, callback);

			Assert::IsTrue(callback->_complete);
			Assert::IsFalse(callback->_success);
		}

		TEST_METHOD(Test_CustomBuildToolFilesInFolders)
		{
		}

		// sourceFileContent - empty string view to skip creating the file on disk
		static com_ptr<IProjectConfigBuilder> MakeProjectWithCustomBuildTool (
			const wchar_t* sourceFileName, const char* sourceFileContent,
			const wchar_t* cbtDescription,
			const wchar_t* cbtCmdLine, IStream* outputStreamUTF16)
		{
			HRESULT hr;

			com_ptr<IVsHierarchy> hier;
			hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			if (sourceFileContent)
				WriteFileOnDisk(tempPath, sourceFileName, sourceFileContent);
			else
				DeleteFileOnDisk(tempPath, sourceFileName);
			auto filePath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\", sourceFileName);
			LPCOLESTR filesToOpen[] = { filePath.get() };
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, filesToOpen, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			auto sourceFile = wil::try_com_query_nothrow<IFileNodeProperties>(hier.try_query<IParentNode>()->FirstChild());

			hr = sourceFile->put_BuildTool(BuildToolKind::CustomBuildTool);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<ICustomBuildToolProperties> cbtProps;
			hr = sourceFile->get_CustomBuildToolProperties(&cbtProps);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = cbtProps->put_Description(wil::make_bstr_nothrow(cbtDescription).get());
			Assert::IsTrue(SUCCEEDED(hr));
			hr = cbtProps->put_CommandLine(wil::make_bstr_nothrow(cbtCmdLine).get());
			Assert::IsTrue(SUCCEEDED(hr));

			auto config = AddDebugProjectConfig(hier);
			auto pane = MakeMockOutputWindowPane(outputStreamUTF16);

			com_ptr<IProjectConfigBuilder> builder;
			hr = MakeProjectConfigBuilder (hier, config, pane, &builder);
			Assert::IsTrue(SUCCEEDED(hr));
			return builder;
		}

		TEST_METHOD(TestCustomBuildToolOutputWithNoEOL)
		{
			HeavyLoad hl;

			com_ptr<IStream> outputStream;
			auto hr = CreateStreamOnHGlobal (NULL, TRUE, &outputStream);
			Assert::IsTrue(SUCCEEDED(hr));
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", "content", nullptr, L"cmd /c type test.xxx", outputStream);
			auto callback = com_ptr(new TestBuildCallback());
			hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			WaitCallbackWithMessageLoop(INFINITE, callback);
			Assert::IsTrue(callback->_complete);
			Assert::IsTrue(callback->_success);
			ULARGE_INTEGER curr;
			hr = outputStream->Seek (LARGE_INTEGER{ .QuadPart = 0 }, STREAM_SEEK_CUR, &curr);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue (curr.QuadPart > 0);
			wil::unique_bstr output;
			hr = MakeBstrFromStreamOnHGlobal (outputStream, &output);
			Assert::IsTrue(SUCCEEDED(hr));
			//if (FAILED(hr))
			//	Assert::Fail(std::to_wstring(hr).c_str());
			Assert::AreEqual(L"content", output.get());
		}

		TEST_METHOD(TestCustomBuildToolOutputWithEOL)
		{
			com_ptr<IStream> outputStream;
			auto hr = CreateStreamOnHGlobal (NULL, TRUE, &outputStream);
			Assert::IsTrue(SUCCEEDED(hr));
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", "content\r\n", nullptr, L"cmd /c type test.xxx", outputStream);
			auto callback = com_ptr(new TestBuildCallback());
			hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			WaitCallbackWithMessageLoop(INFINITE, callback);
			Assert::IsTrue(callback->_complete);
			Assert::IsTrue(callback->_success);
			ULARGE_INTEGER curr;
			hr = outputStream->Seek (LARGE_INTEGER{ .QuadPart = 0 }, STREAM_SEEK_CUR, &curr);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue (curr.QuadPart > 0);
			wil::unique_bstr output;
			hr = MakeBstrFromStreamOnHGlobal (outputStream, &output);
			Assert::IsTrue(SUCCEEDED(hr));
			//if (FAILED(hr))
			//	Assert::Fail(std::to_wstring(hr).c_str());
			Assert::AreEqual(L"content\r\n", output.get());
		}

		TEST_METHOD(TestBuilderDestroyedWhenReleasedWithPendingBuild)
		{
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", { }, nullptr, L"cmd /c pause", nullptr);
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			WaitCallbackWithMessageLoop(1000, callback);
			Assert::IsFalse(callback->_complete);
			ULONG remainingRefCount = builder.detach()->Release();
			Assert::AreEqual((ULONG)0, remainingRefCount);
		}

		TEST_METHOD(TestCustomBuildToolWaitingUserInput)
		{
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", { }, nullptr, L"cmd /c pause", nullptr);
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			WaitCallbackWithMessageLoop(1000, callback);
			Assert::IsFalse(callback->_complete);
			hr = builder->CancelBuild();
			Assert::IsTrue(callback->_complete);
			Assert::IsFalse(callback->_success);
			Assert::IsTrue(SUCCEEDED(hr));
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
		static void CancelAfterAsyncBuildProcessExited (const wchar_t* command, BOOL* complete, BOOL* success)
		{
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", "content", nullptr, command, nullptr);
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
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited (L"cmd /c exit 0", &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}

		class HeavyLoad
		{
			wil::unique_event_failfast event;
			vector_nothrow<wil::unique_handle> threads;

		public:
			HeavyLoad()
			{
				SYSTEM_INFO si;
				GetSystemInfo (&si);
				threads.try_resize(si.dwNumberOfProcessors);
				event.create(wil::EventOptions::ManualReset);
				for (auto& h : threads)
					h.reset(CreateThread(nullptr, 0, ThreadProc, this, 0, nullptr));
			}

			~HeavyLoad()
			{
				event.SetEvent();
				while(!threads.empty())
				{
					WaitForSingleObject(threads.back().get(), INFINITE);
					threads.remove_back();
				}
			}

		private:
			static DWORD WINAPI ThreadProc (void* arg)
			{
				HeavyLoad* _this = (HeavyLoad*)arg;
				DWORD tickStart = GetTickCount();
				while (!_this->event.is_signaled())
					;
				return (DWORD)0;
			}
		};

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode1)
		{
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited (L"cmd /c exit 1", &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}

		static void CancelAfterAsyncBuildProcessExited_NotOnLastCmd (DWORD firstCommandExitCode, BOOL* complete, BOOL* success)
		{
			wchar_t tempPath[MAX_PATH + 1];
			DWORD dwres = GetTempPathW (MAX_PATH + 1, tempPath);
			Assert::IsTrue(dwres > 0);
			wchar_t tempFilename[MAX_PATH];
			UINT uires = GetTempFileNameW (tempPath, L"TST", 0, tempFilename);
			Assert::IsTrue(uires > 0);
			static const wchar_t Format[] = L"cmd /c exit %u\r\ncmd /c del \"%s\"";
			size_t allocLen = _countof(Format) + 10 + wcslen(tempFilename);
			auto cmd = wil::make_hlocal_string_nothrow(nullptr, allocLen);
			Assert::IsNotNull(cmd.get());
			swprintf_s (cmd.get(), allocLen, Format, firstCommandExitCode, tempFilename);
			CancelAfterAsyncBuildProcessExited (cmd.get(), complete, success);

			// Since we canceled the build right after the first command, the second command
			// (the one that deletes the temporary file), shouldn't have been executed.
			BOOL fileExists = PathFileExists(tempFilename);
			DeleteFileW(tempFilename);
			Assert::IsTrue(fileExists);
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode0_NotOnLastCmd)
		{
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited_NotOnLastCmd(0, &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}

		TEST_METHOD(CancelAfterAsyncBuildProcessExitedWithExitCode1_NotOnLastCmd)
		{
			BOOL complete, success;
			CancelAfterAsyncBuildProcessExited_NotOnLastCmd(1, &complete, &success);
			Assert::IsTrue(complete);
			Assert::IsFalse(success);
		}

		TEST_METHOD(TestCustomBuildToolOnlyWhitespaceCommands)
		{
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", { }, nullptr, L"   \r\n   \t   ", nullptr);
			auto callback = com_ptr(new TestBuildCallback());
			auto hr = builder->StartBuild(callback);
			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES), hr);
		}

		TEST_METHOD(TestCustomBuildToolSomeWhitespaceCommands)
		{
			com_ptr<IStream> outputStream;
			auto hr = CreateStreamOnHGlobal (NULL, TRUE, &outputStream);
			Assert::IsTrue(SUCCEEDED(hr));
			static const wchar_t cmdLine[] = L"  cmd /c type test.xxx  \t\r\n   \t   ";
			auto builder = MakeProjectWithCustomBuildTool(L"test.xxx", "content", nullptr, cmdLine, outputStream);
			auto callback = com_ptr(new TestBuildCallback());
			hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			WaitCallbackWithMessageLoop(INFINITE, callback);
			Assert::IsTrue(callback->_complete);
			Assert::IsTrue(callback->_success);
			wil::unique_bstr output;
			hr = MakeBstrFromStreamOnHGlobal (outputStream, &output);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"content", output.get());
		}

		TEST_METHOD(TestPrePostBuildEvents)
		{
			HRESULT hr;

			com_ptr<IVsHierarchy> hier;
			hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto config = AddDebugProjectConfig(hier);

			com_ptr<IProjectConfigPrePostBuildProperties> preBuildProps;
			hr = config->AsProjectConfigProperties()->get_PreBuildProperties(&preBuildProps);
			Assert::IsTrue(SUCCEEDED(hr));
			auto preCmdLine = wil::make_bstr_nothrow(L"cmd /c echo XXX");
			hr = preBuildProps->put_CommandLine(preCmdLine.get());
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IProjectConfigPrePostBuildProperties> postBuildProps;
			hr = config->AsProjectConfigProperties()->get_PostBuildProperties(&postBuildProps);
			Assert::IsTrue(SUCCEEDED(hr));
			auto postCmdLine = wil::make_bstr_nothrow(L"cmd /c echo YYY");
			hr = postBuildProps->put_CommandLine(postCmdLine.get());
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IStream> outputStream;
			hr = CreateStreamOnHGlobal (NULL, TRUE, &outputStream);
			Assert::IsTrue(SUCCEEDED(hr));
			auto pane = MakeMockOutputWindowPane(outputStream);
			com_ptr<IProjectConfigBuilder> builder;
			hr = MakeProjectConfigBuilder (hier, config, pane, &builder);
			Assert::IsTrue(SUCCEEDED(hr));

			auto callback = com_ptr(new TestBuildCallback());
			hr = builder->StartBuild(callback);
			Assert::IsTrue(SUCCEEDED(hr));
			//auto callback = com_ptr(new TestBuildCallback());
			//auto hr = builder->StartBuild(callback);
			//Assert::IsTrue(SUCCEEDED(hr));
			//WaitWithMessageLoop(500, callback);
			//Assert::IsFalse(callback->_complete);
			//ULONG remainingRefCount = builder.detach()->Release();
			//Assert::AreEqual((ULONG)0, remainingRefCount);
			ULONG remainingRefCount = builder.detach()->Release();
		}

		TEST_METHOD(BuildOnlySynchronousSteps)
		{
			auto builder = MakeProjectWithCustomBuildTool (L"test.asm", { }, L"CBT Description", nullptr, nullptr);
			auto hr = builder->StartBuild(nullptr);
			Assert::AreEqual(S_OK, hr);
		}

		TEST_METHOD(BuildFilesInFoldersAndSubfolders)
		{
			com_ptr<IVsUIHierarchy> hier;
			auto hr = MakeProjectNode (TemplatePath_EmptyProject.get(), tempPath, L"BuildFilesInFoldersAndSubfolders.flx",
				CPF_CLONEFILE | CPF_OVERWRITE | CPF_SILENT, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto config = FelixTests::AddDebugProjectConfig(hier);

			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };

			wil::unique_variant folder;
			hr = hier->ExecCommand (VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &folder);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, folder.vt);
			hr = hier->SetProperty (V_VSITEMID(&folder), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"folder"));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier.try_query<IVsProject>()->AddItem (V_VSITEMID(&folder), VSADDITEMOP_CLONEFILE, L"file1.asm", 1, templateasm, NULL, NULL);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant subfolder;
			hr = hier->ExecCommand (V_VSITEMID(&folder), &CMDSETID_StandardCommandSet97, cmdidNewFolder, 0, nullptr, &subfolder);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, subfolder.vt);
			hr = hier->SetProperty (V_VSITEMID(&subfolder), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"subfolder"));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier.try_query<IVsProject>()->AddItem (V_VSITEMID(&subfolder), VSADDITEMOP_CLONEFILE, L"file2.asm", 1, templateasm, NULL, NULL);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_bstr cmdLine;
			hr = MakeSjasmCommandLine (hier, config, nullptr, &cmdLine);
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsNotNull(wcsstr(cmdLine.get(), L" folder\\file1.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), L" folder\\subfolder\\file2.asm"));
		}

		TEST_METHOD(BuildFilesNotInProjectDir)
		{
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\BuildFilesNotInProjectDir");
			auto projDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\projdir");

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, projDir.get(), nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto project = hier.try_query<IVsProject>();

			auto config = FelixTests::AddDebugProjectConfig(hier);

			// file 1 in project dir
			auto file1FullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(projDir, L"\\file1.asm");
			// file 2 in project sub dir
			auto file2FullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(projDir, L"\\subdir\\file2.asm");
			// file 3 outside project dir but on same drive
			auto file3FullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\file3.asm");
			// file 4 on different drive
			auto file4FullPath = wil::make_process_heap_string_failfast(L"D:\\FelixTest\\BuildFilesNotInProjectDir\\file4.asm");
			
			LPCOLESTR files[] = { file1FullPath.get(), file2FullPath.get(), file3FullPath.get(), file4FullPath.get() };
			
			hr = project->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, (ULONG)_countof(files), files, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_bstr cmdLine;
			hr = MakeSjasmCommandLine (hier, config, nullptr, &cmdLine);
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsNotNull(wcsstr(cmdLine.get(), L" file1.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), L" subdir\\file2.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), L" ..\\file3.asm"));
			Assert::IsNotNull(wcsstr(cmdLine.get(), file4FullPath.get()));
		}
	};
}
