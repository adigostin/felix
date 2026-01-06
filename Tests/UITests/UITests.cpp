
#include "pch.h"
#include "shared/com.h"
#include "FelixPackage.h"
#include "../TestsCommon.h"

#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>
namespace VxDTE
{
	#include <dte80.h>
	#include <dte90.h>
}

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace UITests
{
	com_ptr<IUIAutomation> automation;
	com_ptr<VxDTE::_DTE> dte;

	TEST_CLASS(UITests)
	{
		static void StartOrRestart()
		{
			HRESULT hr;

			wil::unique_process_heap_string devenvExe;
			hr = wil::GetEnvironmentVariableW (L"VSAPPIDDIR", devenvExe);
			Assert::IsTrue(SUCCEEDED(hr));
			devenvExe = wil::str_concat_failfast<wil::unique_process_heap_string>(devenvExe, L"devenv.exe");
			const wchar_t* devenvArguments = L"/rootSuffix Exp";

			auto cmdLine = wil::str_concat_failfast<wil::unique_process_heap_string>(L"\"", devenvExe, L"\" ", devenvArguments);
			STARTUPINFO si = { .cb = sizeof(si) };
			wil::unique_process_information pi;
			BOOL bres = CreateProcessW (devenvExe.get(), cmdLine.get(), nullptr, nullptr, TRUE, 0, nullptr, nullptr, &si, &pi);
			Assert::IsTrue(bres);
			DWORD exitCode;
			bres = GetExitCodeProcess (pi.hProcess, &exitCode);
			Assert::IsTrue(bres);
			Assert::AreEqual(STILL_ACTIVE, exitCode);

			com_ptr<VxDTE::_DTE> dte;
			DWORD startTime = GetTickCount();
			while (GetTickCount() - startTime <= 60'000 && !dte)
			{
				if (SUCCEEDED(GetDTE(pi.dwProcessId, &dte)))
					break;
				Sleep(250);
			}

			Assert::IsNotNull(dte.get(), L"Failed to start VS");

			// The window we get from DTE is the main window (even if it's invisible), not the splash window.
			com_ptr<VxDTE::Window> dteMainWindow;
			hr = dte->get_MainWindow(&dteMainWindow);
			Assert::IsTrue(SUCCEEDED(hr));
			long dteMainWindowHWnd;
			hr = dteMainWindow->get_HWnd(&dteMainWindowHWnd);
			Assert::IsTrue(SUCCEEDED(hr));

			auto handle = FindTopLevelWindow(pi.dwProcessId);
			Assert::IsNotNull(handle);
			if (handle != (HWND)(size_t)(DWORD)dteMainWindowHWnd)
			{
				// We have the splash window up.
				com_ptr<IUIAutomationElement> topLevelWindow;
				hr = automation->ElementFromHandle(handle, &topLevelWindow);
				Assert::IsTrue(SUCCEEDED(hr));
				com_ptr<IUIAutomationCondition> condition;
				hr = automation->CreatePropertyCondition(UIA_NamePropertyId, wil::make_variant_bstr_failfast(L"Continue without code"), &condition);
				Assert::IsTrue(SUCCEEDED(hr));
				com_ptr<IUIAutomationElement> cwc;
				hr = topLevelWindow->FindFirst(TreeScope_Descendants, condition, &cwc);
				Assert::IsTrue(SUCCEEDED(hr));
				com_ptr<IUnknown> patternUnk;
				hr = cwc->GetCurrentPattern(UIA_InvokePatternId, &patternUnk);
				Assert::IsTrue(SUCCEEDED(hr));
				com_ptr<IUIAutomationInvokePattern> invpat;
				hr = patternUnk->QueryInterface(IID_PPV_ARGS(&invpat));
				Assert::IsTrue(SUCCEEDED(hr));
				hr = invpat->Invoke();
				Assert::IsTrue(SUCCEEDED(hr));

				// Let's wait for the main window to become the top level window
				DWORD tickStart = GetTickCount();
				while (true)
				{
					handle = FindTopLevelWindow(pi.dwProcessId);
					if (handle == (HWND)(size_t)(DWORD)dteMainWindowHWnd)
						break;
					if (GetTickCount() - tickStart >= 5000)
						Assert::Fail();
					Sleep(100);
				}
			}

			if (IsDebuggerPresent())
				FindAndAttach(pi.dwProcessId, dte);

			::UITests::dte = std::move(dte);
		}

		static void CloseCurrentInstance(bool hard = false)
		{
			if (dte)
			{
				com_ptr<VxDTE::Window> dteMainWindow;
				auto hr = dte->get_MainWindow(&dteMainWindow);
				Assert::IsTrue(SUCCEEDED(hr));
				long dteMainWindowHWnd;
				hr = dteMainWindow->get_HWnd(&dteMainWindowHWnd);
				Assert::IsTrue(SUCCEEDED(hr));
				DWORD processID;
				GetWindowThreadProcessId((HWND)(size_t)(DWORD)dteMainWindowHWnd, &processID);
				wil::unique_handle hProcess (OpenProcess (SYNCHRONIZE, FALSE, processID));

				bool closed = false;
				if (!hard && SUCCEEDED(dte->Quit()))
				{
					DWORD waitRes = WaitForSingleObject(hProcess.get(), IsDebuggerPresent() ? INFINITE : 10000);
					if (waitRes == WAIT_OBJECT_0)
						closed = true;
				}

				if (!closed)
					TerminateProcess(hProcess.get(), 1234);

				dte.reset();
			}
		}

		static void FindAndAttach (DWORD targetVsProcessId, VxDTE::_DTE* targetVSDTE)
		{
			HRESULT hr;

			// We (testhost.exe) are being debugged. Find the VS instance that's debugging us
			// and tell it to debug the target VS too.
			DWORD selfId = GetCurrentProcessId();

			// Get all running processes.
			DWORD cap = 300;
			wil::unique_hlocal_ptr<DWORD[]> pids;
			DWORD processCount;
			while(true)
			{
				pids = wil::make_unique_hlocal_failfast<DWORD[]>(cap);
				DWORD neededBytes;
				BOOL bres = EnumProcesses (pids.get(), cap * sizeof(DWORD), &neededBytes);
				Assert::IsTrue(bres);
				processCount = neededBytes / sizeof(DWORD);
				if (processCount < cap)
					break;
				cap += cap;
			}

			// Retain only the "devenv.exe" processes.
			DWORD vsCount = 0;
			for (DWORD i = 0; i < cap; i++)
			{
				wil::unique_process_handle h (OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, pids.get()[i]));
				if (h.is_valid())
				{
					wil::unique_process_heap_string path;
					hr = wil::QueryFullProcessImageNameW (h.get(), 0, path);
					if (SUCCEEDED(hr))
					{
						auto p = wcsrchr(path.get(), L'\\');
						if (p && !wcscmp(p + 1, L"devenv.exe"))
							pids.get()[vsCount++] = pids.get()[i];
					}
				}
			}

			// Now find the VS process that's debugging us.
			for (DWORD i = 0; i < vsCount; i++)
			{
				DWORD vspid = pids.get()[i];
				if (vspid == targetVsProcessId)
					continue;
				com_ptr<VxDTE::_DTE> dte;
				if (SUCCEEDED(GetDTE(vspid, &dte)))
				{
					com_ptr<VxDTE::Debugger> debugger;
					VxDTE::dbgDebugMode mode;
					com_ptr<VxDTE::Processes> processes;
					com_ptr<IUnknown> newEnum;
					com_ptr<IEnumVARIANT> enumVar;
					if (SUCCEEDED(dte->get_Debugger(&debugger))
						&& SUCCEEDED(debugger->get_CurrentMode(&mode))
						&& mode != VxDTE::dbgDesignMode
						&& SUCCEEDED(debugger->get_DebuggedProcesses(&processes))
						&& SUCCEEDED(processes->_NewEnum(&newEnum))
						&& SUCCEEDED(newEnum->QueryInterface(IID_PPV_ARGS(&enumVar))))
					{
						wil::unique_variant var;
						ULONG fetched;
						while (enumVar->Next(1, &var, &fetched) == S_OK)
						{
							com_ptr<VxDTE::Process> process;
							if (var.vt == VT_DISPATCH && SUCCEEDED(var.pdispVal->QueryInterface(IID_PPV_ARGS(&process))))
							{
								long pid;
								if (SUCCEEDED(process->get_ProcessID(&pid))
									&& (DWORD)pid == selfId)
								{
									// This is the correct VS, so attach and return.
									Attach (targetVsProcessId, targetVSDTE, vspid, debugger);
									return;
								}
							}
						}
					}
				}
			}

			Assert::Fail();
		}

		static void Attach (DWORD targetVSProcessID, VxDTE::_DTE* targetVSDTE, DWORD debuggerVsProcessId, VxDTE::Debugger* debugger)
		{
			HRESULT hr;
			auto dbg3 = wil::com_query_failfast<VxDTE::Debugger3>(debugger);
			com_ptr<VxDTE::Processes> processes;
			hr = dbg3->get_LocalProcesses(&processes);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IUnknown> newEnum;
			hr = processes->_NewEnum(&newEnum);
			Assert::IsTrue(SUCCEEDED(hr));
			auto enumProcesses = wil::com_query_failfast<IEnumVARIANT>(newEnum);
			wil::unique_variant var;
			ULONG fetched;
			while (enumProcesses->Next(1, &var, &fetched) == S_OK)
			{
				if (auto targetProcess = wil::try_com_query_failfast<VxDTE::Process3>(var.pdispVal))
				{
					long pid;
					if (SUCCEEDED(targetProcess->get_ProcessID(&pid))
						&& (DWORD)pid == targetVSProcessID)
					{
						targetVSDTE->put_SuppressUI(VARIANT_TRUE);
						//hr = targetProcess->Attach2(wil::make_variant_bstr_failfast(L"Managed/Native"));
						hr = targetProcess->Attach2(wil::make_variant_bstr_failfast(L"Native"));
						targetVSDTE->put_SuppressUI(VARIANT_FALSE);
						Assert::IsTrue(SUCCEEDED(hr));
						return;
					}
				}
			}

			Assert::Fail();
		}

		static HRESULT GetDTE (DWORD processId, VxDTE::_DTE** ppDTE)
		{
			HRESULT hr;

			//MessageFilter.Register();

			wil::unique_process_heap_string fn;
			hr = wil::GetModuleFileNameW(nullptr, fn); RETURN_IF_FAILED_EXPECTED(hr);
			DWORD dummydw;
			DWORD verlen = GetFileVersionInfoSize (fn.get(), &dummydw); RETURN_LAST_ERROR_IF_EXPECTED(verlen == 0);
			auto verbuffer = wil::make_unique_nothrow<char[]>(verlen);
			BOOL bres = GetFileVersionInfo(fn.get(), 0, verlen, verbuffer.get()); RETURN_IF_WIN32_BOOL_FALSE_EXPECTED(bres);
			void* valbuffer;
			UINT vallen;
			bres = VerQueryValueW (verbuffer.get(), L"\\", &valbuffer, &vallen); RETURN_HR_IF_EXPECTED(E_FAIL, !bres);
			VS_FIXEDFILEINFO* fi = (VS_FIXEDFILEINFO*)valbuffer;
			DWORD vsMajorVersion = fi->dwProductVersionMS >> 16;
			auto progId = wil::str_printf_failfast<wil::unique_process_heap_string> (L"!VisualStudio.DTE.%u.0:%u", vsMajorVersion, processId);

			com_ptr<IUnknown> runningObject;
			com_ptr<IBindCtx> bindCtx;
			com_ptr<IRunningObjectTable> rot;
			com_ptr<IEnumMoniker> enumMonikers;

			hr = CreateBindCtx(0, &bindCtx); RETURN_IF_FAILED_EXPECTED(hr);
			hr = bindCtx->GetRunningObjectTable(&rot); RETURN_IF_FAILED_EXPECTED(hr);
			hr = rot->EnumRunning(&enumMonikers); RETURN_IF_FAILED_EXPECTED(hr);
			
			com_ptr<IMoniker> moniker;
			ULONG numberFetched = 0;
			while (enumMonikers->Next (1, moniker.addressof(), &numberFetched) == S_OK)
			{
				wil::unique_cotaskmem_string name;
				hr = moniker->GetDisplayName(bindCtx, nullptr, &name);
				if (FAILED(hr))
				{
					if (hr == E_ACCESSDENIED)
					{
						// Do nothing, there is something in the ROT that we do not have access to.
					}
					else
						return hr;
				}

				if (name && !wcscmp(name.get(), progId.get()))
				{
					hr = rot->GetObject(moniker, &runningObject);
					Assert::IsTrue(SUCCEEDED(hr));
					Assert::IsNotNull(runningObject.get());
					hr = runningObject->QueryInterface(IID_PPV_ARGS(ppDTE)); RETURN_IF_FAILED_EXPECTED(hr);
					return S_OK;
				}
			}

			return E_FAIL;
		}

		static HWND FindTopLevelWindow (DWORD process_id)
		{
			struct handle_data {
				DWORD process_id;
				HWND window_handle;
			};

			handle_data data;
			data.process_id = process_id;
			data.window_handle = 0;

			auto callback = [](HWND handle, LPARAM lParam) -> BOOL
				{
					handle_data* data = (handle_data*)lParam;
					DWORD process_id = 0;
					GetWindowThreadProcessId(handle, &process_id);
					if (data->process_id != process_id || GetWindow(handle, GW_OWNER) || !IsWindowVisible(handle))
						return TRUE;
					data->window_handle = handle;
					return FALSE;   
				};
			EnumWindows(callback, (LPARAM)&data);
			return data.window_handle;
		}

		// <param name="testDir">The directory in which to create the solution.</param>
		// <param name="solutionName">The name of the solution to create, without extension.</param>
		// <param name="projectName">NULL to create a project with the same name as the solution in the same dir, filename without extension to create project in subdir with different name.</param>
		static std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
			CreateSolutionAndProject (PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName)
		{
			HRESULT hr;
			wil::com_ptr_failfast<IUnknown> solution;
			hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::IsTrue(SUCCEEDED(hr));
			auto sln = solution.query<VxDTE::_Solution>();
			hr = sln->Create(wil::make_bstr_failfast(testDir).get(), wil::make_bstr_failfast(solutionName).get());
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_process_heap_string projDirBuffer, projName;
			PCWSTR projDir;
			if (projectName)
			{
				projDirBuffer = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\", projectName);
				projDir = projDirBuffer.get();
				projName = wil::str_concat_failfast<wil::unique_process_heap_string>(projectName, L".flx");
				Assert::IsTrue(CreateDirectory(projDir, nullptr));
			}
			else
			{
				projDir = testDir;
				projName = wil::str_concat_failfast<wil::unique_process_heap_string>(solutionName, L".flx");
			}

			wil::com_ptr_failfast<VxDTE::Project> proj;
			hr = sln->AddFromTemplate (
				wil::make_bstr_failfast(TemplatePath_TwoConfigsOneFile.get()).get(),
				wil::make_bstr_failfast(projDir).get(),
				wil::make_bstr_failfast(projName.get()).get(), VARIANT_FALSE, &proj);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = sln->SaveAs(wil::make_bstr_failfast(solutionName).get());
			Assert::IsTrue(SUCCEEDED(hr));

			return { std::move(sln), std::move(proj) };
		}

		static void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount)
		{
			HRESULT hr;

			com_ptr<VxDTE::SolutionBuild> solutionBuild;
			hr = sln->get_SolutionBuild(&solutionBuild);
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<VxDTE::SolutionConfiguration> solConfig;
			hr = solutionBuild->get_ActiveConfiguration(&solConfig); 
			Assert::IsTrue(SUCCEEDED(hr));
			hr = solutionBuild->Build(VARIANT_TRUE);
			Assert::IsTrue(SUCCEEDED(hr));

			// LastBuildInfo returns the number of failed projects, despite the parameter name.
			hr = solutionBuild->get_LastBuildInfo(buildFailCount);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		static void __stdcall WilLoggingCallback (const wil::FailureInfo& fi) noexcept
		{
			Assert::Fail(fi.pszMessage);
		}

	public:
		TEST_CLASS_INITIALIZE(UITestsInitialize)
		{
			wil::SetResultLoggingCallback (WilLoggingCallback);

			MakeTemplates (L"FelixTestUI");

			auto hr = CoCreateInstance (__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation));
			Assert::IsTrue(SUCCEEDED(hr));
			StartOrRestart();
		}

		TEST_CLASS_CLEANUP(UITestsCleanup)
		{
			CloseCurrentInstance();
			automation.reset();

			wil::SetResultLoggingCallback(nullptr);
		}

		TEST_METHOD(LaunchVS)
		{
			com_ptr<IUnknown> solution;
			auto hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::IsTrue(SUCCEEDED(hr));
			auto sln = wil::com_query_failfast<VxDTE::_Solution>(solution);
			sln->Close();
		}

		TEST_METHOD(CloneProject)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"CloneProject");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

			com_ptr<IUnknown> solution;
			hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::IsTrue(SUCCEEDED(hr));
			auto sln = wil::com_query_failfast<VxDTE::_Solution>(solution);
			hr = sln->Create(wil::make_bstr_failfast(testPath.get()).get(), wil::make_bstr_failfast(L"test.sln").get());
			Assert::IsTrue(SUCCEEDED(hr));
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			com_ptr<VxDTE::Project> proj;
			hr = sln->AddFromTemplate (
				wil::make_bstr_failfast(TemplatePath_TwoConfigsOneFile.get()).get(),
				wil::make_bstr_failfast(testPath.get()).get(),
				wil::make_bstr_failfast(L"test.flx").get(), VARIANT_TRUE, &proj);
			Assert::IsTrue(SUCCEEDED(hr), wil::str_printf_failfast<wil::unique_process_heap_string>(L"0x%08x", hr).get());
		}

		TEST_METHOD(BuildProject)
		{
			HRESULT hr;

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildProject");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			long buildFailCount;
			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);
		}

		TEST_METHOD(BuildProjectWithError)
		{
			HRESULT hr;

			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"BuildProjectWithError");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			long buildFailCount;
			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);

			WriteFileOnDisk (CombinePath(testPath.get(), L"file.asm").get(), "\tabcde");

			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(1l, buildFailCount);

			hr = dte->ExecuteCommand(wil::make_bstr_failfast(L"View.ErrorList").get());
			auto dte2 = wil::com_query_failfast<VxDTE::DTE2>(dte);
			wil::com_ptr_failfast<VxDTE::ToolWindows> toolWindows;
			hr = dte2->get_ToolWindows(&toolWindows);
			wil::com_ptr_failfast<VxDTE::ErrorList> errorList;
			hr = toolWindows->get_ErrorList(&errorList);
			wil::com_ptr_failfast<VxDTE::ErrorItems> errorItems;
			hr = errorList->get_ErrorItems(&errorItems);
			long errorCount;
			hr = errorItems->get_Count(&errorCount);
			Assert::AreEqual(1l, errorCount);
		}

		TEST_METHOD(OpenSpecificEditor)
		{
			// At some point between VS 17.10 and 17.14, the VS implementation of IVsUIShellOpenDocument::OpenStandardEditor
			// started trying to open our .asm files not with "Source Code (Text) Editor", but with
			// "Common Language Editor Supporting TextMate Bundles" (or at least this was the new default shown in Open With...).
			// The implementation started returning E_NOTIMPL in most cases.
			// The solution was to call OpenSpecificEditor with GUID_TextEditorFactory, instead of OpenStandardEditor.
			// This test verifies this solution. The test fails with the old code that calls OpenStandardEditor,
			// and succeeds with the new code that calls OpenSpecificEditor.

			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"OpenSpecificEditor");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

			auto[sln, proj] = CreateSolutionAndProject(testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			auto hier = proj.query<IVsUIHierarchy>();
			VSITEMID itemid;
			hr = hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			auto vsp2 = proj.query<IVsProject2>();
			com_ptr<IVsWindowFrame> wf;
			hr = vsp2->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = wf->Show();
			Assert::IsTrue(SUCCEEDED(hr));
			hr = sln->Close();
			Assert::IsTrue(SUCCEEDED(hr));

			auto fullSlnPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\test.sln");
			hr = sln->Open(wil::make_bstr_failfast(fullSlnPath.get()).get()); 
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<VxDTE::Projects> projects;
			hr = sln->get_Projects(&projects);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = projects->Item(wil::make_variant_bstr_failfast(L"test.flx"), &proj);
			Assert::IsTrue(SUCCEEDED(hr));

			hier = proj.query<IVsUIHierarchy>();
			hr = hier->ParseCanonicalName(L"file.asm", &itemid);
			Assert::IsTrue(SUCCEEDED(hr));
			vsp2 = proj.query<IVsProject2>();
			hr = vsp2->OpenItem (itemid, LOGVIEWID_Primary, nullptr, &wf); // This would return E_NOTIMPL before the fix.
			Assert::IsTrue(SUCCEEDED(hr));
		}

		struct FileChangeEvents : IVsFileChangeEvents
		{
			ULONG _refCount = 0;
			bool _changed = false;

			#pragma region IUnknown
			virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
			{
				if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<IVsFileChangeEvents>(this, riid, ppvObject))
					return S_OK;

				*ppvObject = nullptr;
				return E_NOINTERFACE;
			}
			virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
			virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
			#pragma endregion

			#pragma region IVsFileChangeEvents
			virtual HRESULT STDMETHODCALLTYPE FilesChanged (DWORD cChanges, LPCOLESTR rgpszFile[], VSFILECHANGEFLAGS rggrfChange[]) override
			{
				_changed = true;
				return S_OK;
			}

			virtual HRESULT STDMETHODCALLTYPE DirectoryChanged (LPCOLESTR pszDirectory) override
			{
				RETURN_HR(E_NOTIMPL);
			}
			#pragma endregion
		};

		TEST_METHOD(AdviseFileChangeNotCalledOnProjectSave)
		{
			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"AdviseFileChangeNotCalledOnProjectSave");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", nullptr);
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			com_ptr<IDispatch> aodisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"TestHelper").get(), &aodisp);
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IFelixTestHelper> ao;
			hr = aodisp->QueryInterface(IID_PPV_ARGS(&ao));
			Assert::IsTrue(SUCCEEDED(hr));

			auto fce = com_ptr(new (std::nothrow) FileChangeEvents()); FAIL_FAST_IF_NULL_ALLOC(fce);
			DWORD cookie;
			hr = ao->AdviseProjectFileChange(proj, fce, &cookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([ao=ao.get(), cookie]() {
				ao->UnadviseProjectFileChange(cookie);
			});

			hr = proj->Save(nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			// give it some time to notice the file changed on disk, and to call our callback
			DWORD tickStart = GetTickCount();
			while (!fce->_changed && GetTickCount() - tickStart < 1000)
			{
				MSG msg;
				while(PeekMessage(&msg,0,0,0,PM_NOREMOVE))
				{
					if (::GetMessage(&msg, NULL, 0, 0) > 0)
						::DispatchMessage(&msg);
				}

				Sleep(10);
			}

			bool changed = fce->_changed;
			hr = CoDisconnectObject(fce, 0);
			Assert::IsTrue(SUCCEEDED(hr));
			fce = nullptr;

			Assert::IsFalse(changed); // this would fail before the fix (IgnoreFile called in ProjectNode::Save)
		}

		TEST_METHOD(NavigateToErrorInFileInSubdir)
		{
			// When the user double-clicks an error in the Error List window, VS goes out of its way to locate the file
			// and find the project it contains. To repro the bug, we need to put the project file in a subdirectory
			// (not next to the .sln file), and the offending file in a subdirectory of its own. In this case, when the user
			// double-clicks the error in the Error List window, VS used to be unable to locate the file due to bugs in
			// ProjectNode::ParseCanonicalName and ProjectNode::IsDocumentInProject (and ultimately in ParseCanonicalName).
			// This test verifies the fixes in these functions.

			HRESULT hr;
			auto testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"NavigateToErrorInFileInSubdir");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			auto delDir = wil::scope_exit([tp=testPath.get()] { std::error_code ec; std::filesystem::remove_all(tp, ec); });

			auto[sln, proj] = CreateSolutionAndProject (testPath.get(), L"test", L"testproj");
			auto close = wil::scope_exit([sln=sln.get()] { sln->Close(); });

			long buildFailCount;
			BuildSolution(sln, &buildFailCount);
			Assert::AreEqual(0l, buildFailCount);

			auto file1Path = CombinePath(testPath.get(), L"testproj\\subdir\\file1.asm");
			WriteFileOnDisk (file1Path.get(), "\t555555");

			com_ptr<VxDTE::ProjectItems> items;
			proj->get_ProjectItems(&items);
			wil::com_ptr_failfast<VxDTE::ProjectItem> item1;
			hr = items->AddFromFile(wil::make_bstr_failfast(file1Path.get()).get(), &item1);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj->Save();
			Assert::IsTrue(SUCCEEDED(hr));

			BuildSolution(sln, &buildFailCount);
			Assert::IsTrue(buildFailCount >= 1);

			hr = dte->ExecuteCommand(wil::make_bstr_failfast(L"View.ErrorList").get());
			auto dte2 = wil::com_query_failfast<VxDTE::DTE2>(dte);
			wil::com_ptr_failfast<VxDTE::ToolWindows> toolWindows;
			hr = dte2->get_ToolWindows(&toolWindows);
			wil::com_ptr_failfast<VxDTE::ErrorList> errorList;
			hr = toolWindows->get_ErrorList(&errorList);
			wil::com_ptr_failfast<VxDTE::ErrorItems> errorItems;
			hr = errorList->get_ErrorItems(&errorItems);
			long errorCount;
			hr = errorItems->get_Count(&errorCount);
			Assert::IsTrue(errorCount >= 1);

			VARIANT v; v.vt = VT_I4; v.lVal = 1;
			com_ptr<VxDTE::ErrorItem> errorItem;
			hr = errorItems->Item(v, &errorItem);
			Assert::IsTrue(SUCCEEDED(hr));

			errorItem->Navigate(); // This call succeeds, even though VS couldn't open the file.

			com_ptr<VxDTE::Document> doc;
			hr = dte->get_ActiveDocument(&doc);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsNotNull(doc.get()); // This would fail before the fix.

			BOOL found = FALSE;
			VSDOCUMENTPRIORITY prio = { };
			VSITEMID itemid = { };
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"subdir\\file1.asm", &found, &prio, &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"subdir/..\\subdir/.\\file1.asm", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);
			hr = proj.query<IVsProject>()->IsDocumentInProject(L"SubDir\\File1.ASM", &found, &prio, &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(found);

			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"subdir\\file1.asm", &itemid); // this would fail too
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"subdir/..\\subdir/.\\file1.asm", &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"SubDir\\File1.ASM", &itemid); // for bonus points
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(GenPrePostInclude_OnlyActiveCfg)
		{
			// Verify that the PreInclude.asm and PostInclude.asm are regenerated when editing the active configuration,
			// and that they are _not_ regenerated when editing an inactive configuration.

			HRESULT hr;

			auto testPath = std::filesystem::path(tempPath) / "GenPrePostInclude_OnlyActiveCfg";
			std::filesystem::create_directory(testPath);
			auto delDir = wil::scope_exit([&testPath] { std::error_code ec; std::filesystem::remove_all(testPath, ec); });

			auto[sln0, proj] = CreateSolutionAndProject (testPath.c_str(), L"test", L"testproj");
			auto close = wil::scope_exit([sln=sln0.get()] { sln->Close(); });

			// Let's not go through the configuration manager since we haven't implemented Project::get_ConfigurationManager yet.
			//Microsoft_VisualStudio_Interop::_SolutionPtr sln = sln0.get();
			//auto configs = sln->SolutionBuild->SolutionConfigurations;
			//Assert::IsTrue(configs->Count >= 2);
			//Assert::AreEqual<void*>(configs->Item[1].GetInterfacePtr(), sln->SolutionBuild->ActiveConfiguration.GetInterfacePtr());

			com_ptr<IVsCfg> cfgs[2];
			ULONG actual;
			VSCFGFLAGS flags;
			hr = proj.query<IVsCfgProvider>()->GetCfgs(2, cfgs[0].addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			// Make a change in the active configuration and verify that the Pre/PostInclude files are generated.
			auto genFilesPath = testPath / L"testproj" / L"GeneratedFiles";
			Assert::IsTrue(std::filesystem::exists(genFilesPath));
			std::filesystem::remove_all(genFilesPath);

			com_ptr<IProjectConfigProperties> props;
			hr = cfgs[0]->QueryInterface(IID_PPV_ARGS(&props));
			Assert::IsTrue(SUCCEEDED(hr));
			com_ptr<IProjectConfigAssemblerProperties> asmProps;
			hr = props->get_AssemblerProperties(&asmProps);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = asmProps->put_BaseAddress(1234);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(std::filesystem::exists(genFilesPath));

			// Now make a change in the inactive configuration.
			std::filesystem::remove_all(genFilesPath);

			hr = cfgs[1]->QueryInterface(IID_PPV_ARGS(&props));
			Assert::IsTrue(SUCCEEDED(hr));
			hr = props->get_AssemblerProperties(&asmProps);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = asmProps->put_BaseAddress(1234);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(!std::filesystem::exists(genFilesPath));
		}
	};
}
