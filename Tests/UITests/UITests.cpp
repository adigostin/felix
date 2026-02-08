
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

	#pragma region DTE Initialization
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
	#pragma endregion

	static void __stdcall WilLoggingCallback (const wil::FailureInfo& fi) noexcept
	{
		Assert::Fail(fi.pszMessage);
	}

	TEST_MODULE_INITIALIZE(UITestsInitialize)
	{
		wil::SetResultLoggingCallback (WilLoggingCallback);

		MakeTemplates (L"FelixTestUI");

		auto hr = CoCreateInstance (__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation));
		Assert::IsTrue(SUCCEEDED(hr));
		StartOrRestart();
	}

	TEST_MODULE_CLEANUP(UITestsCleanup)
	{
		CloseCurrentInstance();
		automation.reset();

		wil::SetResultLoggingCallback(nullptr);
	}

	// <param name="testDir">The directory in which to create the solution.</param>
	// <param name="solutionName">The name of the solution to create, without extension.</param>
	// <param name="projectName">NULL to create a project with the same name as the solution in the same dir, filename without extension to create project in subdir with different name.</param>
	std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
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

	void BuildSolution (VxDTE::_Solution* sln, long* buildFailCount)
	{
		HRESULT hr;

		com_ptr<VxDTE::SolutionBuild> solutionBuild;
		hr = sln->get_SolutionBuild(&solutionBuild);
		Assert::IsTrue(SUCCEEDED(hr));

		Assert::IsTrue(SUCCEEDED(hr));
		hr = solutionBuild->Build(VARIANT_TRUE);
		Assert::IsTrue(SUCCEEDED(hr));

		// LastBuildInfo returns the number of failed projects, despite the parameter name.
		hr = solutionBuild->get_LastBuildInfo(buildFailCount);
		Assert::IsTrue(SUCCEEDED(hr));
	}
}
