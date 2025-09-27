

#include "pch.h"
#include "shared/com.h"
#include "../FelixPackage/FelixPackage.h"
#include "../FelixPackage/Z80Xml.h"
#include "Mocks.h"

#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>
namespace VxDTE
{
	#include <dte80.h>
	#include <dte90.h>
}

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	struct VSInstance
	{
		wil::unique_process_information processInfo;
		com_ptr<VxDTE::_DTE> dte;
	};

	static VSInstance _targetVS;

	TEST_CLASS(UITests)
	{
		static void StartOrRestart (const wchar_t* devenvExe, const wchar_t* devenvArguments, const wchar_t* testDataRoot, const wchar_t* tempRoot)
		{
			auto cmdLine = wil::str_concat_failfast<wil::unique_process_heap_string>(L"\"", devenvExe, L"\" ", devenvArguments);
			STARTUPINFO si = { .cb = sizeof(si) };
			wil::unique_process_information pi;
			BOOL bres = CreateProcessW (devenvExe, cmdLine.get(), nullptr, nullptr, TRUE, 0, nullptr, nullptr, &si, &pi);
			Assert::IsTrue(bres);
			Sleep(5000);
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
				Sleep(500);
			}

			Assert::IsNotNull(dte.get(), L"Failed to start VS");

			if (IsDebuggerPresent())
				FindAndAttach(pi.dwProcessId, dte);

			_targetVS = { std::move(pi), std::move(dte) };
		}

		static void CloseCurrentInstance(bool hard = false)
		{
			_targetVS.dte = nullptr;
			if (_targetVS.processInfo.hProcess)
			{
				if (hard) {
					TerminateProcess(_targetVS.processInfo.hProcess, 1234);
				} else {
					//if (!_vs.CloseMainWindow()) {
					//	TerminateProcess(_targetVS.processInfo.hProcess, 1234);
					//}
					//if (!_vs.WaitForExit(10000)) {
					TerminateProcess(_targetVS.processInfo.hProcess, 1234);
					//}
				}
				_targetVS.processInfo.reset();
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
						hr = targetProcess->Attach2(wil::make_variant_bstr_failfast(L"Managed/Native"));
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

			wil::unique_process_heap_string prefix;
			//hr = wil::GetModuleFileNameExW<wil::unique_process_heap_string>(_processInfo.hProcess, NULL, prefix);
			//Assert::IsTrue(SUCCEEDED(hr));
			//if (!_wcsicmp(PathFindFileName(prefix.get()), L"devenv.exe"))
				prefix = wil::make_process_heap_string_failfast(L"VisualStudio");
            
			//string progId = string.Format("!{0}.DTE.{1}:{2}", prefix, AssemblyVersionInfo.VSVersion, processId);
			auto progId = wil::str_printf_failfast<wil::unique_process_heap_string> (L"!%s.DTE.%s:%u", prefix, L"17.0", processId);
            
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
					rot->GetObject(moniker, &runningObject);
					break;
				}
            }

			Assert::IsNotNull(runningObject.get());
			hr = runningObject->QueryInterface(IID_PPV_ARGS(ppDTE)); RETURN_IF_FAILED_EXPECTED(hr);
			return S_OK;
        }

	public:
		TEST_CLASS_INITIALIZE(UITestsInitialize)
		{
			wil::unique_process_heap_string devenvExe;
			auto hr = wil::GetEnvironmentVariableW (L"VSAPPIDDIR", devenvExe);
			Assert::IsTrue(SUCCEEDED(hr));
			devenvExe = wil::str_concat_failfast<wil::unique_process_heap_string>(devenvExe, L"devenv.exe");
			StartOrRestart (devenvExe.get(), L"/rootSuffix Exp", L"", L"");
		}

		TEST_CLASS_CLEANUP(UITestsCleanup)
		{
			CloseCurrentInstance();
		}

		TEST_METHOD(LaunchVS)
		{
			auto name = wil::make_bstr_failfast(L"FelixProjects");
			com_ptr<IDispatch> obj;
			auto hr = _targetVS.dte->GetObject(name.get(), &obj);
			Assert::IsTrue(SUCCEEDED(hr));
		}
	};
}
