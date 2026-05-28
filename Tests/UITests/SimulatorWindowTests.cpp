
#include "pch.h"
#include "shared/com.h"
#include "UITests.h"
#include "dispids.h"

namespace UITests
{
	TEST_CLASS(SimulatorWindowTests)
	{
		static inline wil::com_ptr_failfast<VxDTE::DTE2> dte;
		static inline wil::com_ptr_failfast<VxDTE::Debugger> debugger;
		static inline wil::unique_process_heap_string testClassPath;

		TEST_CLASS_INITIALIZE(ClassInit)
		{
			HRESULT hr;

			dte = GetDefaultVSInstance();
			hr = dte->get_Debugger(&debugger);
			Assert::AreEqual(S_OK, hr);

			testClassPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"SimulatorWindowTests\\");
			Assert::IsTrue(CreateDirectory(testClassPath.get(), nullptr));
		}

		TEST_CLASS_CLEANUP(ClassCleanup)
		{
			debugger.reset();
			dte.reset();
		}

		static void Bas2Tap (const char* basicText, const wchar_t* tapPath)
		{
			auto basPath = str_concat(testClassPath, L"test.bas");
			WriteFileCreateDirs (basPath, basicText);

			DeleteFile(tapPath);
			auto cmdLine = wil::str_printf_failfast<wil::unique_process_heap_string>(L"bas2tap -a10 %s %s", basPath, tapPath);
			STARTUPINFO si = { .cb = sizeof(si) };
			wil::unique_process_information pi;
			BOOL bres = CreateProcessW (NULL, cmdLine.get(), 0, 0, 0, 0, 0, 0, &si, &pi);
			Assert::IsTrue(bres);
			WaitForSingleObject(pi.hProcess, INFINITE);
			DWORD exitCode;
			bres = GetExitCodeProcess(pi.hProcess, &exitCode);
			Assert::IsTrue(bres);
			Assert::AreEqual(0ul, exitCode);
		}

		TEST_METHOD(SimulatorWindow_TapLoad)
		{
			HRESULT hr;

			static const char basic[] = "10 POKE 32768, 85\r\n20 POKE 32769,170";
			auto tapPath = str_concat(testClassPath, L"test.tap");
			Bas2Tap (basic, tapPath.get());

			wil::com_ptr_failfast<IDispatch> simWindowDisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"SimulatorWindow").get(), &simWindowDisp);
			Assert::AreEqual(S_OK, hr);

			auto simWindow = simWindowDisp.query<ISimulatorWindowAutomationObject>();
			hr = simWindow->OpenTapFile (wil::make_bstr_failfast(tapPath.get()).get(), 10'000);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<ISimulator_> simulator;
			hr = simWindow->GetSimulator(&simulator);
			Assert::AreEqual(S_OK, hr);

			simulator->Break();

			UINT8 data;
			hr = simulator->ReadMemoryBus8(32768, &data);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<UINT8>(85, data);

			hr = simulator->ReadMemoryBus8(32769, &data);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<UINT8>(170, data);
		}

		TEST_METHOD(SimulatorWindow_TapDebug)
		{
		}
	};
}

