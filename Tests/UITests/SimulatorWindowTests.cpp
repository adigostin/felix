
#include "pch.h"
#include "shared/com.h"
#include "UITests.h"

namespace UITests
{
	TEST_CLASS(SimulatorWindowTests)
	{
		static inline wil::com_ptr_failfast<VxDTE::DTE2> dte;
		static inline wil::com_ptr_failfast<VxDTE::Debugger> debugger;
		static inline wil::unique_process_heap_string testClassPath;
		static inline wil::com_ptr_failfast<ISimulatorWindowAutomationObject> simWindow;
		static inline wil::com_ptr_failfast<ISimulator_> simulator;

		TEST_CLASS_INITIALIZE(ClassInit)
		{
			HRESULT hr;

			dte = GetDefaultVSInstance();
			hr = dte->get_Debugger(&debugger);
			Assert::AreEqual(S_OK, hr);

			testClassPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"SimulatorWindowTests\\");
			Assert::IsTrue(CreateDirectory(testClassPath.get(), nullptr));

			wil::com_ptr_failfast<IDispatch> simWindowDisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"SimulatorWindow").get(), &simWindowDisp);
			Assert::AreEqual(S_OK, hr);

			simWindow = simWindowDisp.query<ISimulatorWindowAutomationObject>();

			hr = simWindow->GetSimulator(&simulator);
			Assert::AreEqual(S_OK, hr);
		}

		TEST_CLASS_CLEANUP(ClassCleanup)
		{
			simulator.reset();
			simWindow.reset();
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

		static UINT8 ReadMemoryBus (UINT16 addr)
		{
			UINT8 data;
			auto hr = simulator->ReadMemoryBus8(addr, &data);
			Assert::AreEqual(S_OK, hr);
			return data;
		}

		void WaitDebugMode (VxDTE::dbgDebugMode mode, DWORD timeoutMilliseconds = 5000)
		{
			DWORD tickStart = GetTickCount();
			while(true)
			{
				VxDTE::dbgDebugMode current;
				auto hr = debugger->get_CurrentMode(&current);
				if (SUCCEEDED(hr))
				{
					if (current == mode)
						break;
				}
				else
					Assert::AreEqual(RPC_E_CALL_REJECTED, hr);
				if (GetTickCount() - tickStart >= timeoutMilliseconds)
					Assert::Fail();
				Sleep(20);
			}
		}

		TEST_METHOD(SimulatorWindow_TapLoad)
		{
			HRESULT hr;

			static const char basic[] = "10 POKE 32768, 85\r\n20 POKE 32769,170";
			auto tapPath = str_concat(testClassPath, L"test.tap");
			Bas2Tap (basic, tapPath.get());

			hr = simWindow->OpenTapFile (wil::make_bstr_failfast(tapPath.get()).get(), TRUE);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_OK, simulator->IsTapFileLoading());

			DWORD tickStart = GetTickCount();
			while(simulator->IsTapFileLoading() == S_OK)
			{
				if (GetTickCount() - tickStart >= 5000)
					Assert::Fail();
				Sleep(20);
			}

			simulator->Break();

			Assert::AreEqual<UINT8>(85, ReadMemoryBus(32768));
			Assert::AreEqual<UINT8>(170, ReadMemoryBus(32769));
		}

		TEST_METHOD(SimulatorWindow_TapLoadWhileTapLoading)
		{
			HRESULT hr;

			static const char basic[] = "10 POKE 32768, 85\r\n20 POKE 32769,170";
			auto tapPath = str_concat(testClassPath, L"test.tap");
			Bas2Tap (basic, tapPath.get());

			hr = simWindow->OpenTapFile (wil::make_bstr_failfast(tapPath.get()).get(), FALSE);
			Assert::AreEqual(S_OK, hr);
			
			Assert::AreEqual(S_OK, simulator->IsTapFileLoading());

			Sleep(500);

			hr = simWindow->OpenTapFile (wil::make_bstr_failfast(tapPath.get()).get(), TRUE);
			Assert::AreEqual(S_OK, hr);

			// Second time we called OpenTapFile, we called with maxSpeed=TRUE.
			// It should take much less than a second to load.
			DWORD tickStart = GetTickCount();
			while(simulator->IsTapFileLoading() == S_OK)
			{
				if (!IsDebuggerPresent() && GetTickCount() - tickStart >= 1000)
					Assert::Fail();
				Sleep(20);
			}


			simulator->Break();

			Assert::AreEqual<UINT8>(85, ReadMemoryBus(32768));
			Assert::AreEqual<UINT8>(170, ReadMemoryBus(32769));
		}

		struct TapPlayNotifySink : ITapPlayNotifySink
		{
			ULONG _refCount = 0;
			ULONG startingCalled = 0;
			ULONG completeCalled = 0;

			#pragma region IUnknown
			virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
			{
				if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<ITapPlayNotifySink>(this, riid, ppvObject))
					return S_OK;

				*ppvObject = nullptr;
				return E_NOINTERFACE;
			}

			virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

			virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
			#pragma endregion

			#pragma region ITapPlayNotifySink
			virtual HRESULT STDMETHODCALLTYPE NotifyTapPlayStarting() override
			{
				startingCalled++;
				return S_OK;
			}

			virtual HRESULT STDMETHODCALLTYPE NotifyTapPlayComplete() override
			{
				completeCalled++;
				return S_OK;
			}
			#pragma endregion
		};

		TEST_METHOD(Simulator_TapPlaySinkNotCalledAfterUnadvise)
		{
			HRESULT hr;

			static const char basic[] = "10 POKE 32768, 85\r\n20 POKE 32769,170";
			auto tapPath = str_concat(testClassPath, L"test.tap");
			Bas2Tap (basic, tapPath.get());

			wil::com_ptr_failfast<IConnectionPoint> cp;
			hr = simulator.query<IConnectionPointContainer>()->FindConnectionPoint(__uuidof(ITapPlayNotifySink), &cp);
			Assert::AreEqual(S_OK, hr);

			auto sink = wil::com_ptr_failfast(new TapPlayNotifySink());
			DWORD cookie;
			hr = cp->Advise(sink, &cookie);
			Assert::AreEqual(S_OK, hr);

			hr = simulator->BeginPlayTapFile(wil::make_bstr_failfast(tapPath.get()).get(), FALSE);
			Assert::AreEqual(S_OK, hr);
			
			Assert::AreEqual(1ul, sink->startingCalled);
			Assert::AreEqual(0ul, sink->completeCalled);

			hr = cp->Unadvise(cookie);
			Assert::AreEqual(S_OK, hr);

			Assert::AreEqual(1ul, sink->startingCalled);
			Assert::AreEqual(0ul, sink->completeCalled);

			hr = simulator->CancelPlayTapFile();
			Assert::AreEqual(S_OK, hr);

			Assert::AreEqual(1ul, sink->startingCalled);
			Assert::AreEqual(0ul, sink->completeCalled);
		}

		TEST_METHOD(SimulatorWindow_TapDebug)
		{
			HRESULT hr;

			static const char basic[] = "10 POKE 32768, 85\r\n20 POKE 32769,170";
			auto tapPath = str_concat(testClassPath, L"test.tap");
			Bas2Tap (basic, tapPath.get());

			wil::com_ptr_failfast<VxDTE::Debugger> debugger;
			hr = dte->get_Debugger(&debugger);
			Assert::AreEqual(S_OK, hr);

			VxDTE::dbgDebugMode debugMode;
			hr = debugger->get_CurrentMode(&debugMode);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<UINT>(VxDTE::dbgDesignMode, debugMode);

			hr = simWindow->DebugTapFile (wil::make_bstr_failfast(tapPath.get()).get(), TRUE);
			Assert::AreEqual(S_OK, hr);

			Assert::AreEqual(S_OK, simulator->IsTapFileLoading());

			hr = debugger->get_CurrentMode(&debugMode);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<UINT>(VxDTE::dbgRunMode, debugMode);

			WaitDebugMode(VxDTE::dbgBreakMode);

			Assert::AreEqual<UINT8>(0, ReadMemoryBus(32768));
			Assert::AreEqual<UINT8>(0, ReadMemoryBus(32769));

			dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			// Will go to dbgRunMode and then quickly to dbgDesignMode.

			WaitDebugMode(VxDTE::dbgDesignMode);

			simulator->Break();
			Assert::AreEqual<UINT8>(85, ReadMemoryBus(32768));
			Assert::AreEqual<UINT8>(170, ReadMemoryBus(32769));
		}
	};
}

