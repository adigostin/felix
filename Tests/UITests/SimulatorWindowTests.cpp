
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
			hr = simWindow->OpenTapFile (wil::make_bstr_failfast(tapPath.get()).get(), TRUE);
			Assert::AreEqual(S_OK, hr);
			hr = simWindow->IsTapFileLoading();
			Assert::AreEqual(S_OK, hr);

			DWORD tickStart = GetTickCount();
			while(simWindow->IsTapFileLoading() == S_OK)
			{
				if (GetTickCount() - tickStart >= 5000)
					Assert::Fail();
				Sleep(20);
			}

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

		TEST_METHOD(SimulatorWindow_TapLoadWhileTapLoading)
		{
			HRESULT hr;

			static const char basic[] = "10 POKE 32768, 85\r\n20 POKE 32769,170";
			auto tapPath = str_concat(testClassPath, L"test.tap");
			Bas2Tap (basic, tapPath.get());

			wil::com_ptr_failfast<IDispatch> simWindowDisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"SimulatorWindow").get(), &simWindowDisp);
			Assert::AreEqual(S_OK, hr);

			auto simWindow = simWindowDisp.query<ISimulatorWindowAutomationObject>();
			hr = simWindow->OpenTapFile (wil::make_bstr_failfast(tapPath.get()).get(), FALSE);
			Assert::AreEqual(S_OK, hr);
			
			hr = simWindow->IsTapFileLoading();
			Assert::AreEqual(S_OK, hr);

			Sleep(500);

			hr = simWindow->OpenTapFile (wil::make_bstr_failfast(tapPath.get()).get(), TRUE);
			Assert::AreEqual(S_OK, hr);

			// Second time we called OpenTapFile, we called with maxSpeed=TRUE.
			// It should take much less than a second to load.
			DWORD tickStart = GetTickCount();
			while(simWindow->IsTapFileLoading() == S_OK)
			{
				if (!IsDebuggerPresent() && GetTickCount() - tickStart >= 1000)
					Assert::Fail();
				Sleep(20);
			}


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

			wil::com_ptr_failfast<IDispatch> simWindowDisp;
			hr = dte->GetObject(wil::make_bstr_failfast(L"SimulatorWindow").get(), &simWindowDisp);
			Assert::AreEqual(S_OK, hr);
			auto simWindow = simWindowDisp.query<ISimulatorWindowAutomationObject>();

			wil::com_ptr_failfast<ISimulator_> simulator;
			hr = simWindow->GetSimulator(&simulator);
			Assert::AreEqual(S_OK, hr);

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
		}
	};
}

