#include <CppUnitTest.h>
#include "Simulator.h"
#include "Utilities.h"
#include "shared/com.h"
#include <wil/resource.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace SimulatorTests
{
	TEST_CLASS(TransitionTests)
	{
	public:
		TEST_METHOD(BreakWhileAlreadyPaused)
		{
			HRESULT hr;
			auto simulator = MakeSimulator();

			TestSimulatorEventNotifySink sink (simulator);

			Assert::AreEqual(S_FALSE, simulator->Running_HR());
			Assert::AreEqual(S_FALSE, simulator->Break());
			Assert::AreEqual(S_FALSE, simulator->Running_HR());
			Assert::AreEqual<UINT32>(0, sink.breakpoints);
			Assert::AreEqual<UINT32>(0, sink.breaks);
			Assert::AreEqual<UINT32>(0, sink.simulateOnes);
			Assert::AreEqual<UINT32>(0, sink.resumes);
		}

		TEST_METHOD(BreakRightAfterBreakpoint)
		{
			HRESULT hr;
			auto simulator = MakeSimulator();

			TestSimulatorEventNotifySink sink (simulator);

			const UINT16 breakpointAddress = 0x8000;
			hr = simulator->Reset(breakpointAddress, SpectrumVariant48K);
			Assert::AreEqual(S_OK, hr);

			SIM_BP_COOKIE breakpointCookie;
			hr = simulator->AddBreakpoint(BreakpointType::Code, false, breakpointAddress, &breakpointCookie);
			Assert::AreEqual(S_OK, hr);
			auto removeBreakpoint = wil::scope_exit([&] { simulator->RemoveBreakpoint(breakpointCookie); });

			hr = simulator->Resume(TRUE);
			Assert::AreEqual(S_OK, hr);
			hr = WaitWithMessageLoop([&sink] { return sink.breakpoints != 0; }, 5000);
			Assert::AreEqual(S_OK, hr);

			Assert::AreEqual(S_FALSE, simulator->Break());
			Assert::AreEqual(S_FALSE, simulator->Running_HR());
			Assert::AreEqual<UINT32>(1, sink.breakpoints);
			Assert::AreEqual<UINT32>(0, sink.breaks);
		}

		TEST_METHOD(BreakWhileRunning)
		{
			HRESULT hr;
			auto simulator = MakeSimulator();

			TestSimulatorEventNotifySink sink (simulator);

			hr = simulator->Resume(TRUE);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_OK, simulator->Running_HR());
			Assert::AreEqual(S_OK, simulator->Break());
			Assert::AreEqual(S_FALSE, simulator->Running_HR());
			Assert::AreEqual<UINT32>(0, sink.breakpoints);
			Assert::AreEqual<UINT32>(1, sink.breaks);
		}

		TEST_METHOD(ResumeFromBreakEventHandler)
		{
			struct BreakEventSink : ISimulatorEventNotifySink
			{
				ULONG _refCount = 1;
				AdviseSinkToken _token;
				ISimulator* _simulator;

				BreakEventSink(ISimulator* simulator)
					: _simulator(simulator)
				{
					auto hr = AdviseSink<ISimulatorEventNotifySink>(simulator, this, &_token);
					Assert::AreEqual(S_OK, hr);
				}

				HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
				{
					if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<ISimulatorEventNotifySink>(this, riid, ppvObject))
						return S_OK;

					*ppvObject = nullptr;
					return E_NOINTERFACE;
				}

				ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
				ULONG STDMETHODCALLTYPE Release() override { return --_refCount; }

				HRESULT STDMETHODCALLTYPE NotifySimulatorEvent(ISimulatorEvent*, REFIID riidEvent) override
				{
					if (riidEvent == __uuidof(ISimulatorBreakEvent))
					{
						auto hr = _simulator->Resume(TRUE);
						Assert::AreEqual(E_UNEXPECTED, hr);
					}

					return S_OK;
				}
			};

			auto simulator = MakeSimulator();
			BreakEventSink sink(simulator);

			Assert::AreEqual(S_OK, simulator->Resume(TRUE));
			Assert::AreEqual(S_OK, simulator->Break());
			Assert::AreEqual(S_FALSE, simulator->Running_HR());
		}

		TEST_METHOD(SimulateOneFromBreakEventHandler)
		{
			struct BreakEventSink : ISimulatorEventNotifySink
			{
				ULONG _refCount = 1;
				AdviseSinkToken _token;
				ISimulator* _simulator;
				HRESULT simulateOneResult = E_FAIL;

				BreakEventSink(ISimulator* simulator)
					: _simulator(simulator)
				{
					auto hr = AdviseSink<ISimulatorEventNotifySink>(simulator, this, &_token);
					Assert::AreEqual(S_OK, hr);
				}

				HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
				{
					if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<ISimulatorEventNotifySink>(this, riid, ppvObject))
						return S_OK;

					*ppvObject = nullptr;
					return E_NOINTERFACE;
				}

				ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
				ULONG STDMETHODCALLTYPE Release() override { return --_refCount; }

				HRESULT STDMETHODCALLTYPE NotifySimulatorEvent(ISimulatorEvent*, REFIID riidEvent) override
				{
					if (riidEvent == __uuidof(ISimulatorBreakEvent))
						simulateOneResult = _simulator->SimulateOne();

					return S_OK;
				}
			};

			auto simulator = MakeSimulator();
			BreakEventSink sink(simulator);

			Assert::AreEqual(S_OK, simulator->Resume(TRUE));
			Assert::AreEqual(S_OK, simulator->Break());
			Assert::AreEqual(E_UNEXPECTED, sink.simulateOneResult);
		}

		TEST_METHOD(SimulateOneWhileHalted)
		{
			auto simulator = MakeSimulator();
			const UINT16 haltAddress = 0x8000;
			const UINT8 haltOpcode = 0x76;

			Assert::AreEqual(S_OK, simulator->Reset(haltAddress, SpectrumVariant48K));
			Assert::AreEqual(S_OK, simulator->WriteMemoryBus(haltAddress, sizeof(haltOpcode), &haltOpcode));
			Assert::AreEqual(S_OK, simulator->SimulateOne());
			simulator->SimulateOne();
		}

		TEST_METHOD(ResetAfterBreakingWhileTapIsPlayingDoesNotBreakAgain)
		{
			TempFile tapFile(L".tap");
			tapFile.Write({ 3, 0, 0, 0, 0 });

			auto simulator = MakeSimulator();
			TestSimulatorEventNotifySink sink(simulator);
			HRESULT hr;

			hr = simulator->BeginPlayTapFile(tapFile.Path(), TRUE);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_OK, simulator->IsTapFileLoading());

			hr = simulator->Resume(FALSE);
			Assert::AreEqual(S_OK, hr);
			hr = simulator->Break();
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_FALSE, simulator->Running_HR());
			Assert::AreEqual(S_OK, simulator->IsTapFileLoading());
			Assert::AreEqual<UINT32>(1, sink.breaks);

			hr = simulator->Reset(0, SpectrumVariant48K);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_FALSE, simulator->IsTapFileLoading());
			Assert::AreEqual<UINT32>(1, sink.breaks);
		}
	};
}
