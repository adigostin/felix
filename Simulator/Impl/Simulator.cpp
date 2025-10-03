
#include "pch.h"
#include "SimulatorInternal.h"
#include "Z80CPU.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/com.h"
#include "shared/inplace_function.h"
#include <optional>

#pragma comment (lib, "Shlwapi")

static constexpr UINT WM_MAIN_THREAD_WORK = WM_APP + 0;
static constexpr UINT WM_SCREEN_COMPLETE  = WM_APP + 1;

static ATOM wndClassAtom;

HRESULT STDMETHODCALLTYPE MakeHC91ROM (Bus* memory_bus, Bus* io_bus, const wchar_t* folder, const wchar_t* BinaryFilename, wistd::unique_ptr<IMemoryDevice>* ppDevice);
HRESULT STDMETHODCALLTYPE MakeHC91RAM (Bus* memory_bus, Bus* io_bus, wistd::unique_ptr<IMemoryDevice>* ppDevice);
HRESULT STDMETHODCALLTYPE MakeBeeper (Bus* io_bus, wistd::unique_ptr<IDevice>* ppDevice);

using unique_cotaskmem_bitmapinfo = wil::unique_any<BITMAPINFO*, decltype(&::CoTaskMemFree), ::CoTaskMemFree>;

// ============================================================================

class SimulatorImpl : public ISimulator, IScreenDeviceCompleteEventHandler, IConnectionPointContainer
{
	ULONG _refCount = 0;

	static constexpr wchar_t WndClassName[] = L"Z80Program-{5F067572-7D58-40A1-A471-7FE701460B3B}";

	HWND _hwnd = nullptr;

	LARGE_INTEGER qpFrequency;

	struct running_info
	{
		UINT64 start_time; // tick count of the simulated clock
		LARGE_INTEGER start_time_perf_counter;
	};

	std::optional<running_info> _running_info; // this is used only by the simulator thread
	bool _running = false; // and this only by the main thread
	wil::unique_handle _cpuThread;
	wil::unique_handle _cpu_thread_exit_request;
	com_ptr<ConnectionPointImpl<ISimulatorEventNotifySink>> _eventHandlers;
	com_ptr<IScreenCompleteEventHandler> _screenCompleteHandler;

	using RunOnSimulatorThreadFunction = HRESULT(*)(void*);
	stdext::inplace_function<HRESULT(), 64> _runOnSimulatorThreadFunction;
	HRESULT            _runOnSimulatorThreadResult;
	wil::unique_handle _run_on_simulator_thread_request;
	wil::unique_handle _runOnSimulatorThreadComplete;
	wil::unique_handle _waitableTimer;

	vector_nothrow<stdext::inplace_function<void()>> _mainThreadWorkQueue;
	wil::srwlock _mainThreadQueueLock;

	Bus memoryBus;
	Bus ioBus;
	irq_line_i irq;
	wistd::unique_ptr<IZ80CPU> _cpu;	
	wistd::unique_ptr<IScreenDevice> _screen;
	wistd::unique_ptr<IKeyboardDevice> _keyboard;
	wistd::unique_ptr<IMemoryDevice> _romDevice;
	wistd::unique_ptr<IMemoryDevice> _ramDevice;
	wistd::unique_ptr<IDevice> _beeper;
	vector_nothrow<IDevice*> _active_devices_;
	bool _showCRTSnapshot = false;

	// Passed via WM_SCREEN_COMPLETE from simulator thread to GUI thread while simulation is running.
	unique_cotaskmem_bitmapinfo _screenComplete;

public:
	HRESULT InitInstance (LPCWSTR dir, LPCWSTR romFilename)
	{
		HRESULT hr;
		
		hr = ConnectionPointImpl<ISimulatorEventNotifySink>::CreateInstance(this, &_eventHandlers); RETURN_IF_FAILED(hr);

		hr = MakeZ80CPU(&memoryBus, &ioBus, &irq, &_cpu); RETURN_IF_FAILED(hr);

		hr = MakeScreenDevice(&memoryBus, &ioBus, &irq, this, &_screen); RETURN_IF_FAILED(hr);

		hr = MakeKeyboardDevice(&ioBus, &_keyboard); RETURN_IF_FAILED(hr);
		
		hr = MakeBeeper(&ioBus, &_beeper); RETURN_IF_FAILED(hr);

		hr = MakeHC91ROM (&memoryBus, &ioBus, dir, romFilename, &_romDevice); RETURN_IF_FAILED(hr);
///		hr = _romDevice->AdviseBusAddressRangeChange(this); RETURN_IF_FAILED(hr);

		hr = MakeHC91RAM (&memoryBus, &ioBus, &_ramDevice); RETURN_IF_FAILED(hr);

		bool pushed = _active_devices_.try_push_back({ _screen.get(), _keyboard.get(), _romDevice.get(), _ramDevice.get(), _beeper.get() }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

		QueryPerformanceFrequency(&qpFrequency);

		if (!wndClassAtom)
		{
			WNDCLASS wc = { };
			wc.lpfnWndProc = window_proc;
			wc.hInstance = (HINSTANCE)&__ImageBase;
			wc.lpszClassName = WndClassName;
			wndClassAtom = RegisterClass(&wc); RETURN_LAST_ERROR_IF(!wndClassAtom);
		}

		_hwnd = CreateWindowExW (0, WndClassName, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, (HINSTANCE)&__ImageBase, nullptr); RETURN_LAST_ERROR_IF_NULL(_hwnd);
		SetWindowLongPtr (_hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));

		_run_on_simulator_thread_request.reset(CreateEventW (nullptr, false, false, nullptr)); RETURN_LAST_ERROR_IF_NULL(_run_on_simulator_thread_request);
		_runOnSimulatorThreadComplete.reset(CreateEventW (nullptr, false, false, nullptr)); RETURN_LAST_ERROR_IF_NULL(_runOnSimulatorThreadComplete);

		_waitableTimer.reset(CreateWaitableTimerExW (nullptr, nullptr, 0, SYNCHRONIZE | TIMER_MODIFY_STATE)); RETURN_LAST_ERROR_IF_NULL(_waitableTimer);

		_cpu_thread_exit_request.reset (CreateEvent(nullptr, FALSE, FALSE, nullptr)); RETURN_LAST_ERROR_IF_NULL(_cpu_thread_exit_request);
		_cpuThread.reset (CreateThread(nullptr, 100'000, simulation_thread_proc_static, this, 0, nullptr)); RETURN_LAST_ERROR_IF_NULL(_cpuThread);

		//SetThreadDescription(_cpuThread.get(), L"Simulator");

		return S_OK;
	};

	~SimulatorImpl()
	{
		WI_ASSERT (_eventHandlers->empty());
		WI_ASSERT (!_screenCompleteHandler);

		if (_cpu)
			WI_ASSERT(!_cpu->HasBreakpoints());

		if (_cpuThread)
		{
			::SetEvent(_cpu_thread_exit_request.get());
			WaitForSingleObject(_cpuThread.get(), IsDebuggerPresent() ? INFINITE : 5000);
			_cpuThread.reset();
		}

		if (_hwnd)
		{
			BOOL bres = ::DestroyWindow(_hwnd); LOG_LAST_ERROR_IF(!bres);
			_hwnd = nullptr;
			bres = UnregisterClass (WndClassName, (HINSTANCE)&__ImageBase); LOG_LAST_ERROR_IF(!bres);
		}
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);

		if (   TryQI<IUnknown>(static_cast<ISimulator*>(this), riid, ppvObject)
			|| TryQI<ISimulator>(this, riid, ppvObject)
			|| TryQI<IConnectionPointContainer>(this, riid, ppvObject)
		)
			return S_OK;

		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	static LRESULT CALLBACK window_proc (HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
	{
		if (msg == WM_MAIN_THREAD_WORK)
		{
			auto p = reinterpret_cast<SimulatorImpl*>(GetWindowLongPtr (hwnd, GWLP_USERDATA));
			WI_ASSERT(p);

			auto lock = p->_mainThreadQueueLock.lock_exclusive();
			auto work = p->_mainThreadWorkQueue.remove(p->_mainThreadWorkQueue.begin());
			lock.reset();
			work();
		}
		else if (msg == WM_SCREEN_COMPLETE)
		{
			auto p = reinterpret_cast<SimulatorImpl*>(GetWindowLongPtr (hwnd, GWLP_USERDATA));
			WI_ASSERT(p);

			auto lock = p->_mainThreadQueueLock.lock_exclusive();
			if (p->_screenComplete)
			{
				auto screen = std::move(p->_screenComplete);
				if (p->_screenCompleteHandler)
				{
					auto hr = p->_screenCompleteHandler->OnScreenComplete(screen.get(), { -1, -1 });
					if (SUCCEEDED(hr))
						screen.release();
				}
			}
		}

		return DefWindowProc (hwnd, msg, wparam, lparam);
	}

	static DWORD CALLBACK simulation_thread_proc_static (void* arg)
	{
		return static_cast<SimulatorImpl*>(arg)->simulation_thread_proc();
	}

	UINT64 real_time() const
	{
		//WI_ASSERT(::GetCurrentThreadId() == GetThreadId(_cpuThread.get()));
		//WI_ASSERT(_running_info);
		
		LARGE_INTEGER timeNow;
		QueryPerformanceCounter(&timeNow);
		
		auto perf_counter_delta = timeNow.QuadPart - _running_info->start_time_perf_counter.QuadPart;

		auto tick_count = _running_info->start_time + perf_counter_delta * 3500 / (qpFrequency.QuadPart / 1000);
		return tick_count;
	}

	bool simulate_devices_to (UINT64 time)
	{
		bool res = false;
		while(true)
		{
			bool advanced = false;
			for (auto& d : _active_devices_)
			{
				if (d->Time() < time)
				{
					uint64_t timeBefore = d->Time();
					d->SimulateTo(time);
					advanced |= (d->Time() > timeBefore);
				}
			}

			if (!advanced)
				return res;

			res = true;
		}
	}

	DWORD simulation_thread_proc()
	{
		HANDLE waitHandles[3] = { _run_on_simulator_thread_request.get(), _cpu_thread_exit_request.get(), _waitableTimer.get() };
		bool exit_request = false;
		while(!exit_request)
		{
			// TODO: add support for syncing on multiple devices, needed in case two devices need to sync on
			// the same time _and_ the first is waiting for the second (we'd have an infinite loop with the code as it is now).

			DWORD waitHandleCount = 2;
			DWORD waitTimeout = INFINITE;
			IDevice* device_to_sync_on = nullptr;
			UINT64 time_to_sync_to_ = 0;

			if (_running_info)
			{
				UINT64 rt = real_time();

				#pragma region Catch up with real time
				// Before we attempt simulation, let's check if some devices are lagging far behind the real time.
				// This happens while debugging the VSIX, or it may happen when this thread is starved. 
				// If we have any such device, we "erase" the same length of time from all of the devices;
				// we do this by rebasing the simulation startup time held in the _running_info variable.
				{
					auto slowestTime = _cpu->Time();
					INT64 slowest_offset_from_rt = (UINT64)(_cpu->Time() - rt);
					for (auto& d : _active_devices_)
					{
						INT64 offset_from_rt = (UINT64)(d->Time() - rt);
						if (offset_from_rt < slowest_offset_from_rt)
						{
							slowestTime = d->Time();
							slowest_offset_from_rt = offset_from_rt;
						}
					}

					if (slowest_offset_from_rt < -(INT64)milliseconds_to_ticks(50))
					{
						// Slowest device is more than 50 ms behind real time. Let's see what time offset
						// it would need to be 50 ms _ahead_ of real time, and add that offset to all devices.
						UINT64 tick_offset = rt + milliseconds_to_ticks(50) - slowestTime;
						UINT64 perf_counter_offset = tick_offset * (qpFrequency.QuadPart / 1000) / 3500;
						_running_info->start_time += tick_offset;
						_running_info->start_time_perf_counter.QuadPart += perf_counter_offset;
					}
				}
				#pragma endregion

				// Let's find out how far we can simulate, and try to simulate to that point in time.
				for (auto& d : _active_devices_)
				{
					UINT64 t;
					if (d->NeedSyncWithRealTime(&t)
						&& (!device_to_sync_on || (t < time_to_sync_to_)))
					{
						device_to_sync_on = d;
						time_to_sync_to_ = t;
					}
				}

				while(true)
				{
					// First ask the CPU to simulate itself; then ask the devices
					// to simulate themselves until they catch up with the CPU, not more.
					// We don't want the devices to go far ahead of the CPU; if the CPU will stop at a breakpoint,
					// we'll want to show to the user the devices at a time as close as possible to the CPU time;
					// we can do that only if the devices are at all times behind the CPU or only slightly
					// (a few clock cycles) ahead of it.
					BreakpointsHit bpsHit = { };
					while (_cpu->Time() < time_to_sync_to_)
					{
						bool advanced = _cpu->SimulateOne(&bpsHit);
						if (!advanced || bpsHit.size)
							break;
					}

					// Whatever the outcome of the above simulation, we must first bring the devices close to the CPU time.
					simulate_devices_to(_cpu->Time());

					if (bpsHit.size)
					{
						on_bp_hit(&bpsHit);
						break;
					}

					if (_cpu->Time() >= time_to_sync_to_)
					{
						// Now let's see how long we need to wait for the real time to catch up.
						WI_ASSERT(device_to_sync_on);
						if (time_to_sync_to_ > rt)
						{
							uint64_t hundredsOfNanoseconds = ticks_to_hundreds_of_nanoseconds(time_to_sync_to_ - rt);
							LARGE_INTEGER dueTime = { .QuadPart = -(INT64)hundredsOfNanoseconds };
							BOOL bRes = SetWaitableTimer (_waitableTimer.get(), &dueTime, 0, nullptr, nullptr, FALSE); WI_ASSERT(bRes);
							waitHandleCount = 3;
						}
						else
							waitTimeout = 0; // We're lagging behind real time, so we won't wait.
						break;
					}
				}
			}

			DWORD waitResult = WaitForMultipleObjects (waitHandleCount, waitHandles, FALSE, waitTimeout);
			switch (waitResult)
			{
				case WAIT_TIMEOUT:
				case WAIT_OBJECT_0 + 2: // _waitableTimer
					// Real time has caught up with the simulated time.
					// Let's unblock the device that was waiting for this time point.
					WI_ASSERT(device_to_sync_on);
					if (device_to_sync_on->Time() < time_to_sync_to_ + 1)
					{
						uint64_t timeBefore = device_to_sync_on->Time();
						device_to_sync_on->SimulateTo(time_to_sync_to_ + 1);
						WI_ASSERT(device_to_sync_on->Time() > timeBefore);
					}
					break;

				case WAIT_OBJECT_0: // _run_on_simulator_thread_request
					_runOnSimulatorThreadResult = _runOnSimulatorThreadFunction();
					SetEvent(_runOnSimulatorThreadComplete.get());
					break;

				case WAIT_OBJECT_0 + 1: // _cpu_thread_exit_request
					exit_request = true;
					break;

				default:
					WI_ASSERT(false); // TODO: handle error conditions
			}
		}

		return 0;
	}

	class BreakpointEvent : public ISimulatorBreakpointEvent
	{
		ULONG _refCount = 0;
		BreakpointType _type;
		uint16_t _address;
		vector_nothrow<SIM_BP_COOKIE> _bps;

	public:
		HRESULT InitInstance (BreakpointType type, uint16_t address, SIM_BP_COOKIE const* bps, uint32_t size)
		{
			_type = type;
			_address = address;
			bool r = _bps.try_reserve(size); RETURN_HR_IF(E_OUTOFMEMORY, !r);
			for (uint32_t i = 0; i < size; i++)
				_bps.try_push_back(bps[i]);
			return S_OK;
		}

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(this, riid, ppvObject)
				|| TryQI<ISimulatorEvent>(this, riid, ppvObject)
				|| TryQI<ISimulatorBreakpointEvent>(this, riid, ppvObject)
				)
				return S_OK;

			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IEnumBreakpoints
		virtual BreakpointType GetType() override { return _type; }

		virtual UINT16 GetAddress() override { return _address; }

		virtual ULONG GetBreakpointCount() override { return _bps.size(); }

		virtual HRESULT GetBreakpointAt(ULONG i, SIM_BP_COOKIE* ppKey) override
		{
			*ppKey = _bps[i];
			return S_OK;
		}
		#pragma endregion
	};

	// Called when simulation was running and the CPU reached a breakpoint.
	HRESULT on_bp_hit (const BreakpointsHit* bps)
	{
		for (uint32_t i = 0; i < bps->size; i++)
			WI_ASSERT(bps->bps[i] != 0);

		auto bpsCopy = wil::make_unique_hlocal_nothrow<BreakpointsHit>(*bps); RETURN_IF_NULL_ALLOC_EXPECTED(bpsCopy);

		WI_ASSERT(_running_info);
		_running_info.reset();

		unique_cotaskmem_bitmapinfo screen;
		POINT beam;
		_screen->CopyBuffer (_showCRTSnapshot, screen.addressof(), &beam);

		auto work = [this, bps=std::move(bpsCopy), screen=std::move(screen), beam]() mutable
			{
				WI_ASSERT(_running);
				_running = false;
				if (auto bpEvent = com_ptr(new (std::nothrow) BreakpointEvent());
					bpEvent && SUCCEEDED(bpEvent->InitInstance(BreakpointType::Code, bps->address, bps->bps, bps->size)))
				{
					_eventHandlers->Notify([bpEvent=bpEvent.get()](ISimulatorEventNotifySink* sink)
						{ sink->NotifySimulatorEvent(bpEvent, __uuidof(ISimulatorBreakpointEvent)); });
				}

				_screenComplete = nullptr;
				if (_screenCompleteHandler && screen)
				{
					auto hr = _screenCompleteHandler->OnScreenComplete(screen.get(), beam);
					if (SUCCEEDED(hr))
						screen.release();
				}
			};
		auto lock = _mainThreadQueueLock.lock_exclusive();
		bool pushed = _mainThreadWorkQueue.try_push_back(std::move(work));
		if (pushed)
		{
			BOOL posted = PostMessageW (_hwnd, WM_MAIN_THREAD_WORK, 0, 0);
			if (!posted)
			{
				DWORD le = GetLastError();
				_mainThreadWorkQueue.remove(_mainThreadWorkQueue.end() - 1);
				return HRESULT_FROM_WIN32(le);
			}
		}

		return S_OK;
	}

	template<typename IEvent>
	class SimulatorEvent : public IEvent
	{
		ULONG _refCount = 0;

	public:
		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			RETURN_HR_IF(E_POINTER, !ppvObject);

			if (   TryQI<IUnknown>(this, riid, ppvObject)
				|| TryQI<ISimulatorEvent>(this, riid, ppvObject)
				|| TryQI<IEvent>(this, riid, ppvObject))
				return S_OK;

			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion
	};

	HRESULT SendSimulateOneCompleteEvent()
	{
		WI_ASSERT(!_running);

		using SimulateOneEvent = SimulatorEvent<ISimulatorSimulateOneEvent>;
		auto event = com_ptr(new (std::nothrow) SimulateOneEvent()); RETURN_IF_NULL_ALLOC(event);

		_eventHandlers->Notify([&event](ISimulatorEventNotifySink* sink)
			{ sink->NotifySimulatorEvent(event, __uuidof(event)); });

		return S_OK;
	}

	#pragma region ISimulator
	virtual HRESULT STDMETHODCALLTYPE Reset (uint16_t startAddress) override
	{
		auto hr = RunOnSimulatorThread ([this, startAddress]
			{
				if (_running_info)
				{
					_running_info.value().start_time = 0;
					QueryPerformanceCounter(&_running_info.value().start_time_perf_counter);
				}
				
				_cpu->Reset();
				_cpu->SetPC(startAddress);
				for (auto& d : _active_devices_)
					d->Reset();
				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		_screenComplete = nullptr;

		if (!_running)
		{
			if (_screenCompleteHandler)
			{
				unique_cotaskmem_bitmapinfo screen;
				POINT beam;
				auto hr = _screen->CopyBuffer (_showCRTSnapshot, screen.addressof(), &beam);
				if (SUCCEEDED(hr))
				{
					auto hr = _screenCompleteHandler->OnScreenComplete(screen.get(), beam);
					if (SUCCEEDED(hr))
						screen.release();
				}
			}
		}

		return S_OK;
	}

	HRESULT RunOnSimulatorThread (stdext::inplace_function<HRESULT(), 64> fun)
	{
		_runOnSimulatorThreadFunction = std::move(fun);

		BOOL bres = SetEvent(_run_on_simulator_thread_request.get()); RETURN_LAST_ERROR_IF(!bres);

		DWORD waitRes = WaitForSingleObject(_runOnSimulatorThreadComplete.get(), IsDebuggerPresent() ? INFINITE : 5000);

		_runOnSimulatorThreadFunction = nullptr;

		switch (waitRes)
		{
		case WAIT_OBJECT_0:
			return _runOnSimulatorThreadResult;

		case WAIT_ABANDONED_0:
			RETURN_WIN32(ERROR_ABANDONED_WAIT_0);

		case WAIT_TIMEOUT:
			RETURN_WIN32(ERROR_TIMEOUT);

		case WAIT_FAILED:
			RETURN_LAST_ERROR();

		default:
			RETURN_HR(E_FAIL);
		}
	}
	
	virtual HRESULT STDMETHODCALLTYPE ReadMemoryBus (uint16_t address, uint16_t size, void* to) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);
		for (uint32_t i = 0; i < size; i++)
			((uint8_t*)to)[i] = memoryBus.read(address + i);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE WriteMemoryBus (uint16_t address, uint16_t size, const void* from) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);
		for (uint32_t i = 0; i < size; i++)
			memoryBus.write (address + i, ((uint8_t*)from)[i]);
		return S_OK;
	}
	
	virtual HRESULT STDMETHODCALLTYPE Break() override
	{
		if (!_running)
			return S_FALSE;

		auto hr = RunOnSimulatorThread([this]
			{
				WI_ASSERT(_running_info);
				_running_info.reset();
				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		_running = false;

		using BreakEvent = SimulatorEvent<ISimulatorBreakEvent>;
		if (auto event = com_ptr(new (std::nothrow) BreakEvent()))
			_eventHandlers->Notify([&event] (ISimulatorEventNotifySink* sink)
				{ sink->NotifySimulatorEvent(event, __uuidof(event)); });

		_screenComplete = nullptr;
		if (_screenCompleteHandler)
		{
			unique_cotaskmem_bitmapinfo screen;
			POINT beam;
			auto hr = _screen->CopyBuffer (_showCRTSnapshot, screen.addressof(), &beam);
			if (SUCCEEDED(hr))
			{
				auto hr = _screenCompleteHandler->OnScreenComplete(screen.get(), beam);
				if (SUCCEEDED(hr))
					screen.release();
			}
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Resume (BOOL checkBreakpointsAtCurrentPC) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);

		auto hr = RunOnSimulatorThread([this, checkBreakpointsAtCurrentPC]
			{
				AssertDevicesUpToDateWithCPU();

				auto start_time = _cpu->Time();
				LARGE_INTEGER perf_counter;
				QueryPerformanceCounter(&perf_counter);

				if (!checkBreakpointsAtCurrentPC)
				{
					bool advanced = _cpu->SimulateOne(nullptr);
					WI_ASSERT(advanced);
				}

				_running_info = running_info { .start_time = start_time, .start_time_perf_counter = perf_counter };
				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		_running = true;
		using ResumeEvent = SimulatorEvent<ISimulatorResumeEvent>;
		auto event = com_ptr(new (std::nothrow) ResumeEvent()); RETURN_IF_NULL_ALLOC(event);

		_eventHandlers->Notify([&event](ISimulatorEventNotifySink* sink)
			{ sink->NotifySimulatorEvent(event, __uuidof(event)); });

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Running_HR() override
	{
		return _running ? S_OK : S_FALSE;
	}

	void AssertDevicesUpToDateWithCPU()
	{
		#ifdef _DEBUG
		WI_ASSERT(!_running_info);

		// While simulation is not running devices are supposed to be up to date with the processor's time.
		uint64_t cpuTime = _cpu->Time();
		for (auto& d : _active_devices_)
		{
			uint64_t deviceTime = d->Time();
			if (deviceTime < cpuTime)
			{
				d->SimulateTo(cpuTime);
				uint64_t deviceNewTime = d->Time();
				WI_ASSERT(deviceNewTime == deviceTime);
			}
		}
		#endif
	}

	virtual HRESULT STDMETHODCALLTYPE SimulateOne() override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);

		auto hr = RunOnSimulatorThread ([this]
			{
				AssertDevicesUpToDateWithCPU();

				if (!_cpu->Halted())
				{
					bool advanced = _cpu->SimulateOne(nullptr);
					WI_ASSERT(advanced);
					simulate_devices_to(_cpu->Time());
				}
				else
				{
					do
					{
						bool advanced = _cpu->SimulateOne(nullptr);
						WI_ASSERT(advanced);
						simulate_devices_to(_cpu->Time());
					} while (_cpu->Halted());
				}

				return S_OK;
			});
		RETURN_IF_FAILED(hr);
		
		hr = SendSimulateOneCompleteEvent(); LOG_IF_FAILED(hr);

		_screenComplete = nullptr;
		if (_screenCompleteHandler)
		{
			unique_cotaskmem_bitmapinfo screen;
			POINT beam;
			hr = _screen->CopyBuffer (_showCRTSnapshot, screen.addressof(), &beam);
			if (SUCCEEDED(hr))
			{
				hr = _screenCompleteHandler->OnScreenComplete(screen.get(), beam);
				if (SUCCEEDED(hr))
					screen.release();
			}
		}

		return S_OK;
	}

	/*
	virtual HRESULT STDMETHODCALLTYPE LoadROM (LPCWSTR filename_relative_to_package, uint32_t address, BSTR* absolutePathOut) override
	{
		RETURN_HR_IF(E_NOT_SUPPORTED_WHILE_SIMULATION_RUNNING, _running);

		auto de_filename = wil::make_process_heap_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(de_filename);
		DWORD dwres = GetModuleFileName (hModule, de_filename.get(), MAX_PATH + 1); RETURN_LAST_ERROR_IF(!dwres || (dwres == MAX_PATH + 1));
		auto fnres = PathFindFileName(de_filename.get()); RETURN_HR_IF(E_FAIL, fnres == de_filename.get());
		*fnres = 0;
		auto rom_path = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(rom_path);
		fnres = PathCombineW(rom_path.get(), de_filename.get(), filename_relative_to_package); RETURN_HR_IF(E_FAIL, !fnres);

		com_ptr<IStream> rom_stream;
		auto hr = SHCreateStreamOnFileEx (rom_path.get(), STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &rom_stream); RETURN_IF_FAILED(hr);

		STATSTG stat;
		hr = rom_stream->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		if (stat.cbSize.HighPart)
			RETURN_WIN32(ERROR_FILE_TOO_LARGE);
		if (!stat.cbSize.LowPart)
			RETURN_HR(E_BINARY_FILE_EMPTY);
		if (address > 0xFFFF || stat.cbSize.LowPart > 0x10000 || address + stat.cbSize.LowPart > 0x10000)
			RETURN_HR(E_NOTIMPL);

		auto buffer = wil::make_unique_hlocal_nothrow<uint8_t[]>(stat.cbSize.LowPart);
		ULONG bytes_read;
		hr = rom_stream->Read(buffer.get(), stat.cbSize.LowPart, &bytes_read); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, bytes_read != stat.cbSize.LowPart);

		_rom->write_rom(buffer.get(), stat.cbSize.LowPart);

		if (absolutePathOut)
		{
			*absolutePathOut = SysAllocString(rom_path.get()); RETURN_IF_NULL_ALLOC(*absolutePathOut);
		}

		return S_OK;
	}
	*/
	HRESULT LoadSnapshot (const wchar_t* pFileName)
	{
		#pragma pack (push, 1)
		// https://rk.nvg.ntnu.no/sinclair/faq/fileform.html#SNA
		struct snapshot_file_header
		{
			uint8_t i;
			uint16_t alt_hl;
			uint16_t alt_de;
			uint16_t alt_bc;
			uint16_t alt_af;
			uint16_t hl;
			uint16_t de;
			uint16_t bc;
			uint16_t iy;
			uint16_t ix;
			uint8_t      : 1;
			uint8_t ei   : 1;
			uint8_t iff2 : 1;
			uint8_t      : 5;
			uint8_t r;
			uint16_t af;
			uint16_t sp;
			uint8_t im;
			uint8_t border;
		};
		#pragma pack (pop)

		com_ptr<IStream> stream;
		auto hr = SHCreateStreamOnFileEx (pFileName, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &stream); RETURN_IF_FAILED_EXPECTED(hr);

		STATSTG stat;
		hr = stream->Stat (&stat, STATFLAG_NONAME); RETURN_IF_FAILED_EXPECTED(hr);
		if (stat.cbSize.QuadPart != sizeof(snapshot_file_header) + 48 * 1024)
			return SetErrorInfo(E_FAIL, L"The file has an unrecognized size.\r\n(Only ZX Spectrum 48K files are supported for now.)");

		snapshot_file_header header;
		ULONG read;
		hr = stream->Read (&header, (ULONG)sizeof(header), &read); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_FAIL, read != sizeof(header));

		auto buffer = wil::make_unique_hlocal_nothrow<uint8_t[]>(48 * 1024); RETURN_IF_NULL_ALLOC_EXPECTED(buffer);
		hr = stream->Read(buffer.get(), 48 * 1024, &read); RETURN_IF_FAILED_EXPECTED(hr); RETURN_HR_IF(E_FAIL, read != 48 * 1024);

		unique_cotaskmem_bitmapinfo screen;
		POINT beam;

		// It's ok to catch by reference since the call to RunOnSimulatorThread is blocking (returns when the work is complete).
		hr = RunOnSimulatorThread ([this, &header, buffer=buffer.get(), &screen, &beam]
			{
				HRESULT hr;

				_cpu->Reset();
				for (auto& d : _active_devices_)
					d->Reset();

				_ramDevice->WriteMemory(0x4000, 48 * 1024, buffer);

				z80_register_set regs;
				regs.halted = false;
				regs.i       = header.i;
				regs.alt.hl  = header.alt_hl;
				regs.alt.bc  = header.alt_bc;
				regs.alt.de  = header.alt_de;
				regs.alt.af  = header.alt_af;
				regs.main.hl = header.hl;
				regs.main.de = header.de;
				regs.main.bc = header.bc;
				regs.ix      = header.ix;
				regs.iy      = header.iy;
				regs.iff1    = header.ei;
				regs.r       = header.r;
				regs.main.af = header.af;
				regs.sp      = header.sp;
				regs.im      = header.im;
				regs.pc = memoryBus.read_uint16(regs.sp);
				regs.sp += 2;
				_cpu->SetZ80Registers(&regs);

				if (_running_info)
				{
					_running_info.value().start_time = 0;
					QueryPerformanceCounter(&_running_info.value().start_time_perf_counter);
				}
				else
				{
					hr = _screen->GenerateScreen(); RETURN_IF_FAILED_EXPECTED(hr);
					hr = _screen->CopyBuffer(TRUE, &screen, &beam); RETURN_IF_FAILED_EXPECTED(hr);
				}

				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		// TODO: do something to cause the GUI to refresh windows
		//if (!_running)
		//	SendSimulateOneCompleteEvent();

		_screenComplete = nullptr;
		if (screen && _screenCompleteHandler)
		{
			hr = _screenCompleteHandler->OnScreenComplete(screen.get(), beam);
			if (SUCCEEDED(hr))
				screen.release();
		}

		return S_OK;
	}

	#pragma pack (push, 1)
	// https://worldofspectrum.org/faq/reference/z80format.htm
	struct z80_header
	{
		uint8_t  a;  // 0
		uint8_t  f;  // 1
		uint16_t bc; // 2-3
		uint16_t hl; // 4-5
		uint16_t pc; // 6-7
		uint16_t sp; // 8-9
		uint8_t  i;  // 10
		uint8_t  r;  // 11
		uint8_t  r0         : 1; // 12 b0
		uint8_t  border     : 3; // 12 b1-b3
		uint8_t  samrom     : 1; // 12 b4
		uint8_t  compressed : 1; // 12 b5
		uint8_t             : 2; // 12 b6-b7
		uint16_t de;     // 13-14
		uint16_t alt_bc; // 15-16
		uint16_t alt_de; // 17-18
		uint16_t alt_hl; // 19-20
		uint8_t  alt_a;  // 21
		uint8_t  alt_f;  // 22
		uint16_t iy;     // 23-24
		uint16_t ix;     // 25-26
		uint8_t  ei;     // 27
		uint8_t  iff2;   // 28
		uint8_t  im       : 2; // 29 b0-b1
		uint8_t  issue2   : 1; // 29 b2
		uint8_t  freq2x   : 1; // 29 b3
		uint8_t  vid_sync : 2; // 29 b4-b5
		uint8_t  joystick : 2; // 29 b6-b7
	};
	#pragma pack (pop)

	HRESULT SetMalformedErrorInfo (const wchar_t* detail)
	{
		return SetErrorInfo(E_FAIL, L"Malformed .Z80 file: %s", detail);
	};

	HRESULT DecompressZ80 (const uint8_t* inPtr, const uint8_t* inEnd, uint8_t* outPtr, uint8_t* outEnd, bool terminatorExpected)
	{
		while (true)
		{
			if (inPtr >= inEnd)
			{
				if (terminatorExpected)
					return SetMalformedErrorInfo (L"Format version 1, compressed, no 00EDED00 terminator.");

				return S_OK;
			}

			if ((inPtr + 4 <= inEnd) && !memcmp(inPtr, "\x00\xED\xED\x00", 4))
				return S_OK;
			
			if ((inPtr + 4 <= inEnd) && (inPtr[0] == 0xED) && (inPtr[1] == 0xED))
			{
				inPtr += 2;
				uint8_t repeat = *inPtr++;
				uint8_t value = *inPtr++;
				if (outPtr + repeat > outEnd)
					return SetMalformedErrorInfo (L"Longer than 48K.");
				memset (outPtr, value, repeat);
				outPtr += repeat;
			}
			else
			{
				if (outPtr + 1 > outEnd)
					return SetMalformedErrorInfo (L"Longer than 48K.");
				*outPtr++ = *inPtr++;
			}
		}
	}

	HRESULT LoadZ80V1 (const z80_header* header, const uint8_t* inPtr, const uint8_t* inEnd, uint8_t* outPtr, uint8_t* outEnd)
	{
		if (!header->compressed)
		{
			if (inEnd - inPtr != 48 * 1024)
				return SetMalformedErrorInfo (L"Format version 1, not compressed, size not 48K.");
			memcpy (outPtr, inPtr, 48 * 1024);
		}
		else
		{
			auto hr = DecompressZ80(inPtr, inEnd, outPtr, outEnd, true); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}

	#pragma pack (push, 1)
	struct z80_header_v23
	{
		uint16_t len;      // 30
		uint16_t pc;       // 32
		uint8_t  hardware_mode; // 34
		uint8_t  byte_35;  // 35
		uint8_t  byte_36;  // 36
		uint8_t  r_emu     : 1; // 37 b0
		uint8_t  ldir_emu  : 1; // 37 b1
		uint8_t  ay_in_use : 1; // 37 b2
		uint8_t            : 3; // 37 b3-b5
		uint8_t  fuller    : 1; // 37 b6
		uint8_t  modify_hw : 1; // 37 b7
		uint8_t  out_fffd;       // 38
		uint8_t  sound_chip[16]; // 39-54
		uint16_t t_counter_low;  // 55-56
		uint8_t  t_counter_high; // 57
		uint8_t  : 8;  // 58 - Flag byte used by Spectator
		uint8_t  : 8;  // 59 - 0xff if MGT Rom paged
		uint8_t  : 8;  // 60 - 0xff if Multiface Rom paged.
		uint8_t  : 8;  // 61 - 0xff if 0-8191 is ROM, 0 if RAM
		uint8_t  : 8;  // 62 - 0xff if 8192-16383 is ROM, 0 if RAM
		uint8_t  joystick_key_mapping[10];  // 63
		uint8_t  keys[10];  // 73
		uint8_t  : 8;  // 83
		uint8_t  : 8;  // 84
		uint8_t  : 8;  // 85
		uint8_t  : 8;  // 86
	};
	#pragma pack (pop)

	static inline const wchar_t* const HardwareModeNamesV2[] = {
		L"48K", L"48K+IF1", L"SamRam", L"128K", L"128K+IF1" };
	static inline const wchar_t* const HardwareModeNamesV3[] = {
		L"48K", L"48K+IF1", L"SamRam", L"48K+M.G.T.", L"128K", L"128K+IF1", L"128K+M.G.T." };

	HRESULT LoadZ80V23 (const z80_header_v23* header23, const uint8_t* inPtr, const uint8_t* inEnd, uint8_t* outPtr, uint8_t* outEnd)
	{
		const wchar_t* const* modeNames;
		size_t modeNameCount;
		if (header23->len == 23)
		{
			// Version 2
			modeNames = HardwareModeNamesV2;
			modeNameCount = std::size(HardwareModeNamesV2);
		}
		else if (header23->len == 54 || header23->len == 55)
		{
			// Version 3
			modeNames = HardwareModeNamesV3;
			modeNameCount = std::size(HardwareModeNamesV3);
		}

		if (auto hw = header23->hardware_mode; hw != 0)
			return SetErrorInfo (E_FAIL, L"This file specifies hardware mode %u (%s), "
				"but the simulator currently supports only hardware mode 0 (%s).", 
				hw, (hw < modeNameCount) ? modeNames[hw] : L"??", modeNames[0]);

		while (inPtr < inEnd)
		{
			if (inEnd - inPtr < 3)
				return SetMalformedErrorInfo(L"Block Header too short.");
			uint16_t lengthCompressedData = inPtr[0] | (inPtr[1] << 8);
			uint8_t pageNumber = inPtr[2];
			inPtr += 3;
			
			uint8_t* pagePtr;
			if (pageNumber == 8)
				pagePtr = outPtr; // Spectrum address 0x4000
			else if (pageNumber == 4)
				pagePtr = &outPtr[0x4000]; // Spectrum address 0x8000
			else if (pageNumber == 5)
				pagePtr = &outPtr[0x8000]; // Spectrum address 0xC000
			else
				return SetErrorInfo(E_FAIL, L"Unknown page number (%u) in Spectrum 48K mode.", pageNumber);

			if (lengthCompressedData == 0xFFFF)
			{
				if (inEnd - inPtr < 0x4000)
					return SetErrorInfo (E_FAIL, L"Page %u data too short. Expected %u, found %d.", pageNumber, 0x4000, inEnd - inPtr);
				memcpy (pagePtr, inPtr, 0x4000);
				inPtr += 0x4000;
			}
			else
			{
				auto hr = DecompressZ80 (inPtr, inPtr + lengthCompressedData, pagePtr, pagePtr + 0x4000, false); RETURN_IF_FAILED(hr);
				inPtr += lengthCompressedData;
			}
		}

		return S_OK;
	}

	HRESULT LoadZ80 (const wchar_t* pFileName)
	{
		com_ptr<IStream> stream;
		auto hr = SHCreateStreamOnFileEx (pFileName, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &stream); RETURN_IF_FAILED_EXPECTED(hr);

		STATSTG stat;
		hr = stream->Stat (&stat, STATFLAG_NONAME); RETURN_IF_FAILED_EXPECTED(hr);

		z80_header header;
		ULONG read;
		hr = stream->Read (&header, (ULONG)sizeof(header), &read); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_FAIL, read != sizeof(header));
		uint32_t inSize = stat.cbSize.LowPart - sizeof(header);
		auto inBuffer = wil::unique_hglobal_ptr<uint8_t>((uint8_t*)GlobalAlloc(GMEM_FIXED, inSize)); RETURN_IF_NULL_ALLOC_EXPECTED(inBuffer);
		hr = stream->Read(inBuffer.get(), inSize, &read); RETURN_IF_FAILED_EXPECTED(hr); RETURN_HR_IF(E_FAIL, read != inSize);
		const uint8_t* inPtr = inBuffer.get();
		const uint8_t* inEnd = inPtr + inSize;

		auto outBuffer = wil::unique_hglobal_ptr<uint8_t>((uint8_t*)GlobalAlloc(GMEM_FIXED, 48 * 1024)); RETURN_IF_NULL_ALLOC_EXPECTED(outBuffer);
		uint8_t* outPtr = outBuffer.get();
		uint8_t* outEnd = outPtr + 48 * 1024;

		uint16_t pc;
		if (header.pc != 0)
		{
			hr = LoadZ80V1 (&header, inPtr, inEnd, outPtr, outEnd); RETURN_IF_FAILED_EXPECTED(hr);
			pc = header.pc;
		}
		else
		{
			if (inEnd - inPtr < 2)
				return SetMalformedErrorInfo(L"Header v2/3 too short.");
			uint16_t header23_len = inPtr[0] | (inPtr[1] << 8);
			if (inEnd - inPtr < (2 + header23_len))
				return SetMalformedErrorInfo(L"Header v2/3 too short.");
			auto* header23 = (const z80_header_v23*)inPtr;
			inPtr += (2 + header23_len);
			hr = LoadZ80V23 (header23, inPtr, inEnd, outPtr, outEnd); RETURN_IF_FAILED_EXPECTED(hr);
			pc = header23->pc;
		}

		unique_cotaskmem_bitmapinfo screen;
		POINT beam;

		// It's ok to catch by reference since the call to RunOnSimulatorThread is blocking (returns when the work is complete).
		hr = RunOnSimulatorThread ([this, pc, &header, buffer=outBuffer.get(), &screen, &beam]
			{
				HRESULT hr;

				_cpu->Reset();
				for (auto& d : _active_devices_)
					d->Reset();

				_ramDevice->WriteMemory(0x4000, 48 * 1024, buffer);

				z80_register_set regs;
				regs.halted = false;
				regs.i       = header.i;
				regs.alt.hl  = header.alt_hl;
				regs.alt.bc  = header.alt_bc;
				regs.alt.de  = header.alt_de;
				regs.alt.a   = header.alt_a;
				regs.alt.f.val = header.alt_f;
				regs.main.hl = header.hl;
				regs.main.de = header.de;
				regs.main.bc = header.bc;
				regs.ix      = header.ix;
				regs.iy      = header.iy;
				regs.iff1    = header.ei;
				regs.r       = header.r;
				regs.main.a  = header.a;
				regs.main.f.val = header.f;
				regs.sp      = header.sp;
				regs.im      = header.im;
				regs.pc = pc;
				_cpu->SetZ80Registers(&regs);

				if (_running_info)
				{
					_running_info.value().start_time = 0;
					QueryPerformanceCounter(&_running_info.value().start_time_perf_counter);
				}
				else
				{
					hr = _screen->GenerateScreen(); RETURN_IF_FAILED_EXPECTED(hr);
					hr = _screen->CopyBuffer(TRUE, &screen, &beam); RETURN_IF_FAILED_EXPECTED(hr);
				}

				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		// TODO: do something to cause the GUI to refresh windows
		//if (!_running)
		//	SendSimulateOneCompleteEvent();

		_screenComplete = nullptr;
		if (screen && _screenCompleteHandler)
		{
			hr = _screenCompleteHandler->OnScreenComplete(screen.get(), beam);
			if (SUCCEEDED(hr))
				screen.release();
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE LoadFile (LPCWSTR pFileName) override
	{
		auto* ext = PathFindExtension(pFileName);

		if (!_wcsicmp(ext, L".sna"))
			return LoadSnapshot(pFileName);

		if (!_wcsicmp(ext, L".z80"))
			return LoadZ80(pFileName);

		return SetErrorInfo (E_FAIL, L"The file extension %s is not recognized.", ext);
	}

	virtual HRESULT STDMETHODCALLTYPE LoadBinary (LPCWSTR pFileName, DWORD address, DWORD* loadedSize) override
	{
		HRESULT hr;

		RETURN_HR_IF(E_UNEXPECTED, _running);

		wil::com_ptr_nothrow<IStream> stream;
		hr = SHCreateStreamOnFileEx (pFileName, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &stream);
		if (FAILED(hr))
			return SetErrorInfo (hr, L"Cannot access file.\r\n\r\n%s", pFileName);

		STATSTG stat;
		hr = stream->Stat (&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		if (stat.cbSize.HighPart)
			RETURN_HR(E_BOUNDS);
		if (!stat.cbSize.LowPart)
			return SetErrorInfo(E_BOUNDS, L"Can't load a binary with zero length.\r\n\r\n%s", pFileName);
		if (address + (uint16_t)stat.cbSize.LowPart <= address)
			RETURN_HR(E_BOUNDS);
		DWORD from, to;
		_ramDevice->GetBounds(&from, &to);
		if ((address < from) || (address + stat.cbSize.LowPart > to))
			return SetErrorInfo(E_BOUNDS, L"Binary does not fit in RAM.\r\n\r\n"
				"Binary load address range is 0x%04X...0x%04X.\r\n\r\n"
				"RAM address range = 0x%04X...0x%04X.\r\n\r\n"
				"Try to adjust the BaseAddress value in the project properties."
				, address, address + stat.cbSize.LowPart, from, to);

		uint8_t buffer[128];
		for (uint16_t i = 0; i < stat.cbSize.LowPart; i += sizeof(buffer))
		{
			ULONG read;
			hr = stream->Read (buffer, sizeof(buffer), &read); RETURN_IF_FAILED(hr);
			//_memory.write (address + (uint16_t)i, { buffer, buffer + read });
			_ramDevice->WriteMemory (address + i, read, buffer);
		}

		if (_screenCompleteHandler)
		{
			unique_cotaskmem_bitmapinfo screenBuffer;
			POINT beam;
			hr = RunOnSimulatorThread([&screenBuffer, &beam, this]
			{
				return _screen->CopyBuffer (_showCRTSnapshot, &screenBuffer, &beam);
			});

			if (SUCCEEDED(hr))
			{
				hr = _screenCompleteHandler->OnScreenComplete(screenBuffer.get(), beam); RETURN_IF_FAILED(hr);
				if (SUCCEEDED(hr))
					screenBuffer.release();
			}
		}

		*loadedSize = stat.cbSize.LowPart;
		return S_OK;
	}
	/*
	virtual HRESULT STDMETHODCALLTYPE SaveMemory (IStream* stream) override
	{
		RETURN_HR_IF(E_NOT_SUPPORTED_WHILE_SIMULATION_RUNNING, _running);

		uint8_t buffer[128];
		for (uint32_t i = 0; i < 0x10000; i += sizeof(buffer))
		{
			//_memory.read(i, { buffer, buffer + sizeof(buffer) });
			for (uint32_t ii = 0; ii < sizeof(buffer); ii++)
				buffer[ii] = _memory.read(i + ii);
			ULONG written;
			auto hr = stream->Write (buffer, (ULONG)sizeof(buffer), &written); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}
	*/

	virtual HRESULT STDMETHODCALLTYPE AdviseScreenComplete (IScreenCompleteEventHandler* handler) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _screenCompleteHandler != nullptr);
		_screenCompleteHandler = handler;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseScreenComplete (IScreenCompleteEventHandler* handler) override
	{
		RETURN_HR_IF(E_INVALIDARG, handler != _screenCompleteHandler);
		_screenCompleteHandler = nullptr;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE ProcessKeyDown (uint32_t vkey, uint32_t modifiers) override
	{
		return RunOnSimulatorThread([this, vkey, modifiers] { return _keyboard->ProcessKeyDown(vkey, modifiers); });
	}

	virtual HRESULT STDMETHODCALLTYPE ProcessKeyUp   (uint32_t vkey, uint32_t modifiers) override
	{
		return RunOnSimulatorThread([this, vkey, modifiers] { return _keyboard->ProcessKeyUp(vkey, modifiers); });
	}

	virtual HRESULT STDMETHODCALLTYPE AddBreakpoint (BreakpointType type, bool physicalMemorySpace, UINT64 address, SIM_BP_COOKIE* pCookie) override
	{
		RETURN_HR_IF(E_NOTIMPL, physicalMemorySpace);
		RETURN_HR_IF(E_INVALIDARG, address > 0xFFFF);

		return RunOnSimulatorThread([this, type, address, pCookie]
			{
				auto hr = _cpu->AddBreakpoint(type, (uint16_t)address, pCookie);
				if (FAILED(hr))
					return hr;
				WI_ASSERT(*pCookie != 0);
				return S_OK;
			});
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveBreakpoint (SIM_BP_COOKIE cookie) override
	{
		RETURN_HR_IF(E_INVALIDARG, cookie == 0);
		
		return RunOnSimulatorThread([this, cookie]
			{
				return _cpu->RemoveBreakpoint(cookie);
			});
	}

	virtual HRESULT STDMETHODCALLTYPE HasBreakpoints_HR() override
	{
		return RunOnSimulatorThread([this]
			{
				return _cpu->HasBreakpoints() ? S_OK : S_FALSE;
			});
	}

	virtual HRESULT STDMETHODCALLTYPE GetPC (uint16_t* pc) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);
		*pc = _cpu->GetPC();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetPC (uint16_t pc) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);
		return _cpu->SetPC(pc);
	}

	virtual HRESULT STDMETHODCALLTYPE GetStackStartAddress (UINT16* stackStartAddress) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);
		*stackStartAddress = _cpu->GetStackStartAddress();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetRegisters (z80_register_set* buffer, uint32_t size) override
	{
		RETURN_HR_IF(E_UNEXPECTED, _running);
		_cpu->GetZ80Registers(buffer);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetRegisters (const z80_register_set* buffer, uint32_t size) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE GetShowCRTSnapshot() override
	{
		return _showCRTSnapshot ? S_OK : S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE SetShowCRTSnapshot(BOOL val) override
	{
		if (_showCRTSnapshot != (bool)val)
		{
			_showCRTSnapshot = (bool)val;

			if (_screenCompleteHandler)
			{
				unique_cotaskmem_bitmapinfo screenBuffer;
				POINT beam;
				auto hr = RunOnSimulatorThread([&screenBuffer, &beam, this]
				{
					return _screen->CopyBuffer (_showCRTSnapshot, &screenBuffer, &beam);
				});

				if (SUCCEEDED(hr))
				{
					hr = _screenCompleteHandler->OnScreenComplete(screenBuffer.get(), beam); RETURN_IF_FAILED(hr);
					if (SUCCEEDED(hr))
						screenBuffer.release();
				}
			}
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IScreenDeviceCompleteEventHandler
	virtual void OnScreenDeviceComplete() override
	{
		// No error checking, not even logging, as this function is called 50 times a second
		// and in case of error it would probably freeze the app.

		// TODO: register for this callback when simulation starts running, unregister when simulation paused.
		if (_running_info)
		{
			// This callback is called when the screen device finishes rendering a complete screen. This means
			// the image on the simulated screen is identical to the image in the video memory. Thus CopyBuffer
			// creates the same image regardless of the value of its "BOOL crt" parameter. Let's pass TRUE
			// since the screen is already generated as a BITMAPINFO and it's much faster to simply copy it.
			BOOL crt = TRUE;

			unique_cotaskmem_bitmapinfo screen;
			auto hr = _screen->CopyBuffer (crt, screen.addressof(), nullptr);
			if (SUCCEEDED(hr))
			{
				auto lock = _mainThreadQueueLock.lock_exclusive();
				bool messagePosted = _screenComplete.is_valid();
				_screenComplete = std::move(screen);
				if (!messagePosted)
				{
					BOOL posted = PostMessageW (_hwnd, WM_SCREEN_COMPLETE, 0, 0);
					if (!posted)
					{
						// Ignoring this error condition for now, don't know how to handle it.
					}
				}
			}
		}
	}
	#pragma endregion

	#pragma region IConnectionPointContainer
	virtual HRESULT STDMETHODCALLTYPE EnumConnectionPoints (IEnumConnectionPoints **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE FindConnectionPoint (REFIID riid, IConnectionPoint **ppCP) override
	{
		if (riid == __uuidof(ISimulatorEventNotifySink))
		{
			*ppCP = _eventHandlers;
			(*ppCP)->AddRef();
			return S_OK;
		}

		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion
};

HRESULT MakeSimulator (LPCWSTR dir, LPCWSTR romFilename, ISimulator** sim)
{
	com_ptr<SimulatorImpl> s = new (std::nothrow) SimulatorImpl(); RETURN_IF_NULL_ALLOC(s);
	auto hr = s->InitInstance(dir, romFilename); RETURN_IF_FAILED(hr);
	*sim = s.detach();
	return S_OK;
}
