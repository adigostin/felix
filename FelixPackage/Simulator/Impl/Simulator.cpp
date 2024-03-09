
#include "pch.h"
#include "../Simulator.h"
#include "Bus.h"
#include "Z80CPU.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/TryQI.h"
#include "shared/inplace_function.h"
#include <optional>

#pragma comment (lib, "Shlwapi")

static constexpr UINT WM_MAIN_THREAD_WORK = WM_APP + 0;
static ATOM wndClassAtom;

HRESULT STDMETHODCALLTYPE MakeHC91ROM (Bus* memory_bus, Bus* io_bus, const wchar_t* folder, const wchar_t* BinaryFilename, wistd::unique_ptr<IMemoryDevice>* ppDevice);
HRESULT STDMETHODCALLTYPE MakeHC91RAM (Bus* memory_bus, Bus* io_bus, wistd::unique_ptr<IMemoryDevice>* ppDevice);
HRESULT STDMETHODCALLTYPE MakeBeeper (Bus* io_bus, wistd::unique_ptr<IDevice>* ppDevice);

// ============================================================================

class SimulatorImpl : public ISimulator, IDeviceEventHandler, IScreenDeviceCompleteEventHandler
{
	ULONG _refCount = 0;

	static constexpr wchar_t WndClassName[] = L"Z80Program-{5F067572-7D58-40A1-A471-7FE701460B3B}";

	HWND _hwnd = nullptr;

	LARGE_INTEGER qpFrequency;

	struct running_info
	{
		UINT64 start_time;
		LARGE_INTEGER start_time_perf_counter;
	};

	std::optional<running_info> _running_info; // this is used only by the simulator thread
	bool _running = false; // and this only by the main thread
	wil::unique_handle _cpuThread;
	wil::unique_handle _cpu_thread_exit_request;
	vector_nothrow<com_ptr<ISimulatorEventHandler>> _eventHandlers;
	vector_nothrow<com_ptr<IScreenCompleteEventHandler>> _screenCompleteHandlers;

	using RunOnSimulatorThreadFunction = HRESULT(*)(void*);
	stdext::inplace_function<HRESULT()> _runOnSimulatorThreadFunction;
	HRESULT            _runOnSimulatorThreadResult;
	wil::unique_handle _run_on_simulator_thread_request;
	wil::unique_handle _runOnSimulatorThreadComplete;
	wil::unique_handle _waitableTimer;

	vector_nothrow<stdext::inplace_function<HRESULT()>> _mainThreadWorkQueue;
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

public:
	HRESULT InitInstance (LPCWSTR dir, LPCWSTR romFilename)
	{
		auto hr = MakeZ80CPU(&memoryBus, &ioBus, &irq, &_cpu); RETURN_IF_FAILED(hr);

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

		SetThreadDescription(_cpuThread.get(), L"Simulator");

		return S_OK;
	};

	~SimulatorImpl()
	{
		WI_ASSERT (_eventHandlers.empty());
		WI_ASSERT (_screenCompleteHandlers.empty());

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
		)
			return S_OK;

		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override
	{
		return InterlockedIncrement(&_refCount);
	}

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		WI_ASSERT(_refCount);
		if (_refCount > 1)
			return InterlockedDecrement(&_refCount);
		delete this;
		return 0;
	}
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
			auto hr = work(); LOG_IF_FAILED(hr);
			return hr;
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
					advanced |= (d->SimulateTo(time, this) == S_OK);
			}

			if (!advanced)
				return res;

			res = true;
		}
	}

	void simulate_one_in_break_mode()
	{
		WI_ASSERT(::GetCurrentThreadId() == GetThreadId(_cpuThread.get()));
		WI_ASSERT(!_running_info);

		// The CPU might be waiting for a device, and that device might be waiting for the real time to catch up.
		// This is often the case after we break a running program; see the related comment in Break().
		// So before we attempt anything with the CPU, let's bring the devices up to date with it.
		simulate_devices_to(_cpu->Time());

		// Let's simulate one CPU instruction. Since this is a user-initiated action,
		// we must ask the CPU to ignore breakpoints while executing this particular instruction.
		com_ptr<IUnknown> outcome;
		auto hr = _cpu->SimulateOne(false, this, &outcome);
		if (hr == S_OK)
		{
		}
		else if (hr == SIM_E_UNDEFINED_INSTRUCTION)
		{
			on_undefined_instruction(_cpu->GetPC());
		}
		else
			WI_ASSERT(false); // we're not supposed to get any other outcome
	}

	DWORD simulation_thread_proc()
	{
		// TODO: "To initialize a thread as free-threaded, call CoInitializeEx, specifying COINIT_MULTITHREADED"
		bool exit_request = false;
		while(!exit_request)
		{
			// TODO: add support for syncing on multiple devices, needed in case two devices need to sync on
			// the same time _and_ the first is waiting for the second (we'd have an infinite loop with the code as it is now).
			if (_running_info)
			{
				UINT64 rt = real_time();

				#pragma region Catch up with real time
				// Before we attempt simulation, let's check if some devices are lagging far behind the real time.
				// This happens while debugging the VSIX, or it may happen when this thread is starved. 
				// If we have any such device, we "erase" the same length of time from all of the devices.
				{
					IDevice* slowest = _cpu.get();
					INT64 slowest_offset_from_rt = (UINT64)(_cpu->Time() - rt);
					for (auto& d : _active_devices_)
					{
						INT64 offset_from_rt = (UINT64)(d->Time() - rt);
						if (offset_from_rt < slowest_offset_from_rt)
						{
							slowest = d;
							slowest_offset_from_rt = offset_from_rt;
						}
					}

					if (slowest_offset_from_rt < -(INT64)milliseconds_to_ticks(50))
					{
						// Slowest device is more than 50 ms behind real time. Let's see what time offset
						// it would need to be 50 ms _ahead_ of real time, and add that offset to all devices.
						UINT64 offset = rt + milliseconds_to_ticks(50) - slowest->Time();
						_cpu->SkipTime(offset);
						for (auto& d : _active_devices_)
							d->SkipTime(offset);
					}
				}
				#pragma endregion

				// --------------------------------------------
				// Let's find out how far we can simulate, and try to simulate to that point in time.
				IDevice* device_to_sync_on = nullptr;
				UINT64 time_to_sync_to_ = rt + milliseconds_to_ticks(1000);
				for (auto& d : _active_devices_)
				{
					UINT64 t;
					if (d->NeedSyncWithRealTime(&t)
						&& (!device_to_sync_on || (t < time_to_sync_to_)))
					{
						device_to_sync_on = d;
						time_to_sync_to_ = t;
					}
				};
				//WI_ASSERT(device_to_sync_on);

				DWORD timeout; // Either 0 or INFINITE. For other values, WaitForMultipleObjects would use an imprecise timer.
				LARGE_INTEGER duration_to_sleep; // 100 nanosecond intervals for the precise timer, used only when timeout==INFINITE
				HRESULT outcomeHR;
				com_ptr<IUnknown> outcomeUnk;
				while(true)
				{
					// First ask the CPU to simulate itself; then ask the devices
					// to simulate themselves until they catch up with the CPU, not more.
					// We don't want the devices to go far ahead of the CPU; if the CPU will stop at a breakpoint
					// or an undefined instruction, we'll want to show to the user the devices at a time as close
					// as possible to the CPU time; we can do that only if the devices are at all times behind
					// the CPU or only slightly (a few clock cycles) ahead of it.
					outcomeHR = S_OK;
					while (!device_to_sync_on || (_cpu->Time() < time_to_sync_to_))
					{
						bool check_breakpoints = true;
						outcomeHR = _cpu->SimulateOne(check_breakpoints, this, &outcomeUnk);
						if (outcomeHR != S_OK)
							break;
					}

					// Whatever the outcome, we must first bring the devices close to the CPU time.
					simulate_devices_to(_cpu->Time());

					if (outcomeHR == SIM_E_BREAKPOINT_HIT)
					{
						com_ptr<ISimulatorBreakpointEvent> bps;
						auto hr = outcomeUnk->QueryInterface(&bps); WI_ASSERT(SUCCEEDED(hr));
						hr = on_bp_hit(bps.get()); WI_ASSERT(SUCCEEDED(hr));
						timeout = INFINITE;
						duration_to_sleep.QuadPart = 0; // not used
						break;
					}
					else if (outcomeHR == SIM_E_UNDEFINED_INSTRUCTION)
					{
						on_undefined_instruction(_cpu->GetPC());
						timeout = INFINITE;
						duration_to_sleep.QuadPart = 0; // not used
						break;
					}
					else if (outcomeHR == S_OK)
					{
						// Now let's see how long we need to wait for the real time to catch up.
						WI_ASSERT(device_to_sync_on);
						timeout = 0;
						duration_to_sleep.QuadPart = 0;
						//auto rt = real_time(); // read it again cause we might have simulated a lot
						// Later edit: the line above was causing a lot of trouble, and the benefit was just theoretical.
						if (time_to_sync_to_ > rt)
						{
							uint64_t hundredsOfNs = ticks_to_hundreds_of_nanoseconds(time_to_sync_to_ - rt);
							if (hundredsOfNs)
							{
								timeout = INFINITE;
								duration_to_sleep.QuadPart = -(INT64)hundredsOfNs;
							}
						}
						break;
					}
					else if (outcomeHR == S_FALSE)
					{
					}
					else
						WI_ASSERT(false);
				}

				HANDLE waitHandles[3] = { _run_on_simulator_thread_request.get(), _cpu_thread_exit_request.get() };
				DWORD waitCount = 2;
				if (duration_to_sleep.QuadPart != 0)
				{
					BOOL bRes = SetWaitableTimer (_waitableTimer.get(), &duration_to_sleep, 0, nullptr, nullptr, FALSE); LOG_IF_WIN32_BOOL_FALSE(bRes);
					if (bRes)
					{
						waitHandles[2] = _waitableTimer.get();
						waitCount = 3;
					}
				}
				
				DWORD waitResult = WaitForMultipleObjects (waitCount, waitHandles, FALSE, timeout);
				if (waitResult == WAIT_TIMEOUT 
					|| waitResult == WAIT_OBJECT_0 + 2)
				{
					WI_ASSERT(device_to_sync_on);

					//WI_ASSERT(real_time() >= time_to_sync_to);
					// Real time has caugth up with the simulated time.
					// Let's unblock the device that was waiting for this time point.
					if (device_to_sync_on->Time() < time_to_sync_to_ + 1)
					{
						bool res = device_to_sync_on->SimulateTo(time_to_sync_to_ + 1, this);
						WI_ASSERT(res);
					}
				}
				else if (waitResult == WAIT_OBJECT_0)
				{
					_runOnSimulatorThreadResult = _runOnSimulatorThreadFunction();
					SetEvent(_runOnSimulatorThreadComplete.get());
				}
				else if (waitResult == WAIT_OBJECT_0 + 1)
				{
					exit_request = true;
				}
				else
					WI_ASSERT(false); // TODO: handle error conditions
			}
			else
			{
				const HANDLE waitHandles[] = { _run_on_simulator_thread_request.get(), _cpu_thread_exit_request.get() };
				DWORD waitResult = WaitForMultipleObjects ((DWORD)std::size(waitHandles), waitHandles, FALSE, INFINITE);
				if (waitResult == WAIT_OBJECT_0)
				{
					_runOnSimulatorThreadResult = _runOnSimulatorThreadFunction();
					SetEvent(_runOnSimulatorThreadComplete.get());
				}
				else if (waitResult == WAIT_OBJECT_0 + 1)
				{
					exit_request = true;
				}
				else
					WI_ASSERT(false); // TODO: handle error conditions
			}
		}

		return 0;
	}

	HRESULT STDMETHODCALLTYPE PostWorkToMainThread (stdext::inplace_function<HRESULT()> work) 
	{
		WI_ASSERT(::GetCurrentThreadId() == GetThreadId(_cpuThread.get()));
		auto lock = _mainThreadQueueLock.lock_exclusive();
		bool pushed = _mainThreadWorkQueue.try_push_back(std::move(work)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

		BOOL posted = PostMessageW (_hwnd, WM_MAIN_THREAD_WORK, 0, 0);
		if (!posted)
		{
			DWORD le = GetLastError();
			LOG_LAST_ERROR();
			_mainThreadWorkQueue.remove(_mainThreadWorkQueue.end() - 1);
			return HRESULT_FROM_WIN32(le);
		}

		return S_OK;
	}

	// Called when simulation was running and the CPU reached a breakpoint.
	HRESULT on_bp_hit (ISimulatorBreakpointEvent* bps)
	{
		_running_info.reset();

		auto work = [this, bbps=com_ptr(bps)]() -> HRESULT
		{
			WI_ASSERT(_running);
			_running = false;
			for (uint32_t i = 0; i < _eventHandlers.size(); i++)
			{
				auto hr = _eventHandlers[i]->ProcessSimulatorEvent(bbps.get(), __uuidof(bbps)); LOG_IF_FAILED(hr);
			}

			return S_OK;
		};

		auto hr = PostWorkToMainThread (std::move(work)); RETURN_IF_FAILED(hr);
		
		_screen->generate_all();
		
		return S_OK;
	}

	// Called when simulation was running and the CPU reached an undefined instruction.
	HRESULT on_undefined_instruction (UINT64 address)
	{
		_running_info.reset();

		auto work = [this, address]() -> HRESULT
		{
			WI_ASSERT(_running);
			_running = false;

			for (uint32_t i = 0; i < _eventHandlers.size(); i++)
			{
				WI_ASSERT(false);
				/*
				com_ptr<ISimulatorEventHandler> h;
				auto hr = p->Resolve(&h); LOG_IF_FAILED(hr);
				if (SUCCEEDED(hr))
					h->OnUndefinedInstruction(address);
					*/
			}
	
			return S_OK;
		};

		auto hr = PostWorkToMainThread (std::move(work)); RETURN_IF_FAILED(hr);
		
		_screen->generate_all();
		
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

		virtual ULONG STDMETHODCALLTYPE AddRef() override
		{
			return ++_refCount;
		}

		virtual ULONG STDMETHODCALLTYPE Release() override
		{
			WI_ASSERT(_refCount);
			auto newRefCount = --_refCount;
			if (newRefCount == 0)
				delete this;
			return newRefCount;
		}
		#pragma endregion
	};

	HRESULT SendSimulateOneCompleteEvent()
	{
		WI_ASSERT(!_running);

		using SimulateOneEvent = SimulatorEvent<ISimulatorSimulateOneEvent>;
		auto event = com_ptr(new (std::nothrow) SimulateOneEvent()); RETURN_IF_NULL_ALLOC(event);

		for (uint32_t i = 0; i < _eventHandlers.size(); i++)
			_eventHandlers[i]->ProcessSimulatorEvent(event, __uuidof(event));

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

		//if (!_running)
		//	SendSimulateOneCompleteEvent();

		return S_OK;
	}

	HRESULT RunOnSimulatorThread (stdext::inplace_function<HRESULT()> fun)
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
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_RUNNING, _running);
		RETURN_HR_IF(E_BOUNDS, (uint32_t)address + size > 0x10000);
		for (uint32_t i = 0; i < size; i++)
			((uint8_t*)to)[i] = memoryBus.read(address + i);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE WriteMemoryBus (uint16_t address, uint16_t size, const void* from) override
	{
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_RUNNING, _running);
		RETURN_HR_IF(E_BOUNDS, (uint32_t)address + size > 0x10000);
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
				WI_ASSERT(_running_info.has_value());
				_running_info.reset();
				//_screen->generate_all();
				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		_running = false;
		for (uint32_t i = 0; i < _eventHandlers.size(); i++)
		{
			WI_ASSERT(false);
			/*
			com_ptr<ISimulatorEventHandler> h;
			auto hr = p->Resolve(&h); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
				h->OnSimulationStopped();
			*/
		}
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Resume (BOOL checkBreakpointsAtCurrentPC) override
	{
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_RUNNING, _running);

		auto hr = RunOnSimulatorThread([this, checkBreakpointsAtCurrentPC]
			{
				WI_ASSERT(!_running_info);

				auto start_time = _cpu->Time();
				LARGE_INTEGER perf_counter;
				QueryPerformanceCounter(&perf_counter);

				if (!checkBreakpointsAtCurrentPC)
					simulate_one_in_break_mode(); // and this should return a hr rather than calling events itself

				_running_info = running_info { .start_time = start_time, .start_time_perf_counter = perf_counter };
				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		WI_ASSERT(!_running);
		_running = true;
		using ResumeEvent = SimulatorEvent<ISimulatorResumeEvent>;
		auto event = com_ptr(new (std::nothrow) ResumeEvent()); RETURN_IF_NULL_ALLOC(event);

		for (uint32_t i = 0; i < _eventHandlers.size(); i++)
			_eventHandlers[i]->ProcessSimulatorEvent(event, __uuidof(event));

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Running_HR() override
	{
		return _running ? S_OK : S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE SimulateOne() override
	{
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_RUNNING, _running);

		auto hr = RunOnSimulatorThread ([this]
			{
				if (!_cpu->Halted())
					simulate_one_in_break_mode();
				else
				{
					do
					{
						simulate_one_in_break_mode();
					} while (_cpu->Halted());
				}

				_screen->generate_all();
				return S_OK;
			});
		RETURN_IF_FAILED(hr);
		
		return SendSimulateOneCompleteEvent();
	}
/*
	virtual HRESULT STDMETHODCALLTYPE AdviseSimulatorEvents (ISimulatorEventHandler* handler) override
	{
		RETURN_HR_IF (E_INVALIDARG, _eventHandlers.contains(handler));
		bool added = _eventHandlers.try_push_back(std::move(wr)); RETURN_HR_IF(E_OUTOFMEMORY, !added);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseSimulatorEvents (ISimulatorEventHandler* handler) override
	{
		RETURN_HR_IF (E_INVALIDARG, !_eventHandlers.contains(handler));
		_eventHandlers.erase(it);
		return S_OK;
	}
*/
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
	
	virtual HRESULT STDMETHODCALLTYPE LoadSnapshot (IStream* stream) override
	{
		auto hr = stream->Seek ( { .QuadPart = 0 }, STREAM_SEEK_SET, nullptr); RETURN_IF_FAILED(hr);

		STATSTG stat;
		hr = stream->Stat (&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		if (stat.cbSize.QuadPart != sizeof(snapshot_file_header) + 48 * 1024)
			return SIM_E_SNAPSHOT_FILE_WRONG_SIZE;

		hr = RunOnSimulatorThread ([this, stream]
			{
				HRESULT hr;

				_cpu->Reset();
				for (auto& d : _active_devices_)
					d->Reset();

				snapshot_file_header header;
				ULONG read;
				hr = stream->Read (&header, (ULONG)sizeof(header), &read); RETURN_IF_FAILED(hr);
				if (read != sizeof(header))
					return SIM_E_SNAPSHOT_FILE_WRONG_SIZE;

				uint8_t buffer[128];
				for (uint16_t i = 0; i < 48 * 1024; i += sizeof(buffer))
				{
					ULONG read;
					hr = stream->Read (buffer, sizeof(buffer), &read); RETURN_IF_FAILED(hr);
					if (read != sizeof(buffer))
						return SIM_E_SNAPSHOT_FILE_WRONG_SIZE;
					_ramDevice->WriteMemory(0x4000 + i, sizeof(buffer), buffer);
				}

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
				regs.iff1    = header.iff2;
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
					_screen->generate_all();

				return S_OK;
			});
		RETURN_IF_FAILED(hr);

		return S_OK;
	}
	
	virtual HRESULT STDMETHODCALLTYPE LoadBinary (IStream* stream, uint16_t address) override
	{
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_SIMULATION_RUNNING, _running);

		auto hr = stream->Seek ( { .QuadPart = 0 }, STREAM_SEEK_SET, nullptr); RETURN_IF_FAILED(hr);

		STATSTG stat;
		hr = stream->Stat (&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		if (stat.cbSize.HighPart)
			RETURN_HR(SIM_E_BINARY_FILE_TOO_LARGE);
		if (!stat.cbSize.LowPart)
			RETURN_HR(SIM_E_BINARY_FILE_EMPTY);
		if (stat.cbSize.LowPart > 0x10000)
			RETURN_HR(SIM_E_BINARY_FILE_TOO_LARGE);
		if (address + (uint16_t)stat.cbSize.LowPart <= address)
			RETURN_HR(SIM_E_BINARY_FILE_TOO_LARGE);

		uint8_t buffer[128];
		for (uint16_t i = 0; i < stat.cbSize.LowPart; i += sizeof(buffer))
		{
			ULONG read;
			hr = stream->Read (buffer, sizeof(buffer), &read); RETURN_IF_FAILED(hr);
			//_memory.write (address + (uint16_t)i, { buffer, buffer + read });
			_ramDevice->WriteMemory (address + i, read, buffer);
		}

		hr = this->RunOnSimulatorThread ([this, address]
			{
				_screen->generate_all();
				return S_OK;
			}); RETURN_IF_FAILED(hr);

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
	virtual HRESULT STDMETHODCALLTYPE AdviseDebugEvents (ISimulatorEventHandler* handler) override
	{
		auto it = _eventHandlers.find(handler); RETURN_HR_IF(E_INVALIDARG, it != _eventHandlers.end());
		bool pushed = _eventHandlers.try_push_back(handler); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseDebugEvents (ISimulatorEventHandler* handler) override
	{
		auto it = _eventHandlers.find(handler); RETURN_HR_IF(E_INVALIDARG, it == _eventHandlers.end());
		_eventHandlers.erase(it);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE AdviseScreenComplete (IScreenCompleteEventHandler* handler) override
	{
		bool pushed = _screenCompleteHandlers.try_push_back(handler); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnadviseScreenComplete (IScreenCompleteEventHandler* handler) override
	{
		auto it = _screenCompleteHandlers.find(handler); RETURN_HR_IF(E_INVALIDARG, it == _screenCompleteHandlers.end());
		_screenCompleteHandlers.erase(it);
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

	virtual HRESULT STDMETHODCALLTYPE GetScreenData  (BITMAPINFO** ppBitmapInfo, POINT* pBeamLocation) override
	{
		return _screen->GetScreenData(ppBitmapInfo, pBeamLocation);
	}

	virtual HRESULT STDMETHODCALLTYPE AddBreakpoint (BreakpointType type, bool physicalMemorySpace, UINT64 address, SIM_BP_COOKIE* pCookie) override
	{
		RETURN_HR_IF(E_NOTIMPL, physicalMemorySpace);
		RETURN_HR_IF(E_INVALIDARG, address > 0xFFFF);

		return RunOnSimulatorThread([this, type, address, pCookie]
			{
				return _cpu->AddBreakpoint(type, (uint16_t)address, pCookie);
			});
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveBreakpoint (SIM_BP_COOKIE cookie) override
	{
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
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_SIMULATION_RUNNING, _running);
		*pc = _cpu->GetPC();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetPC (uint16_t pc) override
	{
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_SIMULATION_RUNNING, _running);
		return _cpu->SetPC(pc);
	}

	virtual HRESULT STDMETHODCALLTYPE GetStackStartAddress (UINT16* stackStartAddress) override
	{
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_SIMULATION_RUNNING, _running);
		*stackStartAddress = _cpu->GetStackStartAddress();
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetRegisters (z80_register_set* buffer, uint32_t size) override
	{
		RETURN_HR_IF(SIM_E_NOT_SUPPORTED_WHILE_SIMULATION_RUNNING, _running);
		_cpu->GetZ80Registers(buffer);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE SetRegisters (const z80_register_set* buffer, uint32_t size) override
	{
		RETURN_HR(E_NOTIMPL);
	}
	#pragma endregion

	#pragma region IDeviceEventHandler
	virtual HRESULT STDMETHODCALLTYPE ProcessDeviceEvent (IDeviceEvent* event, REFIID riidEvent) override
	{
		return PostWorkToMainThread([this, event=com_ptr(event), riidEvent]
			{
				for (uint32_t i = 0; i < _eventHandlers.size(); i++)
					_eventHandlers[i]->ProcessSimulatorEvent(event.get(), riidEvent);
				return S_OK;
			});
	}
	#pragma endregion

	#pragma region IScreenDeviceCompleteEventHandler
	virtual void OnScreenComplete() override
	{
		// No error checking, not even logging, as in case of error it would probably freeze the app.
		PostWorkToMainThread([this]
			{
				for (uint32_t i = 0; i < _screenCompleteHandlers.size(); i++)
					_screenCompleteHandlers[i]->OnScreenComplete();
				return S_OK;
			});
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
