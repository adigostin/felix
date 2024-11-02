
#pragma once
#include "../Simulator.h"
#include "shared/vector_nothrow.h"

struct DECLSPEC_NOVTABLE IDevice
{
	virtual ~IDevice() = default;
	virtual void Reset() = 0;
	virtual uint64_t Time() = 0;
	virtual BOOL NeedSyncWithRealTime (UINT64* sync_time) = 0;

	// Returns true if it simulated something (thus advancing the device's time),
	// even if it couldn't simulate to the "requested_time".
	// Returs false if it couldn't simulate anything (thus the device's time didn't change).
	virtual bool SimulateTo (UINT64 requested_time) = 0;
};

struct BreakpointsHit
{
	//BreakpointType type;
	uint16_t address;
	SIM_BP_COOKIE bps[8];
	uint32_t size;
};

// TODO: ICPU shoulnd't derive from IDevice, too little in common.
struct DECLSPEC_NOVTABLE ICPU : IDevice
{
//	virtual HRESULT STDMETHODCALLTYPE GetRegisters (IRegisterGroup** ppRegs) = 0;
	virtual UINT16 GetStackStartAddress() const = 0;
	virtual BOOL STDMETHODCALLTYPE Halted() = 0;
	virtual UINT16 GetPC() const = 0;
	virtual HRESULT SetPC(UINT16 pc) = 0;

	// If the caller wants this function to check for breakpoints, it passes into "bps" a pointer
	// to a caller-allocated structure for receiving information about any breakpoint hit (zero or more).
	// In this case the implementation checks for code breakpoints (BreakpointType::Code)
	// before attempting to execute the instruction at PC, or for data breakpoints (BreakpointType::Data)
	// during the execution of the instruction.
	// If the caller wants this function to ignore breakpoints, is passes NULL for "bps". The caller
	// might want to ignore breakpoints, for example, when execution is stopped and the user does
	// a Step or Go command.
	// 
	// Returns:
	//  - "true" if something was simulated and the CPU's time advanced its Time() by some cycles
	//    (maybe because an instruction was executed, maybe because an interrupt kicked in and jumping to 
	//    the interrupt routine took some cycles, or maybe because the CPU was HALTed and remained so)
	//  - "false" - if nothing was executed and the CPU remained at the same Time()
	//    (maybe waiting for other devices to catch up, maybe code breakpoint).
	virtual bool SimulateOne (BreakpointsHit* bps) = 0;

	virtual HRESULT AddBreakpoint (BreakpointType type, uint16_t address, SIM_BP_COOKIE* pCookie) = 0;
	virtual HRESULT RemoveBreakpoint (SIM_BP_COOKIE cookie) = 0;
	virtual BOOL HasBreakpoints() = 0;
};

/*
struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{A97526D1-239A-4322-ABA8-33CF2690A732}") IBusAddressRangeChangeHandler : IWeakRefSource
{
	virtual HRESULT OnBusAddressRangeChanging (IMemoryDevice* device, uint32_t address, uint32_t size) = 0;
	virtual HRESULT OnBusAddressRangeChanged  (IMemoryDevice* device, uint32_t address, uint32_t size) = 0;
};
*/
struct DECLSPEC_NOVTABLE IMemoryDevice : IDevice
{
	virtual HRESULT GetBounds (DWORD* from, DWORD* to) = 0;
	virtual HRESULT ReadMemory (uint32_t busAddress, uint32_t size, void* dest) = 0;
	virtual HRESULT WriteMemory (uint32_t internalAddress, uint32_t size, const void* bytes) = 0;
};

struct IInterruptingDevice
{
	virtual IDevice* as_device() = 0;
	virtual uint8_t STDMETHODCALLTYPE irq_priority() = 0;
	virtual bool irq_pending (uint64_t& irq_time, uint8_t& irq_address) const = 0;
	virtual void acknowledge_irq() = 0;
};

static constexpr UINT64 milliseconds_to_ticks (DWORD milliseconds)
{
	return milliseconds * 3500;
}

static constexpr DWORD ticks_to_milliseconds (UINT64 ticks)
{
	return (DWORD)(ticks / 3500);
}

static constexpr UINT64 ticks_to_hundreds_of_nanoseconds (UINT64 ticks)
{
	WI_ASSERT(ticks < 400'000);
	return ticks * 10000 / 3500;
}

// ReadResponder and WriteResponder are implemented as pairs of device + handler
// (rather than as interface classes with virtual functions) because some devices
// are present on both buses (memory and IO) and we'd get a conflict when overriding
// once for the memory and once for the IO. We also want to stay away from dynamic_cast.

struct ReadResponder
{
	IDevice* Device;
	uint8_t(*ProcessReadRequest)(IDevice* device, uint16_t address);
};

struct WriteResponder
{
	IDevice* Device;
	void(*ProcessWriteRequest)(IDevice* device, uint16_t address, uint8_t value);
};

struct DECLSPEC_NOVTABLE Bus
{
	// Returns all devices that initiate read requests at the given address.
	//virtual std::span<IDevice* const> read_requesters (uint16_t address) const = 0;

	// All devices that initiate write requests.
	vector_nothrow<IDevice*> write_requesters;

	// All devices that respond to read requests.
	vector_nothrow<ReadResponder> read_responders;

	// All devices that respond to write requests.
	vector_nothrow<WriteResponder> write_responders;
	
	// Reads something from a bus while ignoring any time difference between devices.
	// Useful for debugging only. Simulation-related code should always use try_read_request.
	uint8_t read (uint16_t address)
	{
		uint8_t val = 0xFF;
		for (auto& d : read_responders)
			val &= d.ProcessReadRequest(d.Device, address);
		return val;
	}

	// Comments from read() apply here too.
	void write (uint16_t address, uint8_t value)
	{
		for (auto& d : write_responders)
			d.ProcessWriteRequest(d.Device, address, value);
	}

	void write (uint16_t address, std::initializer_list<uint8_t> values)
	{
		for (size_t i = 0; i < values.size(); i++)
		{
			uint8_t value = *(values.begin() + i);
			for (auto& d : write_responders)
				d.ProcessWriteRequest(d.Device, address + (uint16_t)i, value);
		}
	}

	uint16_t read_uint16 (uint16_t address)
	{
		uint16_t val = 0xFFFF;
		for (auto& d : read_responders)
		{
			uint8_t l = d.ProcessReadRequest(d.Device, address);
			uint8_t h = d.ProcessReadRequest(d.Device, address + 1);
			val &= (l | (h << 8));
		}
		return val;
	}

	void write_uint16 (uint16_t address, uint16_t value)
	{
		uint8_t l = (uint8_t)value;
		uint8_t h = (uint8_t)(value >> 8);
		for (auto& d : write_responders)
		{
			d.ProcessWriteRequest(d.Device, address, l);
			d.ProcessWriteRequest(d.Device, address + 1, h);
		}
	}

	// Tries to performs a read request on the bus.
	// The function checks to see if the devices that respond to read requests at the specified
	// address have simulated themselves at least up to the requested time.
	// If no, it returns false. If yes, it sets the 'value' variable and returns true.
	bool try_read_request (uint16_t address, uint8_t& value, UINT64 requested_time)
	{
		uint8_t temp = 0xFF;
		for (auto& d : read_responders)
		{
			if (d.Device->Time() < requested_time)
			{
				// A read responder is at an earlier time point. Let's try to simulate it
				// up to the requested time. In the vast majority of cases our caller is
				// the CPU and the read responder is a memory, so simulation is possible
				// and is faster when done here.
				// 
				// We don't attempt here anything fancy such as simulating repeatedly
				// in case the device is blocked on some other device that's also blocked.
				// This kind of scenario is taken care of in the simulator class.

				bool advanced = d.Device->SimulateTo(requested_time);
				if (!advanced || (d.Device->Time() < requested_time))
					return false;
			}

			temp &= d.ProcessReadRequest(d.Device, address);
		}

		value = temp;
		return true;
	}

	// Comment from try_read_request() applies here as well.
	bool try_write_request (uint16_t address, uint8_t value, UINT64 requested_time)
	{
		for (auto& d : write_responders)
		{
			if (d.Device->Time() < requested_time)
			{
				// Comment from try_read_request applies here too.
				bool advanced = d.Device->SimulateTo(requested_time);
				if (!advanced || (d.Device->Time() < requested_time))
					return false;
			}
		}

		// We want this operation to write either to all responders, or to none of them.
		// If we got past the above "for", we know we can write to all of them.

		for (auto& d : write_responders)
			d.ProcessWriteRequest(d.Device, address, value);

		return true;
	}

	bool try_read_request (uint16_t address, uint16_t& value, UINT64 requested_time)
	{
		if (!this->try_read_request (address, *(uint8_t*)&value, requested_time))
			return false;
		if (!this->try_read_request (address + 1, *((uint8_t*)&value + 1), requested_time))
			return false;
		return true;
	}

	bool try_write_request (uint16_t address, uint16_t value, UINT64 requested_time)
	{
		if (!this->try_write_request (address, *(uint8_t*)&value, requested_time))
			return false;

		if (!this->try_write_request (address + 1, *((uint8_t*)&value + 1), requested_time))
		{
			WI_ASSERT(false); // TODO: roll back the first write
			return false;
		}

		return true;
	}

	IDevice* writer_behind_of (UINT64 time)
	{
		uint32_t count = this->write_requesters.size();
		for (uint32_t i = 0; i < count; i++)
		{
			IDevice* d = this->write_requesters[i];
			if (d->Time() < time)
				return d;
		}

		return nullptr;
	}
};

struct __declspec(novtable) irq_line_i
{
	vector_nothrow<IInterruptingDevice*> interrupting_devices;

	// Returns true if all devices have caught up with the CPU. Returns false otherwise (some device lags behind).
	//
	// When returning true, "interrupted" is set to true if a device requested an interrupt
	// (in which case the functin also sets "irq_address"), and to false if not.
	//
	bool try_poll_irq_at_time_point (bool& interrupted, uint8_t& irq_address, UINT64 cpu_time)
	{
		for (auto d : interrupting_devices)
		{
			if (d->as_device()->Time() < cpu_time)
				return false;

			UINT64 irq_time;
			if (d->irq_pending(irq_time, irq_address) && ((int32_t)(irq_time - cpu_time) <= 0))
			{
				// TODO: look for the earliest irqs, and if there are more of them at the same time, look at irq priorities.
				d->acknowledge_irq();
				interrupted = true;
				return true;
			}
		}

		// All interrupting devices have caught up with the CPU and none has a pending interrupt at that time point.
		interrupted = false;
		return true;
	}
};

struct IScreenDeviceCompleteEventHandler
{
	// Function called by the screen device every time it completes drawing a screen.
	virtual void OnScreenDeviceComplete() = 0;
};

struct IScreenDevice : IDevice//, IInterruptingDevice
{
	//virtual zx_spectrum_ula_regs regs() const = 0;

	// crt = FALSE   - creates a snapshot of the video memory, including newly drawn parts that the CRT beam hasn't reached yet to put on screen.
	// crt = TRUE    - creates a snapshot of the CRT screen.
	// ppBuffer      - must be freed by the caller when no longer needed, using CoTaskMemFree.
	// pBeamLocation - caller can pass NULL if it doesn't need this information.
	virtual HRESULT CopyBuffer (BOOL crt, OUT BITMAPINFO** ppBuffer, OUT POINT* pBeamLocation) = 0;

	// Generates the entire CRT screen from video memory.
	virtual HRESULT GenerateScreen() = 0;
};
HRESULT STDMETHODCALLTYPE MakeScreenDevice (Bus* memory, Bus* io, irq_line_i* irq, IScreenDeviceCompleteEventHandler* eh, wistd::unique_ptr<IScreenDevice>* ppDevice);

struct IKeyboardDevice : IDevice
{
	virtual HRESULT STDMETHODCALLTYPE ProcessKeyDown (uint32_t vkey, uint32_t modifiers) = 0;
	virtual HRESULT STDMETHODCALLTYPE ProcessKeyUp (uint32_t vkey, uint32_t modifiers) = 0;
};
HRESULT STDMETHODCALLTYPE MakeKeyboardDevice (Bus* io_bus, wistd::unique_ptr<IKeyboardDevice>* ppDevice);
