
#include "pch.h"
#include "Bus.h"

// TODO: split the RAM into two IMemoryDevice's, one for the video memory area, one for the rest.
// This will allow the CPU to work on non-video memory for a long time, without having to wait for
// the screen device to catch up.
// (Either that, or refactor the bus access...)

class HC_RAM : public IMemoryDevice
{
	Bus* _memory_bus;
	Bus* _io_bus;
	UINT64 _time = 0;
	bool _cpm = false; // false - responds to range 4000-FFFF; true - responds to range 0-DFFF
	uint8_t _data[0x10000]; // this one last

public:
	HRESULT InitInstance (Bus* memory_bus, Bus* io_bus)
	{
		_memory_bus = memory_bus;
		_io_bus = io_bus;
		for (size_t i = 0; i < sizeof(_data); i++)
			_data[i] = (uint8_t)rand();

		bool pushed = _memory_bus->read_responders.try_push_back({ this, &process_mem_read_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _memory_bus->write_responders.try_push_back({ this, &process_mem_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _io_bus->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	~HC_RAM()
	{
		_io_bus->write_responders.remove([this](auto& d) { return d.Device == this; });
		_memory_bus->write_responders.remove([this](auto& d) { return d.Device == this; });
		_memory_bus->read_responders.remove([this](auto& d) { return d.Device == this; });
	}

	virtual void STDMETHODCALLTYPE Reset() override
	{
		_time = 0;
		for (size_t i = 0; i < sizeof(_data); i++)
			_data[i] = (uint8_t)rand();
	}

	virtual UINT64 STDMETHODCALLTYPE Time() override { return _time; }

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return false; }

	virtual bool SimulateTo (UINT64 requested_time) override
	{
		// Only a bus write can change the state of this device. If there are still write-capable devices
		// whose timepoint is in the timespan we want to jump over (_time to requested_time), we can't jump.
		if (_memory_bus->writer_behind_of(requested_time))
			return false;
		if (_io_bus->writer_behind_of(requested_time))
			return false;
		_time = requested_time;
		return true;
	}

	static uint8_t process_mem_read_request (IDevice* d, uint16_t address)
	{
		auto* ram = static_cast<HC_RAM*>(d);
		if (!ram->_cpm)
		{
			if (address >= 0x4000)
				return ram->_data[address];
			return 0xFF;
		}
		else
		{
			if (address < 0xE000)
				return ram->_data[address];
			return 0xFF;
		}
	}

	static void process_mem_write_request (IDevice* d, uint16_t address, uint8_t value)
	{
		auto* ram = static_cast<HC_RAM*>(d);
		ram->_data[address] = value;
	}

	static void process_io_write_request (IDevice* d, uint16_t address, uint8_t value)
	{
		auto* ram = static_cast<HC_RAM*>(d);
		if ((address & 0x81) == 0)
			ram->_cpm = value & 2;
	}

	#pragma region IMemoryDevice
	virtual HRESULT GetBounds (DWORD* from, DWORD* to) override
	{
		if (!_cpm)
		{
			*from =  0x4000;
			*to   = 0x10000;
			return S_OK;
		}
		else
		{
			*from = 0;
			*to = 0xE000;
			return S_OK;
		}
	}

	virtual HRESULT ReadMemory (uint32_t address, uint32_t size, void* dest) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT WriteMemory (uint32_t address, uint32_t size, const void* bytes) override
	{
		RETURN_HR_IF(E_BOUNDS, address >= sizeof(_data));
		RETURN_HR_IF(E_BOUNDS, address + size > sizeof(_data));
		memcpy (&_data[address], bytes, size);
		return S_OK;
	}
	#pragma endregion
};

HRESULT STDMETHODCALLTYPE MakeHC91RAM (Bus* memory_bus, Bus* io_bus, wistd::unique_ptr<IMemoryDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<HC_RAM>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(memory_bus, io_bus); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
