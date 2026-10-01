
#include "pch.h"
#include "SimulatorInternal.h"

// TODO: split the RAM into two IMemoryDevice's, one for the video memory area, one for the rest.
// This will allow the CPU to work on non-video memory for a long time, without having to wait for
// the screen device to catch up.
// (Either that, or refactor the bus access...)

class HC_RAM : public IMemoryDevice
{
	Bus* _memory_bus;
	Bus* _io_bus;
	SpectrumVariant _spectrumVariant;
	uint8_t ram_bank = 0; // The RAM bank that responds to the 0xC000-0xFFFF range in Spectrum 128K mode.
	bool locked = false;
	bool _cpm = false; // false - responds to range 4000-FFFF; true - responds to range 0-DFFF
	uint8_t _data[0x20000]; // this one last

public:
	HRESULT InitInstance (Bus* memory_bus, Bus* io_bus, SpectrumVariant variant)
	{
		_memory_bus = memory_bus;
		_io_bus = io_bus;
		_spectrumVariant = variant;
		for (size_t i = 0; i < sizeof(_data); i++)
			_data[i] = (uint8_t)rand();

		bool pushed = _memory_bus->read_responders.try_push_back({ this, &process_mem_read_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _memory_bus->physical_read_responders.try_push_back({ this, &process_physical_mem_read_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _memory_bus->write_responders.try_push_back({ this, &process_mem_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _io_bus->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	~HC_RAM()
	{
		_io_bus->write_responders.remove([this](auto& d) { return d.Device == this; });
		_memory_bus->write_responders.remove([this](auto& d) { return d.Device == this; });
		_memory_bus->physical_read_responders.remove([this](auto& d) { return d.Device == this; });
		_memory_bus->read_responders.remove([this](auto& d) { return d.Device == this; });
	}

	virtual HRESULT STDMETHODCALLTYPE Reset(SpectrumVariant variant) override
	{
		_spectrumVariant = variant;
		_time = 0;
		ram_bank = 0;
		locked = false;
		for (size_t i = 0; i < sizeof(_data); i++)
			_data[i] = (uint8_t)rand();
		return S_OK;
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return false; }

	virtual bool SimulateDeviceTo (UINT64 requested_time) override
	{
		// To bring the device to the requested time, we must first simulate any other device
		// that might change our state, meaning every device that initiates writes to either bus.
		// We "know" some implementation details of the simulator: (1) currently the only thing that writes
		// to these buses is the CPU, and (2) the simulator keeps the CPU ahead of all devices;
		// thus there's no need to do anything here other than advance to the requested time.
		_time = requested_time;
		return true;
	}

	static uint8_t process_mem_read_request (IDevice* d, WORD address)
	{
		auto* ram = static_cast<HC_RAM*>(d);
		if (ram->_spectrumVariant == SpectrumVariant128)
		{
			if (address < 0x4000)
				return 0xFF;

			uint8_t bank;
			if (address < 0x8000)
				bank = 5;
			else if (address < 0xC000)
				bank = 2;
			else
				bank = ram->ram_bank;
			return ram->_data[bank * 0x4000 + (address & 0x3FFF)];
		}
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

	static uint8_t process_physical_mem_read_request (IDevice* d, DWORD address)
	{
		auto* ram = static_cast<HC_RAM*>(d);
		return address < sizeof(ram->_data) ? ram->_data[address] : 0xFF;
	}

	static void process_mem_write_request (IDevice* d, WORD address, uint8_t value)
	{
		auto* ram = static_cast<HC_RAM*>(d);
		if (ram->_spectrumVariant == SpectrumVariant128)
		{
			uint8_t bank;
			if (address < 0x4000)
				return;
			if (address < 0x8000)
				bank = 5;
			else if (address < 0xC000)
				bank = 2;
			else
				bank = ram->ram_bank;
			ram->_data[bank * 0x4000 + (address & 0x3FFF)] = value;
			return;
		}
		ram->_data[address] = value;
	}

	static void process_io_write_request (IDevice* d, WORD address, uint8_t value)
	{
		auto* ram = static_cast<HC_RAM*>(d);
		if (ram->_spectrumVariant == SpectrumVariant128)
		{
			if ((address & 0x8002) == 0 && !ram->locked)
			{
				ram->ram_bank = value & 7;
				if (value & 0x20)
					ram->locked = true;
			}
			return;
		}
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
		return E_NOTIMPL;
	}

	virtual HRESULT WriteMemory (uint32_t address, uint32_t size, const void* bytes) override
	{
		if (address >= 0x10000)
			return E_BOUNDS;
		if (address + size > 0x10000)
			return E_BOUNDS;
		if (_spectrumVariant != SpectrumVariant128)
			memcpy (&_data[address], bytes, size);
		else
		{
			for (uint32_t i = 0; i < size; i++)
			{
				uint32_t logical = address + i;
				uint8_t bank = logical < 0x8000 ? 5 : (logical < 0xC000 ? 2 : ram_bank);
				_data[bank * 0x4000 + (logical & 0x3FFF)] = ((const uint8_t*)bytes)[i];
			}
		}
		return S_OK;
	}
	#pragma endregion
};

HRESULT STDMETHODCALLTYPE MakeHC91RAM (Bus* memory_bus, Bus* io_bus, SpectrumVariant variant, wistd::unique_ptr<IMemoryDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<HC_RAM>(); RETURN_IF_NULL_ALLOC(d);
    auto hr = d->InitInstance(memory_bus, io_bus, variant); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
