
#include "pch.h"
#include "SimulatorInternal.h"
#include "shared/com.h"

class HC_ROM : public IMemoryDevice
{
	Bus* _memory_bus;
	Bus* _io_bus;
	SpectrumVariant _spectrumVariant;
	uint8_t rom_bank = 0;
	bool locked = false;
	bool _cpmSrc = false; // false - reading from index 0 of _data; true - reading from index 24K of _data.
	bool _cpmDst = false; // false - responding to bus address range 0-3FFF; true - responding to bus address range E000-FFFF
	wil::unique_process_heap_string _binaryFilenames[SpectrumVariantCount];
	uint8_t _data[0x8000]; // this one last
	//vector_nothrow<com_ptr<IWeakRef>> _callbacks;

public:
	HRESULT InitInstance (Bus* memory_bus, Bus* io_bus, LPCWSTR const* romFilenames, size_t romFilenameCount, SpectrumVariant variant)
	{
		RETURN_HR_IF(E_POINTER, !romFilenames);
		RETURN_HR_IF(E_INVALIDARG, romFilenameCount != SpectrumVariantCount);
		RETURN_HR_IF(E_INVALIDARG, variant != SpectrumVariant48K && variant != SpectrumVariant128);
		_memory_bus = memory_bus;
		_io_bus = io_bus;
		for (size_t i = 0; i < romFilenameCount; i++)
		{
			if (!romFilenames[i])
				continue;
			_binaryFilenames[i] = wil::make_process_heap_string_nothrow(romFilenames[i]); RETURN_IF_NULL_ALLOC(_binaryFilenames[i]);
		}
		auto hr = LoadROM(variant); RETURN_IF_FAILED(hr);

		bool pushed = _memory_bus->read_responders.try_push_back({ this, &process_mem_read_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _io_bus->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	~HC_ROM()
	{
		//WI_ASSERT(_callbacks.empty());
		_memory_bus->read_responders.remove([this](auto& d) { return d.Device == this; });
		_io_bus->write_responders.remove([this](auto& d) { return d.Device == this; });
	}

	virtual HRESULT STDMETHODCALLTYPE Reset(SpectrumVariant variant) override
	{
		if (_spectrumVariant != variant)
		{
			auto hr = LoadROM(variant); RETURN_IF_FAILED(hr);
			_spectrumVariant = variant;
		}

		_time = 0;
		_cpmSrc = false;
		_cpmDst = false;
		rom_bank = 0;
		locked = false;
		return S_OK;
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return FALSE; }

	virtual bool SimulateDeviceTo (UINT64 requested_time) override
	{
		// Comment from HC_RAM applies here as well.
		_time = requested_time;
		return true;
	}

	static uint8_t process_mem_read_request (IDevice* d, WORD address)
	{
		auto* rom = static_cast<HC_ROM*>(d);
		if (rom->_spectrumVariant == SpectrumVariant128)
			return address < 0x4000 ? rom->_data[rom->rom_bank * 0x4000 + address] : 0xFF;
		uint32_t readOffset = (rom->_cpmSrc ? 0x4000 : 0);
		if (!rom->_cpmDst)
		{
			if (address < 0x4000)
				return rom->_data[address + readOffset];
			return 0xFF;
		}
		else
		{
			if (address >= 0xE000)
				return rom->_data[address - 0xC000 + readOffset];
			return 0xFF;
		}
	}

	static void process_io_write_request (IDevice* d, WORD address, uint8_t value)
	{
		auto* rom = static_cast<HC_ROM*>(d);
		if (rom->_spectrumVariant == SpectrumVariant128)
		{
			if ((address & 0x8002) == 0 && !rom->locked)
			{
				rom->rom_bank = (value >> 4) & 1;
				if (value & 0x20)
					rom->locked = true;
			}
			return;
		}

		if ((address & 0x81) == 0)
		{
			rom->_cpmSrc = value & 1;

			bool newCpmDst = value & 2;
			if (rom->_cpmDst != newCpmDst)
			{
				uint32_t address = rom->_cpmDst ? 0xE000 : 0;
				uint32_t size = rom->_cpmDst ? 0x2000 : 0x4000;
				WI_ASSERT(false);
				/*
				for (auto& c : rom->_callbacks)
				{
					com_ptr<IBusAddressRangeChangeHandler> h;
					auto hr = c->Resolve(&h); LOG_IF_FAILED(hr);
					if (SUCCEEDED(hr))
					{
						hr = h->OnBusAddressRangeChanging(rom, address, size); LOG_IF_FAILED(hr);
					}
				}

				rom->_cpmDst = newCpmDst;

				address = rom->_cpmDst ? 0xE000 : 0;
				size = rom->_cpmDst ? 0x2000 : 0x4000;
				for (auto& c : rom->_callbacks)
				{
					com_ptr<IBusAddressRangeChangeHandler> h;
					auto hr = c->Resolve(&h); LOG_IF_FAILED(hr);
					if (SUCCEEDED(hr))
					{
						hr = h->OnBusAddressRangeChanged(rom, address, size); LOG_IF_FAILED(hr);
					}
				}
				*/
			}
		}
	}

	HRESULT LoadROM (SpectrumVariant variant)
	{
		RETURN_HR_IF(E_INVALIDARG, variant != SpectrumVariant48K && variant != SpectrumVariant128);
		size_t variantIndex = static_cast<size_t>(variant);
		RETURN_HR_IF(E_INVALIDARG, !_binaryFilenames[variantIndex] || !_binaryFilenames[variantIndex].get()[0]);
		LPCWSTR binaryFilename = _binaryFilenames[variantIndex].get();

		com_ptr<IStream> romStream;
		auto hr = SHCreateStreamOnFileEx(binaryFilename, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &romStream); RETURN_IF_FAILED(hr);
		STATSTG stat;
		hr = romStream->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		if (stat.cbSize.HighPart)
			RETURN_WIN32(ERROR_FILE_TOO_LARGE);
		if (variant == SpectrumVariant48K && stat.cbSize.LowPart != 0x4000)
			RETURN_WIN32(ERROR_FILE_CORRUPT);
		if (variant == SpectrumVariant128 && stat.cbSize.LowPart != 0x8000)
			RETURN_WIN32(ERROR_FILE_CORRUPT);

		auto data = wil::make_unique_hlocal_nothrow<uint8_t[]>(0x8000); RETURN_IF_NULL_ALLOC(data);
		memset(data.get(), 0, 0x8000);
		ULONG bytes_read;
		hr = romStream->Read(data.get(), stat.cbSize.LowPart, &bytes_read); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, bytes_read != stat.cbSize.LowPart);

		memcpy(_data, data.get(), sizeof(_data));
		_spectrumVariant = variant;
		return S_OK;
	}

	#pragma region IMemoryDevice
	virtual HRESULT GetBounds (DWORD* from, DWORD* to) override
	{
		if (!_cpmDst)
		{
			*from = 0;
			*to = 0x4000;
			return S_OK;
		}
		else
		{
			*from = 0x0'E000;
			*to   = 0x1'0000;
			return S_OK;
		}
	}

	virtual HRESULT ReadMemory (uint32_t address, uint32_t size, void* dest) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT WriteMemory (uint32_t internalAddress, uint32_t size, const void* bytes) override
	{
		return E_NOTIMPL;
	}
	/*
	virtual HRESULT AdviseBusAddressRangeChange (IBusAddressRangeChangeHandler* handler) override
	{
		com_ptr<IWeakRef> wr;
		auto hr = handler->GetWeakReference(&wr); RETURN_IF_FAILED(hr);
		auto it = _callbacks.find(wr);
		RETURN_HR_IF(E_INVALIDARG, it != _callbacks.end());
		bool pushed = _callbacks.try_push_back(std::move(wr)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	virtual HRESULT UnadviseBusAddressRangeChange (IBusAddressRangeChangeHandler* handler) override
	{
		com_ptr<IWeakRef> wr;
		auto hr = handler->GetWeakReference(&wr); RETURN_IF_FAILED(hr);
		auto it = _callbacks.find(wr);
		RETURN_HR_IF(E_INVALIDARG, it == _callbacks.end());
		_callbacks.erase(it);
		return S_OK;
	}
	*/
	#pragma endregion
};

HRESULT STDMETHODCALLTYPE MakeHC91ROM (Bus* memory_bus, Bus* io_bus, LPCWSTR const* romFilenames, size_t romFilenameCount, SpectrumVariant variant, wistd::unique_ptr<IMemoryDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<HC_ROM>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(memory_bus, io_bus, romFilenames, romFilenameCount, variant); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
