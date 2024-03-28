
#include "pch.h"
#include "Bus.h"
#include "shared/com.h"

class HC_ROM : public IMemoryDevice
{
	Bus* _memory_bus;
	Bus* _io_bus;
	UINT64 _time = 0;
	bool _cpmSrc = false; // false - reading from index 0 of _data; true - reading from index 24K of _data.
	bool _cpmDst = false; // false - responding to bus address range 0-3FFF; true - responding to bus address range E000-FFFF
	wil::unique_hlocal_string _folder;
	uint8_t _data[0x8000]; // this one last
	//vector_nothrow<com_ptr<IWeakRef>> _callbacks;

public:
	HRESULT InitInstance (Bus* memory_bus, Bus* io_bus, const wchar_t* folder, const wchar_t* BinaryFilename)
	{
		_memory_bus = memory_bus;
		_io_bus = io_bus;
		_folder = wil::make_hlocal_string_nothrow(folder); RETURN_IF_NULL_ALLOC(_folder);
		auto binaryPath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(binaryPath);
		auto hr = PathCchCombine (binaryPath.get(), MAX_PATH + 1, _folder.get(), BinaryFilename); RETURN_IF_FAILED(hr);
		com_ptr<IStream> romStream;
		hr = SHCreateStreamOnFileEx (binaryPath.get(), STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &romStream); RETURN_IF_FAILED(hr);
		STATSTG stat;
		hr = romStream->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
		if (stat.cbSize.HighPart)
			RETURN_WIN32(ERROR_FILE_TOO_LARGE);
		if (stat.cbSize.LowPart < 0x4000)
			RETURN_WIN32(ERROR_FILE_CORRUPT);

		auto buffer = wil::make_unique_hlocal_nothrow<uint8_t[]>(stat.cbSize.LowPart);
		ULONG bytes_read;
		hr = romStream->Read(buffer.get(), stat.cbSize.LowPart, &bytes_read); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, bytes_read != stat.cbSize.LowPart);

		memcpy_s(_data, sizeof(_data), buffer.get(), stat.cbSize.LowPart);
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

	virtual void STDMETHODCALLTYPE Reset() override
	{
		_time = 0;
	}

	virtual UINT64 STDMETHODCALLTYPE Time() override { return _time; }

	virtual void SkipTime (uint64_t offset) override
	{
		_time += offset;
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return FALSE; }

	virtual bool SimulateTo (UINT64 requested_time) override
	{
		// Only a bus write can change the state of this device. If there are still write-capable devices
		// whose timepoint is in the timespan we want to jump over (_time to requested_time), we can't jump.
		if (_io_bus->writer_behind_of(requested_time))
			return false;
		_time = requested_time;
		return true;
	}

	static uint8_t process_mem_read_request (IDevice* d, uint16_t address)
	{
		auto* rom = static_cast<HC_ROM*>(d);
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

	static void process_io_write_request (IDevice* d, uint16_t address, uint8_t value)
	{
		auto* rom = static_cast<HC_ROM*>(d);
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

	#pragma region IMemoryDevice
	virtual HRESULT ReadMemory (uint32_t address, uint32_t size, void* dest) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT WriteMemory (uint32_t internalAddress, uint32_t size, const void* bytes) override
	{
		RETURN_HR(E_NOTIMPL);
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

HRESULT STDMETHODCALLTYPE MakeHC91ROM (Bus* memory_bus, Bus* io_bus, const wchar_t* folder, const wchar_t* BinaryFilename, wistd::unique_ptr<IMemoryDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<HC_ROM>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(memory_bus, io_bus, folder, BinaryFilename); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
