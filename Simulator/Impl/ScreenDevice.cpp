
#include "pch.h"
#include "SimulatorInternal.h"
#include <optional>

// http://www.zxdesign.info/vidparam.shtml

// https://docs.microsoft.com/en-us/windows/win32/direct2d/supported-pixel-formats-and-alpha-modes#specifying-a-pixel-format-for-an-id2d1bitmap
// We want DXGI_FORMAT_B8G8R8A8_UNORM. So 32 bits per pixel, or an uint32_t.

static constexpr uint32_t border_size_top = 48;
static constexpr uint32_t border_size_bottom = 56;
static constexpr uint32_t border_size_left_px = 48;
static constexpr uint32_t border_size_right_px = 48;
static constexpr uint32_t border_size_left_ticks = 24;
static constexpr uint32_t border_size_right_ticks = 24;
static constexpr uint32_t vsync_row_count = 16;
static constexpr uint32_t hsync_col_count = 48; // in clock cycles
// TODO: Select variant-specific timing here and in real-time/audio conversion; 128K uses 228 T-states per line and 311 lines per frame.
static constexpr uint32_t ticks_per_row = hsync_col_count + border_size_left_ticks + 128 + border_size_right_ticks;
static constexpr uint32_t rows_per_frame = vsync_row_count + border_size_top + 192 + border_size_bottom;
static constexpr uint32_t irq_offset_from_frame_start = hsync_col_count + border_size_left_ticks;

static constexpr uint32_t screen_width  = border_size_left_px + 256 + border_size_right_px;
static constexpr uint32_t screen_height = border_size_top  + 192 + border_size_bottom;

static constexpr UINT64 max_time_offset = milliseconds_to_ticks(1000);

static constexpr UINT32 ScreenBufferSize = sizeof(BITMAPINFOHEADER) + screen_width * screen_height * 4;

class ScreenDeviceImpl : public IScreenDevice
{
	Bus* memory;
	Bus* io;
	IrqLine* irq;
	SpectrumVariant _variant;
	uint8_t screen_bank = 5;
	bool locked = false;
	std::optional<UINT64> _pending_irq_time;
	uint8_t _border;
	uint32_t frameNumber = 0;
	uint32_t row = 0;
	uint32_t col = 0; // column in clock cycles (one unit equals two pixels)
	uint16_t src_pixel_data_ptr = 0;
	uint16_t src_pixel_attr_ptr = 0;
	wil::unique_process_heap_ptr<BITMAPINFO> _screenData;
	IScreenDeviceCompleteEventHandler* _screenCompleteHandler;

public:
	HRESULT InitInstance (Bus* memory, Bus* io, IrqLine* irq, SpectrumVariant variant, IScreenDeviceCompleteEventHandler* screenCompleteEventHandler)
	{
		this->memory = memory;
		this->io = io;
		this->irq = irq;
		_variant = variant;
		_screenCompleteHandler = screenCompleteEventHandler;

		bool pushed = io->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = irq->interrupting_devices.try_push_back(this); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

		_screenData.reset((BITMAPINFO*)HeapAlloc(GetProcessHeap(), 0, ScreenBufferSize)); RETURN_IF_NULL_ALLOC(_screenData);
		InitBitmapInfoHeader(_screenData.get(), TRUE);

		return S_OK;
	}

	~ScreenDeviceImpl()
	{
		irq->interrupting_devices.remove(static_cast<IDevice*>(this));
		io->write_responders.remove([this](auto& w) { return w.Device == this; });
	}

	static void InitBitmapInfoHeader (BITMAPINFO* bi, BOOL includeBorder)
	{
		bi->bmiHeader.biSize = sizeof(BITMAPINFO);
		bi->bmiHeader.biWidth = includeBorder ? screen_width : 256;
		bi->bmiHeader.biHeight = includeBorder ? screen_height : 192;
		bi->bmiHeader.biPlanes = 1;
		bi->bmiHeader.biBitCount = 32;
		bi->bmiHeader.biCompression = BI_RGB;
		bi->bmiHeader.biSizeImage = 0;
		bi->bmiHeader.biXPelsPerMeter = 2835;
		bi->bmiHeader.biYPelsPerMeter = 2835;
		bi->bmiHeader.biClrUsed = 0;
		bi->bmiHeader.biClrImportant = 0;
	}

	static void process_io_write_request (IDevice* d, WORD address, uint8_t value)
	{
		auto* s = static_cast<ScreenDeviceImpl*>(d);
		if ((address & 1) == 0)
			s->_border = value & 7;
		if (s->_variant == SpectrumVariant128 && (address & 0x8002) == 0 && !s->locked)
		{
			s->screen_bank = (value & 8) ? 7 : 5;
			if (value & 0x20)
				s->locked = true;
		}
	}

	virtual HRESULT STDMETHODCALLTYPE Reset(SpectrumVariant variant) override
	{
		_variant = variant;
		_time = 0;
		_pending_irq_time.reset();
		screen_bank = 5;
		locked = false;
		frameNumber = 0;
		row = 0;
		col = 0;
		src_pixel_data_ptr = 0;
		src_pixel_attr_ptr = 0;
		return S_OK;
	}

	DWORD physical_memory_address (uint16_t address) const
	{
		if (_variant == SpectrumVariant128)
			return screen_bank * 0x4000 + (address & 0x3FFF);
		return address;
	}

	static constexpr uint32_t low_brightness_colors[] =
		{ 0xFF000000, 0xFF0000C0, 0xFFC00000, 0xFFC000C0, 0xFF00C000, 0xFF00C0C0, 0xFFC0C000, 0xFFC0C0C0 };

	static constexpr uint32_t high_brightness_colors[] =
		{ 0xFF000000, 0xFF0000FF, 0xFFFF0000, 0xFFFF00FF, 0xFF00FF00, 0xFF00FFFF, 0xFFFFFF00, 0xFFFFFFFF };

	static uint32_t spectrum_color_to_argb (uint8_t spc, bool brightness)
	{
		_ASSERT(spc < 8);

		if (!brightness)
			return low_brightness_colors[spc];
		else
			return high_brightness_colors[spc];
	}

	static uint32_t* get_dest_pixel (BITMAPINFO* bi, uint32_t row, uint32_t col)
	{
		// We're going to draw the image with StretchDIBits(), which expects the bitmap to be flipped vertically.
		// We're flipping it now, because flipping while drawing with StretchDIBits() seems to be _much_ slower.
		return (uint32_t*)bi->bmiColors + (bi->bmiHeader.biHeight - 1 - row) * bi->bmiHeader.biWidth + col;
	}

	virtual bool SimulateDeviceTo (UINT64 requested_time) override
	{
		_ASSERT (_time < requested_time);

		uint64_t initial_time = _time;

		// TODO: optimize this to work as a state machine with "row" as state.
		while (true)
		{
			if (row < vsync_row_count)
			{
				// V-Sync
				// Jump to where the ULA generates the interrupt (more or less)
				if ((row == 0) && (col < irq_offset_from_frame_start))
				{
					auto requested_offset = requested_time - _time;
					_ASSERT(requested_offset < max_time_offset);
					auto offset_to_irq = irq_offset_from_frame_start - col;
					if (requested_offset < offset_to_irq)
					{
						_time = requested_time;
						col += requested_offset;
						return true;
					}

					_time += offset_to_irq;
					col = irq_offset_from_frame_start;
					if (!_pending_irq_time)
						_pending_irq_time = _time;
				}

				// jump to the end of line, then jump over all V-Sync rows
				auto requested_offset = requested_time - _time;
				_ASSERT(requested_offset < max_time_offset);
				uint32_t offset_to_border_row_0 = (ticks_per_row - col) + (ticks_per_row * (vsync_row_count - row - 1));
				if (requested_offset <= offset_to_border_row_0)
				{
					_time = requested_time;
					auto ticks_into_vsync = col + requested_offset;
					while (ticks_into_vsync >= ticks_per_row)
					{
						ticks_into_vsync -= ticks_per_row;
						row++;
					}
					col = (uint32_t)ticks_into_vsync;
					return true;
				}

				_time += offset_to_border_row_0;
				col = 0;
				row = vsync_row_count;
			}

			if (col < hsync_col_count)
			{
				// H-Sync on any visible row
				// jump to the border area
				auto requested_offset = requested_time - _time;
				_ASSERT(requested_offset < max_time_offset);
				auto offset_to_border_col_0 = hsync_col_count - col;
				if (requested_offset <= offset_to_border_col_0)
				{
					_time = requested_time;
					col += requested_offset;
					return true;
				}

				_time += offset_to_border_col_0;
				col = hsync_col_count;
			}

			if (col < hsync_col_count + border_size_left_ticks)
			{
				// left border from top of screen to bottom of screen
				_ASSERT (requested_time - _time < max_time_offset);
				uint32_t argb = spectrum_color_to_argb (_border, false);
				uint32_t* dest_pixel = get_dest_pixel (_screenData.get(), row - vsync_row_count, (col - hsync_col_count) * 2);
				auto ticks = std::min((uint32_t)(requested_time - _time), hsync_col_count + border_size_left_ticks - col);
				__stosd((PDWORD)dest_pixel, argb, ticks * 2);
				col += ticks;
				_time += ticks;
				if (_time == requested_time)
					return true;
			}

			if (col < hsync_col_count + border_size_left_ticks + 128)
			{
				// border above pixels, or pixels, or border below pixels

				if ((row < vsync_row_count + border_size_top) || (row >= vsync_row_count + border_size_top + 192))
				{
					// border above or below
					_ASSERT (requested_time - _time < max_time_offset);
					uint32_t argb = spectrum_color_to_argb (_border, false);
					uint32_t* dest_pixel = get_dest_pixel (_screenData.get(), row - vsync_row_count, (col - hsync_col_count) * 2);
					uint32_t ticks = std::min((uint32_t)(requested_time - _time), hsync_col_count + border_size_left_ticks + 128 - col);
					__stosd((PDWORD)dest_pixel, argb, ticks * 2);
					col += ticks;
					_time += ticks;
					if (_time == requested_time)
						return true;
				}
				else
				{
					// pixels
					if (col == hsync_col_count + border_size_left_ticks)
					{
						uint32_t y = row - (vsync_row_count + border_size_top);
						src_pixel_data_ptr = 0x4000 | ((y & 7) << 8) | ((y & 0x38) << 2) | ((y & 0xC0) << 5);
						src_pixel_attr_ptr = 0x5800 | ((y >> 3) << 5);
					}

					while (col < hsync_col_count + border_size_left_ticks + 128)
					{
						_ASSERT (requested_time - _time < max_time_offset);

						uint8_t data;
						if (!memory->try_physical_read_request(physical_memory_address(src_pixel_data_ptr), data, _time))
							return false;

						// We assume the attribute is in the same memory area as the pixel, so read it directly.
						uint8_t attr = memory->read_physical(physical_memory_address(src_pixel_attr_ptr));

						uint32_t* dest_pixel = get_dest_pixel (_screenData.get(), row - vsync_row_count, (col - hsync_col_count) * 2);

						bool brightness = attr & 0x40;
						auto ink_color   = spectrum_color_to_argb (attr & 7, brightness);
						auto paper_color = spectrum_color_to_argb ((attr >> 3) & 7, brightness);
						if ((attr & 0x80) && (frameNumber & 16))
							std::swap(ink_color, paper_color);
						for (uint8_t i = 0x80; i; i >>= 1)
							*dest_pixel++ = (data & i) ? ink_color : paper_color;
						src_pixel_data_ptr++;
						src_pixel_attr_ptr++;
						auto requested_offset = requested_time - _time;
						_ASSERT (requested_offset < max_time_offset);
						col += 4;
						_time += 4;
						if (requested_offset <= 4)
							return true;
					}
				}
			}

			if (col >= hsync_col_count + border_size_left_ticks + 128)
			{
				// right border from top to bottom of screen
				uint32_t argb = spectrum_color_to_argb (_border, false);
				uint32_t* dest_pixel = get_dest_pixel (_screenData.get(), row - vsync_row_count, (col - hsync_col_count) * 2);
				uint32_t ticks = std::min((uint32_t)(requested_time - _time), ticks_per_row - col);
				__stosd((PDWORD)dest_pixel, argb, ticks * 2);

				if ((col + ticks == ticks_per_row) && (row == rows_per_frame - 1))
				{
					_time += ticks - 1;
					col += ticks - 1;
					if (_screenCompleteHandler)
						_screenCompleteHandler->OnScreenDeviceComplete();
					_time++;
					col++;
				}
				else
				{
					_time += ticks;
					col += ticks;
				}
				if (col == ticks_per_row)
				{
					col = 0;
					row++;
					if (row == rows_per_frame)
					{
						row = 0;
						frameNumber++;
					}
				}

				if (_time == requested_time)
					return true;
			}

			_ASSERT(col == 0);
		}
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override
	{
		UINT64 this_frame_start_time = _time - (_time % (ticks_per_row * rows_per_frame));
		UINT64 next_frame_start_time = this_frame_start_time + ticks_per_row * rows_per_frame;
		*sync_time = next_frame_start_time - 1;
		return true;
	}

	#pragma region IInterruptingDevice
	virtual uint8_t STDMETHODCALLTYPE irq_priority() override { return 0; }

	virtual bool irq_pending (uint64_t& irq_time, uint8_t& irq_address) const override
	{
		if (_pending_irq_time)
		{
			irq_time = _pending_irq_time.value();
			irq_address = 0xff;
			return true;
		}

		return false;
	}

	virtual void acknowledge_irq() override
	{
		_ASSERT(_pending_irq_time);
		_pending_irq_time.reset();
	}
	/*
	virtual bool next_irq_time (uint32_t* irq_time) const override
	{
		uint32_t this_frame_start_time = _time % (ticks_per_row * rows_per_frame);
		uint32_t this_frame_irq_time = this_frame_start_time + irq_offset_from_frame_start;
		if (this->behind_of(this_frame_irq_time))
		{
			*irq_time = this_frame_irq_time;
			return true;
		}

		uint32_t next_frame_start_time = this_frame_start_time + ticks_per_row * rows_per_frame;
		*irq_time = next_frame_start_time + irq_offset_from_frame_start;
		return true;
	}
	*/
	#pragma endregion
/*
	virtual zx_spectrum_ula_regs regs() const override
	{
		zx_spectrum_ula_regs regs;
		regs.frame_time = (uint32_t) (_time % (ticks_per_row * rows_per_frame));
		regs.line_ticks = regs.frame_time / ticks_per_row;
		regs.col_ticks = regs.frame_time % ticks_per_row; // column in clock cycles (one unit equals two pixels)
		regs.irq = _pending_irq_time.has_value();
		return regs;
	}
*/
	#pragma region IScreenDevice
	virtual HRESULT CopyBuffer (BOOL crt, OUT BITMAPINFO** ppBuffer, OUT POINT* pBeamLocation, BOOL includeBorder = TRUE) override
	{
		constexpr uint32_t cropped_width = 256;
		constexpr uint32_t cropped_height = 192;
		uint32_t width = includeBorder ? screen_width : cropped_width;
		uint32_t height = includeBorder ? screen_height : cropped_height;
		SIZE_T bufferSize = sizeof(BITMAPINFOHEADER) + width * height * 4;
		BITMAPINFO* bi = (BITMAPINFO*)CoTaskMemAlloc(bufferSize); RETURN_IF_NULL_ALLOC(bi);
		InitBitmapInfoHeader(bi, includeBorder);

		if (crt)
		{
			if (includeBorder)
				memcpy(bi, _screenData.get(), ScreenBufferSize);
			else
			{
				const uint32_t* sourcePixels = (const uint32_t*)_screenData->bmiColors;
				uint32_t* destinationPixels = (uint32_t*)bi->bmiColors;
				for (uint32_t row = 0; row < cropped_height; row++)
				{
					const uint32_t* sourceRow = sourcePixels + (row + border_size_bottom) * screen_width + border_size_left_px;
					uint32_t* destinationRow = destinationPixels + row * cropped_width;
					memcpy(destinationRow, sourceRow, cropped_width * sizeof(uint32_t));
				}
			}
		}
		else
			GenerateInternal(bi, includeBorder);

		if (pBeamLocation)
		{
			uint32_t frame_time = (uint32_t)(_time % (ticks_per_row * rows_per_frame));
			pBeamLocation->y = (LONG)(frame_time / ticks_per_row) - (LONG)vsync_row_count;
			pBeamLocation->x = ((LONG)(frame_time % ticks_per_row) - (LONG)hsync_col_count) * 2; // column in clock cycles (one unit equals two pixels)
		}

		*ppBuffer = bi;
		return S_OK;
	}

	void GenerateInternal (BITMAPINFO* bi, BOOL includeBorder)
	{
		uint32_t pixel_row_offset = includeBorder ? border_size_top : 0;
		uint32_t pixel_col_offset = includeBorder ? border_size_left_px : 0;

		if (includeBorder)
		{
			uint32_t border_argb = spectrum_color_to_argb (_border, false);

			for (uint32_t row = 0; row < border_size_top; row++)
				__stosd((unsigned long*)get_dest_pixel(bi, row, 0), border_argb, screen_width);

			for (uint32_t row = border_size_top + 192; row < screen_height; row++)
				__stosd((unsigned long*)get_dest_pixel(bi, row, 0), border_argb, screen_width);

			for (uint32_t y = border_size_top; y <= border_size_top + 192; y++)
			{
				uint32_t* p = get_dest_pixel(bi, y, 0);
				for (uint32_t x = 0; x < border_size_left_px; x++)
					*p++ = border_argb;
				p = get_dest_pixel(bi, y, border_size_left_px + 256);
				for (uint32_t x = 0; x < border_size_right_px; x++)
					*p++ = border_argb;
			}
		}

		// pixels
		for (uint32_t y = 0; y < 192; y++)
		{
			for (uint32_t x = 0; x < 32; x++)
			{
				uint16_t src_pixel_data = 0x4000 | ((y & 7) << 8) | ((y & 0x38) << 2) | ((y & 0xC0) << 5) | x;
				_ASSERT (src_pixel_data < 0x5800);
				uint16_t src_pixel_attr = 0x5800 | ((y >> 3) << 5) | x;
				_ASSERT (src_pixel_attr < 0x5B00);
				uint8_t data = memory->read_physical(physical_memory_address(src_pixel_data));
				uint8_t attr = memory->read_physical(physical_memory_address(src_pixel_attr));
				uint32_t* dest_pixel = get_dest_pixel (bi, y + pixel_row_offset, x * 8 + pixel_col_offset);
				bool brightness = attr & 0x40;
				auto ink_color   = spectrum_color_to_argb (attr & 7, brightness);
				auto paper_color = spectrum_color_to_argb ((attr >> 3) & 7, brightness);
				for (uint8_t i = 0x80; i; i >>= 1)
					*dest_pixel++ = (data & i) ? ink_color : paper_color;
			}
		}
	}

	virtual HRESULT GenerateScreen() override
	{
		GenerateInternal(_screenData.get(), TRUE);
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetPosition(DWORD* row, DWORD* col, DWORD* frameNumber) override
	{
		RETURN_HR_IF(E_POINTER, !row || !col || !frameNumber);
		*frameNumber = this->frameNumber;
		*row = this->row;
		*col = this->col;
		return S_OK;
	}
	#pragma endregion
};

HRESULT STDMETHODCALLTYPE MakeScreenDevice (Bus* memory, Bus* io, IrqLine* irq, SpectrumVariant variant, IScreenDeviceCompleteEventHandler* eh, wistd::unique_ptr<IScreenDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<ScreenDeviceImpl>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(memory, io, irq, variant, eh); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
