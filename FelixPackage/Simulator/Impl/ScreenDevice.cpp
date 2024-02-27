
#include "pch.h"
#include "Bus.h"
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
static constexpr uint32_t hsync_col_count = 48;
static constexpr uint32_t ticks_per_row = hsync_col_count + border_size_left_ticks + 128 + border_size_right_ticks;
static constexpr uint32_t rows_per_frame = vsync_row_count + border_size_top + 192 + border_size_bottom;
static constexpr uint32_t irq_offset_from_frame_start = hsync_col_count + border_size_left_ticks;

static constexpr uint32_t screen_width  = border_size_left_px + 256 + border_size_right_px;
static constexpr uint32_t screen_height = border_size_top  + 192 + border_size_bottom;

static constexpr UINT64 max_time_offset = milliseconds_to_ticks(1000);

class ScreenDeviceImpl : public IScreenDevice, public IInterruptingDevice
{
	Bus* memory;
	Bus* io;
	irq_line_i* irq;
	UINT64 _time = 0;
	std::optional<UINT64> _pending_irq_time;
	uint8_t _border;
	wil::unique_process_heap_ptr<BITMAPINFO> _screenData;
	wil::srwlock _screenDataLock;
	IScreenDeviceCompleteEventHandler* _screenCompleteHandler;

public:
	HRESULT InitInstance (Bus* memory, Bus* io, irq_line_i* irq, IScreenDeviceCompleteEventHandler* screenCompleteHandler)
	{
		this->memory = memory;
		this->io = io;
		this->irq = irq;
		_screenCompleteHandler = screenCompleteHandler;

		bool pushed = io->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = irq->interrupting_devices.try_push_back(this); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

		_screenData.reset((BITMAPINFO*)HeapAlloc(GetProcessHeap(), 0, sizeof(BITMAPINFO) + screen_width * screen_height * 4)); RETURN_IF_NULL_ALLOC(_screenData);
		_screenData->bmiHeader.biSize = sizeof(BITMAPINFO);
		_screenData->bmiHeader.biWidth = screen_width;
		_screenData->bmiHeader.biHeight = screen_height;
		_screenData->bmiHeader.biPlanes = 1;
		_screenData->bmiHeader.biBitCount = 32;
		_screenData->bmiHeader.biCompression = BI_RGB;
		_screenData->bmiHeader.biSizeImage = 0;
		_screenData->bmiHeader.biXPelsPerMeter = 2835;
		_screenData->bmiHeader.biYPelsPerMeter = 2835;
		_screenData->bmiHeader.biClrUsed = 0;
		_screenData->bmiHeader.biClrImportant = 0;

		return S_OK;
	}

	~ScreenDeviceImpl()
	{
		irq->interrupting_devices.remove(static_cast<IInterruptingDevice*>(this));
		io->write_responders.remove([this](auto& w) { return w.Device == this; });
	}

	virtual IDevice* as_device() override { return this; }

	static void process_io_write_request (IDevice* d, uint16_t address, uint8_t value)
	{
		if ((address & 0xFF) == 0xFE)
		{
			auto* s = static_cast<ScreenDeviceImpl*>(d);
			s->_border = value & 7;
		}
	}

	virtual void STDMETHODCALLTYPE Reset() override
	{
		_time = 0;
		_pending_irq_time.reset();
	}
	
	virtual UINT64 STDMETHODCALLTYPE Time() override
	{
		return _time;
	}

	virtual HRESULT STDMETHODCALLTYPE SkipTime (UINT64 offset) override
	{
		_time += offset;
		return S_OK;
	}

	static constexpr uint32_t low_brightness_colors[] = 
		{ 0xFF000000, 0xFF0000C0, 0xFFC00000, 0xFFC000C0, 0xFF00C000, 0xFF00C0C0, 0xFFC0C000, 0xFFC0C0C0 };

	static constexpr uint32_t high_brightness_colors[] = 
		{ 0xFF000000, 0xFF0000FF, 0xFFFF0000, 0xFFFF00FF, 0xFF00FF00, 0xFF00FFFF, 0xFFFFFF00, 0xFFFFFFFF };

	static uint32_t spectrum_color_to_argb (uint8_t spc, bool brightness)
	{
		WI_ASSERT(spc < 8);

		if (!brightness)
			return low_brightness_colors[spc];
		else
			return high_brightness_colors[spc];
	}

	bool try_read_border_color (uint32_t& argb) const
	{
		if (io->writer_behind_of(_time))
			return false;
		argb = spectrum_color_to_argb (_border, false);
		return true;
	};

	uint32_t* get_dest_pixel (uint32_t row, uint32_t col)
	{
		// We're going to draw the image with StretchDIBits(), which expects the bitmap to be flipped vertically.
		// We're flipping it now, because flipping while drawing with StretchDIBits() seems to be _much_ slower.
		return (uint32_t*)_screenData->bmiColors + (screen_height - 1 - row) * screen_width + col;
	}

	virtual HRESULT STDMETHODCALLTYPE SimulateTo (UINT64 requested_time, IDeviceEventHandler* eh) override
	{
		WI_ASSERT (_time < requested_time);

		uint64_t initial_time = _time;

		// When the _time variable wraps to 0 (every ~20 minutes), we'll have a tiny glitch in one frame.
		// It's tiny because 2^32 is nearly perfectly divisible by ticks_per_row * rows_per_frame.
		uint32_t frame_time = (uint32_t)(_time % (ticks_per_row * rows_per_frame));
		uint32_t row = (uint32_t)(frame_time / ticks_per_row);
		uint32_t col = (uint32_t)(frame_time % ticks_per_row); // column in clock cycles (one unit equals two pixels)
		uint32_t argb;

		uint64_t frame_number = _time / (ticks_per_row * rows_per_frame);

		while (true)
		{
			if (row < vsync_row_count)
			{
				// V-Sync
				// Jump to where the ULA generates the interrupt (more or less)
				if ((row == 0) && (col < irq_offset_from_frame_start))
				{
					UINT64 requested_offset = requested_time - _time;
					WI_ASSERT(requested_offset < max_time_offset);
					uint32_t offset_to_irq = irq_offset_from_frame_start - col;
					if (requested_offset < offset_to_irq)
					{
						_time = requested_time;
						return true;
					}

					_time += offset_to_irq;
					col = irq_offset_from_frame_start;
					if (!_pending_irq_time)
						_pending_irq_time = _time;
				}

				// jump to the end of line, then jump over all V-Sync rows
				UINT64 requested_offset = requested_time - _time;
				WI_ASSERT(requested_offset < max_time_offset);
				uint32_t offset_to_border_row_0 = (ticks_per_row - col) + (ticks_per_row * (vsync_row_count - row));
				if (requested_offset <= offset_to_border_row_0)
				{
					_time = requested_time;
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
				UINT64 requested_offset = requested_time - _time;
				WI_ASSERT(requested_offset < max_time_offset);
				uint32_t offset_to_border_col_0 = hsync_col_count - col;
				if (requested_offset <= offset_to_border_col_0)
				{
					_time = requested_time;
					return true;
				}

				_time += offset_to_border_col_0;
				col = hsync_col_count;
			}

			while (col < hsync_col_count + border_size_left_ticks)
			{
				// left border from top of screen to bottom of screen
				WI_ASSERT (requested_time - _time < max_time_offset);
				if (!try_read_border_color(argb))
					return _time != initial_time;
				uint32_t* dest_pixel = get_dest_pixel(row - vsync_row_count, (col - hsync_col_count) * 2);
				dest_pixel[0] = argb;
				dest_pixel[1] = argb;
				col++;
				_time++;
				if (_time == requested_time)
					return true;
			}

			if (col < hsync_col_count + border_size_left_ticks + 128)
			{
				// border above pixels, or pixels, or border below pixels

				if ((row < vsync_row_count + border_size_top) || (row >= vsync_row_count + border_size_top + 192))
				{
					// border above or below
					while (col < hsync_col_count + border_size_left_ticks + 128)
					{
						WI_ASSERT (requested_time - _time < max_time_offset);
						if (!try_read_border_color(argb))
							return _time != initial_time;
						uint32_t* dest_pixel = get_dest_pixel(row - vsync_row_count, (col - hsync_col_count) * 2);
						dest_pixel[0] = argb;
						dest_pixel[1] = argb;
						col++;
						_time++;
						if (_time == requested_time)
							return true;
					}
				}
				else
				{
					// pixels
					while (col < hsync_col_count + border_size_left_ticks + 128)
					{
						WI_ASSERT (requested_time - _time < max_time_offset);

						// src pixel x and y
						uint32_t x = (col - (hsync_col_count + border_size_left_ticks)) / 4;
						//WI_ASSERT (x < 32);
						uint32_t y = row - (vsync_row_count + border_size_top);
						//WI_ASSERT (y < 192);
						uint16_t src_pixel_data = 0x4000 | ((y & 7) << 8) | ((y & 0x38) << 2) | ((y & 0xC0) << 5) | x;
						//WI_ASSERT (src_pixel_data < 0x5800);
						uint16_t src_pixel_attr = 0x5800 | ((y >> 3) << 5) | x;
						//WI_ASSERT (src_pixel_attr < 0x5B00);
						
						uint8_t data, attr;
						if (!memory->try_read_request(src_pixel_data, data, _time))
							return _time != initial_time;
						if (!memory->try_read_request(src_pixel_attr, attr, _time))
							return _time != initial_time;
						//data = memory->read(src_pixel_data);
						//attr = memory->read(src_pixel_attr);

						uint32_t* dest_pixel = get_dest_pixel (row - vsync_row_count, (col - hsync_col_count) * 2);

						bool brightness = attr & 0x40;
						auto ink_color   = spectrum_color_to_argb (attr & 7, brightness);
						auto paper_color = spectrum_color_to_argb ((attr >> 3) & 7, brightness);
						if ((attr & 0x80) && (frame_number & 16))
							std::swap(ink_color, paper_color);
						for (uint8_t i = 0x80; i; i >>= 1)
							*dest_pixel++ = (data & i) ? ink_color : paper_color;
						UINT64 requested_offset = requested_time - _time;
						WI_ASSERT (requested_offset < max_time_offset);
						col += 4;
						_time += 4;
						if (requested_offset <= 4)
							return true;
					}
				}
			}

			while (col < hsync_col_count + border_size_left_ticks + 128 + border_size_right_ticks)
			{
				// right border from top of screen to bottom of screen
				if (!try_read_border_color(argb))
					return _time != initial_time;
				uint32_t* dest_pixel = get_dest_pixel (row - vsync_row_count, (col - hsync_col_count) * 2);
				dest_pixel[0] = argb;
				dest_pixel[1] = argb;

				if ((col == ticks_per_row - 1) && (row == rows_per_frame - 1))
					_screenCompleteHandler->OnScreenComplete();

				col++;
				_time++;
				if (_time == requested_time)
					return true;
			}

			WI_ASSERT(col == ticks_per_row);
			col = 0;
			row++;
			if (row == rows_per_frame)
			{
				row = 0;
				frame_number++;
			}
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
		WI_ASSERT(_pending_irq_time);
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
	virtual HRESULT STDMETHODCALLTYPE GetScreenData (BITMAPINFO** ppBitmapInfo, POINT* pBeamLocation) override
	{
		if (ppBitmapInfo)
			*ppBitmapInfo = _screenData.get();

		if (pBeamLocation)
		{
			uint32_t frame_time = (uint32_t)(_time % (ticks_per_row * rows_per_frame));
			pBeamLocation->y = (LONG)(frame_time / ticks_per_row) - (LONG)vsync_row_count;
			pBeamLocation->x = ((LONG)(frame_time % ticks_per_row) - (LONG)hsync_col_count) * 2;
		}

		return S_OK;
	}

	virtual void generate_all() override
	{
		uint32_t border_argb = spectrum_color_to_argb (_border, false);

		for (uint32_t row = 0; row < border_size_top; row++)
			__stosd((unsigned long*)get_dest_pixel(row, 0), border_argb, screen_width);

		for (uint32_t row = border_size_top + 192; row < screen_height; row++)
			__stosd((unsigned long*)get_dest_pixel(row, 0), border_argb, screen_width);
		
		for (uint32_t y = border_size_top; y <= border_size_top + 192; y++)
		{
			uint32_t* p = get_dest_pixel(y, 0);
			for (uint32_t x = 0; x < border_size_left_px; x++)
				*p++ = border_argb;
			p = get_dest_pixel(y, border_size_left_px + 256);
			for (uint32_t x = 0; x < border_size_right_px; x++)
				*p++ = border_argb;
		}

		// pixels
		for (uint32_t y = 0; y < 192; y++)
		{
			for (uint32_t x = 0; x < 32; x++)
			{
				uint16_t src_pixel_data = 0x4000 | ((y & 7) << 8) | ((y & 0x38) << 2) | ((y & 0xC0) << 5) | x;
				WI_ASSERT (src_pixel_data < 0x5800);
				uint16_t src_pixel_attr = 0x5800 | ((y >> 3) << 5) | x;
				WI_ASSERT (src_pixel_attr < 0x5B00);
				uint8_t data = memory->read(src_pixel_data);
				uint8_t attr = memory->read(src_pixel_attr);
				uint32_t* dest_pixel = get_dest_pixel(y + border_size_top, x * 8 + border_size_left_px);
				bool brightness = attr & 0x40;
				auto ink_color   = spectrum_color_to_argb (attr & 7, brightness);
				auto paper_color = spectrum_color_to_argb ((attr >> 3) & 7, brightness);
				for (uint8_t i = 0x80; i; i >>= 1)
					*dest_pixel++ = (data & i) ? ink_color : paper_color;
			}
		}

		uint32_t frame_time = (uint32_t) (_time % (ticks_per_row * rows_per_frame));
		uint32_t row = frame_time / ticks_per_row;
		uint32_t col = frame_time % ticks_per_row; // column in clock cycles (one unit equals two pixels)
		_screenCompleteHandler->OnScreenComplete();
	}
};

HRESULT STDMETHODCALLTYPE MakeScreenDevice (Bus* memory, Bus* io, irq_line_i* irq, IScreenDeviceCompleteEventHandler* eh, wistd::unique_ptr<IScreenDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<ScreenDeviceImpl>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(memory, io, irq, eh); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
