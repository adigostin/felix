
#include "pch.h"
#include "Bus.h"

struct keyboard : IKeyboardDevice
{
	UINT64 _time = 0;
	Bus* io_bus;
	uint8_t keys_down[8] = { };

	HRESULT InitInstance (Bus* io_bus)
	{
		this->io_bus = io_bus;
		bool pushed = io_bus->read_responders.try_push_back({ this, &process_read_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = io_bus->write_responders.try_push_back({ this, &ProcessWriteRequest }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	~keyboard()
	{
		io_bus->write_responders.remove([this](auto& r) { return r.Device == this; });
		io_bus->read_responders.remove([this](auto& r) { return r.Device == this; });
	}

	virtual void STDMETHODCALLTYPE Reset() override
	{
		_time = 0;
		memset (keys_down, 0, sizeof(keys_down));
	}

	virtual bool SimulateTo (UINT64 requested_time) override
	{
		WI_ASSERT (_time < requested_time);
		_time = requested_time;
		return true;
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

	static uint8_t process_read_request (IDevice* d, uint16_t address)
	{
		if ((address & 0xFF) == 0xFE)
		{
			// keys and Tape In
			auto* kb = static_cast<keyboard*>(d);
			uint8_t result = 0xFF;
			for (uint8_t i = 0; i < 8; i++)
			{
				if (!(address & (1 << (i + 8))))
					result &= ~kb->keys_down[i];
			}

			return result;
		}

		return 0xFF;
	}

	static void ProcessWriteRequest (IDevice* d, uint16_t address, uint8_t value)
	{
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return false; }

	struct key_info { uint8_t index; uint8_t mask; };

	static uint8_t get_key_info (uint32_t vkey, uint32_t modifiers, key_info keys[4])
	{
		switch (vkey)
		{
			case ' ': keys[0] = { 7, 1 << 0 }; return 1;
			case 'A': keys[0] = { 1, 1 << 0 }; return 1;
			case 'B': keys[0] = { 7, 1 << 4 }; return 1;
			case 'C': keys[0] = { 0, 1 << 3 }; return 1;
			case 'D': keys[0] = { 1, 1 << 2 }; return 1;
			case 'E': keys[0] = { 2, 1 << 2 }; return 1;
			case 'F': keys[0] = { 1, 1 << 3 }; return 1;
			case 'G': keys[0] = { 1, 1 << 4 }; return 1;
			case 'H': keys[0] = { 6, 1 << 4 }; return 1;
			case 'I': keys[0] = { 5, 1 << 2 }; return 1;
			case 'J': keys[0] = { 6, 1 << 3 }; return 1;
			case 'K': keys[0] = { 6, 1 << 2 }; return 1;
			case 'L': keys[0] = { 6, 1 << 1 }; return 1;
			case 'M': keys[0] = { 7, 1 << 2 }; return 1;
			case 'N': keys[0] = { 7, 1 << 3 }; return 1;
			case 'O': keys[0] = { 5, 1 << 1 }; return 1;
			case 'P': keys[0] = { 5, 1 << 0 }; return 1;
			case 'Q': keys[0] = { 2, 1 << 0 }; return 1;
			case 'R': keys[0] = { 2, 1 << 3 }; return 1;
			case 'S': keys[0] = { 1, 1 << 1 }; return 1;
			case 'T': keys[0] = { 2, 1 << 4 }; return 1;
			case 'U': keys[0] = { 5, 1 << 3 }; return 1;
			case 'V': keys[0] = { 0, 1 << 4 }; return 1;
			case 'W': keys[0] = { 2, 1 << 1 }; return 1;
			case 'X': keys[0] = { 0, 1 << 2 }; return 1;
			case 'Y': keys[0] = { 5, 1 << 4 }; return 1;
			case 'Z': keys[0] = { 0, 1 << 1 }; return 1;
			case '0': keys[0] = { 4, 1 << 0 }; return 1;
			case '1': keys[0] = { 3, 1 << 0 }; return 1;
			case '2': keys[0] = { 3, 1 << 1 }; return 1;
			case '3': keys[0] = { 3, 1 << 2 }; return 1;
			case '4': keys[0] = { 3, 1 << 3 }; return 1;
			case '5': keys[0] = { 3, 1 << 4 }; return 1;
			case '6': keys[0] = { 4, 1 << 4 }; return 1;
			case '7': keys[0] = { 4, 1 << 3 }; return 1;
			case '8': keys[0] = { 4, 1 << 2 }; return 1;
			case '9': keys[0] = { 4, 1 << 1 }; return 1;
			case VK_RETURN: keys[0] = { 6, 1 << 0 }; return 1;
		}

		if (vkey == VK_BACK)
		{
			// CS + 0
			keys[0] = { 0, 1 << 0 };
			keys[1] = { 4, 1 << 0 };
			return 2;
		}

		if (vkey == VK_OEM_COMMA)
		{
			// SS + N
			keys[0] = { 7, 1 << 1 };
			keys[1] = { 7, 1 << 3 };
			return 2;
		}

		if (vkey == VK_OEM_MINUS)
		{
			// SS + J
			keys[0] = { 7, 1 << 1 };
			keys[1] = { 6, 1 << 3 };
			return 2;
		}

		if (vkey == VK_LEFT)
		{
			// CS + 5
			keys[0] = { 0, 1 << 0 };
			keys[1] = { 3, 1 << 4 };
			return 2;
		}

		if (vkey == VK_DOWN)
		{
			// CS + 6
			keys[0] = { 0, 1 << 0 };
			keys[1] = { 4, 1 << 4 };
			return 2;
		}

		if (vkey == VK_UP)
		{
			// CS + 7
			keys[0] = { 0, 1 << 0 };
			keys[1] = { 4, 1 << 3 };
			return 2;
		}

		if (vkey == VK_RIGHT)
		{
			// CS + 8
			keys[0] = { 0, 1 << 0 };
			keys[1] = { 4, 1 << 2 };
			return 2;
		}

		return 0;
	}

	void process_vk_shift()
	{
		// Caps Shift
		if (GetKeyState(VK_LSHIFT) >> 15)
			keys_down[0] |= 1;
		else
			keys_down[0] &= ~1;

		// Symbol Shift
		if (GetKeyState(VK_RSHIFT) >> 15)
			keys_down[7] |= 2;
		else
			keys_down[7] &= ~2;
	}

	virtual HRESULT STDMETHODCALLTYPE ProcessKeyDown (uint32_t vkey, uint32_t modifiers) override
	{
		key_info keys[4];
		if (uint8_t count = get_key_info (vkey, modifiers, keys))
		{
			for (uint8_t i = 0; i < count; i++)
				keys_down[keys[i].index] |= keys[i].mask;
			return S_OK;;
		}

		if (vkey == VK_SHIFT)
		{
			process_vk_shift();
			return S_OK;;
		}

		if (vkey == VK_OEM_7)
		{
			//  ' or " for US physical keyboard layout

			if (modifiers & MK_SHIFT)
			{
				// If the user generated this with Left Shift (rather than with Right Shift),
				// we have generated CS out of that. We need to clear the CS first.
				keys_down[0] &= ~1; // clear CS
				keys_down[7] |= 2; // SS
				keys_down[5] |= 1; // P
			}
			else
			{
				keys_down[7] |= 2; // SS
				keys_down[4] |= 8; // 7
			}
			return S_OK;;
		}

		if (vkey == VK_OEM_PLUS)
		{
			// = or +
			if (modifiers & MK_SHIFT)
			{
				// If the user generated this with Left Shift (rather than with Right Shift),
				// we have generated CS out of that. We need to clear the CS first.
				keys_down[0] &= ~1; // clear CS
				keys_down[7] |= 2; // SS
				keys_down[6] |= 4; // K
			}
			else
			{
				keys_down[7] |= 2; // SS
				keys_down[6] |= 2; // L
			}
			return S_OK;;
		}

		return S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE ProcessKeyUp (uint32_t vkey, uint32_t modifiers) override
	{
		if (vkey == VK_OEM_PLUS)
		{
			// = or +
			keys_down[7] &= ~2; // SS
			keys_down[6] &= ~4; // K
			keys_down[6] &= ~2; // L

			// Is Left Shift (CS) pressed now?
			if (GetKeyState(VK_LSHIFT) >> 15)
				keys_down[0] |= 1;

			// ... or Right Shift (SS)?
			if (GetKeyState(VK_RSHIFT) >> 15)
				keys_down[7] |= 2;

			return E_NOTIMPL;
		}

		if (vkey == VK_OEM_7)
		{
			//  ' or " for US
			keys_down[7] &= ~2; // SS
			keys_down[5] &= ~1; // P
			keys_down[4] &= ~8; // 7

			// Is Left Shift (CS) pressed now?
			if (GetKeyState(VK_LSHIFT) >> 15)
				keys_down[0] |= 1;

			// ... or Right Shift (SS)?
			if (GetKeyState(VK_RSHIFT) >> 15)
				keys_down[7] |= 2;

			return E_NOTIMPL;
		}

		if (vkey == VK_SHIFT)
		{
			process_vk_shift();
			return E_NOTIMPL;
		}

		key_info keys[4];
		if (uint8_t count = get_key_info (vkey, modifiers, keys))
		{
			for (uint8_t i = 0; i < count; i++)
				keys_down[keys[i].index] &= ~keys[i].mask;
			return E_NOTIMPL;
		}

		return S_FALSE;
	}
};

HRESULT STDMETHODCALLTYPE MakeKeyboardDevice (Bus* io_bus, wistd::unique_ptr<IKeyboardDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<keyboard>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(io_bus); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}

