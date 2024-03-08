
#pragma once

template<typename value_t>
struct hex
{
	value_t const value;
	hex (value_t value) : value(value) { }
};

template<typename value_t>
struct fixed_width_hex
{
	value_t const value;
	fixed_width_hex (value_t value) : value(value) { }
};

// Not null-terminated. User of this class must append a null-terminator when needed.
template<class elem_t>
class basic_string_builder
{
	elem_t* _buffer = nullptr;
	uint32_t _size = 0;
	uint32_t _capacity = 0; // A _capacity equal to -1 represents an out-of-memory condition.

public:
	basic_string_builder() = default;
	
	~basic_string_builder()
	{
		if(_buffer)
			free(_buffer);
	}

	const elem_t* data() const { return _buffer; }
	uint32_t size() const { return _size; }
	bool out_of_memory() const { return _capacity == -1; }

	basic_string_builder& operator<< (elem_t ch)
	{
		if (_capacity == -1)
			return *this;

		if (_size == _capacity)
		{
			uint32_t newcap = _capacity ? (_capacity * 3 / 2) : 10;
			auto* newb = realloc(_buffer, newcap * sizeof(elem_t));
			if (!newb)
			{
				_capacity = -1;
				return *this;
			}

			_buffer = (elem_t*)newb;
			_capacity = newcap;
		}

		_buffer[_size] = ch;
		_size++;
		return *this;
	}

	void append (const elem_t* str, size_t len)
	{
		// TODO: optimize this
		for (size_t i = 0; i < len; i++)
			operator<<(str[i]);
	}

	template<typename from_elem_t>
	basic_string_builder& operator<< (from_elem_t ch) requires (sizeof(from_elem_t) < sizeof(elem_t))
	{
		return operator<<(static_cast<elem_t>(ch));
	}
	
	basic_string_builder& operator<< (const elem_t* nul_terminated)
	{
		while (*nul_terminated)
			operator<<(*nul_terminated++);
		return *this;
	}

	template<typename from_elem_t>
	basic_string_builder& operator<< (const from_elem_t* nul_terminated) requires (sizeof(from_elem_t) < sizeof(elem_t))
	{
		while (*nul_terminated)
			operator<<(static_cast<elem_t>(*nul_terminated++));
		return *this;
	}

	basic_string_builder& operator<< (uint64_t n)
	{
		char str[30];
		sprintf_s (str, "%llu", n);
		return operator<<(str);
	}

	basic_string_builder& operator<< (uint32_t n) { return operator<<((uint64_t)n); }
	basic_string_builder& operator<< (uint16_t n) { return operator<<((uint64_t)n); }
	basic_string_builder& operator<< (uint8_t n) { return operator<<((uint64_t)n); }
	basic_string_builder& operator<< (double n);
	basic_string_builder& operator<< (const basic_string_builder& other);
	basic_string_builder& operator<< (nullptr_t) { return *this; }

	basic_string_builder& operator<< (hex<uint32_t> val)
	{
		if (val.value < 10)
			return operator<<((char)(val.value + '0'));

		char str[20];
		sprintf_s (str, "%X", val.value);
		if (str[0] >= 'A' && str[0] <= 'F')
			operator<<('0');
		operator<<(str);
		operator<<('h');
		return *this;
	}

	basic_string_builder& operator<< (hex<uint16_t> val)
	{
		return operator<< (hex<uint32_t>(val.value));
	}

	basic_string_builder& operator<< (hex<uint8_t> val)
	{
		return operator<< (hex<uint32_t>(val.value));
	}

	basic_string_builder& operator<< (fixed_width_hex<uint8_t> val)
	{
		char str[5];
		sprintf_s (str, "%02X", val.value);
		return operator<<(str);
	}

	basic_string_builder& operator<< (fixed_width_hex<uint16_t> val)
	{
		char str[10];
		sprintf_s (str, "%04X", val.value);
		return operator<<(str);
	}

	bool empty() const { return _size == 0; }

	void clear() { _size = 0; }
};

using string_builder = basic_string_builder<char>;
using wstring_builder = basic_string_builder<wchar_t>;
