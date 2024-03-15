
#pragma once
#include <wil/wistd_type_traits.h>

template<typename T>
class vector_nothrow
{
	T* _ptr = nullptr;
	uint32_t _size = 0;
	uint32_t _capacity = 0;

public:
	using iterator       = T*;
	using const_iterator = const T*;

	class reverse_iterator
	{
		T* ptr;
	public:
		reverse_iterator(iterator it) : ptr(it) { }
		T& operator*() const { return *(ptr - 1); }
		T* operator->() const { return ptr - 1; }
		bool operator==(const reverse_iterator& other) const { return this->ptr == other.ptr; }
		reverse_iterator& operator++() { --ptr; return *this; }
		reverse_iterator operator++(int) { reverse_iterator _Tmp = *this; --ptr; return _Tmp; }
	};

public:
	vector_nothrow() = default;

	vector_nothrow (const vector_nothrow&) = delete;
	vector_nothrow& operator= (const vector_nothrow&) = delete;

	vector_nothrow (vector_nothrow&& from) noexcept
		: _ptr(from._ptr), _size(from._size), _capacity(from._capacity)
	{
		from._ptr = nullptr;
		from._size = 0;
		from._capacity = 0;
	}

	vector_nothrow& operator= (vector_nothrow&& from)
	{
		if (_ptr)
		{
			clear();
			free(_ptr);
		}

		_ptr = from._ptr; from._ptr = nullptr;
		_size = from._size; from._size = 0;
		_capacity = from._capacity; from._capacity = 0;
		return *this;
	}

	~vector_nothrow()
	{
		if (_ptr)
		{
			clear();
			free(_ptr);
		}
	}

	void clear()
	{
		for (uint32_t i = _size - 1; i != -1; i--)
			_ptr[i].~T();
		_size = 0;
	}

	T& front()
	{
		WI_ASSERT(_size);
		return _ptr[0];
	}

	const T& front() const
	{
		WI_ASSERT(_size);
		return _ptr[0];
	}

	T& back()
	{
		WI_ASSERT(_size);
		return _ptr[_size - 1];
	}

	const T& back() const
	{
		WI_ASSERT(_size);
		return _ptr[_size - 1];
	}

	T remove (iterator it)
	{
		WI_ASSERT((it >= _ptr) && (it < _ptr + _size));
		T res = std::move(*it);
		while(it < end() - 1)
		{
			auto temp = it;
			it++;
			*temp = wistd::move(*it);
		}
		(*it).~T();
		_size--;
		return res;
	}

	T remove_back()
	{
		WI_ASSERT(_size);
		T res = std::move(_ptr[_size - 1]);
		_ptr[_size - 1].~T();
		_size--;
		return res;
	}

	void erase (iterator it)
	{
		remove(it);
	}

	bool try_push_back (const T& from)
	{
		if (_size == _capacity)
		{
			if (!try_reserve((_size < 2) ? (_size + 1) : (_size * 3 / 2)))
				return false;
		}

		new (_ptr + _size) T(from);
		_size++;
		return true;
	}

	bool try_push_back (T&& from)
	{
		if (_size == _capacity)
		{
			if (!try_reserve((_size < 2) ? (_size + 1) : (_size * 3 / 2)))
				return false;
		}

		new (_ptr + _size) T(std::move(from));
		_size++;
		return true;
	}

	bool try_push_back (std::initializer_list<T> ilist)
	{
		bool reserved = try_reserve (_size + (uint32_t)ilist.size());
		if (!reserved)
			return false;
		for (auto&& e : ilist)
			try_push_back(std::move(e));
		return true;
	}

	bool try_insert (const_iterator pos, std::initializer_list<T> ilist)
	{
		if (pos == end())
			return try_push_back(ilist);

		// inserting in the middle not yet implemented
		WI_ASSERT(false); return false;
	}

	bool try_resize (uint32_t new_size)
	{
		if (_size > new_size)
		{
			while (_size > new_size)
				remove(end() - 1);
		}
		else if (_size < new_size)
		{
			bool r = try_reserve(new_size);
			if (!r)
				return false;
			while (_size < new_size)
				try_push_back(T());
		}

		return true;
	}

	T* data() { return _ptr; }
	const T* data() const { return _ptr; }
	uint32_t size() const { return _size; }
	uint32_t capacity() const { return _capacity; }
	bool empty() const { return !_size; }

	[[nodiscard]] bool try_reserve (uint32_t new_cap)
	{
		if (new_cap > _capacity)
		{
			T* new_ptr = (T*) malloc(new_cap * sizeof(T));
			if (!new_ptr)
				return false;
			for (uint32_t i = 0; i < _size; i++)
			{
				new (new_ptr + i) T(std::move(_ptr[i]));
				_ptr[i].~T();
			}
			free(_ptr);
			_ptr = new_ptr;
			_capacity = new_cap;
		}

		return true;
	}

	iterator begin() { return _ptr; }
	iterator end() { return _ptr + _size; }

	const_iterator begin() const { return _ptr; }
	const_iterator end() const { return _ptr + _size; }

	reverse_iterator rbegin() { return reverse_iterator(this->end()); }
	reverse_iterator rend() { return reverse_iterator(this->begin()); }

	T& operator[](uint32_t index)
	{
		WI_ASSERT(index < _size);
		return _ptr[index];
	}

	const T& operator[](uint32_t index) const
	{
		WI_ASSERT(index < _size);
		return _ptr[index];
	}

	template<typename U> requires std::equality_comparable_with<T, U>
	iterator find(const U& e)
	{
		auto p = _ptr;
		auto end = _ptr + _size;
		while(p != end)
		{
			if (*p == e)
				return p;
			p++;
		}

		return p;
	}

	template<typename predicate_t>
	iterator find_if (predicate_t pred) //requires std::is_invocable_r_v<bool, predicate_t, const T&>
	{
		auto p = _ptr;
		auto end = _ptr + _size;
		for (; p != end; p++) {
			if (pred(*p)) {
				break;
			}
		}

		return p;
	}

	T remove (const T& e) //requires std::equality_comparable<T>
	{
		auto it = find(e); WI_ASSERT(it != end());
		return remove(it);
	}

	template<typename predicate_t>
	T remove (predicate_t pred) //requires std::is_invocable_r_v<bool, predicate_t, const T&>
	{
		auto it = find_if(pred); WI_ASSERT(it != end());
		return remove(it);
	}
};
