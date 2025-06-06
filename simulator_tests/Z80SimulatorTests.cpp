
#include "CppUnitTest.h"
#include "Impl/Z80CPU.h"
#include "Simulator.h"
#include "shared/string_builder.h"
#include "shared/unordered_map_nothrow.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

struct dummy_irq_line : irq_line_i
{
};


class TestRAM : public IDevice
{
	Bus* _memory_bus;
	UINT64 _time = 0;
	uint8_t _data[0x10000]; // this one last

public:
	HRESULT InitInstance (Bus* memory_bus)
	{
		_memory_bus = memory_bus;
		bool pushed = _memory_bus->read_responders.try_push_back({ this, &process_mem_read_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _memory_bus->write_responders.try_push_back({ this, &process_mem_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	~TestRAM()
	{
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

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { Assert::Fail(); return false; }

	virtual bool SimulateTo (UINT64 requested_time) override
	{
		// Only a bus write can change the state of this device. If there are still write-capable devices
		// whose timepoint is in the timespan we want to jump over (_time to requested_time), we can't jump.
		if (_memory_bus->writer_behind_of(requested_time))
			return false;
		_time = requested_time;
		return true;
	}

	static uint8_t process_mem_read_request (IDevice* d, uint16_t address)
	{
		auto* ram = static_cast<TestRAM*>(d);
		return ram->_data[address];
	}

	static void process_mem_write_request (IDevice* d, uint16_t address, uint8_t value)
	{
		auto* ram = static_cast<TestRAM*>(d);
		ram->_data[address] = value;
	}
};

class TestIODevice : public IDevice
{
	Bus* _io_bus;
	UINT64 _time = 0;
	unordered_map_nothrow<uint16_t, uint8_t> _data;

public:
	HRESULT InitInstance (Bus* io_bus)
	{
		_io_bus = io_bus;
		bool pushed = _io_bus->read_responders.try_push_back({ this, &process_io_read_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = _io_bus->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	~TestIODevice()
	{
		_io_bus->write_responders.remove([this](auto& d) { return d.Device == this; });
		_io_bus->read_responders.remove([this](auto& d) { return d.Device == this; });
	}

	virtual void STDMETHODCALLTYPE Reset() override
	{
		_time = 0;
		_data.clear();
	}

	virtual UINT64 STDMETHODCALLTYPE Time() override { return _time; }

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { Assert::Fail(); return false; }

	virtual bool SimulateTo (UINT64 requested_time) override
	{
		// Only a bus write can change the state of this device. If there are still write-capable devices
		// whose timepoint is in the timespan we want to jump over (_time to requested_time), we can't jump.
		if (_io_bus->writer_behind_of(requested_time))
			return false;
		_time = requested_time;
		return true;
	}

	static uint8_t process_io_read_request (IDevice* d, uint16_t address)
	{
		auto* iod = static_cast<TestIODevice*>(d);
		auto it = iod->_data.find(address);
		if (it != iod->_data.end())
			return it->second;
		return 0xFF;
	}

	static void process_io_write_request (IDevice* d, uint16_t address, uint8_t value)
	{
		auto* iod = static_cast<TestIODevice*>(d);
		auto it = iod->_data.find(address);
		if (it != iod->_data.end())
			it->second = value;
		else
		{
			bool inserted = iod->_data.try_insert({ address, value }); THROW_HR_IF(E_OUTOFMEMORY, !inserted);
		}
	}
};

namespace Z80SimulatorTests
{
	TEST_CLASS(Z80SimulatorTests)
	{
		Bus memory;
		Bus io_bus;
		dummy_irq_line irq_line;
		wistd::unique_ptr<IZ80CPU> cpu;
		wistd::unique_ptr<IDevice> ram;
		wistd::unique_ptr<IDevice> iodevice;
		z80_register_set* regs;

	public:
		Z80SimulatorTests()
		{
			auto hr = MakeZ80CPU (&memory, &io_bus, &irq_line, &cpu); THROW_IF_FAILED(hr);
			regs = cpu->GetRegsPtr();

			auto r = wil::make_unique_nothrow<TestRAM>(); THROW_IF_NULL_ALLOC(r);
			hr = r->InitInstance(&memory); THROW_IF_FAILED(hr);
			ram = std::move(r);

			auto iod = wil::make_unique_nothrow<TestIODevice>(); THROW_IF_NULL_ALLOC(iod);
			hr = iod->InitInstance(&io_bus); THROW_IF_FAILED(hr);
			iodevice = std::move(iod);
		}

		void SimulateOne()
		{
			bool res = cpu->SimulateOne(nullptr);
			Assert::IsTrue(res);
		}

		TEST_METHOD_INITIALIZE(MethodInitialize)
		{
			cpu->Reset();
			ram->Reset();
		}

		TEST_METHOD(Test_NOP)
		{
			memory.write(0, 0); // NOP
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, cpu->GetPC());
			Assert::AreEqual<uint64_t>(4, cpu->Time());
		}

		TEST_METHOD(test_xor_b)
		{
			regs->main.a = 0xff;
			regs->b() = 0xAA;
			memory.write(0, 0xA8); // xor b
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint64_t>(4, cpu->Time());
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			//Assert::AreEqual<uint8_t>(1, regs->main.f.pv); not sure about this one
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.a = 0xff;
			regs->b() = 0x55;
			memory.write(1, 0xA8); // xor b
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint8_t>(0xAA, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			//Assert::AreEqual<uint8_t>(1, regs->main.f.pv); not sure about this one
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
		}

		TEST_METHOD(test_xor_hl)
		{
			regs->main.a = 0xff;
			regs->main.hl = 0x10;
			memory.write(0x10, 0x55);
			memory.write(0, 0xAe); // xor (hl)
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint8_t>(0xAA, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			//Assert::AreEqual<uint8_t>(1, regs->main.f.pv); not sure about this one
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
		}

		TEST_METHOD(test_add_ix_plus_disp)
		{
			regs->main.a = 0x80;
			regs->ix = 0;
			memory.write(0, { 0xDD, 0x86, 0x10 }); // add (ix + 10h)
			memory.write(0x10, 0x33);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xB3, regs->main.a);
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual<uint64_t>(19, cpu->Time());
		}

		TEST_METHOD(test_ld_b_a)
		{
			regs->main.a = 0x55;
			memory.write(0, 0x47); // ld b, a
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(4, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, regs->b());
		}

		TEST_METHOD(test_ld_a_hl)
		{
			regs->main.a = 0;
			regs->main.hl = 0x10;
			memory.write(0x10, 0x55);
			memory.write(0, 0x7e); // ld a, (hl)
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
		}

		TEST_METHOD(test_ld_hl_a)
		{
			regs->main.a = 0x55;
			regs->main.hl = 0x10;
			memory.write(0, 0x77); // ld (hl), a
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(djnz)
		{
			memory.write(0, { 0x10, (uint8_t)-2 }); // label: djnz label
			regs->b() = 10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0, regs->pc);
			Assert::AreEqual<uint64_t>(13, cpu->Time());
			Assert::AreEqual<uint8_t>(9, regs->b());

			cpu->SimulateOne(nullptr); cpu->SimulateOne(nullptr); cpu->SimulateOne(nullptr); cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr); cpu->SimulateOne(nullptr); cpu->SimulateOne(nullptr); cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0, regs->pc);
			Assert::AreEqual<uint64_t>(117, cpu->Time());
			Assert::AreEqual<uint8_t>(1, regs->b());

			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(125, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->b());
		}

		TEST_METHOD(set_1_b)
		{
			memory.write(0, { 0xCB, 0xC8 }); // set 1, b
			regs->b() = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(2, regs->b());
		}

		TEST_METHOD(set_1_hl)
		{
			memory.write(0, { 0xCB, 0xCE }); // set 1, (hl)
			memory.write(0x10, 0);
			regs->main.hl = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
			Assert::AreEqual<uint8_t>(2, memory.read(0x10));
		}

		TEST_METHOD(set_1_ix_plus_disp)
		{
			memory.write(0, { 0xDD, 0xCB, 0x10, 0xCE }); // set 1, (ix + 10h)
			memory.write(0x10, 0);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(23, cpu->Time());
			Assert::AreEqual<uint8_t>(2, memory.read(0x10));
		}

		TEST_METHOD(res_1_b)
		{
			memory.write(0, { 0xCB, 0x88 }); // res 1, b
			regs->b()= 0xFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(0xFD, regs->b());
		}

		TEST_METHOD(res_1_hl)
		{
			memory.write(0, { 0xCB, 0x8E }); // res 1, (hl)
			regs->main.hl = 0x10;
			memory.write(0x10, 0xFF);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
			Assert::AreEqual<uint8_t>(0xFD, memory.read(0x10));
		}
		
		TEST_METHOD(res_1_ix_plus_disp)
		{
			memory.write(0, { 0xDD, 0xCB, 0x10, 0x8E }); // res 1, (ix + 10h)
			memory.write(0x10, 0xFF);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(23, cpu->Time());
			Assert::AreEqual<uint8_t>(0xFD, memory.read(0x10));
		}

		TEST_METHOD(bit_n_c)
		{
			for (uint16_t value = 0; value < 256; value += 255)
			{
				regs->c() = (uint8_t)value;
				for (uint8_t bit = 0; bit < 8; bit++)
				{
					memory.write(0, 0xCB);
					memory.write(1, 0x41 | (bit << 3)); // BIT b, C
					regs->pc = 0;
					regs->main.f.c = bit & 1;
					uint64_t prevTime = cpu->Time();
					cpu->SimulateOne(nullptr);
					Assert::AreEqual<uint16_t>(2, regs->pc);
					Assert::AreEqual<uint64_t>(8, cpu->Time() - prevTime);
					uint8_t andres = value & (1 << bit);
					uint8_t expectedf = andres & (z80_flag::s | z80_flag::r5 | z80_flag::r3)
						| (andres ? 0 : (z80_flag::z | z80_flag::pv))
						| z80_flag::h
						| ((bit & 1) ? z80_flag::c : 0);
					Assert::AreEqual<uint8_t>(expectedf, regs->main.f.val);
				}
			}
		}

		TEST_METHOD(bit_1_hl)
		{
			memory.write(0, { 0xCB, 0x4E }); // bit 1, (hl)
			regs->main.hl = 0x10;
			memory.write(0x10, 0xFF);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(12, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);

			memory.write(2, { 0xCB, 0x4E }); // bit 1, (hl)
			regs->main.hl = 0x10;
			memory.write(0x10, 0);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(24, cpu->Time());
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
		}
		
		TEST_METHOD(bit_1_ix_plus_disp)
		{
			memory.write(0, { 0xDD, 0xCB, 0x10, 0x4E }); // bit 1, (ix+16)
			memory.write(0x10, 0xFF);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(20, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);

			memory.write(4, { 0xDD, 0xCB, 0x10, 0x4E }); // bit 1, (ix+16)
			memory.write(0x10, 0);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(8, regs->pc);
			Assert::AreEqual<uint64_t>(40, cpu->Time());
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
		}

		TEST_METHOD(ld_b_55)
		{
			memory.write(0, { 0x06, 0x55 }); // ld b, 55h
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, regs->b());
		}

		TEST_METHOD(ld_memhl_55)
		{
			memory.write(0, { 0x36, 0x55 }); // ld (hl), 55h
			regs->main.hl = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(10, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(ld_memix_disp_55)
		{
			memory.write(0, { 0xDD, 0x36, 0x10, 0x55 }); // ld (ix + 10h), 55h
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(19, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(add_a_0fh)
		{
			memory.write(0, { 0xC6, 0x0F }); // add a, 0fh
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x0F, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r3, regs->main.f.val);

			regs->pc = 0;
			regs->main.a = 1;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x10, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::h, regs->main.f.val);

			regs->pc = 0;
			regs->main.a = 0x100 - 0xF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::h | z80_flag::c | z80_flag::z, regs->main.f.val);

			regs->pc = 0;
			regs->main.a = 0x7F; // max signed int (127)
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x8E, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r3 | z80_flag::s | z80_flag::pv | z80_flag::h, regs->main.f.val);

			regs->pc = 0;
			regs->main.a = 0x80; // min signed int (-128)
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x8F, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r3 | z80_flag::s, regs->main.f.val);
		}

		TEST_METHOD(add_a_35h)
		{
			memory.write(0, { 0xC6, 0x35 }); // add a, 35h
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x35, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
		}

		TEST_METHOD(add_a_35h_with_carry_and_zero)
		{
			regs->main.a = 0x100 - 0x35;
			memory.write(0, { 0xC6, 0x35 }); // add a, 35h
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
		}

		TEST_METHOD(sbc_a_nn)
		{
			regs->main.a = 0;
			memory.write(0, { 0xDE, 0xFF }); // sbc a, 0ffh
			regs->main.f.val = z80_flag::c;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::c | z80_flag::n | z80_flag::h | z80_flag::z, regs->main.f.val);

			regs->pc = 0;
			regs->main.a = 0;
			regs->main.f.val = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint8_t>(1, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::c | z80_flag::n | z80_flag::h, regs->main.f.val);
		}

		TEST_METHOD(push_bc)
		{
			memory.write(0, 0xC5); // push bc
			regs->main.bc = 0x55AA;
			regs->sp = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint16_t>(0xE, regs->sp);
			Assert::AreEqual<uint16_t>(0x55AA, memory.read_uint16(regs->sp));
			Assert::AreEqual<uint64_t>(11, cpu->Time());
		}

		TEST_METHOD(push_ix)
		{
			memory.write(0, { 0xDD, 0xE5 }); // push ix
			regs->ix = 0x55AA;
			regs->sp = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint16_t>(0xE, regs->sp);
			Assert::AreEqual<uint16_t>(0x55AA, memory.read_uint16(regs->sp));
			Assert::AreEqual<uint64_t>(15, cpu->Time());
		}

		TEST_METHOD(push_af)
		{
			memory.write(0, 0xF5); // push af
			regs->main.af = 0x55AA;
			regs->sp = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint16_t>(0xE, regs->sp);
			Assert::AreEqual<uint16_t>(0x55AA, memory.read_uint16(regs->sp));
			Assert::AreEqual<uint64_t>(11, cpu->Time());
		}

		TEST_METHOD(pop_bc)
		{
			memory.write(0, 0xC1); // pop bc
			regs->sp = 0xE;
			memory.write_uint16(0xE, 0x1234);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint16_t>(0x10, regs->sp);
			Assert::AreEqual<uint16_t>(0x1234, regs->main.bc);
			Assert::AreEqual<uint64_t>(10, cpu->Time());
		}
		
		TEST_METHOD(pop_ix)
		{
			memory.write(0, { 0xDD, 0xE1 }); // pop ix
			regs->sp = 0xE;
			memory.write_uint16(0xE, 0x2345);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint16_t>(0x10, regs->sp);
			Assert::AreEqual<uint16_t>(0x2345, regs->ix);
			Assert::AreEqual<uint64_t>(14, cpu->Time());
		}
		
		TEST_METHOD(pop_af)
		{
			memory.write(0, 0xF1); // pop af
			regs->sp = 0xE;
			memory.write_uint16(0xE, 0x3456);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint16_t>(0x10, regs->sp);
			Assert::AreEqual<uint16_t>(0x3456, regs->main.af);
			Assert::AreEqual<uint64_t>(10, cpu->Time());
		}

		TEST_METHOD(rlca)
		{
			memory.write(0, 7); // rlca
			regs->main.a = 0x55;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xAA, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(4, cpu->Time());

			memory.write(1, 7); // rlca
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
		}

		TEST_METHOD(rrca)
		{
			memory.write(0, 15); // rrca
			regs->main.a = 0x55;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xAA, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(4, cpu->Time());

			memory.write(1, 15); // rrca
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
		}

		TEST_METHOD(rla)
		{
			memory.write(0, 0x17); // rla
			regs->main.a = 0x55;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xAA, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(4, cpu->Time());

			memory.write(1, 0x17); // rla
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x54, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
		}
		
		TEST_METHOD(rra)
		{
			memory.write(0, 0x1f); // rra
			regs->main.a = 0x55;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x2A, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(4, cpu->Time());

			memory.write(1, 0x1f); // rra
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x95, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
		}

		TEST_METHOD(ex_sp_hl)
		{
			memory.write(0, 0xE3); // ex (sp), hl
			regs->main.hl = 0x1234;
			regs->sp = 0x10;
			memory.write_uint16 (0x10, 0x55AA);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(19, cpu->Time());
			Assert::AreEqual<uint16_t>(0x55AA, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x1234, memory.read_uint16(0x10));
		}

		TEST_METHOD(ex_sp_iy)
		{
			memory.write(0, { 0xFD, 0xE3 }); // ex (sp), iy
			regs->iy = 0x1234;
			regs->sp = 0x10;
			memory.write_uint16 (0x10, 0x55AA);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(23, cpu->Time());
			Assert::AreEqual<uint16_t>(0x55AA, regs->iy);
			Assert::AreEqual<uint16_t>(0x1234, memory.read_uint16(0x10));
		}

		TEST_METHOD(ld_a_membc)
		{
			memory.write(0, 0x0a); // ld a, (bc)
			memory.write(0x10, 0x55);
			regs->main.bc = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
		}

		TEST_METHOD(ld_a_memde)
		{
			memory.write(0, 0x1a); // ld a, (de)
			memory.write(0x10, 0x55);
			regs->main.de = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
		}

		TEST_METHOD(ld_membc_a)
		{
			memory.write(0, 2); // ld (bc), a
			regs->main.a = 0x55;
			regs->main.bc = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(ld_memde_a)
		{
			memory.write(0, 0x12); // ld (de), a
			regs->main.a = 0x55;
			regs->main.de = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(7, cpu->Time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD (rst_n)
		{
			for (uint8_t i = 0; i < 8; i++)
			{
				regs->pc = 0x80;
				regs->sp = 0x10;
				memory.write(0x80, (i << 3) | 0xC7); // RST n
				cpu->SimulateOne(nullptr);
				Assert::AreEqual<uint16_t>((i << 3), regs->pc);
				Assert::AreEqual<uint16_t>(0x0E, regs->sp);
				Assert::AreEqual<uint16_t>(0x81, memory.read_uint16(0x0E));
				Assert::AreEqual<uint64_t>(11 * (i + 1), cpu->Time());
			}
		}

		TEST_METHOD (add_hl_bc)
		{
			memory.write(0, 9); // ADD HL, BC
			regs->main.bc = 0x0FFF;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x0FFF, regs->main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::r3, regs->main.f.val);
			Assert::AreEqual<uint64_t>(11, cpu->Time());

			regs->pc = 0;
			regs->main.hl = 1;
			regs->main.bc = 0xFFFF;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0, regs->main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::c | z80_flag::h, regs->main.f.val);

			regs->pc = 0;
			regs->main.hl = 1;
			regs->main.bc = 0x0FFF;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1000, regs->main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::h, regs->main.f.val);
		}

		TEST_METHOD(add_ix_ix)
		{
			memory.write(0, { 0xDD, 0x29 }); // ADD IX, IX
			regs->ix = 0x8000;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0, regs->ix);
			Assert::AreEqual<uint8_t>(z80_flag::c, regs->main.f.val);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
		}

		TEST_METHOD(adc_hl_de)
		{
			memory.write(0, { 0xED, 0x5A }); // ADC HL, DE
			regs->main.hl = 0x8000;
			regs->main.de = 0x7FFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
			Assert::AreEqual<uint16_t>(0xFFFF, regs->main.hl);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->pc = 0;
			regs->main.hl = 0x8000;
			regs->main.de = 0x8000;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0, regs->main.hl);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);

			regs->pc = 0;
			regs->main.hl = 0x8000;
			regs->main.de = 0x7FFF;
			regs->main.f.c = 1;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0, regs->main.hl);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);

			regs->pc = 0;
			regs->main.hl = 0x0FFF;
			regs->main.de = 0;
			regs->main.f.val = z80_flag::c;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1000, regs->main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::h, regs->main.f.val);

			regs->pc = 0;
			regs->main.hl = 0x0FFF;
			regs->main.de = 0;
			regs->main.f.val = 0;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x0FFF, regs->main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::r3, regs->main.f.val);
		}

		TEST_METHOD(sbc_hl_de)
		{
			memory.write(0, { 0xED, 0x52 }); // SBC HL, DE
			regs->main.hl = 0x8000;
			regs->main.de = 0x8000;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
			Assert::AreEqual<uint16_t>(0, regs->main.hl);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->pc = 0;
			regs->main.hl = 0x7FFF;
			regs->main.de = 0x8000;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0xFFFF, regs->main.hl);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);

			regs->pc = 0;
			regs->main.hl = 0x8000;
			regs->main.de = 0x8000;
			regs->main.f.c = 1;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0xFFFF, regs->main.hl);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);

			regs->pc = 0;
			regs->main.hl = 0;
			regs->main.de = 0x7fff;
			regs->main.f.val = z80_flag::c;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0x8000, regs->main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::h | z80_flag::n | z80_flag::c, regs->main.f.val);
		}

		TEST_METHOD(rrc_c)
		{
			memory.write(0, { 0xCB, 0x09 }); // rrc c
			regs->main.bc = 0x55;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0xAA, regs->main.bc);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());

			memory.write(2, { 0xCB, 0x09 }); // rrc c
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0x55, regs->main.bc);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(16, cpu->Time());
		}

		TEST_METHOD(rr_memhl)
		{
			memory.write (0, { 0xCB, 0x1E }); // rr (hl)
			memory.write (0x10, 0x42);
			regs->main.hl = 0x10;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x21, memory.read(0x10));
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());

			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x10, memory.read(0x10));
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(30, cpu->Time());

			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x88, memory.read(0x10));
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(45, cpu->Time());
		}

		TEST_METHOD(sra_memix)
		{
			memory.write (0, { 0xDD, 0xCB, 0x10, 0x2E }); // sra (ix+10h)
			memory.write (0x10, 0x82);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xC1, memory.read(0x10));
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(23, cpu->Time());

			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xE0, memory.read(0x10));
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(46, cpu->Time());
		}

		static void get_sub_sbc_af (int C, int X, int Y, int& res, uint8_t& flags)
		{
			res = X - Y - C;

			//var flags = res & (FlagsSetMask.R3 | FlagsSetMask.R5 | FlagsSetMask.S);
			flags = res & (z80_flag::r3 | z80_flag::r5 | z80_flag::s);

			//if ((res & 0xFF) == 0) flags |= FlagsSetMask.Z;
			if ((res & 0xff) == 0)
				flags |= z80_flag::z;

			//if ((res & 0x10000) != 0) flags |= FlagsSetMask.C;
			if ((res & 0x10000) != 0)
				flags |= z80_flag::c;

			int ri = (int8_t)X - (int8_t)Y - C;
			if (ri >= 0x80 || ri < -0x80)
				//flags |= FlagsSetMask.PV;
				flags |= z80_flag::pv;

			if ((((X & 0x0F) - (res & 0x0F) - C) & 0x10) != 0)
				//flags |= FlagsSetMask.H;
				flags |= z80_flag::h;

			//flags |= FlagsSetMask.N;
			flags |= z80_flag::n;
		}

		TEST_METHOD(sub_sbc_cp_b)
		{
			// --- SUB and SBC flags
			for (int C = 0; C < 2; C++)
			{
				for (int X = 0; X < 0x100; X++)
				{
					for (int Y = 0; Y < 0x100; Y++)
					{
						int res;
						uint8_t flags;
						get_sub_sbc_af (C, X, Y, res, flags);

						if (!C)
						{
							memory.write (0, 0x90); // sub b
							regs->pc = 0;
							regs->main.a = X;
							regs->main.f.val = 0;
							regs->b() = Y;
							cpu->SimulateOne(nullptr);
							if (regs->pc != 1) 
								Assert::Fail();
							if (regs->main.a != (uint8_t)res) 
								Assert::Fail();
							if (regs->main.f.val != flags) 
								Assert::Fail();

							memory.write (0, 0xB8); // cp b
							regs->pc = 0;
							regs->main.a = X;
							regs->main.f.val = 0;
							regs->b() = Y;
							cpu->SimulateOne(nullptr);
							if (regs->pc != 1) 
								Assert::Fail();
							if (regs->main.a != X) 
								Assert::Fail();
							uint8_t flagsCP = flags & ~(z80_flag::r3 | z80_flag::r5)
								| (Y & (z80_flag::r3 | z80_flag::r5));
							if (regs->main.f.val != flagsCP) 
								Assert::Fail();
						}
						
						memory.write (0, 0x98); // sbc b
						regs->pc = 0;
						regs->main.a = X;
						regs->main.f.val = 0;
						regs->main.f.c = C;
						regs->b() = Y;
						cpu->SimulateOne(nullptr);
						if (regs->pc != 1) 
							Assert::Fail();
						if (regs->main.a != (uint8_t)res) 
							Assert::Fail();
						if (regs->main.f.val != flags) 
							Assert::Fail();
					}
				}
			}
		}

		TEST_METHOD(R_REG_TEST)
		{
			memory.write(0, 0); // NOP
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(1, cpu->GetRegsPtr()->r);
		}

		TEST_METHOD(R_REG_TEST_DD_FD)
		{
			memory.write (0, { 0xDD, 0x21, 0, 0 }); // LD IX, 0
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(2, cpu->GetRegsPtr()->r);
			memory.write (0, { 0xFD, 0x21, 0, 0 }); // LD IY, 0
			cpu->SetPC(0);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(4, cpu->GetRegsPtr()->r);
		}

		TEST_METHOD(R_REG_TEST_ED)
		{
			memory.write(0, { 0xED, 0x44 }); // NEG
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(2, cpu->GetRegsPtr()->r);
		}

		TEST_METHOD(R_REG_TEST_CB)
		{
			memory.write(0, { 0xCB, 0x27 }); // SLA A
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(2, cpu->GetRegsPtr()->r);
		}

		TEST_METHOD(R_REG_TEST_DD_CB_FD_CB)
		{
			memory.write(0, { 0xDD, 0xCB, 0x00, 0x4E }); // BIT 1, (IX + 0) 
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(2, cpu->GetRegsPtr()->r);
			memory.write(0, { 0xFD, 0xCB, 0x00, 0x4E }); // BIT 1, (IX + 0) 
			cpu->SetPC(0);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(4, cpu->GetRegsPtr()->r);
		}

		TEST_METHOD(ld_a_i)
		{
			// PDF says: "If an interrupt occurs during execution of this instruction, the Parity flag contains a 0."
			// We're not testing this at the moment.
			memory.write (0, { 0xED, 0x57 });
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(9, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(regs->iff2, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.f.c = 1;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);

			regs->i = 1;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);

			regs->i = 0x80;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
		}

		TEST_METHOD(ld_a_r)
		{
			// PDF says: "If an interrupt occurs during execution of this instruction, the Parity flag contains a 0."
			// We're not testing this at the moment.
			memory.write (0, { 0xED, 0x5F });
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(9, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(regs->iff2, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.f.c = 1;
			regs->r = 0;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);

			regs->r = 1;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);

			regs->r = 0x80;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
		}

		TEST_METHOD(ldi)
		{
			memory.write(0, { 0xED, 0xA0 }); // LDI
			memory.write(2, { 0xED, 0xA0 }); // LDI
			regs->main.hl = 10;
			regs->main.de = 20;
			regs->main.bc = 2;
			regs->main.a = 8;
			memory.write(10, { 0x55, 0xAA });

			regs->main.f.val = 0xFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(16ull, cpu->Time());
			Assert::AreEqual<uint16_t>(11, regs->main.hl);
			Assert::AreEqual<uint16_t>(21, regs->main.de);
			Assert::AreEqual<uint16_t>(1, regs->main.bc);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5); // bit 1 of A+(HL)=0x5D
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x3); // bit 3 of A+(HL)=0x5D
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0x55, memory.read(20));

			regs->main.f.val = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(32ull, cpu->Time());
			Assert::AreEqual<uint16_t>(12, regs->main.hl);
			Assert::AreEqual<uint16_t>(22, regs->main.de);
			Assert::AreEqual<uint16_t>(0, regs->main.bc);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x5); // bit 1 of A+(HL)=0xB2
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3); // bit 3 of A+(HL)=0xB2
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(21));
		}

		TEST_METHOD(ldd)
		{
			memory.write(0, { 0xED, 0xA8 }); // LDD
			memory.write(2, { 0xED, 0xA8 }); // LDD
			regs->main.hl = 11;
			regs->main.de = 21;
			regs->main.bc = 2;
			regs->main.a = 8;
			regs->main.f.val = 0xFF;
			memory.write(10, { 0x55, 0xAA });
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(16ull, cpu->Time());
			Assert::AreEqual<uint16_t>(10, regs->main.hl);
			Assert::AreEqual<uint16_t>(20, regs->main.de);
			Assert::AreEqual<uint16_t>(1, regs->main.bc);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x5); // bit 1 of A+(HL)=0xB2
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3); // bit 3 of A+(HL)=0xB2
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(21));
			regs->main.f.val = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(32ull, cpu->Time());
			Assert::AreEqual<uint16_t>(9, regs->main.hl);
			Assert::AreEqual<uint16_t>(19, regs->main.de);
			Assert::AreEqual<uint16_t>(0, regs->main.bc);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5); // bit 1 of A+(HL)=0x5D
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x3); // bit 3 of A+(HL)=0x5D
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0x55, memory.read(20));
		}

		TEST_METHOD(ldir)
		{
			memory.write(0, { 0xED, 0xB0 }); // LDIR
			regs->main.hl = 10;
			regs->main.de = 20;
			regs->main.bc = 2;
			regs->main.a = 2;
			memory.write(10, { 0x55, 0xAA });

			regs->main.f.val = 0xFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(21ull, cpu->Time());
			Assert::AreEqual<uint16_t>(0, cpu->GetPC());
			Assert::AreEqual<uint16_t>(11, regs->main.hl);
			Assert::AreEqual<uint16_t>(21, regs->main.de);
			Assert::AreEqual<uint16_t>(1, regs->main.bc);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x5); // bit 1 of A+(HL)=0x57
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3); // bit 3 of A+(HL)=0x57
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0x55, memory.read(20));

			regs->main.f.val = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(37ull, cpu->Time());
			Assert::AreEqual<uint16_t>(2, cpu->GetPC());
			Assert::AreEqual<uint16_t>(12, regs->main.hl);
			Assert::AreEqual<uint16_t>(22, regs->main.de);
			Assert::AreEqual<uint16_t>(0, regs->main.bc);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5); // bit 1 of A+(HL)=0xAC
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x3); // bit 3 of A+(HL)=0xAC
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(21));
		}

		TEST_METHOD(lddr)
		{
			memory.write(0, { 0xED, 0xB8 }); // LDDR
			regs->main.hl = 11;
			regs->main.de = 21;
			regs->main.bc = 2;
			regs->main.a = 10;
			memory.write(10, { 0x55, 0xAA });

			regs->main.f.val = 0xFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(21ull, cpu->Time());
			Assert::AreEqual<uint16_t>(0, cpu->GetPC());
			Assert::AreEqual<uint16_t>(10, regs->main.hl);
			Assert::AreEqual<uint16_t>(20, regs->main.de);
			Assert::AreEqual<uint16_t>(1, regs->main.bc);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5); // bit 1 of A+(HL)=0xB4
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3); // bit 3 of A+(HL)=0xB4
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(21));

			regs->main.f.val = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(37ull, cpu->Time());
			Assert::AreEqual<uint16_t>(2, cpu->GetPC());
			Assert::AreEqual<uint16_t>(9, regs->main.hl);
			Assert::AreEqual<uint16_t>(19, regs->main.de);
			Assert::AreEqual<uint16_t>(0, regs->main.bc);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x5); // bit 1 of A+(HL)=0x5F
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x3); // bit 3 of A+(HL)=0x5F
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0x55, memory.read(20));
		}

		TEST_METHOD(add_a_c)
		{
			memory.write(0, 0x81); // ADD A, C
			regs->main.a = 1;
			regs->main.bc = 2;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint64_t>(4, cpu->Time());
			Assert::AreEqual<uint8_t>(3, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.a = 5;
			regs->main.bc = 10;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(15, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.a = 6;
			regs->main.bc = 10;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(16, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.a = 31;
			regs->main.bc = 1;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual(32ui8, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.a = 0x7f;
			regs->main.bc = 1;
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x80, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);

			regs->main.a = 0x80; // -128
			regs->main.bc = 0x80; // -128
			regs->pc = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
		}

		TEST_METHOD(add_a_ix_d)
		{
			memory.write (0, { 0xDD, 0x86, 2 });
			regs->main.a = 1;
			regs->ix = 10;
			memory.write (12, 0x55);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint64_t>(19, cpu->Time());
			Assert::AreEqual<uint8_t>(2, regs->r);
			Assert::AreEqual<uint8_t>(0x56, regs->main.a);
		}

		TEST_METHOD(adc_a_c)
		{
			memory.write(0, 0x89); // ADC A, C
			regs->main.a = 0xFF;
			regs->main.bc = 0;
			regs->main.f.c = 1;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint64_t>(4, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);

			regs->pc = 0;
			regs->main.a = 0xFF;
			regs->main.bc = 1;
			regs->main.f.val = 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
		}

		#pragma region ADD A, Reg - Ported from https://github.com/Dotneteer/spectnetide.git
		TEST_METHOD(ADD_A_B_WorksAsExpected)
		{
			memory.write (0, { 
				0x3E, 0x12, // LD A,12H
				0x06, 0x24, // LD B,24H
				0x80        // ADD A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
		}

		TEST_METHOD(ADD_A_B_HandlesCarryFlag)
		{
			memory.write (0, {
				0x3E, 0xF0, // LD A,F0H
				0x06, 0xF0, // LD B,F0H
				0x80        // ADD A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xE0, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
		}

		TEST_METHOD(ADD_A_B_HandlesZeroFlag)
		{
			memory.write(0, {
				0x3E, 0x82, // LD A,82H
				0x06, 0x7E, // LD B,7EH
				0x80        // ADD A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
		}

		TEST_METHOD(ADD_A_B_HandlesSignFlag)
		{
			memory.write(0, {
				0x3E, 0x44, // LD A,44H
				0x06, 0x42, // LD B,42H
				0x80        // ADD A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x86, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}

		TEST_METHOD(ADD_A_C_WorksAsExpected)
		{
			memory.write(0, {
				0x3E, 0x12, // LD A,12H
				0x0E, 0x24, // LD C,24H
				0x81        // ADD A,C
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}

		TEST_METHOD(ADD_A_D_WorksAsExpected)
		{
			memory.write(0, {
				0x3E, 0x12, // LD A,12H
				0x16, 0x24, // LD D,24H
				0x82        // ADD A,D
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}

		TEST_METHOD(ADD_A_E_WorksAsExpected)
		{
			memory.write(0, {
				0x3E, 0x12, // LD A,12H
				0x1E, 0x24, // LD E,24H
				0x83        // ADD A,E
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}

		TEST_METHOD(ADD_A_H_WorksAsExpected)
		{
			memory.write(0, {
				0x3E, 0x12, // LD A,12H
				0x26, 0x24, // LD H,24H
				0x84        // ADD A,H
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}

		TEST_METHOD(ADD_A_L_WorksAsExpected)
		{
			memory.write(0, {
				0x3E, 0x12, // LD A,12H
				0x2E, 0x24, // LD L,24H
				0x85        // ADD A,L
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}

		TEST_METHOD(ADD_A_HLi_WorksAsExpected)
		{
			memory.write(0, {
				0x3E, 0x12,        // LD A,12H
				0x21, 0x00,  0x10, // LD HL,1000H
				0x86               // ADD A,(HL)
			});
			memory.write(0x1000, 0x24);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}

		TEST_METHOD(ADD_A_A_WorksAsExpected)
		{
			memory.write(0, {
				0x3E, 0x12, // LD A,12H
				0x87        // ADD A,A
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x24, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
		}
		#pragma endregion

		#pragma region ADC A, Reg - Ported from https://github.com/Dotneteer/spectnetide.git
		TEST_METHOD(ADC_A_B_WorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x06, 0x24, // LD B,24H
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x36, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
		}

		TEST_METHOD(ADC_A_B_HandlesCarryFlag)
		{
			memory.write (0, {
				0x3E, 0xF0, // LD A,F0H
				0x06, 0xF0, // LD B,F0H
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xE0, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
		}

		TEST_METHOD(ADC_A_B_HandlesZeroFlag)
		{
			memory.write (0, {
				0x3E, 0x82, // LD A,82H
				0x06, 0x7E, // LD B,7EH
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
		}

		TEST_METHOD(ADC_A_B_HandlesSignFlag)
		{
			memory.write (0, {
				0x3E, 0x44, // LD A,44H
				0x06, 0x42, // LD B,42H
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x86, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
		}

		TEST_METHOD(ADC_A_B_WithCarryWorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x06, 0x24, // LD B,24H
				0x37,       // SCF
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x37, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_B_WithCarryHandlesCarryFlag)
		{
			memory.write (0, {
				0x3E, 0xF0, // LD A,F0H
				0x06, 0xF0, // LD B,F0H
				0x37,       // SCF
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0xE1, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_B_WithCarryHandlesZeroFlag)
		{
			memory.write (0, {
				0x3E, 0x82, // LD A,82H
				0x06, 0x7D, // LD B,7DH
				0x37,       // SCF
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_B_WithCarryHandlesSignFlag)
		{
			memory.write (0, {
				0x3E, 0x44, // LD A,44H
				0x06, 0x42, // LD B,42H
				0x37,       // SCF
				0x88        // ADC A,B
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x87, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_C_WithCarryWorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x0E, 0x24, // LD C,24H
				0x37,       // SCF
				0x89        // ADC A,C
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x37, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_D_WithCarryWorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x16, 0x24, // LD D,24H
				0x37,       // SCF
				0x8A        // ADC A,D
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x37, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_E_WithCarryWorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x1E, 0x24, // LD E,24H
				0x37,       // SCF
				0x8B        // ADC A,E
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x37, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_H_WithCarryWorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x26, 0x24, // LD H,24H
				0x37,       // SCF
				0x8C        // ADC A,H
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x37, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_L_WithCarryWorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x2E, 0x24, // LD L,24H
				0x37,       // SCF
				0x8D        // ADC A,L
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x37, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(22, cpu->Time());
		}

		TEST_METHOD(ADC_A_HLi_WorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12,        // LD A,12H
				0x21, 0x00,  0x10, // LD HL,1000H
				0x37,              // SCF
				0x8E               // ADD A,(HL)
			});
			memory.write(0x1000, 0x24);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x37, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(28, cpu->Time());
		}

		TEST_METHOD(ADC_A_A_WithCarryWorksAsExpected)
		{
			memory.write (0, {
				0x3E, 0x12, // LD A,12H
				0x37,       // SCF
				0x8F        // ADC A,A
			});
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint8_t>(0x25, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
		}
		#pragma endregion

		TEST_METHOD(inc_a)
		{
			memory.write(0, 0x3C); // INC A
			regs->main.a = 0x7E;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x7F, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(0, regs->main.f.h);
			Assert::AreEqual<uint8_t>(1, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(0, regs->main.f.c);
			regs->pc = 0;
			regs->main.f.c = 1;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, regs->main.a);
			Assert::AreEqual<uint8_t>(1, regs->main.f.s);
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(1, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
			regs->pc = 0;
			regs->main.a = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0, regs->main.f.s);
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x5);
			Assert::AreEqual<uint8_t>(1, regs->main.f.h);
			Assert::AreEqual<uint8_t>(0, regs->main.f.x3);
			Assert::AreEqual<uint8_t>(0, regs->main.f.pv);
			Assert::AreEqual<uint8_t>(0, regs->main.f.n);
			Assert::AreEqual<uint8_t>(1, regs->main.f.c);
		}

		TEST_METHOD(dec_a)
		{
			memory.write(0, 0x3D); // DEC A
			regs->main.a = 0x81;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::n, regs->main.f.val);
			regs->pc = 0;
			regs->main.f.c = 1;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x7F, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r5 | z80_flag::h | z80_flag::r3 | z80_flag::pv | z80_flag::n | z80_flag::c, regs->main.f.val);
			regs->pc = 0;
			regs->main.f.c = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x7E, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r5 | z80_flag::r3 | z80_flag::n, regs->main.f.val);
			regs->pc = 0;
			regs->main.a = 1;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n, regs->main.f.val);
			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xFF, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::h | z80_flag::r3 | z80_flag::n, regs->main.f.val);
		}

		TEST_METHOD(neg)
		{
			memory.write(0, { 0xED, 0x44 }); // NEG
			regs->main.a = 3;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xFD, regs->main.a);
			Assert::AreEqual<uint8_t>(0b1011'1011, regs->main.f.val);
			regs->pc = 0;
			regs->main.a = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(0b0100'0010, regs->main.f.val);
			regs->pc = 0;
			regs->main.a = 0x80;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, regs->main.a);
			Assert::AreEqual<uint8_t>(0b1000'0111, regs->main.f.val);
		}

		TEST_METHOD(inc_rr)
		{
			memory.write (0, 3); // INC BC
			SimulateOne();
			Assert::AreEqual<uint16_t>(1, regs->main.bc);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(6, cpu->Time());

			regs->pc = 0;
			memory.write (0, { 0xDD, 0x23 }); // INC IX
			SimulateOne();
			Assert::AreEqual<uint16_t>(1, regs->ix);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(16, cpu->Time());
		}

		TEST_METHOD(dec_rr)
		{
			memory.write (0, 0x1B); // DEC DE
			SimulateOne();
			Assert::AreEqual<uint16_t>(0xFFFF, regs->main.de);
			Assert::AreEqual<uint16_t>(1, regs->pc);
			Assert::AreEqual<uint64_t>(6, cpu->Time());

			regs->pc = 0;
			memory.write (0, { 0xFD, 0x2B }); // DEC IY
			SimulateOne();
			Assert::AreEqual<uint16_t>(0xFFFF, regs->iy);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(16, cpu->Time());
		}

		TEST_METHOD(rlc_b)
		{
			memory.write (0, { 0xCB, 0x00 }); // RLC B
			regs->b() = 0x55;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xAA, regs->b());
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::pv, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x55, regs->b());
			Assert::AreEqual<uint8_t>(z80_flag::pv | z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			regs->b() = 0;
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->b());
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::pv, regs->main.f.val);
		}

		TEST_METHOD(rlc_hl)
		{
			memory.write (0, { 0xCB, 0x06 }); // RLC (HL)
			memory.write (0x10, 0x80);
			regs->main.hl = 0x10;
			SimulateOne();
			Assert::AreEqual<uint8_t>(1, memory.read(0x10));
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::c, regs->main.f.val);
		}

		TEST_METHOD(rlc_ix_plus_d)
		{
			memory.write (0, { 0xDD, 0xCB, 6, 0x06 }); // RLC (IX + 6)
			memory.write (0x10, 0x80);
			regs->ix = 10;
			SimulateOne();
			Assert::AreEqual<uint8_t>(1, memory.read(0x10));
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(23, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			memory.write (0, { 0xDD, 0xCB, 0x80, 0x06 }); // RLC (IX - 128)
			memory.write (0x10, 0x80);
			regs->ix = 128 + 0x10;
			SimulateOne();
			Assert::AreEqual<uint8_t>(1, memory.read(0x10));
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(23 + 23, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::c, regs->main.f.val);
		}

		TEST_METHOD(rrc_m)
		{
			memory.write (0, { 0xCB, 0x0F }); // RRC A
			regs->main.a = 0xAA;
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::pv, regs->main.f.val);

			cpu->Reset();
			regs->main.a = 0x55;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xAA, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::pv | z80_flag::c, regs->main.f.val);
			Assert::AreEqual<uint64_t>(8, cpu->Time());

			cpu->Reset();
			regs->main.a = 0;
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::pv, regs->main.f.val);

			cpu->Reset();
			memory.write (0, { 0xCB, 0x0E }); // RRC (HL)
			memory.write (0x10, 1);
			regs->main.hl = 0x10;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, memory.read(0x10));
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::c, regs->main.f.val);

			cpu->Reset();
			memory.write (0, { 0xDD, 0xCB, 0x80, 0x0E }); // RRC (IX - 128)
			memory.write (0x10, 1);
			regs->ix = 128 + 0x10;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, memory.read(0x10));
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(23, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::c, regs->main.f.val);
		}

		TEST_METHOD(rl_b)
		{
			memory.write (0, { 0xCB, 0x10 }); // RL B
			regs->b() = 0x55;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xAA, regs->b());
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::pv, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x54, regs->b());
			Assert::AreEqual<uint8_t>(z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xA9, regs->b());
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::pv, regs->main.f.val);

			regs->pc = 0;
			regs->b() = 0;
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint8_t>(1, regs->b());
			Assert::AreEqual<uint8_t>(0, regs->main.f.val);
		}

		TEST_METHOD(rr_m)
		{
			memory.write (0, { 0xCB, 0x1E }); // RR (HL)
			memory.write (0x10, 1);
			regs->main.hl = 0x10;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(15, cpu->Time());
			Assert::AreEqual<uint8_t>(0, memory.read(0x10));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::pv | z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, memory.read(0x10));
			Assert::AreEqual<uint8_t>(z80_flag::s, regs->main.f.val);

			cpu->Reset();
			regs->iy = 8;
			memory.write (0, { 0xFD, 0xCB, 8, 0x1E }); // RR (IY + 8)
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x40, memory.read(0x10));
			Assert::AreEqual<uint8_t>(0, regs->main.f.val);

			regs->pc = 0;
			regs->main.f.c = 1;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xA0, memory.read(0x10));
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::pv , regs->main.f.val);
		}

		TEST_METHOD(sla_h)
		{
			memory.write (0, { 0xCB, 0x24 }); // SLA H
			regs->h() = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(0xFE, regs->h());
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::c, regs->main.f.val);
		}

		TEST_METHOD(sra_a)
		{
			memory.write (0, { 0xCB, 0x2F }); // SRA A
			regs->main.a = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(0xFF, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::pv | z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			regs->main.a = 1;
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->main.a);
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::pv | z80_flag::c, regs->main.f.val);
		}

		TEST_METHOD(sll_e)
		{
			memory.write (0, { 0xCB, 0x33 }); // SLL E
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(1, regs->e());
			Assert::AreEqual<uint8_t>(0, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x0F, regs->e());
			Assert::AreEqual<uint8_t>(z80_flag::r3 | z80_flag::pv, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xFF, regs->e());

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0xFF, regs->e());
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::pv | z80_flag::c, regs->main.f.val);
		}

		TEST_METHOD(srl_d)
		{
			memory.write (0, { 0xCB, 0x3A }); // SRL D
			regs->d() = 0xFF;
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(0x7F, regs->d());
			Assert::AreEqual<uint8_t>(z80_flag::r5 | z80_flag::r3 | z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x0F, regs->d());
			Assert::AreEqual<uint8_t>(z80_flag::r3 | z80_flag::pv | z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->d());
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::pv | z80_flag::c, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0, regs->d());
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::pv, regs->main.f.val);
		}

		TEST_METHOD(rld)
		{
			memory.write(0, { 0xED, 0x6F }); // RLD
			regs->main.a = 0x7A;
			memory.write(0x5000, 0x31);
			regs->main.hl = 0x5000;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
			Assert::AreEqual<uint8_t>(0x73, regs->main.a);
			Assert::AreEqual<uint8_t>(0x1A, memory.read(0x5000));
			Assert::AreEqual<uint8_t>(z80_flag::r5, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x71, regs->main.a);
			Assert::AreEqual<uint8_t>(0xA3, memory.read(0x5000));
			Assert::AreEqual<uint8_t>(z80_flag::r5 | z80_flag::pv, regs->main.f.val);
		}

		TEST_METHOD(rrd)
		{
			memory.write(0, { 0xED, 0x67 }); // RRD
			regs->main.a = 0x84;
			memory.write(0x5000, 0x20);
			regs->main.hl = 0x5000;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(18, cpu->Time());
			Assert::AreEqual<uint8_t>(0x80, regs->main.a);
			Assert::AreEqual<uint8_t>(0x42, memory.read(0x5000));
			Assert::AreEqual<uint8_t>(z80_flag::s, regs->main.f.val);

			regs->pc = 0;
			SimulateOne();
			Assert::AreEqual<uint8_t>(0x82, regs->main.a);
			Assert::AreEqual<uint8_t>(0x04, memory.read(0x5000));
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::pv, regs->main.f.val);
		}

		TEST_METHOD(jp_nn)
		{
			memory.write (0, { 0xC3, 0x34, 0x12 }); // JP 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual<uint64_t>(10, cpu->Time());
		}

		TEST_METHOD(jp_nz_nn)
		{
			memory.write (0, { 0xC2, 0x34, 0x12 }); // JP NZ, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.z = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jp_z_nn)
		{
			memory.write (0, { 0xCA, 0x34, 0x12 }); // JP Z, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.z = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jp_nc_nn)
		{
			memory.write (0, { 0xD2, 0x34, 0x12 }); // JP NC, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.c = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jp_c_nn)
		{
			memory.write (0, { 0xDA, 0x34, 0x12 }); // JP C, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.c = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jp_po_nn)
		{
			memory.write (0, { 0xE2, 0x34, 0x12 }); // JP PO, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.pv = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jp_pe_nn)
		{
			memory.write (0, { 0xEA, 0x34, 0x12 }); // JP PE, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.pv = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jp_p)
		{
			memory.write (0, { 0xF2, 0x34, 0x12 }); // JP P, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.s = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jp_m_nn)
		{
			memory.write (0, { 0xFA, 0x34, 0x12 }); // JP M, 1234h
			SimulateOne();
			Assert::AreEqual<uint16_t>(3, regs->pc);
			Assert::AreEqual(10ui64, cpu->Time());
			regs->pc = 0;
			regs->main.f.s = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual(20ui64, cpu->Time());
		}

		TEST_METHOD(jr_e)
		{
			memory.write (0, { 0x18, 0x7F }); // JR +127
			SimulateOne();
			Assert::AreEqual<uint16_t>(2 + 127, regs->pc);
			Assert::AreEqual (12ui64, cpu->Time());

			regs->pc = 0;
			memory.write (0, { 0x18, 0 }); // JR +0
			SimulateOne();
			Assert::AreEqual ((uint16_t)2, regs->pc);
			Assert::AreEqual (24ui64, cpu->Time());

			regs->pc = 0;
			memory.write (0, { 0x18, 0x80 }); // JR -128
			SimulateOne();
			Assert::AreEqual ((uint16_t)(2 - 128), regs->pc);
			Assert::AreEqual (36ui64, cpu->Time());
		}

		TEST_METHOD(jr_nz_e)
		{
			memory.write (0, { 0x20, 100 }); // JR NZ, +100
			SimulateOne();
			Assert::AreEqual<uint16_t>(2 + 100, regs->pc);
			Assert::AreEqual(12ui64, cpu->Time());
			
			cpu->Reset();
			regs->main.f.z = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual(7ui64, cpu->Time());
		}

		TEST_METHOD(jr_z_e)
		{
			memory.write (0, { 0x28, 256 - 100 }); // JR NZ, -100
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual(7ui64, cpu->Time());

			cpu->Reset();
			regs->main.f.z = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>((uint16_t)(2 - 100), regs->pc);
			Assert::AreEqual(12ui64, cpu->Time());
		}

		TEST_METHOD(jr_nc_e)
		{
			memory.write (0, { 0x30, 100 }); // JR NC, +100
			SimulateOne();
			Assert::AreEqual<uint16_t>(2 + 100, regs->pc);
			Assert::AreEqual(12ui64, cpu->Time());

			cpu->Reset();
			regs->main.f.c = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual(7ui64, cpu->Time());
		}

		TEST_METHOD(jr_c_e)
		{
			memory.write (0, { 0x38, 256 - 100 }); // JR C, -100
			SimulateOne();
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual(7ui64, cpu->Time());

			cpu->Reset();
			regs->main.f.c = 1;
			SimulateOne();
			Assert::AreEqual<uint16_t>((uint16_t)(2 - 100), regs->pc);
			Assert::AreEqual(12ui64, cpu->Time());
		}

		TEST_METHOD(jp_hl_ix_iy)
		{
			memory.write (0, 0xE9); // JP (HL)
			SimulateOne();
			Assert::AreEqual<uint16_t>(0, regs->pc);
			Assert::AreEqual(4ui64, cpu->Time());

			cpu->Reset();
			memory.write (0, { 0xDD, 0xE9 }); // JP (IX)
			SimulateOne();
			Assert::AreEqual<uint16_t>(0, regs->pc);
			Assert::AreEqual(8ui64, cpu->Time());

			cpu->Reset();
			memory.write (0, { 0xFD, 0xE9 }); // JP (IY)
			SimulateOne();
			Assert::AreEqual<uint16_t>(0, regs->pc);
			Assert::AreEqual(8ui64, cpu->Time());
		}

		TEST_METHOD(djnz_e)
		{
			memory.write (0, { 0x10, 256 - 128 }); // DJNZ -128
			SimulateOne();
			Assert::AreEqual(255ui8, regs->b());
			Assert::AreEqual((uint16_t)(2 - 128), regs->pc);
			Assert::AreEqual(13ui64, cpu->Time());

			cpu->Reset();
			regs->b() = 1;
			SimulateOne();
			Assert::AreEqual(0ui8, regs->b());
			Assert::AreEqual((uint16_t)2, regs->pc);
			Assert::AreEqual(8ui64, cpu->Time());

			cpu->Reset();
			memory.write (0, { 0x10, 127 }); // DJNZ +127
			regs->b() = 255;
			SimulateOne();
			Assert::AreEqual(254ui8, regs->b());
			Assert::AreEqual((uint16_t)(2 + 127), regs->pc);
			Assert::AreEqual(13ui64, cpu->Time());

			cpu->Reset();
			regs->b() = 1;
			SimulateOne();
			Assert::AreEqual(0ui8, regs->b());
			Assert::AreEqual((uint16_t)2, regs->pc);
			Assert::AreEqual(8ui64, cpu->Time());
		}

		TEST_METHOD(call_nn)
		{
			memory.write (0, { 0xCD, 0x34, 0x12 }); // CALL 1234h
			SimulateOne();
			Assert::AreEqual<uint64_t>(17, cpu->Time());
			Assert::AreEqual<uint16_t>(0x1234, regs->pc);
			Assert::AreEqual<uint16_t>(0xFFFE, regs->sp);
			Assert::AreEqual<uint16_t>(3, memory.read_uint16(0xFFFE));
		}

		TEST_METHOD(call_cc_nn)
		{
			static const uint8_t ccmask[] = { z80_flag::z, z80_flag::c, z80_flag::pv, z80_flag::s };

			for (uint8_t i = 0; i < 8; i++)
			{
				// call must be taken
				cpu->Reset();
				memory.write (0, { (uint8_t)(0xC4 | (i << 3)), 0x34, 0x12 }); // CALL cc, 1234h
				if (i & 1)
				{
					// Z, C, PE, M
					regs->main.f.val |= ccmask[i / 2];
				}
				SimulateOne();
				Assert::AreEqual<uint16_t>(0x1234, regs->pc);
				Assert::AreEqual<uint64_t>(17, cpu->Time());
				Assert::AreEqual<uint16_t>(0xFFFE, regs->sp);
				Assert::AreEqual<uint16_t>(3, memory.read_uint16(0xFFFE));

				// call must not be taken
				cpu->Reset();
				if ((i & 1) == 0)
				{
					// NZ, NC, PO, P
					regs->main.f.val |= ccmask[i / 2];
				}
				SimulateOne();
				Assert::AreEqual<uint16_t>(3, regs->pc);
				Assert::AreEqual<uint64_t>(10, cpu->Time());
				Assert::AreEqual<uint16_t>(0, regs->sp);
			}
		}

		TEST_METHOD(ret)
		{
			memory.write(0, 0xC9); // RET
			memory.write(0xFFFE, { 0x34, 0x12 }); // return address is 0x1234
			regs->sp = 0xFFFE;
			SimulateOne();
			Assert::AreEqual (10ui64, cpu->Time());
			Assert::AreEqual (0x1234ui16, regs->pc);
			Assert::AreEqual (0ui16, regs->sp);
		}

		TEST_METHOD(ret_cc)
		{
			static const uint8_t ccmask[] = { z80_flag::z, z80_flag::c, z80_flag::pv, z80_flag::s };

			memory.write(0xFFFE, { 0x34, 0x12 }); // return address is 0x1234

			for (uint8_t i = 0; i < 8; i++)
			{
				// must return
				cpu->Reset();
				memory.write (0, (uint8_t)(0xC0 | (i << 3))); // RET cc
				regs->sp = 0xFFFE;
				if (i & 1)
				{
					// Z, C, PE, M
					regs->main.f.val |= ccmask[i / 2];
				}
				SimulateOne();
				Assert::AreEqual<uint16_t>(0x1234, regs->pc);
				Assert::AreEqual<uint16_t>(0, regs->sp);
				Assert::AreEqual<uint64_t>(11, cpu->Time());

				// must not return
				cpu->Reset();
				regs->sp = 0xFFFE;
				if ((i & 1) == 0)
				{
					// NZ, NC, PO, P
					regs->main.f.val |= ccmask[i / 2];
				}
				SimulateOne();
				Assert::AreEqual<uint16_t>(1, regs->pc);
				Assert::AreEqual<uint64_t>(5, cpu->Time());
				Assert::AreEqual<uint16_t>(0xFFFE, regs->sp);
			}
		}

		TEST_METHOD(reti)
		{
			memory.write (0, { 0xED, 0x4D }); // RETI
			memory.write (0xFFFE, { 0x34, 0x12 });
			regs->sp = 0xFFFE;
			SimulateOne();
			Assert::AreEqual (14ui64, cpu->Time());
			Assert::AreEqual (0x1234ui16, regs->pc);
			Assert::AreEqual (0ui16, regs->sp);
		}

		TEST_METHOD(retn)
		{
			memory.write (0, { 0xED, 0x45 }); // RETN
			memory.write (0xFFFE, { 0x34, 0x12 });
			regs->sp = 0xFFFE;
			regs->iff1 = 1;
			SimulateOne();
			Assert::AreEqual (14ui64, cpu->Time());
			Assert::AreEqual (0x1234ui16, regs->pc);
			Assert::AreEqual (0ui16, regs->sp);
			Assert::AreEqual<bool>(0, regs->iff1);

			cpu->Reset();
			regs->sp = 0xFFFE;
			regs->iff2 = 1;
			SimulateOne();
			Assert::AreEqual (14ui64, cpu->Time());
			Assert::AreEqual (0x1234ui16, regs->pc);
			Assert::AreEqual (0ui16, regs->sp);
			Assert::AreEqual<bool>(1, regs->iff1);
		}

		TEST_METHOD(in_a_n)
		{
			memory.write (0, { 0xDB, 0x10 }); // IN A, (10h)
			regs->main.a = 0x80; // A is put on the upper 8 bits, so address will be 0x8010
			io_bus.write (0x8010, 0x55);
			SimulateOne();
			Assert::AreEqual<uint64_t>(11, cpu->Time());
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint8_t>(0x55, regs->main.a);
		}

		TEST_METHOD(out_n_a)
		{
			memory.write (0, { 0xD3, 0x10 }); // OUT A, (10h)
			regs->main.a = 0x80; // A is put on the upper 8 bits, so address will be 0x8010
			SimulateOne();
			Assert::AreEqual<uint64_t>(11, cpu->Time());
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint8_t>(0x80, io_bus.read (0x8010));
		}

		TEST_METHOD(in_r_c)
		{
			io_bus.write (0x1234, 0xA9);

			for (uint8_t reg = 0; reg < 8; reg++)
			{
				if (reg == 6)
					continue;

				cpu->Reset();
				memory.write (0, { 0xED, (uint8_t)(0x40 | (reg << 3)) }); // IN reg, (C)
				regs->main.bc = 0x1234;
				SimulateOne();
				uint8_t val = regs->r8(reg, hl_ix_iy::hl);
				Assert::AreEqual<uint8_t>(0xA9, val);
				Assert::AreEqual<uint64_t>(12, cpu->Time());
				Assert::AreEqual<uint16_t>(2, regs->pc);
				uint8_t flags_from_a9 = z80_flag::s | z80_flag::r5 | z80_flag::r3 | z80_flag::pv;
				Assert::AreEqual<uint8_t>(flags_from_a9, regs->main.f.val);
			}
		}

		TEST_METHOD(out_c_r)
		{
			for (uint8_t reg = 0; reg < 8; reg++)
			{
				if (reg == 6)
					continue;

				cpu->Reset();
				iodevice->Reset();
				memory.write (0, { 0xED, (uint8_t)(0x41 | (reg << 3)) }); // OUT (C), reg
				regs->r8(reg, hl_ix_iy::hl) = 0x55;
				regs->main.bc = 0x1234;
				SimulateOne();
				Assert::AreEqual<uint64_t>(12, cpu->Time());
				Assert::AreEqual<uint16_t>(2, regs->pc);
				uint8_t expected;
				if (reg == 0)
					expected = 0x12;
				else if (reg == 1)
					expected = 0x34;
				else
					expected = 0x55;
				Assert::AreEqual<uint8_t>(expected, io_bus.read (0x1234));
			}
		}

		TEST_METHOD(ini)
		{
			memory.write(0, { 0xED, 0xA2 }); // INI
			memory.write(2, { 0xED, 0xA2 }); // INI
			regs->main.hl = 10;     // memory address is 10
			regs->main.bc = 0x02FE; // io address is 0xFE

			io_bus.write(regs->main.bc, 0x55); // data is 0x55
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual(16ull, cpu->Time());
			Assert::AreEqual<uint16_t>(11, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, memory.read(10));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
			
			io_bus.write(regs->main.bc, 0xAA); // data is 0xAA
			regs->main.f.val = 0;
			SimulateOne();
			Assert::AreEqual(32ull, cpu->Time());
			Assert::AreEqual<uint16_t>(12, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(11));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}
		
		TEST_METHOD(ind)
		{
			memory.write(0, { 0xED, 0xAA }); // IND
			memory.write(2, { 0xED, 0xAA }); // IND
			regs->main.hl = 11;
			regs->main.bc = 0x02FE;

			io_bus.write(regs->main.bc, 0x55); // data is 0x55
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual(16ull, cpu->Time());
			Assert::AreEqual<uint16_t>(10, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, memory.read(11));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));

			io_bus.write(regs->main.bc, 0xAA); // data is 0xAA
			regs->main.f.val = 0;
			SimulateOne();
			Assert::AreEqual(32ull, cpu->Time());
			Assert::AreEqual<uint16_t>(9, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(10));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}
		
		TEST_METHOD(inir)
		{
			memory.write(0, { 0xED, 0xB2 }); // INIR
			regs->main.hl = 10;
			regs->main.bc = 0x02FE;

			io_bus.write(regs->main.bc, 0x55); // data is 0x55
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual(21ull, cpu->Time());
			Assert::AreEqual<uint16_t>(0, cpu->GetPC());
			Assert::AreEqual<uint16_t>(11, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, memory.read(10));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));

			io_bus.write(regs->main.bc, 0xAA); // data is 0xAA
			SimulateOne();
			Assert::AreEqual(37ull, cpu->Time());
			Assert::AreEqual<uint16_t>(2, cpu->GetPC());
			Assert::AreEqual<uint16_t>(12, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(11));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}
		
		TEST_METHOD(indr)
		{
			memory.write(0, { 0xED, 0xBA }); // INDR
			regs->main.hl = 11;
			regs->main.bc = 0x02FE;
			
			io_bus.write(regs->main.bc, 0x55); // data is 0x55
			regs->main.f.val = 0xFF;
			SimulateOne();
			Assert::AreEqual(21ull, cpu->Time());
			Assert::AreEqual<uint16_t>(0, cpu->GetPC());
			Assert::AreEqual<uint16_t>(10, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, memory.read(11));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));

			io_bus.write(regs->main.bc, 0xAA); // data is 0xAA
			SimulateOne();
			Assert::AreEqual(37ull, cpu->Time());
			Assert::AreEqual<uint16_t>(2, cpu->GetPC());
			Assert::AreEqual<uint16_t>(9, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, memory.read(10));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}

		TEST_METHOD(outi)
		{
			memory.write (0, { 0xED, 0xA3 }); // OUTI
			memory.write (2, { 0xED, 0xA3 }); // OUTI
			regs->main.hl = 10; // memory address is 10
			regs->main.bc = 0x02FE; // io address is 0xFE

			memory.write (regs->main.hl, 0x55); // data is 0x55
			regs->main.f.val = 0xff;
			SimulateOne();
			Assert::AreEqual(16ull, cpu->Time());
			Assert::AreEqual<uint16_t>(11, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, io_bus.read(0x01FE));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));

			memory.write (regs->main.hl, 0xAA); // data is 0xAA
			regs->main.f.val = 0;
			SimulateOne();
			Assert::AreEqual(32ull, cpu->Time());
			Assert::AreEqual<uint16_t>(12, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, io_bus.read(0x00FE));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}

		TEST_METHOD(outd)
		{
			memory.write (0, { 0xED, 0xAB }); // OUTD
			memory.write (2, { 0xED, 0xAB }); // OUTD
			regs->main.hl = 11;
			regs->main.bc = 0x02FE;

			memory.write (regs->main.hl, 0x55); // data is 0x55
			regs->main.f.val = 0xff;
			SimulateOne();
			Assert::AreEqual(16ull, cpu->Time());
			Assert::AreEqual<uint16_t>(10, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, io_bus.read(0x01FE));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));

			memory.write (regs->main.hl, 0xAA); // data is 0xAA
			regs->main.f.val = 0;
			SimulateOne();
			Assert::AreEqual(32ull, cpu->Time());
			Assert::AreEqual<uint16_t>(9, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, io_bus.read(0x00FE));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}

		TEST_METHOD(otir)
		{
			memory.write (0, { 0xED, 0xB3 }); // OTIR
			memory.write (2, { 0xED, 0xB3 }); // OTIR
			regs->main.hl = 10; // memory address is 10
			regs->main.bc = 0x02FE; // io address is 0xFE

			memory.write (regs->main.hl, 0x55); // data is 0x55
			regs->main.f.val = 0xff;
			SimulateOne();
			Assert::AreEqual(21ull, cpu->Time());
			Assert::AreEqual<uint16_t>(0, cpu->GetPC());
			Assert::AreEqual<uint16_t>(11, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, io_bus.read(0x01FE));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));

			memory.write (regs->main.hl, 0xAA); // data is 0xAA
			SimulateOne();
			Assert::AreEqual(37ull, cpu->Time());
			Assert::AreEqual<uint16_t>(2, cpu->GetPC());
			Assert::AreEqual<uint16_t>(12, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, io_bus.read(0x00FE));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}

		TEST_METHOD(otdr)
		{
			memory.write (0, { 0xED, 0xBB }); // OTDR
			memory.write (2, { 0xED, 0xBB }); // OTDR
			regs->main.hl = 11;
			regs->main.bc = 0x02FE;

			memory.write (regs->main.hl, 0x55); // data is 0x55
			regs->main.f.val = 0xff;
			SimulateOne();
			Assert::AreEqual(21ull, cpu->Time());
			Assert::AreEqual<uint16_t>(0, cpu->GetPC());
			Assert::AreEqual<uint16_t>(10, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x01FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0x55, io_bus.read(0x01FE));
			Assert::AreEqual<uint8_t>(z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));

			memory.write (regs->main.hl, 0xAA); // data is 0xAA
			SimulateOne();
			Assert::AreEqual(37ull, cpu->Time());
			Assert::AreEqual<uint16_t>(2, cpu->GetPC());
			Assert::AreEqual<uint16_t>(9, regs->main.hl);
			Assert::AreEqual<uint16_t>(0x00FE, regs->main.bc);
			Assert::AreEqual<uint8_t>(0xAA, io_bus.read(0x00FE));
			Assert::AreEqual<uint8_t>(z80_flag::z | z80_flag::n | z80_flag::c, regs->main.f.val & ~(z80_flag::h | z80_flag::pv));
		}
	};
}
