
#include "CppUnitTest.h"
#include "Impl/Z80CPU.h"
#include "Simulator.h"
#include "shared/string_builder.h"

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

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return false; }

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

namespace Z80SimulatorTests
{
	TEST_CLASS(Z80SimulatorTests)
	{
		Bus memory;
		Bus io_bus;
		dummy_irq_line irq_line;
		wistd::unique_ptr<IZ80CPU> cpu;
		wistd::unique_ptr<IDevice> ram;
		z80_register_set* regs;

	public:
		Z80SimulatorTests()
		{
			auto hr = MakeZ80CPU (&memory, &io_bus, &irq_line, &cpu); THROW_IF_FAILED(hr);
			regs = cpu->GetRegsPtr();

			auto r = wil::make_unique_nothrow<TestRAM>(); THROW_IF_NULL_ALLOC(r);
			hr = r->InitInstance(&memory); THROW_IF_FAILED(hr);
			ram = std::move(r);
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

		TEST_METHOD(bit_1_b)
		{
			memory.write(0, { 0xCB, 0x48 }); // bit 1, b
			regs->b()= 0xFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(2, regs->pc);
			Assert::AreEqual<uint64_t>(8, cpu->Time());
			Assert::AreEqual<uint8_t>(0, regs->main.f.z);

			memory.write(2, { 0xCB, 0x48 }); // bit 1, b
			regs->b()= 0;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(4, regs->pc);
			Assert::AreEqual<uint64_t>(16, cpu->Time());
			Assert::AreEqual<uint8_t>(1, regs->main.f.z);
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
			memory.write(0, 9);
			regs->main.bc = 0x0FFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0x0FFF, regs->main.hl);
			Assert::AreEqual<uint8_t>(0, regs->main.f.val);

			regs->pc = 0;
			regs->main.hl = 1;
			regs->main.bc = 0xFFFF;
			cpu->SimulateOne(nullptr);
			Assert::AreEqual<uint16_t>(0, regs->main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::c | z80_flag::h, regs->main.f.val);
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
			auto* regs = cpu->GetRegsPtr();
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
	};
}
