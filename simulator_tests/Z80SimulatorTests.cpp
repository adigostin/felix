
#include "CppUnitTest.h"
#include "cpu.h"
#include "memory_bus.h"
#include "shared/string_builder.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

template<>
std::wstring Microsoft::VisualStudio::CppUnitTestFramework::ToString<uint16_t>(const uint16_t& q)
{
	wstring_builder<50> sb;
	sb << hex<uint16_t>(q);
	return { sb.data(), sb.size() };
}

struct memory_bus_256B_ram : memory_bus_i
{
	using memory_bus_i::read;
	using memory_bus_i::write;

	device_i* const _active_devices[1];
	uint8_t _ram[256];

	memory_bus_256B_ram (cpu_device_i* cpu)
		: _active_devices{ cpu }
	{ }

	virtual void reset() override { }

	virtual uint8_t read(uint16_t address) const override
	{
		WI_ASSERT(address < std::size(_ram));
		return _ram[address];
	}

	virtual void write(uint16_t address, uint8_t value) override
	{
		WI_ASSERT(address < std::size(_ram));
		_ram[address] = value;
	}

	virtual device_i* const* read_responders (uint16_t address, uint8_t* count_out) const override
	{
		*count_out = 0;
		return nullptr;
	}

	virtual std::span<device_i* const> active_devices_connected(uint16_t address) const override
	{
		return _active_devices;
	}

	virtual std::span<device_i* const> write_requesters (uint16_t address) const override
	{
		return _active_devices;
	}
};

struct dummy_io_bus : io_bus_i
{
	virtual void reset() override { }

	virtual uint8_t read(uint16_t address) const override { return 0xff; }
	
	virtual void write(uint16_t address, uint8_t value) override { }

	virtual device_i* const* read_responders (uint16_t address, uint8_t* count_out) const override
	{
		*count_out = 0;
		return nullptr;
	}

	virtual std::span<device_i* const> active_devices_connected(uint16_t address) const override
	{
		return { };
	}

	virtual std::span<device_i* const> write_requesters (uint16_t address) const override
	{
		return { };
	}
};

struct dummy_irq_line : irq_line_i
{
	virtual std::span<interrupting_device_i* const> interrupting_devices() const override
	{
		return { };
	}
};

namespace Z80SimulatorTests
{
	TEST_CLASS(Z80SimulatorTests)
	{
		std::unique_ptr<cpu_device_i> cpu;
		memory_bus_256B_ram memory;
		dummy_io_bus io_bus;
		dummy_irq_line irq_line;
		z80_register_set& regs;

	public:
		Z80SimulatorTests()
			: cpu(make_cpu(&memory, &io_bus, &irq_line))
			, memory(cpu.get())
			, regs(cpu->get_regs())
		{ }

		TEST_METHOD_INITIALIZE(MethodInitialize)
		{
			cpu->reset_device_and_time();
			memset (memory._ram, 0, sizeof(memory._ram));
		}

		TEST_METHOD(Test_NOP)
		{
			memory.write(0, 0); // NOP
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(4, cpu->time());
		}

		TEST_METHOD(test_xor_b)
		{
			regs.main.a = 0xff;
			regs.b() = 0xAA;
			memory.write(0, 0xA8); // xor b
			cpu->simulate_one(false);
			Assert::AreEqual<uint64_t>(4, cpu->time());
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint8_t>(0x55, regs.main.a);
			Assert::AreEqual<uint8_t>(0, regs.main.f.s);
			Assert::AreEqual<uint8_t>(0, regs.main.f.z);
			Assert::AreEqual<uint8_t>(0, regs.main.f.h);
			//Assert::AreEqual<uint8_t>(1, regs.main.f.pv); not sure about this one
			Assert::AreEqual<uint8_t>(0, regs.main.f.n);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);

			regs.main.a = 0xff;
			regs.b() = 0x55;
			memory.write(1, 0xA8); // xor b
			cpu->simulate_one(false);
			Assert::AreEqual<uint64_t>(8, cpu->time());
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint8_t>(0xAA, regs.main.a);
			Assert::AreEqual<uint8_t>(1, regs.main.f.s);
			Assert::AreEqual<uint8_t>(0, regs.main.f.z);
			Assert::AreEqual<uint8_t>(0, regs.main.f.h);
			//Assert::AreEqual<uint8_t>(1, regs.main.f.pv); not sure about this one
			Assert::AreEqual<uint8_t>(0, regs.main.f.n);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
		}

		TEST_METHOD(test_xor_hl)
		{
			regs.main.a = 0xff;
			regs.main.hl = 0x10;
			memory.write(0x10, 0x55);
			memory.write(0, 0xAe); // xor (hl)
			cpu->simulate_one(false);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint8_t>(0xAA, regs.main.a);
			Assert::AreEqual<uint8_t>(1, regs.main.f.s);
			Assert::AreEqual<uint8_t>(0, regs.main.f.z);
			Assert::AreEqual<uint8_t>(0, regs.main.f.h);
			//Assert::AreEqual<uint8_t>(1, regs.main.f.pv); not sure about this one
			Assert::AreEqual<uint8_t>(0, regs.main.f.n);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
		}

		TEST_METHOD(test_add_ix_plus_disp)
		{
			regs.main.a = 0x80;
			regs.ix = 0;
			memory.write(0, { 0xDD, 0x86, 0x10 }); // add (ix + 10h)
			memory.write(0x10, 0x33);
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0xB3, regs.main.a);
			Assert::AreEqual<uint16_t>(3, regs.pc);
			Assert::AreEqual<uint64_t>(19, cpu->time());
		}

		TEST_METHOD(test_ld_b_a)
		{
			regs.main.a = 0x55;
			memory.write(0, 0x47); // ld b, a
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(4, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, regs.b());
		}

		TEST_METHOD(test_ld_a_hl)
		{
			regs.main.a = 0;
			regs.main.hl = 0x10;
			memory.write(0x10, 0x55);
			memory.write(0, 0x7e); // ld a, (hl)
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, regs.main.a);
		}

		TEST_METHOD(test_ld_hl_a)
		{
			regs.main.a = 0x55;
			regs.main.hl = 0x10;
			memory.write(0, 0x77); // ld (hl), a
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(djnz)
		{
			memory.write(0, { 0x10, (uint8_t)-2 }); // label: djnz label
			regs.b() = 10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0, regs.pc);
			Assert::AreEqual<uint64_t>(13, cpu->time());
			Assert::AreEqual<uint8_t>(9, regs.b());

			cpu->simulate_one(false); cpu->simulate_one(false); cpu->simulate_one(false); cpu->simulate_one(false);
			cpu->simulate_one(false); cpu->simulate_one(false); cpu->simulate_one(false); cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0, regs.pc);
			Assert::AreEqual<uint64_t>(117, cpu->time());
			Assert::AreEqual<uint8_t>(1, regs.b());

			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(125, cpu->time());
			Assert::AreEqual<uint8_t>(0, regs.b());
		}

		TEST_METHOD(set_1_b)
		{
			memory.write(0, { 0xCB, 0xC8 }); // set 1, b
			regs.b() = 0;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());
			Assert::AreEqual<uint8_t>(2, regs.b());
		}

		TEST_METHOD(set_1_hl)
		{
			memory.write(0, { 0xCB, 0xCE }); // set 1, (hl)
			regs.main.hl = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(15, cpu->time());
			Assert::AreEqual<uint8_t>(2, memory.read(0x10));
		}

		TEST_METHOD(set_1_ix_plus_disp)
		{
			memory.write(0, { 0xDD, 0xCB, 0x10, 0xCE }); // set 1, (ix + 10h)
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(23, cpu->time());
			Assert::AreEqual<uint8_t>(2, memory.read(0x10));
		}

		TEST_METHOD(res_1_b)
		{
			memory.write(0, { 0xCB, 0x88 }); // res 1, b
			regs.b()= 0xFF;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());
			Assert::AreEqual<uint8_t>(0xFD, regs.b());
		}

		TEST_METHOD(res_1_hl)
		{
			memory.write(0, { 0xCB, 0x8E }); // res 1, (hl)
			regs.main.hl = 0x10;
			memory.write(0x10, 0xFF);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(15, cpu->time());
			Assert::AreEqual<uint8_t>(0xFD, memory.read(0x10));
		}
		
		TEST_METHOD(res_1_ix_plus_disp)
		{
			memory.write(0, { 0xDD, 0xCB, 0x10, 0x8E }); // res 1, (ix + 10h)
			memory.write(0x10, 0xFF);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(23, cpu->time());
			Assert::AreEqual<uint8_t>(0xFD, memory.read(0x10));
		}

		TEST_METHOD(bit_1_b)
		{
			memory.write(0, { 0xCB, 0x48 }); // bit 1, b
			regs.b()= 0xFF;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());
			Assert::AreEqual<uint8_t>(0, regs.main.f.z);

			memory.write(2, { 0xCB, 0x48 }); // bit 1, b
			regs.b()= 0;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(16, cpu->time());
			Assert::AreEqual<uint8_t>(1, regs.main.f.z);
		}

		TEST_METHOD(bit_1_hl)
		{
			memory.write(0, { 0xCB, 0x4E }); // bit 1, (hl)
			regs.main.hl = 0x10;
			memory.write(0x10, 0xFF);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(12, cpu->time());
			Assert::AreEqual<uint8_t>(0, regs.main.f.z);

			memory.write(2, { 0xCB, 0x4E }); // bit 1, (hl)
			regs.main.hl = 0x10;
			memory.write(0x10, 0);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(24, cpu->time());
			Assert::AreEqual<uint8_t>(1, regs.main.f.z);
		}
		
		TEST_METHOD(bit_1_ix_plus_disp)
		{
			memory.write(0, { 0xDD, 0xCB, 0x10, 0x4E }); // bit 1, (ix+16)
			memory.write(0x10, 0xFF);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(20, cpu->time());
			Assert::AreEqual<uint8_t>(0, regs.main.f.z);

			memory.write(4, { 0xDD, 0xCB, 0x10, 0x4E }); // bit 1, (ix+16)
			memory.write(0x10, 0);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(8, regs.pc);
			Assert::AreEqual<uint64_t>(40, cpu->time());
			Assert::AreEqual<uint8_t>(1, regs.main.f.z);
		}

		TEST_METHOD(ld_b_55)
		{
			memory.write(0, { 0x06, 0x55 }); // ld b, 55h
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, regs.b());
		}

		TEST_METHOD(ld_memhl_55)
		{
			memory.write(0, { 0x36, 0x55 }); // ld (hl), 55h
			regs.main.hl = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(10, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(ld_memix_disp_55)
		{
			memory.write(0, { 0xDD, 0x36, 0x10, 0x55 }); // ld (ix + 10h), 55h
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(19, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(add_a_0fh)
		{
			memory.write(0, { 0xC6, 0x0F }); // add a, 0fh
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x0F, regs.main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r3, regs.main.f.val);

			regs.pc = 0;
			regs.main.a = 1;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x10, regs.main.a);
			Assert::AreEqual<uint8_t>(z80_flag::h, regs.main.f.val);

			regs.pc = 0;
			regs.main.a = 0x100 - 0xF;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0, regs.main.a);
			Assert::AreEqual<uint8_t>(z80_flag::h | z80_flag::c | z80_flag::z, regs.main.f.val);

			regs.pc = 0;
			regs.main.a = 0x7F; // max signed int (127)
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x8E, regs.main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r3 | z80_flag::s | z80_flag::pv | z80_flag::h, regs.main.f.val);

			regs.pc = 0;
			regs.main.a = 0x80; // min signed int (-128)
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x8F, regs.main.a);
			Assert::AreEqual<uint8_t>(z80_flag::r3 | z80_flag::s, regs.main.f.val);
		}

		TEST_METHOD(add_a_35h)
		{
			memory.write(0, { 0xC6, 0x35 }); // add a, 35h
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x35, regs.main.a);
			Assert::AreEqual<uint8_t>(0, regs.main.f.z);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
		}

		TEST_METHOD(add_a_35h_with_carry_and_zero)
		{
			regs.main.a = 0x100 - 0x35;
			memory.write(0, { 0xC6, 0x35 }); // add a, 35h
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0, regs.main.a);
			Assert::AreEqual<uint8_t>(1, regs.main.f.z);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
		}

		TEST_METHOD(sbc_a_nn)
		{
			regs.main.a = 0;
			memory.write(0, { 0xDE, 0xFF }); // sbc a, 0ffh
			regs.main.f.val = z80_flag::c;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0, regs.main.a);
			Assert::AreEqual<uint8_t>(z80_flag::c | z80_flag::n | z80_flag::h | z80_flag::z, regs.main.f.val);

			regs.pc = 0;
			regs.main.a = 0;
			regs.main.f.val = 0;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint8_t>(1, regs.main.a);
			Assert::AreEqual<uint8_t>(z80_flag::c | z80_flag::n | z80_flag::h, regs.main.f.val);
		}

		TEST_METHOD(push_bc)
		{
			memory.write(0, 0xC5); // push bc
			regs.main.bc = 0x55AA;
			regs.sp = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint16_t>(0xE, regs.sp);
			Assert::AreEqual<uint16_t>(0x55AA, memory.read_uint16(regs.sp));
			Assert::AreEqual<uint64_t>(11, cpu->time());
		}

		TEST_METHOD(push_ix)
		{
			memory.write(0, { 0xDD, 0xE5 }); // push ix
			regs.ix = 0x55AA;
			regs.sp = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint16_t>(0xE, regs.sp);
			Assert::AreEqual<uint16_t>(0x55AA, memory.read_uint16(regs.sp));
			Assert::AreEqual<uint64_t>(15, cpu->time());
		}

		TEST_METHOD(push_af)
		{
			memory.write(0, 0xF5); // push af
			regs.main.af = 0x55AA;
			regs.sp = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint16_t>(0xE, regs.sp);
			Assert::AreEqual<uint16_t>(0x55AA, memory.read_uint16(regs.sp));
			Assert::AreEqual<uint64_t>(11, cpu->time());
		}

		TEST_METHOD(pop_bc)
		{
			memory.write(0, 0xC1); // pop bc
			regs.sp = 0xE;
			memory.write_uint16(0xE, 0x1234);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint16_t>(0x10, regs.sp);
			Assert::AreEqual<uint16_t>(0x1234, regs.main.bc);
			Assert::AreEqual<uint64_t>(10, cpu->time());
		}
		
		TEST_METHOD(pop_ix)
		{
			memory.write(0, { 0xDD, 0xE1 }); // pop ix
			regs.sp = 0xE;
			memory.write_uint16(0xE, 0x2345);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint16_t>(0x10, regs.sp);
			Assert::AreEqual<uint16_t>(0x2345, regs.ix);
			Assert::AreEqual<uint64_t>(14, cpu->time());
		}
		
		TEST_METHOD(pop_af)
		{
			memory.write(0, 0xF1); // pop af
			regs.sp = 0xE;
			memory.write_uint16(0xE, 0x3456);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint16_t>(0x10, regs.sp);
			Assert::AreEqual<uint16_t>(0x3456, regs.main.af);
			Assert::AreEqual<uint64_t>(10, cpu->time());
		}

		TEST_METHOD(rlca)
		{
			memory.write(0, 7); // rlca
			regs.main.a = 0x55;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0xAA, regs.main.a);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(4, cpu->time());

			memory.write(1, 7); // rlca
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x55, regs.main.a);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());
		}

		TEST_METHOD(rrca)
		{
			memory.write(0, 15); // rrca
			regs.main.a = 0x55;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0xAA, regs.main.a);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(4, cpu->time());

			memory.write(1, 15); // rrca
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x55, regs.main.a);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());
		}

		TEST_METHOD(rla)
		{
			memory.write(0, 0x17); // rla
			regs.main.a = 0x55;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0xAA, regs.main.a);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(4, cpu->time());

			memory.write(1, 0x17); // rla
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x54, regs.main.a);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());
		}
		
		TEST_METHOD(rra)
		{
			memory.write(0, 0x1f); // rra
			regs.main.a = 0x55;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x2A, regs.main.a);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(4, cpu->time());

			memory.write(1, 0x1f); // rra
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x95, regs.main.a);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());
		}

		TEST_METHOD(ex_sp_hl)
		{
			memory.write(0, 0xE3); // ex (sp), hl
			regs.main.hl = 0x1234;
			regs.sp = 0x10;
			memory.write_uint16 (0x10, 0x55AA);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(19, cpu->time());
			Assert::AreEqual<uint16_t>(0x55AA, regs.main.hl);
			Assert::AreEqual<uint16_t>(0x1234, memory.read_uint16(0x10));
		}

		TEST_METHOD(ex_sp_iy)
		{
			memory.write(0, { 0xFD, 0xE3 }); // ex (sp), iy
			regs.iy = 0x1234;
			regs.sp = 0x10;
			memory.write_uint16 (0x10, 0x55AA);
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(23, cpu->time());
			Assert::AreEqual<uint16_t>(0x55AA, regs.iy);
			Assert::AreEqual<uint16_t>(0x1234, memory.read_uint16(0x10));
		}

		TEST_METHOD(ld_a_membc)
		{
			memory.write(0, 0x0a); // ld a, (bc)
			memory.write(0x10, 0x55);
			regs.main.bc = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, regs.main.a);
		}

		TEST_METHOD(ld_a_memde)
		{
			memory.write(0, 0x1a); // ld a, (de)
			memory.write(0x10, 0x55);
			regs.main.de = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, regs.main.a);
		}

		TEST_METHOD(ld_membc_a)
		{
			memory.write(0, 2); // ld (bc), a
			regs.main.a = 0x55;
			regs.main.bc = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD(ld_memde_a)
		{
			memory.write(0, 0x12); // ld (de), a
			regs.main.a = 0x55;
			regs.main.de = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(1, regs.pc);
			Assert::AreEqual<uint64_t>(7, cpu->time());
			Assert::AreEqual<uint8_t>(0x55, memory.read(0x10));
		}

		TEST_METHOD (rst_n)
		{
			for (uint8_t i = 0; i < 8; i++)
			{
				regs.pc = 0x80;
				regs.sp = 0x10;
				memory.write(0x80, (i << 3) | 0xC7); // RST n
				cpu->simulate_one(false);
				Assert::AreEqual<uint16_t>((i << 3), regs.pc);
				Assert::AreEqual<uint16_t>(0x0E, regs.sp);
				Assert::AreEqual<uint16_t>(0x81, memory.read_uint16(0x0E));
				Assert::AreEqual<uint64_t>(11 * (i + 1), cpu->time());
			}
		}

		TEST_METHOD (add_hl_bc)
		{
			memory.write(0, 9);
			regs.main.bc = 0x0FFF;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0x0FFF, regs.main.hl);
			Assert::AreEqual<uint8_t>(0, regs.main.f.val);

			regs.pc = 0;
			regs.main.hl = 1;
			regs.main.bc = 0xFFFF;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0, regs.main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::c | z80_flag::h, regs.main.f.val);
		}

		TEST_METHOD(adc_hl_de)
		{
			memory.write(0, { 0xED, 0x5A }); // ADC HL, DE
			regs.main.hl = 0x8000;
			regs.main.de = 0x7FFF;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(15, cpu->time());
			Assert::AreEqual<uint16_t>(0xFFFF, regs.main.hl);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);

			regs.pc = 0;
			regs.main.hl = 0x8000;
			regs.main.de = 0x8000;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0, regs.main.hl);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);

			regs.pc = 0;
			regs.main.hl = 0x8000;
			regs.main.de = 0x7FFF;
			regs.main.f.c = 1;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0, regs.main.hl);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
		}

		TEST_METHOD(sbc_hl_de)
		{
			memory.write(0, { 0xED, 0x52 }); // SBC HL, DE
			regs.main.hl = 0x8000;
			regs.main.de = 0x8000;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(15, cpu->time());
			Assert::AreEqual<uint16_t>(0, regs.main.hl);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);

			regs.pc = 0;
			regs.main.hl = 0x7FFF;
			regs.main.de = 0x8000;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0xFFFF, regs.main.hl);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);

			regs.pc = 0;
			regs.main.hl = 0x8000;
			regs.main.de = 0x8000;
			regs.main.f.c = 1;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0xFFFF, regs.main.hl);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);

			regs.pc = 0;
			regs.main.hl = 0;
			regs.main.de = 0x7fff;
			regs.main.f.val = z80_flag::c;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0x8000, regs.main.hl);
			Assert::AreEqual<uint8_t>(z80_flag::s | z80_flag::h | z80_flag::n | z80_flag::c, regs.main.f.val);
		}

		TEST_METHOD(rrc_c)
		{
			memory.write(0, { 0xCB, 0x09 }); // rrc c
			regs.main.bc = 0x55;
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0xAA, regs.main.bc);
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(8, cpu->time());

			memory.write(2, { 0xCB, 0x09 }); // rrc c
			cpu->simulate_one(false);
			Assert::AreEqual<uint16_t>(0x55, regs.main.bc);
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(16, cpu->time());
		}

		TEST_METHOD(rr_memhl)
		{
			memory.write (0, { 0xCB, 0x1E }); // rr (hl)
			memory.write (0x10, 0x42);
			regs.main.hl = 0x10;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x21, memory.read(0x10));
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(15, cpu->time());

			regs.pc = 0;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x10, memory.read(0x10));
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(30, cpu->time());

			regs.pc = 0;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0x88, memory.read(0x10));
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(2, regs.pc);
			Assert::AreEqual<uint64_t>(45, cpu->time());
		}

		TEST_METHOD(sra_memix)
		{
			memory.write (0, { 0xDD, 0xCB, 0x10, 0x2E }); // sra (ix+10h)
			memory.write (0x10, 0x82);
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0xC1, memory.read(0x10));
			Assert::AreEqual<uint8_t>(0, regs.main.f.c);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(23, cpu->time());

			regs.pc = 0;
			cpu->simulate_one(false);
			Assert::AreEqual<uint8_t>(0xE0, memory.read(0x10));
			Assert::AreEqual<uint8_t>(1, regs.main.f.c);
			Assert::AreEqual<uint16_t>(4, regs.pc);
			Assert::AreEqual<uint64_t>(46, cpu->time());
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
							regs.pc = 0;
							regs.main.a = X;
							regs.main.f.val = 0;
							regs.b() = Y;
							cpu->simulate_one(false);
							if (regs.pc != 1) 
								Assert::Fail();
							if (regs.main.a != (uint8_t)res) 
								Assert::Fail();
							if (regs.main.f.val != flags) 
								Assert::Fail();

							memory.write (0, 0xB8); // cp b
							regs.pc = 0;
							regs.main.a = X;
							regs.main.f.val = 0;
							regs.b() = Y;
							cpu->simulate_one(false);
							if (regs.pc != 1) 
								Assert::Fail();
							if (regs.main.a != X) 
								Assert::Fail();
							if (regs.main.f.val != flags) 
								Assert::Fail();
						}
						
						memory.write (0, 0x98); // sbc b
						regs.pc = 0;
						regs.main.a = X;
						regs.main.f.val = 0;
						regs.main.f.c = C;
						regs.b() = Y;
						cpu->simulate_one(false);
						if (regs.pc != 1) 
							Assert::Fail();
						if (regs.main.a != (uint8_t)res) 
							Assert::Fail();
						if (regs.main.f.val != flags) 
							Assert::Fail();
					}
				}
			}
		}

	};
}
