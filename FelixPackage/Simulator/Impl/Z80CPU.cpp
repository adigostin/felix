
#include "pch.h"
#include "Z80CPU.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/com.h"

// For the P/V flag calculation, the "Parity/Overflow Flag" paragraph in the Z80 pdf has a good explanation.

// From http://z80.info/z80info.htm:
// d) About the R register:
// This is not really an undocumented feature, although I have never seen any thorough description
// of it anywhere. The R register is a counter that is updated every instruction, where DD, FD, ED 
// and CB are to be regarded as separate instructions. So shifted instruction will increase R by two. 
// There's an interesting exception: doubly-shifted opcodes, the DDCB and FDCB ones, increase R by 
// two too. LDI increases R by two, LDIR increases it by 2 times BC, as does LDDR etcetera.
// The sequence LD R,A/LD A,R increases A by two, except for the highest bit: this bit of the R
// register is never changed.

#pragma region pre-calculated values and flags
// From https://github.com/Dotneteer/spectnetide/blob/master/v2/Core/Spect.Net.SpectrumEmu/Cpu/Z80AluHelpers.cs

struct daa_table_t
{
	uint16_t s_DaaResults[0x800];

	constexpr daa_table_t()
	{
		for (uint32_t b = 0; b < 0x100; b++)
		{
			uint32_t hNibble = b >> 4;
			uint32_t lNibble = b & 0x0F;

			for (uint8_t H = 0; H <= 1; H++)
			{
				for (uint8_t N = 0; N <= 1; N++)
				{
					for (uint8_t C = 0; C <= 1; C++)
					{
						// --- Calculate DIFF and the new value of C Flag
						uint8_t diff = 0x00;
						uint8_t cAfter = 0;
						if (C == 0)
						{
							if (hNibble >= 0 && hNibble <= 9 && lNibble >= 0 && lNibble <= 9)
							{
								diff = H == 0 ? 0x00 : 0x06;
							}
							else if (hNibble >= 0 && hNibble <= 8 && lNibble >= 0x0A && lNibble <= 0xF)
							{
								diff = 0x06;
							}
							else if (hNibble >= 0x0A && hNibble <= 0x0F && lNibble >= 0 && lNibble <= 9 && H == 0)
							{
								diff = 0x60;
								cAfter = 1;
							}
							else if (hNibble >= 9 && hNibble <= 0x0F && lNibble >= 0x0A && lNibble <= 0xF)
							{
								diff = 0x66;
								cAfter = 1;
							}
							else if (hNibble >= 0x0A && hNibble <= 0x0F && lNibble >= 0 && lNibble <= 9)
							{
								if (H == 1) diff = 0x66;
								cAfter = 1;
							}
						}
						else
						{
							// C == 1
							cAfter = 1;
							if (lNibble >= 0 && lNibble <= 9)
							{
								diff = H == 0 ? 0x60 : 0x66;
							}
							else if (lNibble >= 0x0A && lNibble <= 0x0F)
							{
								diff = 0x66;
							}
						}

						// --- Calculate new value of H Flag
						uint8_t hAfter = 0;
						if (lNibble >= 0x0A && lNibble <= 0x0F && N == 0
							|| lNibble >= 0 && lNibble <= 5 && N == 1 && H == 1)
						{
							hAfter = 1;
						}

						// --- Calculate new value of register A
						uint8_t A = (N == 0 ? b + diff : b - diff) & 0xFF;

						// --- Calculate other flags
						uint8_t aPar = 0;
						uint8_t val = A;
						for (uint8_t i = 0; i < 8; i++)
						{
							aPar += val & 0x01;
							val >>= 1;
						}

						// --- Calculate result
						uint8_t fAfter =
							(A & z80_flag::r3) |
							(A & z80_flag::r5) |
							((A & 0x80) != 0 ? z80_flag::s : 0) |
							(A == 0 ? z80_flag::z : 0) |
							(aPar % 2 == 0 ? z80_flag::pv : 0) |
							(N == 1 ? z80_flag::n : 0) |
							(hAfter == 1 ? z80_flag::h : 0) |
							(cAfter == 1 ? z80_flag::c : 0);

						uint16_t result = (uint16_t) (A << 8 | (byte) fAfter);
						uint8_t fBefore = (byte) (H * 4 + N * 2 + C);
						uint32_t idx = (fBefore << 8) + b;
						s_DaaResults[idx] = result;
					}
				}
			}
		}
	}
};

static constexpr daa_table_t daa_table;

struct s_z_pv_flags_t
{
	uint8_t flags[0x100];

	s_z_pv_flags_t()
	{
		for (uint32_t a = 0; a < 0x100; a++)
		{
			uint8_t pv = (__popcnt(a) & 1) ? 0 : z80_flag::pv;
			flags[a] = (a & (z80_flag::r3 | z80_flag::r5 | z80_flag::s)) | pv;
		}
		flags[0] |= z80_flag::z;
	}
};

// S is set if the Accumulator is negative after an operation; otherwise, it is reset.
// Z is set if the Accumulator is 0 after an operation; otherwise, it is reset.
// P/V is set if the parity of the Accumulator is even after an operation; otherwise, it is reset.
static const s_z_pv_flags_t s_z_pv_flags;
#pragma endregion

class cpu : public IZ80CPU
{
	using handler_t    = bool(cpu::*)(hl_ix_iy xy, uint8_t opcode);
	using ed_handler_t = bool(cpu::*)(uint8_t opcode);
	using cb_handler_t = bool(cpu::*)(hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode);

	Bus*       memory;
	Bus*       io;
	irq_line_i*  irq;

	UINT64 cpu_time = 0;
	z80_register_set regs = { };

	// The "EI" instruction sets this to 2.
	// After simulate_one() executes any instruction, it checks this value and:
	//  - if 2, it decrements it to 1;
	//  - if 1, it decrements it to 0 and sets IFF in the registers;
	//  - if 0, it leaves it unchanged.
	uint8_t _ei_countdown = 0;

	// The stack starts at SP and ends at this address.
	// Whenever SP is assigned, this is set to the same value;
	uint16_t _start_of_stack = 0;

	SIM_BP_COOKIE _nextBpCookie = 1;
	unordered_map_nothrow<uint16_t, vector_nothrow<SIM_BP_COOKIE>> code_bps;

public:
	HRESULT InitInstance (Bus* memory, Bus* io, irq_line_i* irq)
	{
		this->memory = memory;
		this->io = io;
		this->irq = irq;
		bool pushed = memory->write_requesters.try_push_back(this); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = io->write_requesters.try_push_back(this); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		return S_OK;
	}

	~cpu()
	{
		memory->write_requesters.erase(memory->write_requesters.find(this));
		io->write_requesters.erase(io->write_requesters.find(this));
	}

	uint8_t decode_u8()
	{
		uint8_t val = memory->read(regs.pc);
		regs.pc++;
		return val;
	}

	uint16_t decode_u16()
	{
		uint16_t val = memory->read_uint16(regs.pc);
		regs.pc += 2;
		return val;
	}

	// Calculates the address of (hl), or (ix+d), or (iy+d)
	uint16_t decode_mem_hl (hl_ix_iy xy)
	{
		if (xy == hl_ix_iy::hl)
			return regs.main.hl;
	
		int8_t disp = (int8_t)decode_u8();
		uint16_t addr = (xy == hl_ix_iy::ix) ? regs.ix : regs.iy;
		uint16_t dispu16 = (uint16_t)(int16_t)disp;
		addr = addr + dispu16;
		return addr;
	}

	static __forceinline bool same_sign(uint8_t a, uint8_t b)
	{
		return !((a ^ b) & 0x80);
	}

	static __forceinline bool same_sign(uint16_t a, uint16_t b)
	{
		return !((a ^ b) & 0x8000);
	}

	void do_reg_a_operation (uint8_t operation, uint8_t other)
	{
		uint8_t before = regs.main.a;
		switch (operation)
		{
		case 0: // add
			// For addition, operands with different signs never cause overflow. When adding operands
			// with similar signs and the result contains a different sign, the Overflow Flag is set
			regs.main.a += other;
			regs.main.f.val = (regs.main.a & 0xA8) // S, R5, R3
				| (regs.main.a ? 0 : z80_flag::z)  // Z
				| (((before & 0xF) + (other & 0xF)) & 0x10) // H
				| ((same_sign(before, other) && !same_sign(before, regs.main.a)) ? z80_flag::pv : 0) // P/V
				| 0 // N
				| (regs.main.a < before); // C
			break;

		case 1: // adc
			regs.main.a += (other + regs.main.f.c);
			regs.main.f.val = (regs.main.a & 0xA8)  // S, R5, R3
				| (regs.main.a ? 0 : z80_flag::z)   // Z
				| (((before & 0xF) + (other & 0xF) + regs.main.f.c) & 0x10) // H
				| ((same_sign(before, other) && !same_sign(before, regs.main.a)) ? z80_flag::pv : 0) // P/V
				| 0 // N
				| ((((uint32_t)before + (uint32_t)other + (uint32_t)regs.main.f.c) >> 8) & 1); // C
			break;

		case 2: // sub
			regs.main.a -= other;
			regs.main.f.val = (regs.main.a & 0xA8) // S, R5, R3
				| (regs.main.a ? 0 : z80_flag::z)  // Z
				| (((before & 0x0F) - (other & 0x0F)) & 0x10) // H
				| ((!same_sign(before, other) && !same_sign(before, regs.main.a)) ? z80_flag::pv : 0) // P/V
				| z80_flag::n // N
				| (regs.main.a > before); // C
			break;

		case 3: // sbc
			regs.main.a -= (other + regs.main.f.c);
			regs.main.f.val = (regs.main.a & 0xA8) // S, R5, R3
				| (regs.main.a ? 0 : z80_flag::z)  // Z
				| (((before & 0x0F) - (other & 0x0F) - regs.main.f.c) & 0x10) // H
				| ((!same_sign(before, other) && !same_sign(before, regs.main.a)) ? z80_flag::pv : 0) // P/V
				| z80_flag::n // N
				| ((((uint32_t)before - (uint32_t)other - (uint32_t)regs.main.f.c) >> 8) & 1); // C
			break;

		case 4: // and
			regs.main.a &= other;
			regs.main.f.val = s_z_pv_flags.flags[regs.main.a] | z80_flag::h;
			break;

		case 5: // xor
			regs.main.a ^= other;
			regs.main.f.val = s_z_pv_flags.flags[regs.main.a];
			break;

		case 6: // or
			regs.main.a |= other;
			regs.main.f.val = s_z_pv_flags.flags[regs.main.a];
			break;

		default: // cp
			uint8_t after = regs.main.a - other;
			regs.main.f.val = (after & 0x80) // S
				| (other & 0x28) // R5, R3
				| (after ? 0 : z80_flag::z)  // Z
				| (((before & 0x0F) - (other & 0x0F)) & 0x10) // H
				| ((!same_sign(before, other) && !same_sign(before, after)) ? z80_flag::pv : 0) // P/V
				| z80_flag::n // N
				| (after > before); // C
			break;
		}
	}

	bool condition_met (uint8_t cc) const
	{
		switch(cc)
		{
			case 0:  return !regs.main.f.z;  // nz
			case 1:  return  regs.main.f.z;  // z
			case 2:  return !regs.main.f.c;  // nc
			case 3:  return  regs.main.f.c;  // c
			case 4:  return !regs.main.f.pv; // po
			case 5:  return  regs.main.f.pv; // pe
			case 6:  return !regs.main.f.s;  // p
			default: return  regs.main.f.s;  // m
		}
	}

	// nop
	bool sim_00 (hl_ix_iy xy, uint8_t opcode)
	{
		cpu_time += 4;
		return true;
	}

	// ld bc/de/hl/sp, nn
	bool sim_01 (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t val = decode_u16();
		uint8_t i = (opcode & 0x30) >> 4;
		regs.bc_de_hl_sp (xy, i) = val;
		if (i == 3)
			_start_of_stack = val;
		cpu_time += 10;
		return true;
	}

	// ld (bc)/(de), a
	bool sim_02 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 1;
		uint16_t addr = i ? regs.main.de : regs.main.bc;
		if (!memory->try_write_request(addr, regs.main.a, cpu_time))
			return false;
		cpu_time += 7;
		return true;
	}

	// inc bc/de/hl/sp
	bool sim_03 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		regs.bc_de_hl_sp(xy, i)++;
		cpu_time += ((xy == hl_ix_iy::hl) ? 6 : 10);
		return true;
	}
	
	// inc r
	bool sim_04 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 3) & 7;
		uint8_t before;
		uint8_t after;
		if (i != 6)
		{
			before = regs.r8(i, xy);
			after = before + 1;
			regs.r8(i, xy) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 4 : 8);
		}
		else
		{
			uint16_t addr = decode_mem_hl(xy);
			if (!memory->try_read_request(addr, before, cpu_time))
				return false;
			after = before + 1;
			memory->write (addr, after);
			cpu_time += ((xy == hl_ix_iy::hl) ? 11 : 23);
		}

		regs.main.f.val = (after & 0xA8) // S, X5, X3
			| (after ? 0 : z80_flag::z) // Z
			| (((before & 0x10) != (after & 0x10)) ? z80_flag::h : 0) // H
			| ((before == 0x7f) ? z80_flag::pv : 0) // P/V
			| (regs.main.f.val & 1); // C
		
		return true;
	}

	// dec r
	bool sim_05 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 3) & 7;
		uint8_t before;
		uint8_t after;
		if (i != 6)
		{
			before = regs.r8(i, xy);
			after = before - 1;
			regs.r8(i, xy) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 4 : 8);
		}
		else
		{
			uint16_t addr = decode_mem_hl(xy);
			if (!memory->try_read_request(addr, before, cpu_time))
				return false;
			after = before - 1;
			memory->write (addr, after);
			cpu_time += ((xy == hl_ix_iy::hl) ? 11 : 23);
		}

		regs.main.f.val = (after & 0xA8) // S, X5, X3
			| (after ? 0 : z80_flag::z) // Z
			| (((before & 0x10) != (after & 0x10)) ? z80_flag::h : 0) // H
			| ((before == 0x80) ? z80_flag::pv : 0) // P/V
			| z80_flag::n // N
			| (regs.main.f.val & 1); // C
		return true;
	}

	// ld b/c/d/e/h/l/(hl)/a, n
	bool sim_06 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 3) & 7;
		if (i != 6)
		{
			uint8_t val = decode_u8();
			regs.r8(i, xy) = val;
			cpu_time += ((xy == hl_ix_iy::hl) ? 7 : 11);
		}
		else
		{
			uint16_t addr = decode_mem_hl(xy);
			uint8_t val = decode_u8();
			if (!memory->try_write_request(addr, val, cpu_time))
				return false;
			cpu_time += (xy == hl_ix_iy::hl) ? 10 : 19;
		}

		return true;
	}

	// rlca
	bool sim_07 (hl_ix_iy xy, uint8_t opcode)
	{
		regs.main.a = (regs.main.a << 1) | (regs.main.a >> 7);
		regs.main.f.h = 0;
		regs.main.f.n = 0;
		regs.main.f.c = regs.main.a & 1;
		cpu_time += 4;
		return true;
	}

	// ex af, af'
	bool sim_08 (hl_ix_iy xy, uint8_t opcode)
	{
		std::swap(regs.main.af, regs.alt.af);
		cpu_time += 4;
		return true;
	}

	// add hl, bc/de/hl/sp
	bool sim_09 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		uint16_t before = regs.hl(xy);
		uint16_t other = regs.bc_de_hl_sp(xy, i);
		uint16_t after = before + other;
		regs.hl(xy) = after;
		cpu_time += ((xy != hl_ix_iy::hl) ? 15 : 11);
		regs.main.f.val = (regs.main.f.val & (z80_flag::s | z80_flag::z | z80_flag::pv)) // S, Z, P/V
			| ((after >> 8) & (z80_flag::r3 | z80_flag::r5)) // R3, R5
			| ((((before & 0x0FFF) + (other & 0x0FFF)) >> 8) & z80_flag::h) // H
			| 0 // N
			| ((after < before) ? z80_flag::c : 0); // C
		return true;
	}

	// ld a, (bc)/(de)
	bool sim_0a (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 1;
		uint16_t addr = i ? regs.main.de : regs.main.bc;
		uint8_t value;
		if (!memory->try_read_request(addr, value, cpu_time))
			return false;
		regs.main.a = value;
		cpu_time += 7;
		return true;
	}

	// dec bc/de/hl/sp
	bool sim_0b (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		regs.bc_de_hl_sp(xy, i)--;
		cpu_time += ((xy == hl_ix_iy::hl) ? 6 : 10);
		return true;
	}

	// rrca
	bool sim_0f (hl_ix_iy xy, uint8_t opcode)
	{
		regs.main.a = (regs.main.a >> 1) | (regs.main.a << 7);
		regs.main.f.h = 0;
		regs.main.f.n = 0;
		regs.main.f.c = (regs.main.a & 0x80) ? 1 : 0;
		cpu_time += 4;
		return true;
	}

	// djnz e
	bool sim_10 (hl_ix_iy xy, uint8_t opcode)
	{
		int8_t e = (int8_t)decode_u8();
		regs.b()--;
		if (regs.b())
		{
			regs.pc = regs.pc + (uint16_t)(int16_t)e;
			cpu_time += 13;
		}
		else
			cpu_time += 8;
		return true;
	}

	// rla
	bool sim_17 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t before = regs.main.a;
		regs.main.a = (regs.main.a << 1) | regs.main.f.c;
		regs.main.f.h = 0;
		regs.main.f.n = 0;
		regs.main.f.c = (before & 0x80) ? 1 : 0;
		cpu_time += 4;
		return true;
	}

	// jr e
	bool sim_18 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t e = decode_u8();
		regs.pc = regs.pc + (uint16_t)(int16_t)(int8_t)e;
		cpu_time += 12;
		return true;
	}

	// rra
	bool sim_1f (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t before = regs.main.a;
		regs.main.a = (regs.main.a >> 1) | (regs.main.f.c ? 0x80 : 0);
		regs.main.f.h = 0;
		regs.main.f.n = 0;
		regs.main.f.c = before & 1;
		cpu_time += 4;
		return true;
	}

	// jr nz/z/nc/c, e
	bool sim_20 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t c = (opcode >> 3) & 3;
		uint8_t e = decode_u8();
		bool jump;
		switch(c)
		{
			case 0: jump = !regs.main.f.z; break; // nz
			case 1: jump =  regs.main.f.z; break; // z
			case 2: jump = !regs.main.f.c; break; // nc
			case 3: jump =  regs.main.f.c; break; // c
		}

		if (jump)
		{
			regs.pc = regs.pc + (uint16_t)(int16_t)(int8_t)e;
			cpu_time += 12;
		}
		else
			cpu_time += 7;
		return true;
	}

	// ld (nn), hl
	bool sim_22 (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t addr = decode_u16();
		if (!memory->try_write_request (addr, regs.hl(xy), cpu_time))
			return false;
		cpu_time += 16;
		return true;
	}

	// daa
	bool sim_27 (hl_ix_iy xy, uint8_t opcode)
	{
		size_t daaIndex = regs.main.a + (((regs.main.f.val & 3) + ((regs.main.f.val >> 2) & 4)) << 8);
		regs.main.af = daa_table.s_DaaResults[daaIndex];
		return true;
	}

	// ld hl, (nn)
	bool sim_2a (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t addr = decode_u16();
		uint16_t val;
		if (!memory->try_read_request (addr, val, cpu_time))
			return false;
		regs.hl(xy) = val;
		cpu_time += 16;
		return true;
	}

	// cpl
	bool sim_2f (hl_ix_iy xy, uint8_t opcode)
	{
		regs.main.a = ~regs.main.a;
		regs.main.f.h = 1;
		regs.main.f.n = 1;
		regs.main.f.val = regs.main.f.val & ~(z80_flag::r3 | z80_flag::r5)
			| (regs.main.a & (z80_flag::r3 | z80_flag::r5));
		cpu_time += 4;
		return true;
	}

	// ld (nn), a
	bool sim_32 (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t addr = decode_u16();
		if (!memory->try_write_request (addr, regs.main.a, cpu_time))
			return false;
		cpu_time += 13;
		return true;
	}

	// scf
	bool sim_37 (hl_ix_iy xy, uint8_t opcode)
	{
		regs.main.f.h = 0;
		regs.main.f.n = 0;
		regs.main.f.c = 1;
		regs.main.f.val = regs.main.f.val & ~(z80_flag::r3 | z80_flag::r5)
			| (regs.main.a & (z80_flag::r3 | z80_flag::r5));
		cpu_time += 4;
		return true;
	}

	// la a, (nn)
	bool sim_3a (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t addr = decode_u16();
		uint8_t val;
		if (!memory->try_read_request(addr, val, cpu_time))
			return false;
		regs.main.a = val;
		cpu_time += 13;
		return true;
	}

	// ccf
	bool sim_3f (hl_ix_iy xy, uint8_t opcode)
	{
		regs.main.f.h = regs.main.f.c;
		regs.main.f.n = 0;
		regs.main.f.c = !regs.main.f.c;
		regs.main.f.val = regs.main.f.val & ~(z80_flag::r3 | z80_flag::r5)
			| (regs.main.a & (z80_flag::r3 | z80_flag::r5));
		cpu_time += 4;
		return true;
	}

	// ld r, r
	bool sim_40 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t dst = (opcode >> 3) & 7;
		uint8_t src = opcode & 7;
		if (dst == 6)
		{
			// ld (hl)/(ix+d)/(iy+d), r8
			uint8_t value = regs.r8(src, hl_ix_iy::hl); // source reg is here the regular h/l; the DD/FD prefix applies only to the destination
			uint16_t dest_addr = decode_mem_hl(xy);
			if (!memory->try_write_request(dest_addr, value, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 7 : 19);
		}
		else if (src == 6)
		{
			// ld r8, (hl)/(ix+d)/(iy+d)
			uint16_t src_addr = decode_mem_hl(xy);
			uint8_t value;
			if (!memory->try_read_request(src_addr, value, cpu_time))
				return false;
			regs.r8(dst, hl_ix_iy::hl) = value; // dest reg is here the regular h/l; the DD/FD prefix applies only to the source
			cpu_time += ((xy == hl_ix_iy::hl) ? 7 : 19);
		}
		else
		{
			regs.r8(dst, xy) = regs.r8(src, xy); // prefix applies to both source and dest
			cpu_time += ((xy == hl_ix_iy::hl) ? 4 : 8);
		}
		return true;
	}

	// halt
	bool sim_76 (hl_ix_iy xy, uint8_t opcode)
	{
		regs.halted = true;
		cpu_time += 4;
		return true;
	}

	// add/adc/sub/sbc/and/xor/or/cp b/c/d/e/h/l/(hl)/a
	bool sim_80 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t operation = (opcode >> 3) & 7;
		uint8_t other;
		if (i != 6)
		{
			other = regs.r8(i, xy);
			cpu_time += ((xy == hl_ix_iy::hl) ? 4 : 8);
		}
		else
		{
			uint16_t addr = decode_mem_hl(xy);
			if (!memory->try_read_request(addr, other, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 7 : 19);
		}

		do_reg_a_operation (operation, other);
		return true;
	}

	// ret nz/z/nc/c/po/pe/p/m
	bool sim_c0 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t cc = (opcode >> 3) & 7;
		if (condition_met(cc))
		{
			uint16_t addr = memory->read_uint16(regs.sp);
			regs.sp += 2;
			regs.pc = addr;
			cpu_time += 11;
		}
		else
			cpu_time += 5;
		return true;
	}

	// pop bc/de/hl/af
	bool sim_c1 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		// Some games write to video memory using push/pop. We have to simulate this, not write through.
		uint16_t value;
		if (!memory->try_read_request(regs.sp, value, cpu_time))
			return false;
		regs.bc_de_hl_af(xy, i) = value;
		regs.sp += 2;
		cpu_time += ((xy == hl_ix_iy::hl) ? 10 : 14);
		return true;
	}

	// jp nz/z/nc/c/po/pe/p/m, nn
	bool sim_c2 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t cc = (opcode >> 3) & 7;
		uint16_t addr = decode_u16();
		if (condition_met(cc))
			regs.pc = addr;
		cpu_time += 10;
		return true;
	}

	// jp nn
	bool sim_c3 (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t val = decode_u16();
		regs.pc = val;
		cpu_time += 10;
		return true;
	}

	// call nz/z/nc/c/po/pe/p/m, nn
	bool sim_c4 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t cc = (opcode >> 3) & 7;
		uint16_t addr = decode_u16();
		if (condition_met(cc))
		{
			// Let's not go too far with the simulation and do a simple write of the return address.
			memory->write_uint16 (regs.sp - 2, regs.pc);
			regs.sp -= 2;
			regs.pc = addr;
			cpu_time += 17;
		}
		else
			cpu_time += 10;

		return true;
	}

	// push bc/de/hl/af
	bool sim_c5 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		uint16_t value = regs.bc_de_hl_af(xy, i);
		// Some games write to video memory using push/pop. We have to simulate this, not write through.
		if (!memory->try_write_request(regs.sp - 2, value, cpu_time))
			return false;
		regs.sp -= 2;
		cpu_time += ((xy == hl_ix_iy::hl) ? 11 : 15);
		return true;
	}

	// add/adc/sub/sbc/and/xor/or/cp imm8
	bool sim_c6 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t other = decode_u8();
		uint8_t operation = (opcode >> 3) & 7;
		do_reg_a_operation (operation, other);
		cpu_time += 7;
		return true;
	}

	// rst 0/8/10h/...
	bool sim_c7 (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t addr = opcode & 0x38;
		// Let's not go too far with the simulation and do a simple write of the return address.
		memory->write_uint16 (regs.sp - 2, regs.pc);
		regs.sp -= 2;
		regs.pc = addr;
		cpu_time += 11;
		return true;
	}

	// ret
	bool sim_c9 (hl_ix_iy xy, uint8_t opcode)
	{
		if (regs.sp == _start_of_stack)
		{
			// The Z80 code does a RET while already at what we believe to be the starting (highest) address of the stack.
			// In this case it's very likely we have a wrong assumption about the starting address of the stack.
			// So let's adjust it.
			// This scenario happens when we simulate the code at 0x1030 from the Spectrum 48K ROM.
			_start_of_stack += 2;
		}

		// Let's not go too far with the simulation and do a simple read of the return address.
		uint16_t addr = memory->read_uint16(regs.sp);
		regs.sp += 2;
		regs.pc = addr;
		cpu_time += 10;
		return true;
	}

	// call nn
	bool sim_cd (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t addr = decode_u16();
		// Let's not go too far with the simulation and do a simple write of the return address.
		memory->write_uint16 (regs.sp - 2, regs.pc);
		regs.sp -= 2;
		regs.pc = addr;
		cpu_time += 17;
		return true;
	}

	// out (n), a
	bool sim_d3 (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t val = decode_u8();
		if (!io->try_write_request(val | (regs.main.a << 8), regs.main.a, cpu_time))
			return false;
		cpu_time += 11;
		return true;
	}

	// exx
	bool sim_d9 (hl_ix_iy xy, uint8_t opcode)
	{
		std::swap<uint16_t>(regs.main.bc, regs.alt.bc);
		std::swap<uint16_t>(regs.main.de, regs.alt.de);
		std::swap<uint16_t>(regs.main.hl, regs.alt.hl);
		cpu_time += 4;
		return true;
	}

	// in a, (n)
	bool sim_db (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t val = decode_u8();
		if (!io->try_read_request (val | (regs.main.a << 8), regs.main.a, cpu_time))
			return false;
		cpu_time += 11;
		return true;
	}

	// ex (sp), hl/ix/iy
	bool sim_e3 (hl_ix_iy xy, uint8_t opcode)
	{
		uint16_t other;
		if (!memory->try_read_request (regs.sp, other, cpu_time))
			return false;
		uint16_t& reg = regs.hl(xy);
		memory->write_uint16(regs.sp, reg);
		reg = other;
		cpu_time += (xy == hl_ix_iy::hl) ? 19 : 23;
		return true;
	}

	// jp hl/ix/iy
	bool sim_e9 (hl_ix_iy xy, uint8_t opcode)
	{
		regs.pc = regs.hl(xy);
		cpu_time += ((xy == hl_ix_iy::hl) ? 4 : 8);
		return true;
	}

	// ex de, hl
	bool sim_eb (hl_ix_iy xy, uint8_t opcode)
	{
		// The real CPU ignores any IX/IY prefix on this one.
		uint16_t temp = regs.main.hl;
		regs.main.hl = regs.main.de;
		regs.main.de = temp;
		cpu_time += 4;
		return true;
	}

	// di
	bool sim_f3 (hl_ix_iy xy, uint8_t opcode)
	{
		regs.iff1 = 0;
		_ei_countdown = 0;
		cpu_time += 4;
		return true;
	}

	// ei
	bool sim_fb (hl_ix_iy xy, uint8_t opcode)
	{
		_ei_countdown = 2;
		cpu_time += 4;
		return true;
	}

	// ld sp, hl/ix/iy
	bool sim_f9 (hl_ix_iy xy, uint8_t opcode)
	{
		regs.sp = regs.hl(xy);
		cpu_time += 6;
		_start_of_stack = regs.sp;
		return true;
	}

	static constexpr handler_t dispatch[] = {
		&sim_00,  &sim_01,  &sim_02,  &sim_03,  &sim_04,  &sim_05,  &sim_06,  &sim_07,  // 00 - 07
		&sim_08,  &sim_09,  &sim_0a,  &sim_0b,  &sim_04,  &sim_05,  &sim_06,  &sim_0f,  // 08 - 0f
		&sim_10,  &sim_01,  &sim_02,  &sim_03,  &sim_04,  &sim_05,  &sim_06,  &sim_17,  // 10 - 17
		&sim_18,  &sim_09,  &sim_0a,  &sim_0b,  &sim_04,  &sim_05,  &sim_06,  &sim_1f,  // 18 - 1f
		&sim_20,  &sim_01,  &sim_22,  &sim_03,  &sim_04,  &sim_05,  &sim_06,  &sim_27,  // 20 - 27
		&sim_20,  &sim_09,  &sim_2a,  &sim_0b,  &sim_04,  &sim_05,  &sim_06,  &sim_2f,  // 28 - 2f
		&sim_20,  &sim_01,  &sim_32,  &sim_03,  &sim_04,  &sim_05,  &sim_06,  &sim_37,  // 30 - 37
		&sim_20,  &sim_09,  &sim_3a,  &sim_0b,  &sim_04,  &sim_05,  &sim_06,  &sim_3f,  // 38 - 3f

		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  // 40 - 47
		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  // 48 - 4f
		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  // 50 - 57
		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  // 58 - 5f
		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  // 60 - 67
		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  // 68 - 6f
		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_76,  &sim_40,  // 70 - 77
		&sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  &sim_40,  // 78 - 7f

		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // 80 - 87
		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // 88 - 8f
		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // 90 - 97
		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // 98 - 9f
		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // a0 - a7
		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // a8 - af
		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // b0 - b7
		&sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  &sim_80,  // b8 - bf

		&sim_c0,  &sim_c1,  &sim_c2,  &sim_c3,  &sim_c4,  &sim_c5,  &sim_c6,  &sim_c7,  // c0 - c7
		&sim_c0,  &sim_c9,  &sim_c2,  nullptr, &sim_c4,  &sim_cd,  &sim_c6,  &sim_c7,  // c8 - cf
		&sim_c0,  &sim_c1,  &sim_c2,  &sim_d3,  &sim_c4,  &sim_c5,  &sim_c6,  &sim_c7,  // d0 - d7
		&sim_c0,  &sim_d9,  &sim_c2,  &sim_db,  &sim_c4,  nullptr, &sim_c6,  &sim_c7,  // d8 - df
		&sim_c0,  &sim_c1,  &sim_c2,  &sim_e3,  &sim_c4,  &sim_c5,  &sim_c6,  &sim_c7,  // e0 - e7
		&sim_c0,  &sim_e9,  &sim_c2,  &sim_eb,  &sim_c4,  nullptr, &sim_c6,  &sim_c7,  // e8 - ef
		&sim_c0,  &sim_c1,  &sim_c2,  &sim_f3,  &sim_c4,  &sim_c5,  &sim_c6,  &sim_c7,  // f0 - f7
		&sim_c0,  &sim_f9,  &sim_c2,  &sim_fb,  &sim_c4,  nullptr, &sim_c6,  &sim_c7,  // f8 - ff
	};

	#pragma region ED group

	// in r, (c)
	bool sim_ed40 (uint8_t opcode)
	{
		uint8_t value;
		if (!io->try_read_request(regs.main.bc, value, cpu_time))
			return false;
		uint8_t i = (opcode >> 3) & 7;
		regs.r8(i, hl_ix_iy::hl) = value;
		regs.main.f.val = s_z_pv_flags.flags[value];
		cpu_time += 12;
		return true;
	}

	// out (c), r
	bool sim_ed41 (uint8_t opcode)
	{
		uint8_t i = (opcode >> 3) & 7;
		uint8_t value = regs.r8(i, hl_ix_iy::hl);
		if (!io->try_write_request(regs.main.bc, value, cpu_time))
			return false;
		cpu_time += 12;
		return true;
	}

	// neg
	bool sim_ed44 (uint8_t opcode)
	{
		uint8_t before = regs.main.a;
		regs.main.a = 0 - regs.main.a;
		regs.main.f.val = (regs.main.a & 0xA8) // S, R5, R3
			| (regs.main.a ? 0 : z80_flag::z) // Z
			| ((0 - (before & 0xF)) & 0x10) // H
			| ((before == 0x80) ? z80_flag::pv : 0) // P/V
			| z80_flag::n // N
			| (before ? z80_flag::c : 0); // C
		cpu_time += 8;
		return true;
	}

	// retn
	bool sim_ed45 (uint8_t opcode)
	{
		regs.iff1 = regs.iff2;
		cpu_time += 4;
		return sim_c9(hl_ix_iy::hl, 0xc9);
	}

	// im 0
	bool sim_ed46 (uint8_t opcode)
	{
		regs.im = 0;
		cpu_time += 8;
		return true;
	}

	// ld i, a
	bool sim_ed47 (uint8_t opcode)
	{
		regs.i = regs.main.a;
		cpu_time += 9;
		return true;
	}

	// ld r, a
	bool sim_ed4f (uint8_t opcode)
	{
		regs.r = regs.main.a;
		cpu_time += 9;
		return true;
	}

	// im 1
	bool sim_ed56 (uint8_t opcode)
	{
		regs.im = 1;
		cpu_time += 8;
		return true;
	}

	// ld a, i (ED 57)
	// ld a, r (ED 5F)
	bool sim_ed57 (uint8_t opcode)
	{
		regs.main.a = (opcode == 0x57) ? regs.i : regs.r;
		regs.main.f.val = (regs.main.a & 0xA8) // S, X5, X3
			| (regs.main.a ? 0 : 0x40) // Z
			| (regs.iff2 ? 4 : 0) // P/V
			| (regs.main.f.val & 1); // C
		// TODO: If an interrupt occurs during execution of this instruction, the parity flag contains a 0.
		cpu_time += 9;
		return true;
	}

	// im 2
	bool sim_ed5e (uint8_t opcode)
	{
		regs.im = 2;
		cpu_time += 8;
		return true;
	}

	// sbc hl, bc/de/hl/sp
	bool sim_ed42 (uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		uint16_t before = regs.main.hl;
		uint16_t other = regs.bc_de_hl_sp(hl_ix_iy::hl, i);
		regs.main.hl -= (other + regs.main.f.c);
		cpu_time += 15;
		regs.main.f.val = ((regs.main.hl >> 8) & 0x80) // S
			| (regs.main.hl ? 0 : z80_flag::z) // Z
			| (regs.main.hl & 0x28) // R3, R5
			| ((((before & 0x0FFF) - (other & 0x0FFF) - regs.main.f.c) & 0x1000) >> 8) // H
			| ((!same_sign(before, other) && !same_sign(before, regs.main.hl)) ? z80_flag::pv : 0) // P/V
			| z80_flag::n // N
			| ((((uint32_t)before - (uint32_t)other - (uint32_t)regs.main.f.c) >> 16) & 1); // C
		return true;
	}

	// adc hl, bc/de/hl/sp
	bool sim_ed4a (uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		uint16_t before = regs.main.hl;
		uint16_t other = regs.bc_de_hl_sp(hl_ix_iy::hl, i);
		regs.main.hl += (other + regs.main.f.c);
		cpu_time += 15;
		regs.main.f.val = ((regs.main.hl >> 8) & 0x80) // S
			| (regs.main.hl ? 0 : z80_flag::z) // Z
			| ((regs.main.hl >> 8) & (z80_flag::r5 | z80_flag::r3)) // R3, R5
			| ((((before & 0x0FFF) + (other & 0xFFF) + regs.main.f.c) & 0x1000) >> 8) // H
			| ((same_sign(before, other) && !same_sign(before, regs.main.hl)) ? z80_flag::pv : 0) // P/V
			| 0 // N
			| ((((uint32_t)before + (uint32_t)other + (uint32_t)regs.main.f.c) >> 16) & 1); // C
		return true;
	}

	// ld (nn), bc/de/hl/sp
	bool sim_ed43 (uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		uint16_t addr = decode_u16();
		uint16_t val = regs.bc_de_hl_sp(hl_ix_iy::hl, i);
		if (!memory->try_write_request(addr, val, cpu_time))
			return false;
		cpu_time += 20;
		return true;
	}

	// ld bc/de/hl/sp, (nn)
	bool sim_ed4B (uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		uint16_t addr = decode_u16();
		uint16_t value;
		if (!memory->try_read_request(addr, value, cpu_time))
			return false;
		regs.bc_de_hl_sp(hl_ix_iy::hl, i) = value;
		if (i == 3)
			_start_of_stack = value;
		cpu_time += 20;
		return true;
	}

	// reti
	bool sim_ed4d (uint8_t ocode)
	{
		// As far as I can tell, this instruction is different from the RET instruction only
		// in hardware signaling. Code below does a simple RET.
		cpu_time += 4;
		return sim_c9(hl_ix_iy::hl, 0xc9);
	}

	// rrd/rld
	bool sim_ed67 (uint8_t opcode)
	{
		bool is_rld = opcode & 8;
		uint8_t mem;
		if (!memory->try_read_request(regs.main.hl, mem, cpu_time))
			return false;
		uint8_t new_a;
		if (is_rld)
		{
			new_a = (regs.main.a & 0xF0) | (mem >> 4);
			mem = (mem << 4) | (regs.main.a & 0x0F);
		}
		else
		{
			new_a = (regs.main.a & 0xF0) | (mem & 0x0F);
			mem = (mem >> 4) | (regs.main.a << 4);
		}
		if (!memory->try_write_request(regs.main.hl, mem, cpu_time))
			return false;
		regs.main.a = new_a;
		regs.main.f.val = (regs.main.f.val & z80_flag::c) | s_z_pv_flags.flags[regs.main.a];
		cpu_time += 18;
		return true;
	}

	// LDI/LDD/LDIR/LDDR
	bool sim_eda0 (uint8_t opcode)
	{
		uint8_t data;
		if (!memory->try_read_request(regs.main.hl, data, cpu_time))
			return false;
		if (!memory->try_write_request(regs.main.de, data, cpu_time))
			return false;

		uint16_t increment = (opcode & 8) ? -1 : 1;
		regs.main.hl += increment;
		regs.main.de += increment;
		regs.main.bc--;
		cpu_time += 16;
		regs.main.f.val = (regs.main.f.val & (z80_flag::s | z80_flag::z | z80_flag::c))
			| (((data + regs.main.a) & 2) ? z80_flag::r5 : 0)
			| (((data + regs.main.a) & 8) ? z80_flag::r3 : 0)
			| (regs.main.bc ? z80_flag::pv : 0);
		bool repeat = opcode & 0x10;
		if (repeat && regs.main.bc)
		{
			regs.pc -= 2;
			cpu_time += 5;
		}

		return true;
	}

	// CPI/CPD/CPIR/CPDR
	bool sim_eda1 (uint8_t opcode)
	{
		uint8_t memhl;
		if (!memory->try_read_request(regs.main.hl, memhl, cpu_time))
			return false;

		uint16_t increment = (opcode & 8) ? -1 : 1;
		regs.main.hl += increment;
		regs.main.bc--;
		uint8_t diff = regs.main.a - memhl;
		uint8_t hf = ((regs.main.a & 0xF) - (memhl & 0xF)) & 0x10; // H
		uint8_t n = regs.main.a - memhl - (hf ? 1 : 0);
		regs.main.f.val = (diff & 0x80) // S
			| (diff ? 0 : z80_flag::z)  // Z
			| (n & z80_flag::r5)
			| hf // H
			| (n & z80_flag::r3)
			| (regs.main.bc ? z80_flag::pv : 0) // P/V
			| z80_flag::n
			| (regs.main.f.val & z80_flag::c);
		bool repeat = opcode & 0x10;
		if (repeat && regs.main.bc && diff)
		{
			cpu_time += 21;
			regs.pc -= 2;
		}
		else
		{
			cpu_time += 16;
		}

		return true;
	}

	// INI/IND/INIR/INDR
	bool sim_eda2 (uint8_t opcode)
	{
		uint8_t value;
		if (!io->try_read_request (regs.main.bc, value, cpu_time))
			return false;
		if (!memory->try_write_request (regs.main.hl, value, cpu_time))
			return false;

		uint16_t increment = (opcode & 8) ? -1 : 1;
		regs.main.hl += increment;
		regs.b()--;
		regs.main.f.val = regs.b() & (z80_flag::s | z80_flag::r5 | z80_flag::r3) // S, R5, R3
			| (regs.b() ? 0 : z80_flag::z) // Z
			| z80_flag::n // N
			| (regs.main.f.val & 1); // C
		cpu_time += 16;
		bool repeat = opcode & 0x10;
		if (repeat && regs.b())
		{
			regs.pc -= 2;
			cpu_time += 5;
		}
		return true;
	}

	// OUTI/OUTD/OTIR/OTDR
	bool sim_eda3 (uint8_t opcode)
	{
		uint8_t value;
		if (!memory->try_read_request (regs.main.hl, value, cpu_time))
			return false;
		if (!io->try_write_request (regs.main.bc - 256, value, cpu_time))
			return false;

		uint16_t increment = (opcode & 8) ? -1 : 1;
		regs.main.hl += increment;
		regs.b()--;
		regs.main.f.val = regs.b() & (z80_flag::s | z80_flag::r5 | z80_flag::r3) // S, R5, R3
			| (regs.b() ? 0 : z80_flag::z) // Z
			| z80_flag::n // N
			| (regs.main.f.val & 1); // C
		cpu_time += 16;
		bool repeat = opcode & 0x10;
		if (repeat && regs.b())
		{
			regs.pc -= 2;
			cpu_time += 5;
		}
		return true;

	}

	static constexpr ed_handler_t dispatch_ed[256] = {
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 00 - 07
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 08 - 0f
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 10 - 17
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 18 - 1f
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 20 - 27
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 28 - 2f
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 30 - 37
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 38 - 3f

		&sim_ed40, &sim_ed41,  &sim_ed42, &sim_ed43, &sim_ed44, &sim_ed45, &sim_ed46, &sim_ed47, // 40 - 47
		&sim_ed40, &sim_ed41,  &sim_ed4a, &sim_ed4B, nullptr,   &sim_ed4d, nullptr,   &sim_ed4f, // 48 - 4f
		&sim_ed40, &sim_ed41,  &sim_ed42, &sim_ed43, nullptr,   nullptr,   &sim_ed56, &sim_ed57, // 50 - 57
		&sim_ed40, &sim_ed41,  &sim_ed4a, &sim_ed4B, nullptr,   nullptr,   &sim_ed5e, &sim_ed57, // 58 - 5f
		&sim_ed40, &sim_ed41,  &sim_ed42, &sim_ed43, nullptr,   nullptr,   nullptr,   &sim_ed67, // 60 - 67
		&sim_ed40, &sim_ed41,  &sim_ed4a, &sim_ed4B, nullptr,   nullptr,   nullptr,   &sim_ed67, // 68 - 6f
		nullptr,   &sim_ed41,  &sim_ed42, &sim_ed43, nullptr,   nullptr,   nullptr,   nullptr,   // 70 - 77
		&sim_ed40, &sim_ed41,  &sim_ed4a, &sim_ed4B, nullptr,   nullptr,   nullptr,   nullptr,   // 78 - 7f

		nullptr,     nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 80 - 87
		nullptr,     nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 88 - 8f
		nullptr,     nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 90 - 97
		nullptr,     nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 98 - 9f
		&sim_eda0, &sim_eda1, &sim_eda2, &sim_eda3,  nullptr,  nullptr,  nullptr,  nullptr,  // a0 - a7
		&sim_eda0, &sim_eda1, &sim_eda2, &sim_eda3,  nullptr,  nullptr,  nullptr,  nullptr,  // a8 - af
		&sim_eda0, &sim_eda1, &sim_eda2, &sim_eda3,  nullptr,  nullptr,  nullptr,  nullptr,  // b0 - b7
		&sim_eda0, &sim_eda1, &sim_eda2, &sim_eda3,  nullptr,  nullptr,  nullptr,  nullptr,  // b8 - bf

		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // c0 - c7
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // c8 - cf
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // d0 - d7
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // d8 - df
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // e0 - e7
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // e8 - ef
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // f0 - f7
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // f8 - ff
	};
	#pragma endregion

	#pragma region CB group
	static void do_cb_rotate_shift_operation (uint8_t& before, uint8_t& after, uint8_t& c, uint8_t opcode)
	{
		switch(opcode & 0x38)
		{
			case 0: // RLC
				after = (before << 1) | (before >> 7);
				c = before >> 7;
				break;
			case 8: // RRC
				after = (before >> 1) | (before << 7);
				c = before & 1;
				break;
			case 0x10: // RL
				after = (before << 1) | c;
				c = before >> 7;
				break;
			case 0x18: // RR
				after = (before >> 1) | (c << 7);
				c = before & 1;
				break;
			case 0x20: // SLA
				after = before << 1;
				c = before >> 7;
				break;
			case 0x28: // SRA
				after = (uint8_t)((int8_t)before >> 1);
				c = before & 1;
				break;
			case 0x30: // SLL (undocumented)
				after = (before << 1) | 1;
				c = before >> 7;
				break;
			case 0x38: // SRL
				after = before >> 1;
				c = before & 1;
				break;
			default:
				FAIL_FAST();
		}
	}

	// rlc/rrc/rl/rr/sla/sra/sll/srl r8
	bool sim_cb00 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t before, after;
		uint8_t c = regs.main.f.c;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			before = regs.r8(i, hl_ix_iy::hl);
			do_cb_rotate_shift_operation(before, after, c, opcode);
			regs.r8(i, hl_ix_iy::hl) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, before, cpu_time))
				return false;
			do_cb_rotate_shift_operation(before, after, c, opcode);
			if (!memory->try_write_request(memhlxy_addr, after, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.val = s_z_pv_flags.flags[after] | c;
		return true;
	}

	// bit b, r8
	bool sim_cb40 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t reg = opcode & 7;
		uint8_t mask = 1 << ((opcode >> 3) & 7);
		uint8_t value;

		if (xy == hl_ix_iy::hl)
		{
			if (reg != 6)
			{
				value = regs.r8(reg, hl_ix_iy::hl);
				cpu_time += 8;
			}
			else
			{
				if (!memory->try_read_request(memhlxy_addr, value, cpu_time))
					return false;
				cpu_time += 12;
			}
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, value, cpu_time))
				return false;
			cpu_time += 20;
		}

		regs.main.f.val = s_z_pv_flags.flags[value & mask] | z80_flag::h | (regs.main.f.val & 1);
		return true;
	}

	// res/set b, r8
	bool sim_cb80 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		// Let's not bother with the undocumented instructions from DDCB and FDCB.

		uint8_t reg = opcode & 7;
		uint8_t bit = (opcode >> 3) & 7;
		if (reg != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			if (opcode & 0x40)
				regs.r8(reg, hl_ix_iy::hl) |= (1 << bit);
			else
				regs.r8(reg, hl_ix_iy::hl) &= ~(1 << bit);
			cpu_time += 8;
		}
		else
		{
			uint8_t value;
			if (!memory->try_read_request(memhlxy_addr, value, cpu_time))
				return false;
			if (opcode & 0x40)
				value |= (1 << bit);
			else
				value &= ~(1 << bit);
			if (!memory->try_write_request(memhlxy_addr, value, cpu_time))
				return false;
		
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		return true;
	}
	#pragma endregion

	virtual bool SimulateOne (BreakpointsHit* bps) override
	{
		if (bps)
			bps->size = 0;

		if (regs.iff1)
		{
			bool interrupted;
			uint8_t irq_address;
			if (!irq->try_poll_irq_at_time_point(interrupted, irq_address, cpu_time))
				return false;

			if (interrupted)
			{
				regs.halted = false;

				if (regs.im == 0)
				{
					LOG_HR_MSG(E_NOTIMPL, "IM0 is not yet implemented");
				}
				else if (regs.im == 1)
				{
					if (!memory->try_write_request (regs.sp - 2, regs.pc, cpu_time))
						return false;
					regs.sp -= 2;
					regs.pc = 0x38;
					regs.iff1 = false;
					cpu_time += 13;
					return true;
				}
				else
				{
					uint16_t addr = (regs.i << 8) | (irq_address & 0xFE);
					if (!memory->try_read_request (addr, addr, cpu_time))
						return false;
					if (!memory->try_write_request (regs.sp - 2, regs.pc, cpu_time))
						return false;
					regs.sp -= 2;
					regs.pc = addr;
					regs.iff1 = false;
					cpu_time += 19;
					return true;
				}
			}
		}

		if (regs.halted)
		{
			cpu_time += 4;
			regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);
			return true;
		}

		if (bps)
		{
			auto it = code_bps.find(regs.pc);
			if (it != code_bps.end())
			{
				bps->address = regs.pc;
				bps->size = std::min((uint32_t)_countof(BreakpointsHit::bps), it->second.size());
				memcpy(bps->bps, it->second.data(), bps->size * sizeof(SIM_BP_COOKIE));
				return false;
			}
		}

		uint8_t opcode;
		bool b = memory->try_read_request (regs.pc, opcode, cpu_time);
		if (!b)
			return false;

		uint16_t oldpc = regs.pc;
		uint8_t oldr = regs.r;
		regs.pc++;
		regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);

		hl_ix_iy xy = hl_ix_iy::hl;
		if (opcode == 0xDD)
		{
			xy = hl_ix_iy::ix;
			opcode = decode_u8();
			regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);
		}
		else if (opcode == 0xFD)
		{
			xy = hl_ix_iy::iy;
			opcode = decode_u8();
			regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);
		}

		bool executed;
		if (opcode == 0xed)
		{
			opcode = decode_u8();
			if (xy == hl_ix_iy::hl)
				regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);
			auto handler = dispatch_ed[opcode];
			if (!handler)
			{
				cpu_time += 8;
				return true;
			}

			executed = (this->*handler)(opcode);
		}
		else if (opcode == 0xcb)
		{
			uint16_t memhlxy_addr = decode_mem_hl(xy);
			opcode = decode_u8();
			if (xy == hl_ix_iy::hl)
				regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);
			cb_handler_t handler;
			if ((opcode & 0xC0) == 0x00)
				handler = &cpu::sim_cb00;
			else if ((opcode & 0xC0) == 0x40)
				handler = &cpu::sim_cb40;
			else
				handler = &cpu::sim_cb80;
			executed = (this->*handler) (xy, memhlxy_addr, opcode);
		}
		else
		{
			auto handler = dispatch[opcode];
			WI_ASSERT(handler);
			executed = (this->*handler)(xy, opcode);
		}

		if (!executed)
		{
			regs.pc = oldpc;
			regs.r = oldr;
			return false;
		}

		if (_ei_countdown)
		{
			_ei_countdown--;
			if (_ei_countdown == 0)
				regs.iff1 = 1;
		}

		return true;
	}

	// ========================================================================

	virtual UINT64 Time() override
	{
		return cpu_time;
	}

	virtual BOOL NeedSyncWithRealTime (UINT64* sync_time) override { return false; }

	virtual bool SimulateTo (UINT64 requested_time) override
	{
		WI_ASSERT(false); return { };
	}

	virtual void Reset() override
	{
		memset (&regs, 0, sizeof(regs));
		_start_of_stack = 0;
		cpu_time = 0;
	}

	virtual void GetZ80Registers (z80_register_set* pRegs) override
	{
		*pRegs = regs;
	}

	virtual void SetZ80Registers (const z80_register_set* pRegs) override
	{
		regs = *pRegs;
	}
	
	#ifdef SIM_TESTS
	virtual z80_register_set* GetRegsPtr() override { return &regs; }
	#endif

	virtual UINT16 GetStackStartAddress() const override { return _start_of_stack; }
	
	virtual BOOL STDMETHODCALLTYPE Halted() override { return regs.halted; }
	
	virtual UINT16 GetPC() const override { return regs.pc; }

	virtual HRESULT SetPC (UINT16 pc) override
	{
		regs.pc = pc;
		return S_OK;
	}

	virtual HRESULT AddBreakpoint (BreakpointType type, uint16_t address, SIM_BP_COOKIE* pCookie) override
	{
		if (type == BreakpointType::Code)
		{
			auto it = code_bps.find(address);
			if (it == code_bps.end())
			{
				bool added = code_bps.try_insert({ address, vector_nothrow<SIM_BP_COOKIE>{ } });
				if (!added)
					return E_OUTOFMEMORY;

				auto it = code_bps.find(address);
				added = it->second.try_push_back(_nextBpCookie);
				if (!added)
				{
					code_bps.remove(it);
					return E_OUTOFMEMORY;
				}
				*pCookie = _nextBpCookie;
				_nextBpCookie++;
			}
			else
			{
				bool added = it->second.try_push_back(_nextBpCookie);
				if (!added)
					return E_OUTOFMEMORY;
				*pCookie = _nextBpCookie;
				_nextBpCookie++;
			}

			return S_OK;
		}
		else
			return E_NOTIMPL;
	}

	virtual HRESULT RemoveBreakpoint (SIM_BP_COOKIE cookie) override
	{
		for (auto it = code_bps.begin(); it != code_bps.end(); it++)
		{
			if (auto it1 = it->second.find(cookie); it1 != it->second.end())
			{
				it->second.erase(it1);
				if (it->second.empty())
					code_bps.erase(it);
				return S_OK;
			}
		}

		// TODO: search data breakpoints

		return E_INVALIDARG;
	}

	virtual BOOL HasBreakpoints() override
	{
		return code_bps.size();
	}
};

HRESULT STDMETHODCALLTYPE MakeZ80CPU (Bus* memory, Bus* io, irq_line_i* irq, wistd::unique_ptr<IZ80CPU>* ppCPU)
{
	auto d = wil::make_unique_nothrow<cpu>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(memory, io, irq); RETURN_IF_FAILED(hr);
	*ppCPU = std::move(d);
	return S_OK;
}
