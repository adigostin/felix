
#pragma once

enum class z80_reg16 { bc, de, hl, alt_bc, alt_de, alt_hl, sp, pc, ix, iy, count };
static inline const char* const z80_reg16_names[] = { "BC", "DE", "HL", "BC'", "DE'", "HL'", "SP", "PC", "IX", "IY" };

enum class z80_reg8 { b, c, d, e, h, l, a, f, alt_b, alt_c, alt_d, alt_e, alt_h, alt_l, alt_a, alt_f, i, r, count };
static inline const char* const z80_reg8_names[] = { "B", "C", "D", "E", "H", "L", "A", "F", "B'", "C'", "D'", "E'", "H'", "L'", "A'", "F'", "I", "R" };

enum class hl_ix_iy { hl, ix, iy };

struct z80_flag
{ 
	static constexpr byte c  = 0x01;
	static constexpr byte n  = 0x02;
	static constexpr byte pv = 0x04;
	static constexpr byte r3 = 0x08; // copy of bit 3 of the result
	static constexpr byte h  = 0x10;
	static constexpr byte r5 = 0x20; // copy of bit 5 of the result
	static constexpr byte z  = 0x40;
	static constexpr byte s  = 0x80;
};

struct z80_register_set
{
	struct bc_de_hl_af_set
	{
		uint16_t bc;
		uint16_t de;
		uint16_t hl;

		union
		{
			uint16_t af;

			struct
			{
				union
				{
					struct
					{
						uint8_t c  : 1;
						uint8_t n  : 1;
						uint8_t pv : 1;
						uint8_t x3 : 1; // copy of bit 3 of the result
						uint8_t h  : 1;
						uint8_t x5 : 1; // copy of bit 5 of the result
						uint8_t z  : 1;
						uint8_t s  : 1;
					};
					uint8_t val;
				} f;

				uint8_t a;
			};
		};
	};

	bc_de_hl_af_set main;
	bc_de_hl_af_set alt;

	uint16_t pc;
	uint16_t sp;
	uint16_t ix;
	uint16_t iy;
	bool iff1;
	bool halted;
	uint8_t im;

	uint8_t i;
	uint8_t r;

	uint8_t& b() { return *((uint8_t*)&main.bc + 1); }
	uint8_t& c() { return *(uint8_t*)&main.bc; }
	uint8_t& d() { return *((uint8_t*)&main.de + 1); }
	uint8_t& e() { return *(uint8_t*)&main.de; }
	uint8_t& h() { return *((uint8_t*)&main.hl + 1); }
	uint8_t& l() { return *(uint8_t*)&main.hl; }
	uint8_t& ixh() { return *((uint8_t*)&ix + 1); }
	uint8_t& ixl() { return *(uint8_t*)&ix; }
	uint8_t& iyh() { return *((uint8_t*)&iy + 1); }
	uint8_t& iyl() { return *(uint8_t*)&iy; }

	uint16_t& hl (hl_ix_iy xy)
	{
		if (xy == hl_ix_iy::hl)
			return main.hl;
		else if (xy == hl_ix_iy::ix)
			return ix;
		else if (xy == hl_ix_iy::iy)
			return iy;
		else
			FAIL_FAST();
	}

	uint8_t& r8 (uint8_t i, hl_ix_iy xy)
	{
		switch(i)
		{
			case 0: return b();
			case 1: return c();
			case 2: return d();
			case 3: return e();
			case 4:
				if (xy == hl_ix_iy::hl)
					return h();
				else if (xy == hl_ix_iy::ix)
					return ixh();
				else
					return iyh();
			case 5:
				if (xy == hl_ix_iy::hl)
					return l();
				else if (xy == hl_ix_iy::ix)
					return ixl();
				else
					return iyl();

			case 7: return main.a;
			default: FAIL_FAST();
		}
	}

	uint16_t& bc_de_hl_sp (hl_ix_iy xy, uint8_t i)
	{
		switch(i)
		{
			case 0: return main.bc;
			case 1: return main.de;
			case 2:
				if (xy == hl_ix_iy::hl)
					return main.hl;
				else if (xy == hl_ix_iy::ix)
					return ix;
				else if (xy == hl_ix_iy::iy)
					return iy;
				else
					FAIL_FAST();
			case 3: return sp;
			default: FAIL_FAST();
		}
	}

	uint16_t& bc_de_hl_af (hl_ix_iy xy, uint8_t i)
	{
		switch(i)
		{
			case 0: return main.bc;
			case 1: return main.de;
			case 2:
				if (xy == hl_ix_iy::hl)
					return main.hl;
				else if (xy == hl_ix_iy::ix)
					return ix;
				else if (xy == hl_ix_iy::iy)
					return iy;
				else
					FAIL_FAST();
			case 3: return main.af;
			default: FAIL_FAST();
		}
	}

	uint16_t& reg (z80_reg16 r)
	{
		switch (r)
		{
			case z80_reg16::bc: return main.bc;
			case z80_reg16::de: return main.de;
			case z80_reg16::hl: return main.hl;
			case z80_reg16::alt_bc: return alt.bc;
			case z80_reg16::alt_de: return alt.de;
			case z80_reg16::alt_hl: return alt.hl;
			case z80_reg16::sp: return sp;
			case z80_reg16::pc: return pc;
			case z80_reg16::ix: return ix;
			case z80_reg16::iy: return iy;
			default:
				FAIL_FAST();
		}
	}
};
