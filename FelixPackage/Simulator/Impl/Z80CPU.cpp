
#include "pch.h"
#include "Z80CPU.h"
#include "shared/unordered_map_nothrow.h"
#include "shared/TryQI.h"
#include <optional>

// For the P/V flag calculation, the "Parity/Overflow Flag" paragraph in the Z80 pdf has a good explanation.

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
/*
class Z80RegisterGroup : public IZ80RegisterGroup
{
	ULONG _refCount = 0;
	z80_register_set _regs;

public:
	HRESULT InitInstance (const z80_register_set& regs)
	{
		_regs = regs;
		return S_OK;
	}

	#pragma region IUnknown
	virtual ULONG STDMETHODCALLTYPE AddRef() override
	{
		return InterlockedIncrement(&_refCount);
	}

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		WI_ASSERT(_refCount);
		if (_refCount > 1)
			return InterlockedDecrement(&_refCount);
		delete this;
		return 0;
	}

	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (riid == __uuidof(IZ80RegisterGroup))
		{
			*ppvObject = static_cast<IZ80RegisterGroup*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}
	#pragma endregion

	#pragma region IRegisterGroup
	virtual HRESULT STDMETHODCALLTYPE GetGroupName (BSTR* pName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual UINT32 STDMETHODCALLTYPE GetRegisterCount() override
	{
		return 20;
	}

	virtual HRESULT STDMETHODCALLTYPE GetRegisterAt (UINT32 i, UINT8* pBitSize, BSTR* pName, UINT32* pValue, UINT8* pBeginSubgroup) override
	{
		switch(i)
		{
		case 0:
			*pBitSize = 8;
			*pName = SysAllocString(L"A"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.a;
			*pBeginSubgroup = TRUE;
			return S_OK;		
		case 1:
			*pBitSize = 16;
			*pName = SysAllocString(L"BC"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.bc;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 2:
			*pBitSize = 16;
			*pName = SysAllocString(L"DE"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.de;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 3:
			*pBitSize = 16;
			*pName = SysAllocString(L"HL"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.hl;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 4:
			*pBitSize = 1;
			*pName = SysAllocString(L"S"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.f.s;
			*pBeginSubgroup = TRUE;
			return S_OK;
		case 5:
			*pBitSize = 1;
			*pName = SysAllocString(L"Z"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.f.z;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 6:
			*pBitSize = 1;
			*pName = SysAllocString(L"H"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.f.h;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 7:
			*pBitSize = 1;
			*pName = SysAllocString(L"P/V"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.f.pv;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 8:
			*pBitSize = 1;
			*pName = SysAllocString(L"N"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.f.n;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 9:
			*pBitSize = 1;
			*pName = SysAllocString(L"C"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.main.f.c;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 10:
			*pBitSize = 16;
			*pName = SysAllocString(L"AF'"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.alt.af;
			*pBeginSubgroup = TRUE;
			return S_OK;
		case 11:
			*pBitSize = 16;
			*pName = SysAllocString(L"BC'"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.alt.bc;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 12:
			*pBitSize = 16;
			*pName = SysAllocString(L"DE'"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.alt.de;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 13:
			*pBitSize = 16;
			*pName = SysAllocString(L"HL'"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.alt.hl;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 14:
			*pBitSize = 16;
			*pName = SysAllocString(L"IX"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.ix;
			*pBeginSubgroup = TRUE;
			return S_OK;
		case 15:
			*pBitSize = 16;
			*pName = SysAllocString(L"IY"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.iy;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 16:
			*pBitSize = 16;
			*pName = SysAllocString(L"PC"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.pc;
			*pBeginSubgroup = TRUE;
			return S_OK;
		case 17:
			*pBitSize = 16;
			*pName = SysAllocString(L"SP"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.sp;
			*pBeginSubgroup = FALSE;
			return S_OK;
		case 18:
			*pBitSize = 8;
			*pName = SysAllocString(L"I"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.i;
			*pBeginSubgroup = TRUE;
			return S_OK;
		case 19:
			*pBitSize = 8;
			*pName = SysAllocString(L"R"); RETURN_IF_NULL_ALLOC(*pName);
			*pValue = _regs.r;
			*pBeginSubgroup = FALSE;
			return S_OK;

		default:
			RETURN_HR(E_BOUNDS);
		}

	}
	#pragma endregion

	#pragma region IZ80RegisterGroup
	virtual UINT16 GetPC() override { return _regs.pc; }
	virtual UINT16 GetSP() override { return _regs.sp; }
	#pragma endregion
};
*/
class BreakpointCollection : public ISimulatorBreakpointEvent
{
	ULONG _refCount = 0;
	BreakpointType _type;
	uint16_t _address;
	vector_nothrow<SIM_BP_COOKIE> _bps;

public:
	HRESULT InitInstance (BreakpointType type, uint16_t address, SIM_BP_COOKIE const* bps, uint32_t size)
	{
		_type = type;
		_address = address;
		bool r = _bps.try_reserve(size); RETURN_HR_IF(E_OUTOFMEMORY, !r);
		for (uint32_t i = 0; i < size; i++)
			_bps.try_push_back(bps[i]);
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);

		if (TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<ISimulatorEvent>(this, riid, ppvObject)
			|| TryQI<ISimulatorBreakpointEvent>(this, riid, ppvObject))
			return S_OK;

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override
	{
		return InterlockedIncrement(&_refCount);
	}

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		WI_ASSERT(_refCount);
		if (_refCount > 1)
			return InterlockedDecrement(&_refCount);
		delete this;
		return 0;
	}
	#pragma endregion

	#pragma region IEnumBreakpoints
	virtual BreakpointType GetType() override { return _type; }

	virtual UINT16 GetAddress() override { return _address; }

	virtual ULONG GetBreakpointCount() override { return _bps.size(); }

	virtual HRESULT GetBreakpointAt(ULONG i, SIM_BP_COOKIE* ppKey) override { *ppKey = _bps[i]; return S_OK; }
	#pragma endregion
};

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

	std::optional<uint16_t> _addr_of_irq_ret_addr;

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
			regs.main.f.val = (after & 0xA8) // S, R5, R3
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
			case 0:  return !regs.main.f.z; // nz
			case 1:  return  regs.main.f.z; // z
			case 2:  return !regs.main.f.c; // nc
			case 3:  return  regs.main.f.c; // c
			case 4:  WI_ASSERT(false); // return !regs.main.f.pv; // po not yet generated correctly
			case 5:  WI_ASSERT(false); // return  regs.main.f.pv; // pe not yet generated correctly
			case 6:  return !regs.main.f.s; // p
			default: return  regs.main.f.s; // m
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

	// inc/dec r
	bool sim_04 (hl_ix_iy xy, uint8_t opcode)
	{
		bool dec = opcode & 1;
		uint8_t i = (opcode >> 3) & 7;
		
		uint8_t before;
		uint8_t after;
		if (i != 6)
		{
			before = regs.r8(i, xy);
			after = before + (dec ? (uint8_t)0xff : (uint8_t)1);
			regs.r8(i, xy) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 4 : 8);
		}
		else
		{
			uint16_t addr = decode_mem_hl(xy);
			if (!memory->try_read_request(addr, before, cpu_time))
				return false;
			after = before + (dec ? (uint8_t)0xff : (uint8_t)1);
			memory->write (addr, after);
			cpu_time += ((xy == hl_ix_iy::hl) ? 11 : 23);
		}

		regs.main.f.s = ((after & 0x80) ? 1 : 0);
		regs.main.f.z = (after ? 0 : 1);
		regs.main.f.h = (before & 0x10) != (after & 0x10);
		regs.main.f.pv = dec ? (before == 0x80) : (before == 0x7f);
		regs.main.f.n = dec ? 1 : 0;
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
		cpu_time += 11;
		regs.main.f.h = ((before & 0x0FFF) + (other & 0x0FFF)) >> 12;
		regs.main.f.n = 0;
		regs.main.f.c = after < before;
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
		cpu_time += 6;
		return true;
	}

	// dec bc/de/hl/sp
	bool sim_0b (hl_ix_iy xy, uint8_t opcode)
	{
		uint8_t i = (opcode >> 4) & 3;
		regs.bc_de_hl_sp(xy, i)--;
		cpu_time += 6;
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
			if (_addr_of_irq_ret_addr && regs.sp == *_addr_of_irq_ret_addr)
				_addr_of_irq_ret_addr.reset();

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

		if (_addr_of_irq_ret_addr && regs.sp == *_addr_of_irq_ret_addr)
			_addr_of_irq_ret_addr.reset();

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
		cpu_time == (xy == hl_ix_iy::hl) ? 4 : 8;
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
		&sim_00,  &sim_01,  &sim_02,  &sim_03,  &sim_04,  &sim_04,  &sim_06,  &sim_07,  // 00 - 07
		&sim_08,  &sim_09,  &sim_0a,  &sim_0b,  &sim_04,  &sim_04,  &sim_06,  &sim_0f,  // 08 - 0f
		&sim_10,  &sim_01,  &sim_02,  &sim_03,  &sim_04,  &sim_04,  &sim_06,  &sim_17,  // 10 - 17
		&sim_18,  &sim_09,  &sim_0a,  &sim_0b,  &sim_04,  &sim_04,  &sim_06,  &sim_1f,  // 18 - 1f
		&sim_20,  &sim_01,  &sim_22,  &sim_03,  &sim_04,  &sim_04,  &sim_06,  &sim_27,  // 20 - 27
		&sim_20,  &sim_09,  &sim_2a,  &sim_0b,  &sim_04,  &sim_04,  &sim_06,  &sim_2f,  // 28 - 2f
		&sim_20,  &sim_01,  &sim_32,  &sim_03,  &sim_04,  &sim_04,  &sim_06,  &sim_37,  // 30 - 37
		&sim_20,  &sim_09,  &sim_3a,  &sim_0b,  &sim_04,  &sim_04,  &sim_06,  &sim_3f,  // 38 - 3f

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
		regs.main.f.s = (regs.main.a & 0x80) ? 1 : 0;
		regs.main.f.z = regs.main.a ? 0 : 1;
		regs.main.f.h = 0;
		// TODO: P/V
		regs.main.f.n = 0;
		cpu_time += 12;
		return true;
	}

	// neg
	bool sim_ed44 (uint8_t opcode)
	{
		uint8_t before = regs.main.a;
		regs.main.a = 0 - regs.main.a;
		regs.main.f.s = (regs.main.a & 0x80) ? 1 : 0;
		regs.main.f.z = regs.main.a ? 0 : 1;
		// TODO: H
		regs.main.f.pv = (before == 0x80);
		regs.main.f.n = 1;
		regs.main.f.c = !!before;
		cpu_time += 8;
		return true;
	}

	// retn
	bool sim_ed45 (uint8_t opcode)
	{
		// TODO: implement this
		//regs.iff1 = regs.iff2;
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

	// ld a, i
	bool sim_ed57 (uint8_t opcode)
	{
		regs.main.a = regs.i;
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

	// ld a, r
	bool sim_ed5f (uint8_t opcode)
	{
		regs.main.a = regs.r;
		cpu_time += 9;
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
			| (regs.main.hl & 0x28) // R3, R5
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

	// rrd 
	bool sim_ed67 (uint8_t opcode)
	{
		WI_ASSERT(false);
		return false;
	}

	// rld
	bool sim_ed6f (uint8_t opcode)
	{
		uint8_t mem;
		if (!memory->try_read_request(regs.main.hl, mem, cpu_time))
			return false;
		uint8_t new_a = (regs.main.a & 0xF0) | (mem >> 4);
		mem = (mem << 4) | (regs.main.a & 0x0F);
		if (!memory->try_write_request(regs.main.hl, mem, cpu_time))
			return false;
		regs.main.a = new_a;
		regs.main.f.val = (regs.main.f.val & z80_flag::c) | s_z_pv_flags.flags[regs.main.a];
		cpu_time += 18;
		return true;
	}

	// ldi/ldd/ldir/lddr
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
		bool repeat = opcode & 0x10;
		if (repeat && regs.main.bc)
		{
			cpu_time += 21;
			regs.pc -= 2;
		}
		else
		{
			cpu_time += 16;
			regs.main.f.h = 0;
			regs.main.f.pv = 0;
			regs.main.f.n = 0;
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

	static constexpr ed_handler_t dispatch_ed[256] = {
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 00 - 07
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 08 - 0f
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 10 - 17
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 18 - 1f
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 20 - 27
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 28 - 2f
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 30 - 37
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 38 - 3f

		&sim_ed40, nullptr,  &sim_ed42, &sim_ed43, &sim_ed44, &sim_ed45, &sim_ed46, &sim_ed47, // 40 - 47
		&sim_ed40, nullptr,  &sim_ed4a, &sim_ed4B, nullptr,   &sim_ed4d, nullptr,   &sim_ed4f, // 48 - 4f
		&sim_ed40, nullptr,  &sim_ed42, &sim_ed43, nullptr,   nullptr,   &sim_ed56, &sim_ed57, // 50 - 57
		&sim_ed40, nullptr,  &sim_ed4a, &sim_ed4B, nullptr,   nullptr,   &sim_ed5e, &sim_ed5f, // 58 - 5f
		&sim_ed40, nullptr,  &sim_ed42, &sim_ed43, nullptr,   nullptr,   nullptr,   &sim_ed67, // 60 - 67
		&sim_ed40, nullptr,  &sim_ed4a, &sim_ed4B, nullptr,   nullptr,   nullptr,   &sim_ed6f, // 68 - 6f
		nullptr,   nullptr,  &sim_ed42, &sim_ed43, nullptr,   nullptr,   nullptr,   nullptr,   // 70 - 77
		&sim_ed40, nullptr,  &sim_ed4a, &sim_ed4B, nullptr,   nullptr,   nullptr,   nullptr,   // 78 - 7f

		nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 80 - 87
		nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 88 - 8f
		nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 90 - 97
		nullptr,   nullptr,   nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 98 - 9f
		&sim_eda0, &sim_eda1, nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // a0 - a7
		&sim_eda0, &sim_eda1, nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // a8 - af
		&sim_eda0, &sim_eda1, nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // b0 - b7
		&sim_eda0, &sim_eda1, nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // b8 - bf

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
	// rlc r8
	bool sim_cb00 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t value;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			value = regs.r8(i, hl_ix_iy::hl);
			value = (value << 1) | (value >> 7);
			regs.r8(i, hl_ix_iy::hl) = value;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, value, cpu_time))
				return false;
			value = (value << 1) | (value >> 7);
			if (!memory->try_write_request(memhlxy_addr, value, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.s = (value & 0x80) ? 1 : 0;
		regs.main.f.z = value ? 0 : 1;
		regs.main.f.h = 0;
		// TODO: P/V
		regs.main.f.n = 0;
		regs.main.f.c = value & 1;
		return true;
	}

	// rrc r8
	bool sim_cb08 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t value;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			value = regs.r8(i, hl_ix_iy::hl);
			value = (value >> 1) | (value << 7);
			regs.r8(i, hl_ix_iy::hl) = value;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, value, cpu_time))
				return false;
			value = (value >> 1) | (value << 7);
			if (!memory->try_write_request(memhlxy_addr, value, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.s = (value & 0x80) ? 1 : 0;
		regs.main.f.z = value ? 0 : 1;
		regs.main.f.h = 0;
		// TODO: P/V
		regs.main.f.n = 0;
		regs.main.f.c = value >> 7;
		return true;
	}

	// rl r8
	bool sim_cb10 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t before, after;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			before = regs.r8(i, hl_ix_iy::hl);
			after = (before << 1) | regs.main.f.c;
			regs.r8(i, hl_ix_iy::hl) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, before, cpu_time))
				return false;
			after= (before << 1) | regs.main.f.c;
			if (!memory->try_write_request(memhlxy_addr, after, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.val = s_z_pv_flags.flags[after] | (before >> 7);
		return true;
	}

	// rr r8
	bool sim_cb18 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t before, after;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			before = regs.r8(i, hl_ix_iy::hl);
			after = (before >> 1) | (regs.main.f.c << 7);
			regs.r8(i, hl_ix_iy::hl) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, before, cpu_time))
				return false;
			after = (before >> 1) | (regs.main.f.c << 7);
			if (!memory->try_write_request(memhlxy_addr, after, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.val = s_z_pv_flags.flags[after] | (before & 1);
		return true;
	}

	// sla r8
	bool sim_cb20 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t before, after;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			before = regs.r8(i, hl_ix_iy::hl);
			after = before << 1;
			regs.r8(i, hl_ix_iy::hl) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, before, cpu_time))
				return false;
			after = before << 1;
			if (!memory->try_write_request(memhlxy_addr, after, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.val = s_z_pv_flags.flags[after] | (before >> 7);
		return true;
	}

	// sra r8
	bool sim_cb28 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t before, after;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			before = regs.r8(i, hl_ix_iy::hl);
			after = (uint8_t)((int8_t)before >> 1);
			regs.r8(i, hl_ix_iy::hl) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, before, cpu_time))
				return false;
			after = (uint8_t)((int8_t)before >> 1);
			if (!memory->try_write_request(memhlxy_addr, after, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.s = 0;
		regs.main.f.z = after ? 0 : 1;
		regs.main.f.h = 0;
		// TODO: P/V
		regs.main.f.n = 0;
		regs.main.f.c = before & 1;
		return true;
	}

	// srl r8
	bool sim_cb38 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		uint8_t i = opcode & 7;
		uint8_t before, after;
		if (i != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			before = regs.r8(i, hl_ix_iy::hl);
			after = before >> 1;
			regs.r8(i, hl_ix_iy::hl) = after;
			cpu_time += ((xy == hl_ix_iy::hl) ? 8 : 23);
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, before, cpu_time))
				return false;
			after = before >> 1;
			if (!memory->try_write_request(memhlxy_addr, after, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 15 : 23);
		}

		regs.main.f.s = 0;
		regs.main.f.z = after ? 0 : 1;
		regs.main.f.h = 0;
		// TODO: P/V
		regs.main.f.n = 0;
		regs.main.f.c = before & 1;
		return true;
	}

	// bit b, r8
	bool sim_cb40 (hl_ix_iy xy, uint16_t memhlxy_addr, uint8_t opcode)
	{
		// Let's not bother with the undocumented instructions from DDCB and FDCB.

		uint8_t reg = opcode & 7;
		uint8_t mask = 1 << ((opcode >> 3) & 7);
		uint8_t value;
		if (reg != 6)
		{
			// Let's not bother with the undocumented instructions from DDCB and FDCB.
			value = regs.r8(reg, hl_ix_iy::hl);
			cpu_time += 8;
		}
		else
		{
			if (!memory->try_read_request(memhlxy_addr, value, cpu_time))
				return false;
			cpu_time += ((xy == hl_ix_iy::hl) ? 12 : 20);
		}

		regs.main.f.z = (value & mask) ? 0 : 1;
		regs.main.f.h = 1;
		regs.main.f.n = 0;

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

	static constexpr cb_handler_t dispatch_cb[256] = {
		&sim_cb00, &sim_cb00, &sim_cb00, &sim_cb00, &sim_cb00, &sim_cb00, &sim_cb00, &sim_cb00, // 00 - 07
		&sim_cb08, &sim_cb08, &sim_cb08, &sim_cb08, &sim_cb08, &sim_cb08, &sim_cb08, &sim_cb08, // 08 - 0f
		&sim_cb10, &sim_cb10, &sim_cb10, &sim_cb10, &sim_cb10, &sim_cb10, &sim_cb10, &sim_cb10, // 10 - 17
		&sim_cb18, &sim_cb18, &sim_cb18, &sim_cb18, &sim_cb18, &sim_cb18, &sim_cb18, &sim_cb18, // 18 - 1f
		&sim_cb20, &sim_cb20, &sim_cb20, &sim_cb20, &sim_cb20, &sim_cb20, &sim_cb20, &sim_cb20, // 20 - 27
		&sim_cb28, &sim_cb28, &sim_cb28, &sim_cb28, &sim_cb28, &sim_cb28, &sim_cb28, &sim_cb28, // 28 - 2f
		nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  nullptr,  // 30 - 37
		&sim_cb38, &sim_cb38, &sim_cb38, &sim_cb38, &sim_cb38, &sim_cb38, &sim_cb38, &sim_cb38, // 38 - 3f

		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 40 - 47
		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 48 - 4f
		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 50 - 57
		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 58 - 5f
		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 60 - 67
		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 68 - 6f
		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 70 - 77
		&sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, &sim_cb40, // 78 - 7f

		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // 80 - 87
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // 88 - 8f
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // 90 - 97
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // 98 - 9f
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // a0 - a7
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // a8 - af
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // b0 - b7
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // b8 - bf

		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // c0 - c7
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // c8 - cf
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // d0 - d7
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // d8 - df
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // e0 - e7
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // e8 - ef
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // f0 - f7
		&sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, &sim_cb80, // f8 - ff
	};
	#pragma endregion

	virtual HRESULT STDMETHODCALLTYPE SimulateOne (BOOL check_breakpoints, IDeviceEventHandler* eh, IUnknown** outcome) override
	{
		if (regs.iff1)
		{
			bool interrupted;
			uint8_t irq_address;
			if (!irq->try_poll_irq_at_time_point(interrupted, irq_address, cpu_time))
				return S_FALSE;

			if (interrupted)
			{
				regs.halted = false;

				if (regs.im == 0)
				{
					WI_ASSERT(false); // not yet implemented
				}
				else if (regs.im == 1)
				{
					if (!memory->try_write_request (regs.sp - 2, regs.pc, cpu_time))
						return S_FALSE;
					regs.sp -= 2;
					regs.pc = 0x38;
					regs.iff1 = false;
					cpu_time += 13;
					WI_ASSERT(!_addr_of_irq_ret_addr); // nested interrupts not supported for now
					_addr_of_irq_ret_addr = regs.sp;
					return S_OK;
				}
				else if (regs.im == 2)
				{
					uint16_t addr = (regs.i << 8) | (irq_address & 0xFE);
					if (!memory->try_read_request (addr, addr, cpu_time))
						return S_FALSE;
					if (!memory->try_write_request (regs.sp - 2, regs.pc, cpu_time))
						return S_FALSE;
					regs.sp -= 2;
					regs.pc = addr;
					regs.iff1 = false;
					cpu_time += 19;
					WI_ASSERT(!_addr_of_irq_ret_addr); // nested interrupts not supported for now
					_addr_of_irq_ret_addr = regs.sp;
					return S_OK;
				}
				else
					WI_ASSERT(false);
			}
		}

		if (regs.halted)
		{
			cpu_time += 4;
			regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);
			return S_OK;
		}

		if (check_breakpoints)
		{
			auto it = code_bps.find(regs.pc);
			if (it != code_bps.end())
			{
				com_ptr<BreakpointCollection> bps = new (std::nothrow) BreakpointCollection(); RETURN_IF_NULL_ALLOC(bps);
				auto hr = bps->InitInstance (BreakpointType::Code, regs.pc, it->second.data(), it->second.size()); RETURN_IF_FAILED(hr);
				*outcome = bps.detach();
				return SIM_E_BREAKPOINT_HIT;
			}
		}

		uint8_t opcode;
		bool b = memory->try_read_request (regs.pc, opcode, cpu_time);
		if (!b)
			return S_FALSE;

		uint16_t oldpc = regs.pc;
		uint8_t oldr = regs.r;
		regs.pc++;
		regs.r = (regs.r & 0x80) | ((regs.r + 1) & 0x7f);

		hl_ix_iy xy = hl_ix_iy::hl;
		if (opcode == 0xDD)
		{
			xy = hl_ix_iy::ix;
			opcode = decode_u8();
		}
		else if (opcode == 0xFD)
		{
			xy = hl_ix_iy::iy;
			opcode = decode_u8();
		}

		bool executed;
		if (opcode == 0xed)
		{
			opcode = decode_u8();
			auto handler = dispatch_ed[opcode];
			if (!handler)
			{
				regs.pc = oldpc;
				regs.r = oldr;
				return SIM_E_UNDEFINED_INSTRUCTION;
			}

			executed = (this->*handler)(opcode);
		}
		else if (opcode == 0xcb)
		{
			uint16_t memhlxy_addr = decode_mem_hl(xy);
			opcode = decode_u8();
			auto handler = dispatch_cb[opcode];
			if (!handler)
			{
				regs.pc = oldpc;
				regs.r = oldr;
				return SIM_E_UNDEFINED_INSTRUCTION;
			}

			executed = (this->*handler) (xy, memhlxy_addr, opcode);
		}
		else
		{
			auto handler = dispatch[opcode];
			if (!handler)
			{
				regs.pc = oldpc;
				regs.r = oldr;
				return SIM_E_UNDEFINED_INSTRUCTION;
			}

			executed = (this->*handler)(xy, opcode);
		}

		if (!executed)
		{
			regs.pc = oldpc;
			regs.r = oldr;
			return S_FALSE;
		}

		if (_ei_countdown)
		{
			_ei_countdown--;
			if (_ei_countdown == 0)
				regs.iff1 = 1;
		}

		return S_OK;
	}

	// ========================================================================

	virtual UINT64 Time() override
	{
		return cpu_time;
	}

	virtual HRESULT STDMETHODCALLTYPE SkipTime (UINT64 offset) override
	{
		cpu_time += offset;
		return S_OK;
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return false; }

	virtual HRESULT SimulateTo (UINT64 requested_time, IDeviceEventHandler* eh) override
	{
		WI_ASSERT(false); return { };
	}

	virtual void STDMETHODCALLTYPE Reset() override
	{
		memset (&regs, 0, sizeof(regs));
		_start_of_stack = 0;
		_addr_of_irq_ret_addr.reset();
		cpu_time = 0;
	}
	/*
	virtual HRESULT STDMETHODCALLTYPE GetRegisters (IRegisterGroup** ppRegs) override
	{
		com_ptr<Z80RegisterGroup> rg = new (std::nothrow) Z80RegisterGroup(); RETURN_IF_NULL_ALLOC(rg);
		auto hr = rg->InitInstance(regs); RETURN_IF_FAILED(hr);
		*ppRegs = rg.detach();
		return S_OK;
	}
*/
	virtual void GetZ80Registers (z80_register_set* pRegs) override
	{
		*pRegs = regs;
	}

	virtual void SetZ80Registers (const z80_register_set* pRegs) override
	{
		regs = *pRegs;
	}
	
	virtual UINT16 GetStackStartAddress() const override { return _start_of_stack; }
	
	virtual BOOL STDMETHODCALLTYPE Halted() override { return regs.halted; }
	
	virtual UINT16 STDMETHODCALLTYPE GetPC() const override { return regs.pc; }

	virtual HRESULT STDMETHODCALLTYPE SetPC (UINT16 pc) override
	{
		regs.pc = pc;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE AddBreakpoint (BreakpointType type, uint16_t address, SIM_BP_COOKIE* pCookie) override
	{
		if (type == BreakpointType::Code)
		{
			auto it = code_bps.find(address);
			if (it == code_bps.end())
			{
				bool added = code_bps.try_insert({ address, vector_nothrow<SIM_BP_COOKIE>{ } });
				RETURN_HR_IF(E_OUTOFMEMORY, !added);

				auto it = code_bps.find(address);
				added = it->second.try_push_back(_nextBpCookie);
				if (!added)
				{
					code_bps.remove(it);
					RETURN_HR(E_OUTOFMEMORY);
				}
				*pCookie = _nextBpCookie;
				_nextBpCookie++;
			}
			else
			{
				bool added = it->second.try_push_back(_nextBpCookie);
				RETURN_HR_IF(E_OUTOFMEMORY, !added);
				*pCookie = _nextBpCookie;
				_nextBpCookie++;
			}

			return S_OK;
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE RemoveBreakpoint (SIM_BP_COOKIE cookie) override
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

		RETURN_HR(E_INVALIDARG);
	}

	virtual BOOL STDMETHODCALLTYPE HasBreakpoints() override
	{
		return code_bps.size();
	}

	// Let's keep this code, maybe we'll eventually try to build a real call stack, as used by high-level languages.
	/*
	virtual bool build_call_stack (cpu_call_stack_helper_i* helper, uint16_t* ptrs_to_ret_addrs, uint32_t buffer_size, uint32_t* actual_size) const override
	{
		// We try to build a call stack using some heuristics and instruction decoding.

		if (actual_size)
			*actual_size = 0;

		auto add_ptr_to_ret_addr = [ptrs_to_ret_addrs, buffer_size, actual_size] (uint32_t& entry_index, uint16_t ptr_to_ret_addr)
		{
			if (entry_index < buffer_size)
				ptrs_to_ret_addrs[entry_index] = ptr_to_ret_addr;

			entry_index++;

			if (actual_size)
				*actual_size = entry_index;
		};

		// An array entry specifies the end (highest address) of the stack frame.
		// If the entry is not the last, the entry is also a pointer where the caller's return address was written.
		// 
		// The first stack frame (the one that appears topmost in the Call Stack window)
		// starts at SP and ends at the first entry of the array.
		// 
		// The array is never empty; it contains at least the start of the stack
		// (i.e., its highest address, the one that the Z80 code initialized for example with LD SP, nn.)

		uint16_t sp = regs.sp;
		uint32_t frame_index = 0;

		// Read addresses one by one starting at 'sp' and check if the instruction before the address is a call.
		// (Building the call stack like this is a simple process, but we might have some junk bytes there
		// that happen to look like a call/rst instruction; the tradeoff is acceptable for now.)
		uint32_t frame_size_words = 0;
		while (true)
		{
			if (((_start_of_stack == 0) && (sp == 0))
				|| (_start_of_stack && (sp >= _start_of_stack)))
			{
				// Add end address of last frame (= bottom frame in the Call Stack window).
				add_ptr_to_ret_addr(frame_index, _start_of_stack);
				return true;
			}

			if (_addr_of_irq_ret_addr && sp == *_addr_of_irq_ret_addr)
			{
				add_ptr_to_ret_addr (frame_index, sp);
				sp += 2;
				frame_size_words = 0;
			}
			else
			{
				uint16_t value = memory->read_uint16(sp);

				if ((value >= 1) && ((memory->read(value - 1) & 0xC7) == 0xC7))
				{
					// rst 0/8/...
					add_ptr_to_ret_addr (frame_index, sp);
					sp += 2;
					frame_size_words = 0;
				}
				else if ((value >= 3) && ((memory->read(value - 3) == 0xCD) || ((memory->read(value - 3) & 0xC7) == 0xC4)))
				{
					// "call nn" or "call cc, nn"
					add_ptr_to_ret_addr (frame_index, sp);
					sp += 2;
					frame_size_words = 0;
				}
				else if (helper->is_known_code_location(value))
				{
					// We didn't decode a call with this address as return address, but the application tells us this is known code.
					add_ptr_to_ret_addr (frame_index, sp);
					sp += 2;
					frame_size_words = 0;
				}
				else
				{
					// not a call

					frame_size_words++;
					if ((frame_size_words == 10) || (sp >= 0xFFFE))
					{
						// We probably went haywire with our algorithm. TODO: handle this.
						WI_ASSERT(false);
						return false;
					}

					sp += 2;
				}
			}
		}
	}
	*/
};

HRESULT STDMETHODCALLTYPE MakeZ80CPU (Bus* memory, Bus* io, irq_line_i* irq, wistd::unique_ptr<IZ80CPU>* ppCPU)
{
	auto d = wil::make_unique_nothrow<cpu>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(memory, io, irq); RETURN_IF_FAILED(hr);
	*ppCPU = std::move(d);
	return S_OK;
}
