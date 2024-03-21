
#include "pch.h"
#include "DebugEngine.h"
#include "../FelixPackage.h"
#include "shared/string_builder.h"
#include "shared/OtherGuids.h"
#include "shared/z80_register_set.h"
#include "shared/com.h"

struct reg8
{
	hl_ix_iy const xy;
	uint8_t  const i;
	int8_t         xydisp;

	reg8 (hl_ix_iy xy, uint8_t i, const uint8_t*& ptr)
		: xy(xy), i(i)
	{
		if ((i == 6) && (xy != hl_ix_iy::hl))
		{
			xydisp = (int8_t)ptr[0];
			ptr++;
		}
	}

	reg8 (hl_ix_iy xy, uint8_t i, int8_t xydisp)
		: xy(xy), i(i), xydisp(xydisp)
	{ }

	reg8 (const reg8&) = delete;
	reg8& operator= (const reg8&) = delete;
};

struct reg16
{
	hl_ix_iy const xy;
	uint8_t  const i;
	bool     const af_instead_of_sp;

	reg16 (hl_ix_iy xy, uint8_t i, bool af_instead_of_sp = false)
		: xy(xy), i(i), af_instead_of_sp(af_instead_of_sp)
	{ }

	reg16(const reg16&) = delete;
	reg16& operator=(const reg16&) = delete;
};

class opersb : public wstring_builder
{
	using base = wstring_builder;

public:
	using base::base;
	using base::operator<<;

	opersb& operator<< (const reg8& r)
	{
		switch(r.i)
		{
			case 0:
				operator<<('B');
				return *this;

			case 1:
				operator<<('C');
				return *this;

			case 2:
				operator<<('D');
				return *this;

			case 3:
				operator<<('E');
				return *this;

			case 4:
				if (r.xy == hl_ix_iy::hl)
					operator<<('H');
				else if (r.xy == hl_ix_iy::ix)
					operator<<("IXH");
				else
					operator<<("IYH");
				return *this;

			case 5:
				if (r.xy == hl_ix_iy::hl)
					operator<<('L');
				else if (r.xy == hl_ix_iy::ix)
					operator<<("IXL");
				else
					operator<<("IYL");
				return *this;

			case 6:
				if (r.xy == hl_ix_iy::hl)
					operator<<("(HL)");
				else
					*this << ((r.xy == hl_ix_iy::ix) ? "(IX" : "(IY")
						<< ((r.xydisp >= 0) ? " + " : " - ")
						<< hex<uint8_t>((r.xydisp >= 0) ? (uint8_t)r.xydisp : (uint8_t)-r.xydisp)
						<< ')';
				return *this;

			default:
				operator<<('A');
				return *this;
		}
	}

	opersb& operator<< (const reg16& r)
	{
		switch (r.i)
		{
			case 0:
				operator<<("BC");
				return *this;

			case 1:
				operator<<("DE");
				return *this;

			case 2:
				if (r.xy == hl_ix_iy::hl)
					operator<<("HL");
				else if (r.xy == hl_ix_iy::ix)
					operator<<("IX");
				else
					operator<<("IY");
				return *this;

			default:
				operator<<(r.af_instead_of_sp ? "AF" : "SP");
				return *this;
		}
	}

	opersb& operator<< (hl_ix_iy xy)
	{
		if (xy == hl_ix_iy::hl)
			operator<<("HL");
		else if (xy == hl_ix_iy::ix)
			operator<<("IX");
		else
			operator<<("IY");
		return *this;
	}

	opersb& operator<< (char ch)
	{
		base::operator<<(ch);
		return *this;
	}

	opersb& operator<< (const char* str)
	{
		base::operator<<(str);
		return *this;
	}

	opersb& operator<< (const wchar_t* str)
	{
		base::operator<<(str);
		return *this;
	}

	opersb& operator<< (uint8_t n)
	{
		base::operator<<(n);
		return *this;
	}

	opersb& operator<< (hex<uint16_t> val)
	{
		base::operator<<(val);
		return *this;
	}
};

class Z80Disassembly : public IDebugDisassemblyStream2
{
	ULONG _refCount = 0;
	DISASSEMBLY_STREAM_SCOPE _dwScope;
	wil::com_ptr_nothrow<IDebugProgram2> _program;
	com_ptr<ISimulator> _simulator;
	//wil::com_ptr_nothrow<IZ80BlockInstruction> _blockInstruction;
	uint16_t _address;

public:
	HRESULT InitInstance (DISASSEMBLY_STREAM_SCOPE dwScope, IDebugProgram2* program, IDebugCodeContext2* pCodeContext, IDebugDisassemblyStream2** to)
	{
		com_ptr<IFelixCodeContext> fcc;
		auto hr = pCodeContext->QueryInterface(&fcc); RETURN_IF_FAILED(hr);

		hr = serviceProvider->QueryService(SID_Simulator, &_simulator); RETURN_IF_FAILED(hr);
		_dwScope = dwScope;
		_program = program;
		_address = (uint16_t)fcc->Address();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = NULL;

		if (TryQI<IUnknown>(this, riid, ppvObject)
			|| TryQI<IDebugDisassemblyStream2>(this, riid, ppvObject))
			return S_OK;

		if (   riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IInspectable
			|| riid == IID_INoMarshal
			|| riid == IID_IMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions)
			return E_NOINTERFACE;

		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	static constexpr const char* const r8[] = { "B", "C", "D", "E", "H", "L", "(HL)", "A" };
	static constexpr const char* const cc[] = { "NZ", "Z", "NC", "C", "PO", "PE", "P", "M" };
	static constexpr const char* const al8_instr[] = { "ADD", "ADC", "SUB", "SBC", "AND", "XOR", "OR", "CP" };
	static constexpr const char* const cb_instrs_0[] = { "RLC", "RRC", "RL", "RR", "SLA", "SRA", "SLL", "SRL" };
	static constexpr const char* const cb_instrs_40[] = { "BIT", "RES", "SET" };

	// This function assumes that memory at "bytes" contains enough bytes for the longest Z80 instruction (four bytes or so).
	// It returns the length of the decoded instruction.
	uint8_t DecodeInstruction (uint16_t address, const uint8_t* bytes, const char*& opcode, opersb& operands, bool useOperandSymbols)
	{
		auto ptr = bytes;

		hl_ix_iy xy = hl_ix_iy::hl;
		if (ptr[0] == 0xDD)
		{
			xy = hl_ix_iy::ix;
			ptr++;
		}
		else if (ptr[0] == 0xFD)
		{
			xy = hl_ix_iy::iy;
			ptr++;
		}

		if (ptr[0] == 0)
		{
			opcode = "NOP";
			ptr++;
		}
		else if ((ptr[0] & 0xCF) == 1)
		{
			// ld bc/de/hl/sp, nn
			opcode = "LD";
			auto reg = reg16(xy, (ptr[0] >> 4) & 3);
			uint16_t val = ptr[1] | (ptr[2] << 8);
			wil::unique_bstr symbol;
			uint16_t offset;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(val, SK_Both, nullptr, &symbol, &offset, nullptr)))
				operands << reg << ", " << symbol.get() << " + " << offset << " ; " << hex<uint16_t>(val);
			else
				operands << reg << ", " << hex<uint16_t>(val);
			ptr += 3;
		}
		else if (ptr[0] == 2)
		{
			opcode = "LD";
			operands << "(BC), A";
			ptr++;
		}
		else if ((ptr[0] & 0xC7) == 0x3)
		{
			// inc/dec bc/de/hl/sp
			opcode = (ptr[0] & 8) ? "DEC" : "INC";
			operands << reg16(xy, (ptr[0] >> 4) & 3);
			ptr++;
		}
		else if ((ptr[0] & 0xC6) == 4)
		{
			// INC/DEC r8
			opcode = (ptr[0] & 1) ? "DEC" : "INC";
			uint8_t i = (ptr[0] >> 3) & 7;
			ptr++;
			operands << reg8(xy, i, ptr);
		}
		else if ((ptr[0] & 0xC7) == 6)
		{
			// ld r8, imm8
			opcode = "LD";
			uint8_t i = (ptr[0] >> 3) & 7;
			ptr++;
			operands << reg8(xy, i, ptr) << ", " << hex<uint8_t>(*ptr++);
		}
		else if (ptr[0] == 7)
		{
			opcode = "RLCA";
			ptr++;
		}
		else if (ptr[0] == 8)
		{
			opcode = "EX";
			operands << "AF, AF'";
			ptr++;
		}
		else if (ptr[0] == 0x0A)
		{
			opcode = "LD";
			operands << "A, (BC)";
			ptr++;
		}
		else if (ptr[0] == 0xF)
		{
			opcode = "RRCA";
			ptr++;
		}
		else if (ptr[0] == 0x12)
		{
			opcode = "LD";
			operands << "(DE), A";
			ptr++;
		}
		else if (ptr[0] == 0x17)
		{
			opcode = "RLA";
			ptr++;
		}
		else if (ptr[0] == 0x1A)
		{
			opcode = "LD";
			operands << "A, (DE)";
			ptr++;
		}
		else if (ptr[0] == 0x1F)
		{
			opcode = "RRA";
			ptr++;
		}
		else if ((ptr[0] == 0x10) || (ptr[0] == 0x18))
		{
			// djnz/jr e
			opcode = (ptr[0] == 0x10) ? "DJNZ" : "JR";
			uint16_t dest = (uint16_t)(address + 2 + (uint16_t)(int16_t)(int8_t)ptr[1]);
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(dest, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << symbol.get() << " ; ";
			operands << hex<uint16_t>(dest);
			ptr += 2;
		}
		else if ((ptr[0] & 0xE7) == 0x20)
		{
			// jr cc, e
			opcode = "JR";
			uint16_t dest = (uint16_t)(address + 2 + (uint16_t)(int16_t)(int8_t)ptr[1]);
			operands << cc[(ptr[0] >> 3) & 3] << ", ";
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(dest, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << symbol.get() << " ; ";
			operands << hex<uint16_t>(dest);
			ptr += 2;
		}
		else if (ptr[0] == 0x22)
		{
			// LD (nn), HL
			uint16_t dest = ptr[1] | (ptr[2] << 8);
			opcode = "LD";
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(dest, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << '(' << symbol.get() << "), " << xy << " ; " << hex<uint16_t>(dest);
			else
				operands << '(' << hex<uint16_t>(dest) << "), " << xy;
			ptr += 3;
		}
		else if (ptr[0] == 0x27)
		{
			// DAA
			opcode = "DAA";
			ptr++;
		}
		else if (ptr[0] == 0x2A)
		{
			// LD HL, (nn)
			uint16_t src = ptr[1] | (ptr[2] << 8);
			opcode = "LD";
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(src, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << xy << ", (" << symbol.get() << ") ; " << hex<uint16_t>(src);
			else
				operands << xy << ", (" << hex<uint16_t>(src) << ")";
			ptr += 3;
		}
		else if (ptr[0] == 0x2F)
		{
			// CPL
			opcode = "CPL";
			ptr++;
		}
		else if (ptr[0] == 0x32)
		{
			// ld (imm16), a
			opcode = "LD";
			uint16_t dest = ptr[1] | (ptr[2] << 8);
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(dest, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << '(' << symbol.get() << "), A ; " << hex<uint16_t>(dest);
			else
				operands << '(' << hex<uint16_t>(dest) << "), A";
			ptr += 3;
		}
		else if (ptr[0] == 0x37)
		{
			opcode = "SCF";
			ptr++;
		}
		else if (ptr[0] == 0x3A)
		{
			// LD A, (nn)
			opcode = "LD";
			operands << "A, (" << hex<uint16_t>(ptr[1] | (ptr[2] << 8)) << ")";
			ptr += 3;
		}
		else if (ptr[0] == 0x3F)
		{
			opcode = "CCF";
			ptr++;
		}
		else if (((ptr[0] & 0xC0) == 0x40) && (ptr[0] != 0x76))
		{
			// ld r8, r8
			opcode = "LD";
			uint8_t idst = (ptr[0] >> 3) & 7;
			uint8_t isrc = ptr[0] & 7;
			ptr++;
			// If one operand is (hl), any DD/FD prexif applies only to that operand, and the other operand is without prefix.
			// If no operand is (hl), any DD/FD prefix applies to both operands, and that's an undocumented instruction.
			hl_ix_iy xydst = xy;
			hl_ix_iy xysrc = xy;
			if (idst == 6)
				xysrc = hl_ix_iy::hl;
			else if (isrc == 6)
				xydst = hl_ix_iy::hl;
			operands << reg8(xydst, idst, ptr) << ", " << reg8(xysrc, isrc, ptr);
		}
		else if (ptr[0] == 0x76)
		{
			opcode = "HALT";
			ptr++;
		}
		else if ((ptr[0] & 0xC0) == 0x80)
		{
			// add/adc/sub/sbc/and/xor/or/cp r8/(hl)/(ix+d)/(iy+d)
			opcode = al8_instr[(ptr[0] >> 3) & 7];
			uint8_t i = ptr[0] & 7;
			ptr++;
			operands << "A, " << reg8(xy, i, ptr);
		}
		else if ((ptr[0] & 0xC7) == 0xC0)
		{
			// ret nz/z/nc/c/po/pe/p/m
			opcode = "RET";
			operands << cc[(ptr[0] >> 3) & 7];
			ptr++;
		}
		else if ((ptr[0] & 0xCB) == 0xC1)
		{
			// pop/push bc/de/hl/af
			opcode = (ptr[0] & 4) ? "PUSH" : "POP";
			uint8_t i = (ptr[0] >> 4) & 3;
			operands << reg16(xy, i, true);
			ptr++;
		}
		else if ((ptr[0] & 0xC7) == 0xC2)
		{
			// jp nz/z/nc/c/po/pe/p/m, nn
			opcode = "JP";
			operands << cc[(ptr[0] >> 3) & 7] << ", " << hex<uint16_t>(ptr[1] | (ptr[2] << 8));
			ptr += 3;
		}
		else if (ptr[0] == 0xC3)
		{
			// jp nn
			opcode = "JP";
			uint16_t dest = ptr[1] | (ptr[2] << 8);
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(dest, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << symbol.get() << " ; " << hex<uint16_t>(dest);
			else
				operands << hex<uint16_t>(dest);
			ptr += 3;
		}
		else if ((ptr[0] & 0xC7) == 0xC4)
		{
			// call nz/z/nc/c/po/pe/p/m, nn
			opcode = "CALL";
			operands << cc[(ptr[0] >> 3) & 7] << ", " << hex<uint16_t>(ptr[1] | (ptr[2] << 8));
			ptr += 3;
		}
		else if ((ptr[0] & 0xC7) == 0xC6)
		{
			// add/adc/sub/sbc/and/xor/or/cp a, imm8
			opcode = al8_instr[(ptr[0] >> 3) & 7];
			operands << "A, " << hex<uint8_t>(ptr[1]);
			ptr += 2;
		}
		else if ((ptr[0] & 0xC7) == 0xC7)
		{
			// RST #
			uint16_t addr = ptr[0] & 0x38;
			opcode = "RST";
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(addr, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << symbol.get() << " ; " << hex<uint16_t>(addr);
			else
				operands << hex<uint16_t>(addr);
			ptr++;
		}
		else if (ptr[0] == 0xC9)
		{
			// ret
			opcode = "RET";
			ptr++;
		}
		else if (ptr[0] == 0xCD)
		{
			// call nn
			opcode = "CALL";
			uint16_t dest = ptr[1] | (ptr[2] << 8);
			wil::unique_bstr symbol;
			if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(dest, SK_Both, nullptr, &symbol, nullptr, nullptr)))
				operands << symbol.get() << " ; " << hex<uint16_t>(dest);
			else
				operands << hex<uint16_t>(dest);
			ptr += 3;
		}
		else if ((ptr[0] & 0xCF) == 9)
		{
			// add hl, bc/de/hl/sp
			opcode = "ADD";
			operands << xy << ", " << reg16(xy, (ptr[0] >> 4) & 3);
			ptr++;
		}
		else if (ptr[0] == 0xD3)
		{
			opcode = "OUT";
			operands << '(' << hex<uint8_t>(ptr[1]) << "), A";
			ptr += 2;
		}
		else if (ptr[0] == 0xD9)
		{
			opcode = "EXX";
			ptr++;
		}
		else if (ptr[0] == 0xDB)
		{
			opcode = "IN";
			operands << "A, (" << hex<uint8_t>(ptr[1]) << ')';
			ptr += 2;
		}
		else if (ptr[0] == 0xE3)
		{
			opcode = "EX";
			operands << "(SP), " << xy;
			ptr++;
		}
		else if (ptr[0] == 0xE9)
		{
			opcode = "JP";
			operands << '(' << xy << ')';
			ptr++;
		}
		else if (ptr[0] == 0xEB)
		{
			opcode = "EX";
			operands << "DE, HL";
			ptr++;
		}
		else if (ptr[0] == 0xF3)
		{
			opcode = "DI";
			ptr++;
		}
		else if (ptr[0] == 0xF9)
		{
			// ld sp, hl/ix/iy
			opcode = "LD";
			operands << "SP, " << xy;
			ptr++;
		}
		else if (ptr[0] == 0xFB)
		{
			opcode = "EI";
			ptr++;
		}
		else if (ptr[0] == 0xED)
		{
			// The CPU ignores any DD/FD prefix before ED.
			if ((ptr[1] & 0xC7) == 0x40)
			{
				opcode = "IN";
				operands << r8[(ptr[1] >> 3) & 7] << ", (BC)";
				ptr += 2;
			}
			else if ((ptr[1] & 0xCF) == 0x43)
			{
				// ld (nn), bc/de/hl/sp
				opcode = "LD";
				uint16_t dest = ptr[2] | (ptr[3] << 8);
				uint8_t reg = (ptr[1] >> 4) & 3;
				wil::unique_bstr symbol;
				if (useOperandSymbols && SUCCEEDED(GetSymbolFromAddress(dest, SK_Both, nullptr, &symbol, nullptr, nullptr)))
					operands << '(' << symbol.get() << "), " << reg16(xy, reg) << " ; " << hex<uint16_t>(dest);
				else
					operands << '(' << hex<uint16_t>(dest) << "), " << reg16(xy, reg);
				ptr += 4;
			}
			else if (ptr[1] == 0x44)
			{
				opcode = "NEG";
				ptr += 2;
			}
			else if ((ptr[1] & 0xCF) == 0x4B)
			{
				// ld bc/de/hl/sp, (nn)
				opcode = "LD";
				operands << reg16(xy, (ptr[1] >> 4) & 3) << ", (" << hex<uint16_t>(ptr[2] | (ptr[3] << 8)) << ')';
				ptr += 4;
			}
			else if (ptr[1] == 0x45)
			{
				opcode = "RETN";
				ptr += 2;
			}
			else if (ptr[1] == 0x46)
			{
				opcode = "IM";
				operands << "0";
				ptr += 2;
			}
			else if (ptr[1] == 0x4D)
			{
				opcode = "RETI";
				ptr += 2;
			}
			else if (ptr[1] == 0x56)
			{
				opcode = "IM";
				operands << "1";
				ptr += 2;
			}
			else if (ptr[1] == 0x5e)
			{
				opcode = "IM";
				operands << "2";
				ptr += 2;
			}
			else if ((ptr[1] & 0xCF) == 0x42)
			{
				// sbc hl, bc/de/hl/sp
				opcode = "SBC";
				operands << "HL, " << reg16(xy, (ptr[1] >> 4) & 3);
				ptr += 2;
			}
			else if (ptr[1] == 0x47)
			{
				opcode = "LD";
				operands << "I, A";
				ptr += 2;
			}
			else if (ptr[1] == 0x4F)
			{
				opcode = "LD";
				operands << "R, A";
				ptr += 2;
			}
			else if (ptr[1] == 0x57)
			{
				opcode = "LD";
				operands << "A, I";
				ptr += 2;
			}
			else if (ptr[1] == 0x5F)
			{
				opcode = "LD";
				operands << "A, R";
				ptr += 2;
			}
			else if (ptr[1] == 0x67)
			{
				// rrd
				opcode = "RRD";
				ptr += 2;
			}
			else if (ptr[1] == 0x6F)
			{
				// rld
				opcode = "RLD";
				ptr += 2;
			}
			else if (ptr[1] == 0xA0)
			{
				opcode = "LDI";
				ptr += 2;
			}
			else if (ptr[1] == 0xA1)
			{
				opcode = "CPI";
				ptr += 2;
			}
			else if (ptr[1] == 0xA8)
			{
				opcode = "LDD";
				ptr += 2;
			}
			else if (ptr[1] == 0xA9)
			{
				opcode = "CPD";
				ptr += 2;
			}
			else if (ptr[1] == 0xB0)
			{
				opcode = "LDIR";
				ptr += 2;
			}
			else if (ptr[1] == 0xB1)
			{
				opcode = "CPIR";
				ptr += 2;
			}
			else if (ptr[1] == 0xB8)
			{
				opcode = "LDDR";
				ptr += 2;
			}
			else if (ptr[1] == 0xB9)
			{
				opcode = "CPDR";
				ptr += 2;
			}
			else
			{
				opcode = "??";
				ptr += 2;
			}
		}
		else if (ptr[0] == 0xCB)
		{
			ptr++;
			
			// As per https://clrhome.org/table/, if there's a DD/FD before CB, the CPU reads a displacement before reading the opcode.
			int8_t xydisp;
			if (xy != hl_ix_iy::hl)
				xydisp = (int8_t)*ptr++;

			if ((ptr[0] & 0xC0) == 0)
			{
				opcode = cb_instrs_0[(ptr[0] >> 3) & 7];
				uint8_t i = ptr[0] & 7;
				ptr++;
				operands << reg8(xy, i, (xy != hl_ix_iy::hl) ? xydisp : 0);
			}
			else
			{
				opcode = cb_instrs_40[(ptr[0] >> 6) - 1];
				uint8_t bit = (ptr[0] >> 3) & 7;
				uint8_t i = ptr[0] & 7;
				ptr++;
				operands << bit << ", " << reg8(xy, i, (xy != hl_ix_iy::hl) ? xydisp : 0);
			}
		}
		else
		{
			opcode = "??";
			ptr++;
		}

		return (uint8_t)(ptr - bytes);
	}

	HRESULT GetSymbolFromAddress(
		__RPC__in uint16_t address,
		__RPC__in SymbolKind searchKind,
		__RPC__deref_out_opt SymbolKind* foundKind,
		__RPC__deref_out_opt BSTR* foundSymbol,
		__RPC__deref_out_opt UINT16* foundOffset,
		__RPC__deref_out_opt IDebugModule2** foundModule)
	{
		wil::com_ptr_nothrow<IDebugModule2> m;
		auto hr = ::GetModuleAtAddress (_program.get(), address, &m); RETURN_IF_FAILED_EXPECTED(hr);

		wil::com_ptr_nothrow<IZ80Module> z80m;
		hr = m->QueryInterface(&z80m); RETURN_IF_FAILED(hr);

		wil::com_ptr_nothrow<IZ80Symbols> syms;
		hr = z80m->GetSymbols(&syms); RETURN_IF_FAILED_EXPECTED(hr);

		hr = syms->GetSymbolAtAddress(address, searchKind, foundKind, foundSymbol, foundOffset); RETURN_IF_FAILED_EXPECTED(hr);

		if (foundModule)
			*foundModule = m.detach();

		return S_OK;
	}
	
	#pragma region IDebugDisassemblyStream2
	virtual HRESULT __stdcall Read(DWORD dwInstructions, DISASSEMBLY_STREAM_FIELDS dwFields, DWORD* pdwInstructionsRead, DisassemblyData* prgDisassembly) override
	{
		HRESULT hr;

		DWORD& i = *pdwInstructionsRead;
		for (i = 0; i < dwInstructions; i++)
		{
			auto dd = &prgDisassembly[i];
			dd->dwFields = 0;

			if (dwFields & DSF_ADDRESS) // 1
			{
				wchar_t str[16];
				swprintf_s(str, L"%04X", _address);
				dd->bstrAddress = SysAllocString(str);
				dd->dwFields |= DSF_ADDRESS;
			}

			if (dwFields & DSF_ADDRESSOFFSET) // 2
			{
				// If we set this information, the IDE shows this instead of what we set for DSF_ADDRESS. We don't want this.
				//dd->bstrAddressOffset = SysAllocString(L"<addressoffset>");
				//dd->dwFields |= DSF_ADDRESSOFFSET;
			}

///			uint32_t blockAddr, blockSize;
///			if (!_blockInstruction
///				|| FAILED(_blockInstruction->GetRange(&blockAddr, &blockSize))
///				|| ((uint32_t)_address < blockAddr) || ((uint32_t)_address >= blockAddr + blockSize))
///			{
///				_blockInstruction = nullptr;
///				// Purposefully ignoring any returned error here. Below in this function We check _blockInstruction before using it.
///				_program->TryDecodeBlockInstruction((uint32_t)_address, nullptr, &_blockInstruction);
///			}
///
			uint32_t decodedLength = 0;
			wil::unique_bstr bstrOpcode;
			wil::unique_bstr bstrOperands;
///			if (_blockInstruction)
///			{
///				hr = _blockInstruction->Decode((uint32_t)_address, &decodedLength, &bstrOpcode, &bstrOperands);
///			}

			if (!decodedLength)
			{
				const char* opcode;
				opersb operands;
				wistd::unique_ptr<uint8_t[]> decodeBuffer;
				static constexpr uint32_t longest_expected = 4;
				decodeBuffer = wil::make_unique_nothrow<uint8_t[]>(longest_expected); 
				hr = _simulator->ReadMemoryBus(_address, (uint16_t)longest_expected, decodeBuffer.get()); RETURN_IF_FAILED(hr);
				decodedLength = DecodeInstruction (_address, decodeBuffer.get(), opcode, operands, dwFields & DSF_OPERANDS_SYMBOLS);
				WI_ASSERT(decodedLength <= longest_expected);
				hr = MakeBstrFromString (opcode, &bstrOpcode); RETURN_IF_FAILED(hr);
				bstrOperands.reset (SysAllocStringLen(operands.data(), operands.size())); RETURN_IF_NULL_ALLOC(bstrOperands);
			}

			// If the instruction we decoded starts before PC and ends after it,
			// we're doing something wrong, so let's replace it with some question marks.
			uint16_t pc;
			hr = _simulator->GetPC(&pc); RETURN_IF_FAILED(hr);
			if (pc > _address && pc < _address + decodedLength)
			{
				bstrOpcode = wil::make_bstr_nothrow(L"??");
				bstrOperands = nullptr;
				decodedLength = pc - _address;
			}

			if (dwFields & DSF_CODEBYTES) // 4
			{
				uint32_t opcodeLength = std::min(8u, decodedLength);
				auto decodeBuffer = wil::make_unique_nothrow<uint8_t[]>(opcodeLength); RETURN_IF_NULL_ALLOC(decodeBuffer);
				hr = _simulator->ReadMemoryBus (_address, opcodeLength, decodeBuffer.get()); RETURN_IF_FAILED(hr);
				wchar_t buffer[64];
				swprintf_s (buffer, L"%02X", decodeBuffer.get()[0]);
				for (uint32_t i = 1; i < opcodeLength; i++)
				{
					wcscat_s (buffer, L" ");
					wchar_t b[5];
					swprintf_s (b, L"%02X", decodeBuffer.get()[i]);
					wcscat_s (buffer, b);
				}

				dd->bstrCodeBytes = SysAllocString(buffer);
				dd->dwFields |= DSF_CODEBYTES;
			}

			if (dwFields & DSF_OPCODE) // 8
			{
				dd->bstrOpcode = bstrOpcode.release(); // TODO: free if rest of function fails
				dd->dwFields |= DSF_OPCODE;
			}

			if (dwFields & DSF_OPERANDS) // 0x10
			{
				dd->bstrOperands = bstrOperands.release(); // TODO: free if rest of function fails
				dd->dwFields |= DSF_OPERANDS;
			}

			if (dwFields & DSF_SYMBOL) // 0x20
			{
				wil::unique_bstr symbol;
				hr = GetSymbolFromAddress(_address, SK_Both, nullptr, &symbol, nullptr, nullptr);
				if (SUCCEEDED(hr) && symbol)
				{
					dd->bstrSymbol = symbol.release();
					dd->dwFields |= DSF_SYMBOL;
				}
			}

			if (dwFields & DSF_CODELOCATIONID) // 0x40
			{
				dd->uCodeLocationId = _address;
				dd->dwFields |= DSF_CODELOCATIONID;
			}

			if (dwFields & DSF_POSITION) // 0x80
			{
				//dd->posBeg = ...;
				//dd->posEnd = ...;
				//dd->dwFields |= DSF_POSITION;
			}

			if (dwFields & DSF_DOCUMENTURL) // 0x100
			{
				// ...
			}

			if (dwFields & DSF_BYTEOFFSET) // 0x200
			{
				// The number of bytes the instruction is from the beginning of the code line.
				dd->dwByteOffset = 0;
				dd->dwFields |= DSF_BYTEOFFSET;
			}

			if (dwFields & DSF_FLAGS) // 0x400
			{
				dd->dwFlags = 0;

				if (_simulator->Running_HR() == S_FALSE)
				{
					// Silly Disassembly Window sometimes calls this function while code is running, that's why the 'if'.
					z80_register_set regs;
					hr = _simulator->GetRegisters(&regs, (uint32_t)sizeof(regs)); RETURN_IF_FAILED(hr);
					if (regs.pc == _address)
						dd->dwFlags |= DF_INSTRUCTION_ACTIVE;
				}

				dd->dwFields |= DSF_FLAGS;
			}

			_address += decodedLength;
		}

		return S_OK;
	}

	virtual HRESULT __stdcall Seek(SEEK_START dwSeekStart, IDebugCodeContext2* pCodeContext, UINT64 uCodeLocationId, INT64 iInstructions) override
	{
///		_blockInstruction = nullptr;

		if (dwSeekStart == SEEK_START_CURRENT)
		{
		}
		else if (dwSeekStart == SEEK_START_CODELOCID)
		{
			_address = (uint16_t)uCodeLocationId;
		}
		else if (dwSeekStart == SEEK_START_CODECONTEXT)
		{
			wil::com_ptr_nothrow<IFelixCodeContext> fcc;
			auto hr = pCodeContext->QueryInterface(&fcc); RETURN_IF_FAILED(hr);
			_address = (uint16_t)fcc->Address();
		}
		else
			RETURN_HR(E_NOTIMPL);

		if (iInstructions < 0)
		{
			// Seek backwards
			// TODO: seek instructions instead of bytes
			if (_address >= (uint16_t)-iInstructions)
			{
				_address -= (uint16_t)-iInstructions;
				return S_OK;
			}
			else
			{
				_address = 0;
				return S_FALSE;
			}
		}
		else
		{
			// Seek forward
			for (INT64 i = 0; i < iInstructions; i++)
			{
				uint8_t bytes[6];
				_simulator->ReadMemoryBus (_address, sizeof(bytes), bytes);
				const char* dummy_opcode;
				opersb dummy_operands;
				uint8_t instructionLength = DecodeInstruction (_address, bytes, dummy_opcode, dummy_operands, false);
				if ((uint32_t)_address + instructionLength >= 0x10000)
					return S_FALSE;
				_address += instructionLength;
			}

			return S_OK;
		}
	}

	virtual HRESULT __stdcall GetCodeLocationId(IDebugCodeContext2* pCodeContext, UINT64* puCodeLocationId) override
	{
		wil::com_ptr_nothrow<IFelixCodeContext> fcc;
		auto hr = pCodeContext->QueryInterface(&fcc); RETURN_IF_FAILED(hr);
		*puCodeLocationId = fcc->Address();
		return S_OK;
	}

	virtual HRESULT __stdcall GetCodeContext(UINT64 uCodeLocationId, IDebugCodeContext2** ppCodeContext) override
	{
		bool physicalMemorySpace = false;
		return MakeDebugContext (physicalMemorySpace, uCodeLocationId, _program.get(), IID_PPV_ARGS(ppCodeContext));
	}

	virtual HRESULT __stdcall GetCurrentLocation(UINT64* puCodeLocationId) override
	{
		*puCodeLocationId = _address;
		return S_OK;
	}

	virtual HRESULT __stdcall GetDocument(BSTR bstrDocumentUrl, IDebugDocument2** ppDocument) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT __stdcall GetScope(DISASSEMBLY_STREAM_SCOPE* pdwScope) override
	{
		*pdwScope = _dwScope;
		return S_OK;
	}

	virtual HRESULT __stdcall GetSize(UINT64* pnSize) override
	{
		*pnSize = 0x10000;
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeDisassemblyStream (DISASSEMBLY_STREAM_SCOPE dwScope, IDebugProgram2* program, IDebugCodeContext2* pCodeContext, IDebugDisassemblyStream2** to)
{
	auto p = com_ptr(new (std::nothrow) Z80Disassembly()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance (dwScope, program, pCodeContext, to); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
