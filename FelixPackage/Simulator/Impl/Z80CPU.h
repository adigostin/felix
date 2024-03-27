
#pragma once
#include "shared/z80_register_set.h"
#include "Bus.h"

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{97DCFDDE-9DCD-4FDB-A373-15FFB9EDF488}") IZ80CPU : ICPU
{
	virtual void GetZ80Registers (z80_register_set* pRegs) = 0;
	virtual void SetZ80Registers (const z80_register_set* pRegs) = 0;

	#ifdef SIM_TESTS
	virtual z80_register_set* GetRegsPtr() = 0;
	#endif
};

HRESULT STDMETHODCALLTYPE MakeZ80CPU (Bus* memory, Bus* io, irq_line_i* irq, wistd::unique_ptr<IZ80CPU>* ppCPU);
