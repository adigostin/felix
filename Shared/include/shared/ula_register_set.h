
#pragma once

struct zx_spectrum_ula_regs
{
	uint32_t frame_time;
	uint16_t line_ticks;
	uint16_t col_ticks;
	bool irq;
};

