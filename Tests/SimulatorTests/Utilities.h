
#pragma once
#include "Simulator.h"
#include "shared/com.h"

inline constexpr UINT64 TicksPerFrame = 3500000 / 50;

uint8_t ReadMemory(ISimulator* simulator, uint16_t address);
void RunUntilTime(ISimulator* simulator, UINT64 targetTime, DWORD timeoutMilliseconds = 10000);

struct ScreenImage
{
	wil::unique_process_heap_ptr<uint32_t[]> pixels;
	DWORD width = 0;
	DWORD height = 0;

	unsigned long GetPixel(DWORD x, DWORD y) const;
};

ScreenImage CaptureScreen(ISimulator* simulator, BOOL includeBorder);
