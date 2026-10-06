#include <CppUnitTest.h>
#include "Simulator.h"
#include "Impl/SimulatorInternal.h"
#include "Utilities.h"
#include "shared/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace Z80SimulatorTests
{
	static constexpr UINT64 TStatesPerLine48K = 224;
	static constexpr UINT64 FirstScreenLine48K = 64;

	static void Initialize48KRedInkBluePaper(ISimulator* simulator, const uint8_t* program, uint16_t programSize)
	{
		HRESULT hr = simulator->Reset(0x8000, SpectrumVariant48K);
		Assert::AreEqual(S_OK, hr);
		uint8_t videoMemory[0x1B00] = {};
		memset(videoMemory + 0x1800, 0x0A, 0x300); // Red ink, blue paper.
		hr = simulator->WriteMemoryBus(0x4000, sizeof(videoMemory), videoMemory);
		Assert::AreEqual(S_OK, hr);
		hr = simulator->WriteMemoryBus(0x8000, programSize, program);
		Assert::AreEqual(S_OK, hr);
	}

	TEST_CLASS(ULA48KTests)
	{
	public:
		TEST_METHOD(SaveScreen)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_OK, simulator->Reset(32768, SpectrumVariant48K));

			UINT64 time;
			hr = simulator->GetTime(&time);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<UINT64>(0, time);

			static const uint8_t program[] = {
				0x3E, 0x01,             // LD A, 0x01
				0xD3, 0xFE,             // OUT (0xFE), A
				0x21, 0x00, 0x40,       // LD HL, 0x4000
				0x11, 0x01, 0x40,       // LD DE, 0x4001
				0x01, 0xFF, 0x17,       // LD BC, 0x17FF
				0xAF,                   // XOR A
				0x77,                   // LD (HL), A
				0xED, 0xB0,             // LDIR
				0x21, 0x00, 0x58,       // LD HL, 0x5800
				0x11, 0x01, 0x58,       // LD DE, 0x5801
				0x01, 0xFF, 0x02,       // LD BC, 0x02FF
				0x3E, 0x10,             // LD A, 0x10
				0x77,                   // LD (HL), A
				0xED, 0xB0,             // LDIR
				0x18, 0xFE,             // JR $-2
			};
			Assert::AreEqual(S_OK, simulator->WriteMemoryBus(32768, static_cast<uint16_t>(sizeof(program)), program));

			const UINT16 loopAddress = static_cast<UINT16>(32768 + sizeof(program) - 2);
			for (;;)
			{
				UINT16 pc;
				Assert::AreEqual(S_OK, simulator->GetPC(&pc));
				if (pc == loopAddress)
					break;
				Assert::AreEqual(S_OK, simulator->SimulateOne());
			}

			auto image = CaptureScreen(simulator.get(), TRUE);
			bool borderIsBlue = true;
			bool screenIsRed = true;
			for (DWORD y = 0; y < image.height; y++)
			{
				for (DWORD x = 0; x < image.width; x++)
				{
					uint32_t pixel = image.pixels[static_cast<size_t>(y) * image.width + x];
					bool isScreenPixel = x >= 48 && x < 304 && y >= 48 && y < 240; // 256 x 192 active pixels
					if (isScreenPixel)
						screenIsRed &= pixel == 0xFFC00000; // Opaque red
					else
						borderIsBlue &= pixel == 0xFF0000C0; // Opaque blue
				}
			}

			Assert::IsTrue(borderIsBlue);
			Assert::IsTrue(screenIsRed);
		}

		TEST_METHOD(BitmapAddressingAndPixelBitOrder)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);

			static const uint8_t program[] = {
				0x3E, 0x80, 0x32, 0x00, 0x40, // LD A,80h; LD (4000h),A: pixel x=0, y=0
				0x3E, 0x40, 0x32, 0x00, 0x41, // LD A,40h; LD (4100h),A: pixel x=1, y=1
				0x3E, 0x20, 0x32, 0x20, 0x40, // LD A,20h; LD (4020h),A: pixel x=2, y=8
				0x3E, 0x08, 0x32, 0x00, 0x47, // LD A,08h; LD (4700h),A: pixel x=4, y=7
				0x18, 0xFE                    // JR $-2
			};
			Initialize48KRedInkBluePaper(simulator, program, sizeof(program));
			for (size_t i = 0; i < 8; i++)
				simulator->SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x40, ReadMemory(simulator, 0x4100));
			Assert::AreEqual<uint8_t>(0x20, ReadMemory(simulator, 0x4020));
			Assert::AreEqual<uint8_t>(0x08, ReadMemory(simulator, 0x4700));

			RunUntilTime(simulator, TicksPerFrame);
			auto image = CaptureScreen(simulator, FALSE);
			Assert::AreEqual(0xFFC00000ul, image.GetPixel(0, 0)); // Opaque normal red ink.
			Assert::AreEqual(0xFF0000C0ul, image.GetPixel(1, 0)); // Opaque normal blue paper.
			Assert::AreEqual(0xFFC00000ul, image.GetPixel(1, 1)); // Opaque normal red ink.
			Assert::AreEqual(0xFFC00000ul, image.GetPixel(2, 8)); // Opaque normal red ink.
			Assert::AreEqual(0xFF0000C0ul, image.GetPixel(3, 8)); // Opaque normal blue paper.
			Assert::AreEqual(0xFFC00000ul, image.GetPixel(4, 7)); // Opaque normal red ink.
		}

		TEST_METHOD(AttributeCellAndBrightColors)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);

			static const uint8_t program[] = {
				0x3E, 0x80,                   // LD A,80h: set the leftmost bitmap pixel.
				0x32, 0x00, 0x40,             // LD (4000h),A: set pixel x=0.
				0x32, 0x01, 0x40,             // LD (4001h),A: set pixel x=8.
				0x3E, 0x74,                   // LD A,74h: bright green ink, bright yellow paper.
				0x32, 0x00, 0x58,             // LD (5800h),A
				0x18, 0xFE                    // JR $-2
			};
			Initialize48KRedInkBluePaper(simulator, program, sizeof(program));
			for (size_t i = 0; i < 5; i++)
				simulator->SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x80, ReadMemory(simulator, 0x4001));
			Assert::AreEqual<uint8_t>(0x74, ReadMemory(simulator, 0x5800));

			RunUntilTime(simulator, TicksPerFrame);
			auto image = CaptureScreen(simulator, FALSE);
			Assert::AreEqual(0xFF00FF00ul, image.GetPixel(0, 0)); // Opaque bright green ink.
			Assert::AreEqual(0xFFFFFF00ul, image.GetPixel(1, 0)); // Opaque bright yellow paper.
			Assert::AreEqual(0xFFC00000ul, image.GetPixel(8, 0)); // Opaque normal red ink.
			Assert::AreEqual(0xFF0000C0ul, image.GetPixel(9, 0)); // Opaque normal blue paper.
		}

		TEST_METHOD(BorderColorFromPortFE)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);

			static const uint8_t program[] = {
				0x3E, 0x02,             // LD A,2: red border
				0x32, 0x00, 0xC0,       // LD (C000h),A for a memory-side-effect assertion
				0xD3, 0xFE,             // OUT (FEh),A
				0x18, 0xFE              // JR $-2
			};
			Initialize48KRedInkBluePaper(simulator, program, sizeof(program));
			simulator->SimulateOne();
			simulator->SimulateOne();
			simulator->SimulateOne();
			Assert::AreEqual<uint8_t>(0x02, ReadMemory(simulator, 0xC000));

			// Allow a subsequent complete 48K frame to use the newly selected border color.
			RunUntilTime(simulator, 2 * TicksPerFrame);
			auto image = CaptureScreen(simulator, TRUE);
			Assert::AreEqual(0xFFC00000ul, image.GetPixel(0, 0)); // Opaque normal red border.
			Assert::AreEqual(0xFFC00000ul, image.GetPixel(24, 100)); // Opaque normal red border.
		}

		TEST_METHOD(CrtSnapshotShowsPixelAsBeamReachesItsScanline)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);

			static const uint8_t program[] = {
				0x3E, 0x80,                   // LD A,80h
				0x32, 0x88, 0x40,             // LD (4088h),A: set pixel x=64, y=32
				0x18, 0xFE                    // JR $-2
			};
			Initialize48KRedInkBluePaper(simulator, program, sizeof(program));
			Assert::AreEqual(S_OK, simulator->SetShowCRTSnapshot(TRUE));
			simulator->SimulateOne();
			simulator->SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, ReadMemory(simulator, 0x4088));

			const UINT64 targetScanline = FirstScreenLine48K + 32;
			RunUntilTime(simulator, (targetScanline - 8) * TStatesPerLine48K);
			auto beforeBeam = CaptureScreen(simulator, FALSE);
			Assert::AreNotEqual(0xFFC00000ul, beforeBeam.GetPixel(64, 32));

			RunUntilTime(simulator, (targetScanline + 8) * TStatesPerLine48K);
			auto afterBeam = CaptureScreen(simulator, FALSE);
			Assert::AreEqual(0xFFC00000ul, afterBeam.GetPixel(64, 32));
		}

		TEST_METHOD(CrtSnapshotShowsHorizontalBeamProgress)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);

			static const uint8_t program[] = {
				0x3E, 0x80,                   // LD A,80h
				0x32, 0x82, 0x40,             // LD (4082h),A: set pixel x=16, y=32
				0x32, 0x99, 0x40,             // LD (4099h),A: set pixel x=200, y=32
				0x18, 0xFE                    // JR $-2
			};
			Initialize48KRedInkBluePaper(simulator, program, sizeof(program));
			Assert::AreEqual(S_OK, simulator->SetShowCRTSnapshot(TRUE));
			simulator->SimulateOne();
			simulator->SimulateOne();
			simulator->SimulateOne();
			Assert::AreEqual<uint8_t>(0x80, ReadMemory(simulator, 0x4082));
			Assert::AreEqual<uint8_t>(0x80, ReadMemory(simulator, 0x4099));

			const UINT64 scanlineStart = (FirstScreenLine48K + 32) * TStatesPerLine48K;
			RunUntilTime(simulator, scanlineStart + 128);
			auto betweenPixels = CaptureScreen(simulator, FALSE);
			Assert::AreEqual(0xFFC00000ul, betweenPixels.GetPixel(16, 32));
			Assert::AreNotEqual(0xFFC00000ul, betweenPixels.GetPixel(200, 32));

			RunUntilTime(simulator, scanlineStart + 192);
			auto afterRightPixel = CaptureScreen(simulator, FALSE);
			Assert::AreEqual(0xFFC00000ul, afterRightPixel.GetPixel(16, 32));
			Assert::AreEqual(0xFFC00000ul, afterRightPixel.GetPixel(200, 32));
		}

		TEST_METHOD(PositionAdvancesAcrossExactFrameBoundaries)
		{
			struct TestIrqLine : IrqLine { };

			constexpr UINT64 frameTicks = 224 * 312;
			Bus memory;
			Bus io;
			TestIrqLine irq;
			wistd::unique_ptr<IScreenDevice> screen;
			Assert::AreEqual(S_OK, MakeScreenDevice(&memory, &io, &irq, SpectrumVariant48K, nullptr, &screen));
			Assert::AreEqual(S_OK, screen->Reset(SpectrumVariant48K));

			DWORD row;
			DWORD col;
			DWORD frameNumber;
			Assert::AreEqual(S_OK, screen->GetPosition(&row, &col, &frameNumber));
			Assert::AreEqual<DWORD>(0, row);
			Assert::AreEqual<DWORD>(0, col);
			Assert::AreEqual<DWORD>(0, frameNumber);

			for (DWORD frame = 1; frame <= 16; ++frame)
			{
				Assert::IsTrue(screen->SimulateDeviceTo(frame * frameTicks));
				Assert::AreEqual(S_OK, screen->GetPosition(&row, &col, &frameNumber));
				Assert::AreEqual<DWORD>(0, row);
				Assert::AreEqual<DWORD>(0, col);
				Assert::AreEqual<DWORD>(frame, frameNumber);
			}
		}
	};
}
