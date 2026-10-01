
#include <CppUnitTest.h>
#include "Simulator.h"
#include "shared/com.h"
#include "shared/vector_nothrow.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace Z80SimulatorTests
{
	class TempFile
	{
		wchar_t _filename[MAX_PATH];

	public:
		TempFile(LPCWSTR fileExtension)
		{
			wchar_t tempDirectory[MAX_PATH];
			DWORD tempDirectoryLength = GetTempPathW(_countof(tempDirectory), tempDirectory);
			Assert::IsTrue(tempDirectoryLength != 0 && tempDirectoryLength < _countof(tempDirectory));
			Assert::IsTrue(GetTempFileNameW(tempDirectory, L"z80", 0, _filename) != 0);
			Assert::IsTrue(DeleteFileW(_filename));
			auto extension = wcsrchr(_filename, L'.');
			Assert::IsNotNull(extension);
			wcscpy_s(extension, _countof(_filename) - (extension - _filename), fileExtension);
		}

		~TempFile()
		{
			DeleteFileW(_filename);
		}

		void Write(const vector_nothrow<uint8_t>& contents)
		{
			wil::unique_hfile file(CreateFileW(_filename, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
			Assert::IsTrue(file.is_valid());

			DWORD bytesWritten = 0;
			Assert::IsTrue(WriteFile(file.get(), contents.data(), static_cast<DWORD>(contents.size()), &bytesWritten, nullptr) != FALSE);
			Assert::AreEqual<DWORD>(static_cast<DWORD>(contents.size()), bytesWritten);
		}

		LPCWSTR Path() const { return _filename; }
	};

	static void AppendWord(vector_nothrow<uint8_t>& bytes, uint16_t value)
	{
		bytes.try_push_back(static_cast<uint8_t>(value));
		bytes.try_push_back(static_cast<uint8_t>(value >> 8));
	}

	static void AppendUncompressedPage(vector_nothrow<uint8_t>& bytes, uint8_t page, uint8_t value)
	{
		AppendWord(bytes, 0xFFFF); // Z80 marker for an uncompressed 16 KB page
		bytes.try_push_back(page);
		for (uint32_t i = 0; i < 0x4000; i++) // One 16 KB page
			bytes.try_push_back(value);
	}

	static uint8_t ReadMemory(ISimulator* simulator, uint16_t address)
	{
		uint8_t value = 0;
		HRESULT hr = simulator->ReadMemoryBus8(address, &value);
		Assert::AreEqual(S_OK, hr);
		return value;
	}

	class TestProgram
	{
	public:
		vector_nothrow<uint8_t> bytes;
		size_t instructionCount = 0;

		void Emit(std::initializer_list<uint8_t> instruction)
		{
			Assert::IsTrue(bytes.try_insert(bytes.end(), instruction));
			++instructionCount;
		}

		void LdBC(uint16_t value)
		{
			Emit({ 0x01, static_cast<uint8_t>(value), static_cast<uint8_t>(value >> 8) });
		}

		void LdA(uint8_t value) { Emit({ 0x3E, value }); }
		void OutC() { Emit({ 0xED, 0x79 }); }
		void StoreA(uint16_t address) { Emit({ 0x32, static_cast<uint8_t>(address), static_cast<uint8_t>(address >> 8) }); }
		void LoadA(uint16_t address) { Emit({ 0x3A, static_cast<uint8_t>(address), static_cast<uint8_t>(address >> 8) }); }
	};

	static void RunProgram(ISimulator* simulator, uint16_t address, const TestProgram& program, SpectrumVariant variant)
	{
		Assert::IsTrue(program.bytes.size() <= UINT16_MAX);
		HRESULT hr = simulator->Reset(address, variant);
		Assert::AreEqual(S_OK, hr);
		hr = simulator->WriteMemoryBus(address, static_cast<uint16_t>(program.bytes.size()), program.bytes.data());
		Assert::AreEqual(S_OK, hr);
		for (size_t i = 0; i < program.instructionCount; i++)
		{
			hr = simulator->SimulateOne();
			Assert::AreEqual(S_OK, hr);
		}
	}

	static void Create128KSimulator(TempFile& romFile, com_ptr<ISimulator>& simulator)
	{
		vector_nothrow<uint8_t> rom;
		Assert::IsTrue(rom.try_resize(0x8000));
		memset(rom.data(), 0x31, 0x4000);
		memset(rom.data() + 0x4000, 0x42, 0x4000);
		romFile.Write(rom);

		LPCWSTR romFilenames[SpectrumVariantCount] = {};
		romFilenames[SpectrumVariant128] = romFile.Path();
		HRESULT hr = MakeSimulator(romFilenames, _countof(romFilenames), SpectrumVariant128, &simulator);
		Assert::AreEqual(S_OK, hr);
	}

	static void AppendZ80V23Header(vector_nothrow<uint8_t>& snapshot, uint16_t headerLength, uint8_t hardwareMode)
	{
		Assert::IsTrue(snapshot.try_resize(30));
     memset(snapshot.data(), 0, 30);
		AppendWord(snapshot, headerLength);
		vector_nothrow<uint8_t> extendedHeader;
		Assert::IsTrue(extendedHeader.try_resize(headerLength));
       memset(extendedHeader.data(), 0, headerLength);
		extendedHeader[0] = 0x34;
		extendedHeader[1] = 0x12;
		extendedHeader[2] = hardwareMode;
		for (auto byte : extendedHeader)
			Assert::IsTrue(snapshot.try_push_back(byte));
	}

	static void AppendCompressedPage(vector_nothrow<uint8_t>& snapshot, uint8_t page, uint32_t decompressedSize, uint8_t value)
	{
		vector_nothrow<uint8_t> compressed;
		while (decompressedSize)
		{
            uint8_t count = static_cast<uint8_t>(decompressedSize > 255 ? 255 : decompressedSize);
			Assert::IsTrue(compressed.try_insert(compressed.end(), { 0xED, 0xED, count, value }));
			decompressedSize -= count;
		}
		Assert::IsTrue(compressed.size() < 0xFFFF);
		AppendWord(snapshot, static_cast<uint16_t>(compressed.size()));
		Assert::IsTrue(snapshot.try_push_back(page));
		for (auto byte : compressed)
			Assert::IsTrue(snapshot.try_push_back(byte));
	}

	static uint32_t ReadBitmapPixel(SAFEARRAY* image, LONG expectedWidth, LONG expectedHeight, LONG x, LONG y)
	{
		void* data;
		HRESULT hr = SafeArrayAccessData(image, &data);
		Assert::AreEqual(S_OK, hr);
		auto unaccess = wil::scope_exit([image] { SafeArrayUnaccessData(image); });

		BITMAPFILEHEADER fileHeader;
		memcpy(&fileHeader, data, sizeof(fileHeader));
		Assert::AreEqual<WORD>(0x4D42, fileHeader.bfType);
		BITMAPINFOHEADER header;
		memcpy(&header, static_cast<const uint8_t*>(data) + sizeof(fileHeader), sizeof(header));
		Assert::AreEqual<LONG>(expectedWidth, header.biWidth);
		Assert::AreEqual<LONG>(expectedHeight, header.biHeight < 0 ? -header.biHeight : header.biHeight);
		LONG bitmapY = header.biHeight > 0 ? header.biHeight - 1 - y : y;
		uint32_t pixel;
		memcpy(&pixel, static_cast<const uint8_t*>(data) + fileHeader.bfOffBits + (bitmapY * header.biWidth + x) * sizeof(pixel), sizeof(pixel));
		return pixel;
	}

	TEST_CLASS(SimulatorAPITests)
	{
	public:
		TEST_METHOD(SaveScreen)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);

			hr = simulator->Reset(32768, SpectrumVariant48K); // Load the test program at the start of RAM
			Assert::AreEqual(S_OK, hr);
			hr = simulator->SetSpeed(UINT32_MAX);
			Assert::AreEqual(S_OK, hr);

			UINT64 time;
			hr = simulator->GetTime(&time);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<UINT64>(0, time);

			// Make border blue and pixels red.
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
			hr = simulator->WriteMemoryBus(32768, static_cast<uint16_t>(sizeof(program)), program);
			Assert::AreEqual(S_OK, hr);

			hr = simulator->Resume(FALSE);
			Assert::AreEqual(S_OK, hr);
			auto stopSimulation = wil::scope_exit([&simulator] { simulator->Break(); });

			// Wait for the simulator to render the entire screen. At 3.5MHz and 50fps that's about 70K T-Cycles.
			DWORD start = GetTickCount();
			for (;;)
			{
				hr = simulator->GetTime(&time);
				Assert::AreEqual(S_OK, hr);
				if (time >= 70000)
					break;
				if (GetTickCount() - start >= 5000)
					Assert::Fail(L"Simulator did not reach 70000 t-states.");
				Sleep(1);
			}

			hr = simulator->Break();
			Assert::IsTrue(SUCCEEDED(hr));
			stopSimulation.release();

			unique_safearray image;
			hr = simulator->SaveScreen(TRUE, &image);
			Assert::AreEqual(S_OK, hr);

			LONG lowerBound;
			LONG upperBound;
			hr = SafeArrayGetLBound(image.get(), 1, &lowerBound);
			Assert::AreEqual(S_OK, hr);
			hr = SafeArrayGetUBound(image.get(), 1, &upperBound);
			Assert::AreEqual(S_OK, hr);

			void* data;
			hr = SafeArrayAccessData(image.get(), &data);
			Assert::AreEqual(S_OK, hr);
			auto unaccess = wil::scope_exit([&image] { SafeArrayUnaccessData(image.get()); });

			const auto* bmp = static_cast<const uint8_t*>(data);
			BITMAPFILEHEADER fileHeader;
			memcpy(&fileHeader, bmp, sizeof(fileHeader));
			Assert::AreEqual<WORD>(0x4D42, fileHeader.bfType);
			BITMAPINFOHEADER header;
			memcpy(&header, bmp + sizeof(fileHeader), sizeof(header));
			Assert::AreEqual<LONG>(352, header.biWidth); // Screen width including border
			LONG imageHeight = header.biHeight > 0 ? header.biHeight : -header.biHeight;
			Assert::AreEqual<LONG>(296, imageHeight); // Screen height including border
			Assert::AreEqual<WORD>(32, header.biBitCount);
			Assert::IsTrue(fileHeader.bfOffBits + header.biSizeImage <= static_cast<DWORD>(upperBound - lowerBound + 1));

			bool borderIsBlue = true;
			bool screenIsRed = true;
			for (LONG y = 0; y < imageHeight; y++)
			{
				LONG bitmapY = header.biHeight > 0 ? imageHeight - 1 - y : y;
				for (LONG x = 0; x < header.biWidth; x++)
				{
					uint32_t pixel;
					memcpy(&pixel, bmp + fileHeader.bfOffBits + (bitmapY * header.biWidth + x) * sizeof(pixel), sizeof(pixel));
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

		TEST_METHOD(LoadZ80V1Uncompressed)
		{
			vector_nothrow<uint8_t> snapshot;
			snapshot.try_resize(30 + 0xC000); // 30-byte v1 header followed by 48 KB of RAM
			snapshot[6] = 0x23; // PC low byte
			snapshot[7] = 0x81; // PC high byte
			snapshot[30 + 0x0000] = 0x14; // First byte at 0x4000
			snapshot[30 + 0x4000] = 0x25; // First byte at 0x8000
			snapshot[30 + 0xBFFF] = 0x36; // Last byte at 0xFFFF

			TempFile file (L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::AreEqual(S_OK, hr);

			UINT16 pc;
			SpectrumVariant variant;
			Assert::AreEqual(S_OK, simulator->GetPC(&pc));
			Assert::AreEqual<UINT16>(0x8123, pc);
			Assert::AreEqual(S_OK, simulator->GetVariant(&variant));
			Assert::IsTrue(variant == SpectrumVariant48K);
			Assert::AreEqual<uint8_t>(0x14, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0x25, ReadMemory(simulator.get(), 0x8000));
			Assert::AreEqual<uint8_t>(0x36, ReadMemory(simulator.get(), 0xFFFF));
		}

		TEST_METHOD(LoadZ80V1Compressed)
		{
			vector_nothrow<uint8_t> snapshot;
			snapshot.try_resize(30); // v1 header
			snapshot[6] = 0x34; // PC low byte
			snapshot[7] = 0x92; // PC high byte
			snapshot[12] = 0x20; // Compression flag
			for (size_t remaining = 0xC000; remaining != 0;)
			{
				uint8_t count = static_cast<uint8_t>(remaining > 200 ? 200 : remaining); // RLE count, excluding zero
				snapshot.try_insert(snapshot.end(), { 0xED, 0xED, count, 0x42 });
				remaining -= count;
			}
			snapshot.try_insert(snapshot.end(), { 0x00, 0xED, 0xED, 0x00 }); // End-of-data marker

			TempFile file (L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::AreEqual(S_OK, hr);

			UINT16 pc;
			Assert::AreEqual(S_OK, simulator->GetPC(&pc));
			Assert::AreEqual<UINT16>(0x9234, pc);
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator.get(), 0x8000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator.get(), 0xFFFF));
		}

		TEST_METHOD(LoadZ80V2_48K)
		{
			vector_nothrow<uint8_t> snapshot;
			snapshot.try_resize(30); // v1 header before extended header
			AppendWord(snapshot, 23); // v2 extended-header length
			vector_nothrow<uint8_t> extendedHeader;
			extendedHeader.try_resize(23);
			extendedHeader[0] = 0x45; // PC low byte
			extendedHeader[1] = 0xA2; // PC high byte
			extendedHeader[2] = 0; // 48K machine type
			for (auto byte : extendedHeader)
				snapshot.try_push_back(byte);
			AppendUncompressedPage(snapshot, 8, 0x18); // 48K page at 0x4000
			AppendUncompressedPage(snapshot, 4, 0x24); // 48K page at 0x8000
			AppendUncompressedPage(snapshot, 5, 0x35); // 48K page at 0xC000

			TempFile file (L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::AreEqual(S_OK, hr);

			UINT16 pc;
			SpectrumVariant variant;
			Assert::AreEqual(S_OK, simulator->GetPC(&pc));
			Assert::AreEqual<UINT16>(0xA245, pc);
			Assert::AreEqual(S_OK, simulator->GetVariant(&variant));
			Assert::IsTrue(variant == SpectrumVariant48K);
			Assert::AreEqual<uint8_t>(0x18, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0x24, ReadMemory(simulator.get(), 0x8000));
			Assert::AreEqual<uint8_t>(0x35, ReadMemory(simulator.get(), 0xC000));
		}

		TEST_METHOD(LoadZ80V3_128K)
		{
			vector_nothrow<uint8_t> snapshot;
			snapshot.try_resize(30); // v1 header before extended header
			AppendWord(snapshot, 54); // v3 extended-header length
			vector_nothrow<uint8_t> extendedHeader;
			extendedHeader.try_resize(54);
			extendedHeader[0] = 0x67; // PC low byte
			extendedHeader[1] = 0xB3; // PC high byte
			extendedHeader[2] = 4; // 128K machine type
			extendedHeader[3] = 3; // 128K paging configuration
			for (auto byte : extendedHeader)
				snapshot.try_push_back(byte);
			for (uint8_t page = 3; page <= 10; page++) // Eight RAM pages in a 128K snapshot
				AppendUncompressedPage(snapshot, page, static_cast<uint8_t>(0x20 + page));

			TempFile file (L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::AreEqual(S_OK, hr);

			UINT16 pc;
			SpectrumVariant variant;
			Assert::AreEqual(S_OK, simulator->GetPC(&pc));
			Assert::AreEqual<UINT16>(0xB367, pc);
			Assert::AreEqual(S_OK, simulator->GetVariant(&variant));
			Assert::IsTrue(variant == SpectrumVariant128);
			Assert::AreEqual<uint8_t>(0x28, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0x25, ReadMemory(simulator.get(), 0x8000));
			Assert::AreEqual<uint8_t>(0x26, ReadMemory(simulator.get(), 0xC000));
		}

		TEST_METHOD(PagingSelectsAllEightRamBanksAndFixedBankAliases)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			for (uint8_t bank = 0; bank < 8; bank++)
			{
				program.LdBC(0x7FFD);
				program.LdA(bank);
				program.OutC();
				program.LdA(static_cast<uint8_t>(0xA0 + bank));
				program.StoreA(0xC000);
			}
			for (uint8_t bank = 0; bank < 8; bank++)
			{
				program.LdBC(0x7FFD);
				program.LdA(bank);
				program.OutC();
				program.LoadA(0xC000);
				program.StoreA(static_cast<uint16_t>(0x9000 + bank));
			}
			RunProgram(simulator.get(), 0xA000, program, SpectrumVariant128);

			for (uint8_t bank = 0; bank < 8; bank++)
				Assert::AreEqual<uint8_t>(static_cast<uint8_t>(0xA0 + bank), ReadMemory(simulator.get(), static_cast<uint16_t>(0x9000 + bank)));
			Assert::AreEqual<uint8_t>(0xA5, ReadMemory(simulator.get(), 0x4000)); // Fixed slot aliases bank 5
			Assert::AreEqual<uint8_t>(0xA2, ReadMemory(simulator.get(), 0x8000)); // Fixed slot aliases bank 2
		}

		TEST_METHOD(PagingPortPartialDecodeAndRomSelection)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			program.LdBC(0x7FFD);
			program.LdA(1);
			program.OutC();
			program.LdA(0x51);
			program.StoreA(0xC000);

			program.LdBC(0x1234); // Bits 15 and 1 are clear; other address bits vary.
			program.LdA(6);
			program.OutC();
			program.LdA(0x66);
			program.StoreA(0xC000);

			program.LdBC(0x9234); // A15 set: this address must not page memory.
			program.LdA(1);
			program.OutC();
			program.LdA(0xEE);
			program.StoreA(0xC000);
			program.LdBC(0x1236); // A1 set: this address must not page memory.
			program.LdA(1);
			program.OutC();
			program.LdA(0xEF);
			program.StoreA(0xC000);

			program.LdBC(0x7FFD);
			program.LdA(1);
			program.OutC();
			program.LoadA(0xC000);
			program.StoreA(0x9000);

			program.LdBC(0x1234);
			program.LdA(0x16); // Keep bank 6 selected while choosing ROM 1.
			program.OutC();
			RunProgram(simulator.get(), 0xA000, program, SpectrumVariant128);

			Assert::AreEqual<uint8_t>(0xEF, ReadMemory(simulator.get(), 0xC000));
			Assert::AreEqual<uint8_t>(0x51, ReadMemory(simulator.get(), 0x9000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator.get(), 0x0000));
		}

		TEST_METHOD(PagingScreenBankIsIndependentOfCpuRamBank)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			program.LdBC(0x7FFD);
			program.LdA(7);
			program.OutC();
			program.LdA(0x80);
			program.StoreA(0xC000);
			program.LdA(0x02);
			program.StoreA(0xD800);

			program.LdA(3);
			program.OutC();
			program.LdA(0x33);
			program.StoreA(0xC000);
			program.LdA(0x0B); // CPU bank 3, display bank 7.
			program.OutC();
			RunProgram(simulator.get(), 0xA000, program, SpectrumVariant128);

			Assert::AreEqual<uint8_t>(0x33, ReadMemory(simulator.get(), 0xC000));
			unique_safearray image;
			HRESULT hr = simulator->SaveScreen(FALSE, &image);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<uint32_t>(0xFFC00000, ReadBitmapPixel(image.get(), 256, 192, 0, 0));
		}

		TEST_METHOD(PagingLockPreventsFurtherRamRomAndScreenChanges)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			program.LdBC(0x7FFD);
			program.LdA(5);
			program.OutC();
			program.LdA(0);
			program.StoreA(0xC000);
			program.StoreA(0xD800);
			program.LdA(0x55);
			program.StoreA(0xC000);

			program.LdA(7);
			program.OutC();
			program.LdA(0x80);
			program.StoreA(0xC000);
			program.LdA(0x02);
			program.StoreA(0xD800);

			program.LdA(0x3D); // Select bank 5, display bank 7 and ROM 1, then lock paging.
			program.OutC();
			program.LdA(0xA5);
			program.StoreA(0xC000);
			program.LdA(0x02); // Attempt to change RAM, ROM and display selection after locking.
			program.OutC();
			program.LdA(0xEE);
			program.StoreA(0xC000);
			RunProgram(simulator.get(), 0xA000, program, SpectrumVariant128);

			Assert::AreEqual<uint8_t>(0xEE, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator.get(), 0x0000));
			unique_safearray image;
			HRESULT hr = simulator->SaveScreen(FALSE, &image);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<uint32_t>(0xFFC00000, ReadBitmapPixel(image.get(), 256, 192, 0, 0));
		}

		TEST_METHOD(SaveScreenWithoutBorderReturnsDestroyableUnlockedArray)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			hr = simulator->Reset(0x8000, SpectrumVariant48K);
			Assert::AreEqual(S_OK, hr);

			uint8_t pixelByte = 0x80;
			uint8_t attribute = 0x02;
			Assert::AreEqual(S_OK, simulator->WriteMemoryBus(0x4000, 1, &pixelByte));
			Assert::AreEqual(S_OK, simulator->WriteMemoryBus(0x5800, 1, &attribute));

			unique_safearray image;
			hr = simulator->SaveScreen(FALSE, &image);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<uint32_t>(0xFFC00000, ReadBitmapPixel(image.get(), 256, 192, 0, 0));
			Assert::AreEqual<USHORT>(0, image.get()->cLocks);
			SAFEARRAY* ownedByTest = image.release();
			Assert::AreEqual(S_OK, SafeArrayDestroy(ownedByTest));
		}

		TEST_METHOD(LoadCompressedZ80V2_48K)
		{
			vector_nothrow<uint8_t> snapshot;
			AppendZ80V23Header(snapshot, 23, 0);
			AppendCompressedPage(snapshot, 8, 0x4000, 0x18);
			AppendCompressedPage(snapshot, 4, 0x4000, 0x24);
			AppendCompressedPage(snapshot, 5, 0x4000, 0x35);

			TempFile file(L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			Assert::AreEqual(S_OK, simulator->LoadFile(file.Path()));
			Assert::AreEqual<uint8_t>(0x18, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0x24, ReadMemory(simulator.get(), 0x8000));
			Assert::AreEqual<uint8_t>(0x35, ReadMemory(simulator.get(), 0xC000));
		}

		TEST_METHOD(LoadCompressedZ80V3_128K)
		{
			vector_nothrow<uint8_t> snapshot;
			AppendZ80V23Header(snapshot, 54, 4);
			for (uint8_t page = 3; page <= 10; page++)
				AppendCompressedPage(snapshot, page, 0x4000, static_cast<uint8_t>(0x30 + page - 3));

			TempFile file(L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			Assert::AreEqual(S_OK, simulator->LoadFile(file.Path()));
			Assert::AreEqual<uint8_t>(0x35, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0x32, ReadMemory(simulator.get(), 0x8000));
			Assert::AreEqual<uint8_t>(0x30, ReadMemory(simulator.get(), 0xC000));
		}

		TEST_METHOD(LoadZ80RejectsCompressedPageWithShortOutput)
		{
			vector_nothrow<uint8_t> snapshot;
			AppendZ80V23Header(snapshot, 23, 0);
			AppendCompressedPage(snapshot, 8, 0x3FFF, 0x18);
			AppendCompressedPage(snapshot, 4, 0x4000, 0x24);
			AppendCompressedPage(snapshot, 5, 0x4000, 0x35);

			TempFile file(L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::IsTrue(FAILED(hr));
		}

		TEST_METHOD(LoadZ80RejectsCompressedPageWithExcessOutput)
		{
			vector_nothrow<uint8_t> snapshot;
			AppendZ80V23Header(snapshot, 23, 0);
			AppendCompressedPage(snapshot, 8, 0x4001, 0x18);
			AppendCompressedPage(snapshot, 4, 0x4000, 0x24);
			AppendCompressedPage(snapshot, 5, 0x4000, 0x35);

			TempFile file(L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::IsTrue(FAILED(hr));
		}

		TEST_METHOD(LoadZ80RejectsTruncatedCompressedPage)
		{
			vector_nothrow<uint8_t> snapshot;
			AppendZ80V23Header(snapshot, 23, 0);
			AppendWord(snapshot, 4);
			Assert::IsTrue(snapshot.try_push_back(8));
			Assert::IsTrue(snapshot.try_insert(snapshot.end(), { 0xED, 0xED }));

			TempFile file(L".z80");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::IsTrue(FAILED(hr));
		}

		TEST_METHOD(LoadSNAAfterSelecting128KResetsTo48K)
		{
			vector_nothrow<uint8_t> snapshot;
			snapshot.try_resize(27 + 48 * 1024);

			TempFile file(L".sna");
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);

			hr = simulator->Reset(0, SpectrumVariant128);
			Assert::AreEqual(S_OK, hr);
			file.Write(snapshot);
			hr = simulator->LoadFile(file.Path());
			Assert::AreEqual(S_OK, hr);

			SpectrumVariant variant;
			hr = simulator->GetVariant(&variant);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual((int)SpectrumVariant48K, (int)variant);
			Assert::AreEqual<uint8_t>(0, ReadMemory(simulator.get(), 0x4000));
			Assert::AreEqual<uint8_t>(0, ReadMemory(simulator.get(), 0x8000));
			Assert::AreEqual<uint8_t>(0, ReadMemory(simulator.get(), 0xFFFF));
		}

	};
}
