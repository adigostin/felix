
#include <CppUnitTest.h>
#include "Simulator.h"
#include "Utilities.h"
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
			Assert::IsTrue(WriteFile(file.get(), contents.data(), contents.size(), &bytesWritten, nullptr) != FALSE);
			Assert::AreEqual<DWORD>(contents.size(), bytesWritten);
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
	};

	class TestScreenCompleteHandler : public IScreenCompleteEventHandler
	{
		ULONG _refCount = 0;

	public:
		UINT32 callCount = 0;

		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<IScreenCompleteEventHandler>(this, riid, ppvObject))
				return S_OK;

			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }

		virtual HRESULT STDMETHODCALLTYPE OnScreenComplete(BITMAPINFO* bi, POINT, UINT64, LONGLONG) override
		{
			CoTaskMemFree(bi);
			++callCount;
			return S_OK;
		}
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

	TEST_CLASS(SimulatorAPITests)
	{
	public:
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
			Assert::AreEqual<uint8_t>(0x14, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x25, ReadMemory(simulator, 0x8000));
			Assert::AreEqual<uint8_t>(0x36, ReadMemory(simulator, 0xFFFF));
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
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator, 0x8000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator, 0xFFFF));
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
			Assert::AreEqual<uint8_t>(0x18, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x24, ReadMemory(simulator, 0x8000));
			Assert::AreEqual<uint8_t>(0x35, ReadMemory(simulator, 0xC000));
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
			Assert::AreEqual<uint8_t>(0x28, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x25, ReadMemory(simulator, 0x8000));
			Assert::AreEqual<uint8_t>(0x26, ReadMemory(simulator, 0xC000));
		}

		TEST_METHOD(PagingSelectsAllEightRamBanksAndFixedBankAliases)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			static constexpr uint8_t bankValues[] = { 0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7 };
			for (uint8_t bank = 0; bank < 8; bank++)
			{
				program.Emit({ 0x01, 0xFD, 0x7F });        // LD BC,7FFDh
				program.Emit({ 0x3E, bank });              // LD A,bank
				program.Emit({ 0xED, 0x79 });              // OUT (C),A
				program.Emit({ 0x3E, bankValues[bank] });  // LD A,A0h+bank
				program.Emit({ 0x32, 0x00, 0xC0 });        // LD (C000h),A
			}
			for (uint8_t bank = 0; bank < 8; bank++)
			{
				program.Emit({ 0x01, 0xFD, 0x7F });  // LD BC,7FFDh
				program.Emit({ 0x3E, bank });        // LD A,bank
				program.Emit({ 0xED, 0x79 });        // OUT (C),A
				program.Emit({ 0x3A, 0x00, 0xC0 });  // LD A,(C000h)
				program.Emit({ 0x32, bank, 0x90 });  // LD (9000h+bank),A
			}
			RunProgram(simulator, 0xA000, program, SpectrumVariant128);

			for (uint8_t bank = 0; bank < 8; bank++)
				Assert::AreEqual<uint8_t>(static_cast<uint8_t>(0xA0 + bank), ReadMemory(simulator, static_cast<uint16_t>(0x9000 + bank)));
			Assert::AreEqual<uint8_t>(0xA5, ReadMemory(simulator, 0x4000)); // Fixed slot aliases bank 5
			Assert::AreEqual<uint8_t>(0xA2, ReadMemory(simulator, 0x8000)); // Fixed slot aliases bank 2
		}

		TEST_METHOD(PagingPortPartialDecodeAndRomSelection)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			program.Emit({ 0x01, 0xFD, 0x7F });  // LD BC,7FFDh
			program.Emit({ 0x3E, 0x01 });        // LD A,1
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0x51 });        // LD A,51h
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A

			program.Emit({ 0x01, 0x34, 0x12 });  // LD BC,1234h: bits 15 and 1 are clear; other address bits vary.
			program.Emit({ 0x3E, 0x06 });        // LD A,6
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0x66 });        // LD A,66h
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A

			program.Emit({ 0x01, 0x34, 0x92 });  // LD BC,9234h: A15 set, so this address must not page memory.
			program.Emit({ 0x3E, 0x01 });        // LD A,1
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0xEE });        // LD A,EEh
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A
			program.Emit({ 0x01, 0x36, 0x12 });  // LD BC,1236h: A1 set, so this address must not page memory.
			program.Emit({ 0x3E, 0x01 });        // LD A,1
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0xEF });        // LD A,EFh
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A

			program.Emit({ 0x01, 0xFD, 0x7F });  // LD BC,7FFDh
			program.Emit({ 0x3E, 0x01 });        // LD A,1
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3A, 0x00, 0xC0 });  // LD A,(C000h)
			program.Emit({ 0x32, 0x00, 0x90 });  // LD (9000h),A

			program.Emit({ 0x01, 0x34, 0x12 });  // LD BC,1234h
			program.Emit({ 0x3E, 0x16 });        // LD A,16h: keep bank 6 selected while choosing ROM 1.
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			RunProgram(simulator, 0xA000, program, SpectrumVariant128);

			Assert::AreEqual<uint8_t>(0xEF, ReadMemory(simulator, 0xC000));
			Assert::AreEqual<uint8_t>(0x51, ReadMemory(simulator, 0x9000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator, 0x0000));
		}

		TEST_METHOD(PagingScreenBankIsIndependentOfCpuRamBank)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			program.Emit({ 0x01, 0xFD, 0x7F });  // LD BC,7FFDh
			program.Emit({ 0x3E, 0x07 });        // LD A,7
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0x80 });        // LD A,80h
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A
			program.Emit({ 0x3E, 0x02 });        // LD A,2
			program.Emit({ 0x32, 0x00, 0xD8 });  // LD (D800h),A

			program.Emit({ 0x3E, 0x03 });        // LD A,3
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0x33 });        // LD A,33h
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A
			program.Emit({ 0x3E, 0x0B });        // LD A,0Bh: CPU bank 3, display bank 7.
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			RunProgram(simulator, 0xA000, program, SpectrumVariant128);

			Assert::AreEqual<uint8_t>(0x33, ReadMemory(simulator, 0xC000));
			auto image = CaptureScreen(simulator, FALSE);
			Assert::AreEqual<uint32_t>(0xFFC00000, image.GetPixel(0, 0));
		}

		TEST_METHOD(PagingLockPreventsFurtherRamRomAndScreenChanges)
		{
			TempFile romFile(L".rom");
			com_ptr<ISimulator> simulator;
			Create128KSimulator(romFile, simulator);

			TestProgram program;
			program.Emit({ 0x01, 0xFD, 0x7F });  // LD BC,7FFDh
			program.Emit({ 0x3E, 0x05 });        // LD A,5
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0x00 });        // LD A,0
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A
			program.Emit({ 0x32, 0x00, 0xD8 });  // LD (D800h),A
			program.Emit({ 0x3E, 0x55 });        // LD A,55h
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A

			program.Emit({ 0x3E, 0x07 });        // LD A,7
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0x80 });        // LD A,80h
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A
			program.Emit({ 0x3E, 0x02 });        // LD A,2
			program.Emit({ 0x32, 0x00, 0xD8 });  // LD (D800h),A

			program.Emit({ 0x3E, 0x3D });        // LD A,3Dh: select bank 5, display bank 7 and ROM 1, then lock paging.
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0xA5 });        // LD A,A5h
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A
			program.Emit({ 0x3E, 0x02 });        // LD A,2: attempt to change RAM, ROM and display selection after locking.
			program.Emit({ 0xED, 0x79 });        // OUT (C),A
			program.Emit({ 0x3E, 0xEE });        // LD A,EEh
			program.Emit({ 0x32, 0x00, 0xC0 });  // LD (C000h),A
			RunProgram(simulator, 0xA000, program, SpectrumVariant128);

			Assert::AreEqual<uint8_t>(0xEE, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x42, ReadMemory(simulator, 0x0000));
			auto image = CaptureScreen(simulator, FALSE);
			Assert::AreEqual<uint32_t>(0xFFC00000, image.GetPixel(0, 0));
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
			Assert::AreEqual<uint8_t>(0x18, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x24, ReadMemory(simulator, 0x8000));
			Assert::AreEqual<uint8_t>(0x35, ReadMemory(simulator, 0xC000));
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
			Assert::AreEqual<uint8_t>(0x35, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0x32, ReadMemory(simulator, 0x8000));
			Assert::AreEqual<uint8_t>(0x30, ReadMemory(simulator, 0xC000));
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
			Assert::AreEqual<uint8_t>(0, ReadMemory(simulator, 0x4000));
			Assert::AreEqual<uint8_t>(0, ReadMemory(simulator, 0x8000));
			Assert::AreEqual<uint8_t>(0, ReadMemory(simulator, 0xFFFF));
		}

		TEST_METHOD(ResetWhilePausedNotifiesScreenComplete)
		{
			HRESULT hr;
			wil::com_ptr_failfast<ISimulator> simulator;
			MakeSimulator (nullptr, 0, SpectrumVariant48K, &simulator);

			auto screenHandler = wil::com_ptr_failfast(new TestScreenCompleteHandler());
			hr = simulator->AdviseScreenComplete(screenHandler);
			Assert::AreEqual(S_OK, hr);
			auto unadvise = wil::scope_exit([&] { simulator->UnadviseScreenComplete(screenHandler); });

			static constexpr UINT16 startAddress = 0x1234;
			hr = simulator->Reset(startAddress, SpectrumVariant48K);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual<UINT32>(1, screenHandler->callCount);
			Assert::AreEqual(S_FALSE, simulator->Running_HR());

			UINT16 pc = 0;
			Assert::AreEqual(S_OK, simulator->GetPC(&pc));
			Assert::AreEqual(startAddress, pc);
			UINT64 time = UINT64_MAX;
			Assert::AreEqual(S_OK, simulator->GetTime(&time));
			Assert::AreEqual<UINT64>(0, time);
		}

		static void VerifyResetWhileRunning(ISimulator* simulator, uint32_t speed)
		{
			HRESULT hr = simulator->SetSpeed(speed);
			Assert::AreEqual(S_OK, hr);
			hr = simulator->Resume(FALSE);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_OK, simulator->Running_HR());

			hr = simulator->Reset(0x1234, SpectrumVariant128);
			Assert::AreEqual(S_OK, hr);

			hr = simulator->Break();
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual(S_FALSE, simulator->Running_HR());

			SpectrumVariant variant;
			Assert::AreEqual(S_OK, simulator->GetVariant(&variant));
			Assert::AreEqual((int)SpectrumVariant128, (int)variant);
		}

		TEST_METHOD(ResetWhileRunning)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			VerifyResetWhileRunning(simulator.get(), 100);
		}

		TEST_METHOD(ResetWhileRunningAtMaxSpeed)
		{
			com_ptr<ISimulator> simulator;
			HRESULT hr = MakeSimulator(nullptr, 0, SpectrumVariant48K, &simulator);
			Assert::AreEqual(S_OK, hr);
			VerifyResetWhileRunning(simulator.get(), UINT32_MAX);
		}

	};
}
