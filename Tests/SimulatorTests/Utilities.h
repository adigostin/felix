
#pragma once
#include <CppUnitTest.h>
#include "Simulator.h"
#include "shared/com.h"

namespace SimulatorTests
{
	using namespace Microsoft::VisualStudio::CppUnitTestFramework;

	class TestSimulatorEventNotifySink : public ISimulatorEventNotifySink
	{
		ULONG _refCount = 1;
		AdviseSinkToken _token;

	public:
		DWORD breakpoints = 0;
		DWORD breaks = 0;
		DWORD simulateOnes = 0;
		DWORD resumes = 0;

		TestSimulatorEventNotifySink(ISimulator* simulator)
		{
			auto hr = AdviseSink<ISimulatorEventNotifySink>(simulator, this, &_token);
			Assert::AreEqual(S_OK, hr);
		}

		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<ISimulatorEventNotifySink>(this, riid, ppvObject))
				return S_OK;

			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }
		virtual ULONG STDMETHODCALLTYPE Release() override { return --_refCount; }

		virtual HRESULT STDMETHODCALLTYPE NotifySimulatorEvent(ISimulatorEvent*, REFIID riidEvent) override
		{
			if (riidEvent == __uuidof(ISimulatorBreakpointEvent))
				breakpoints++;
			else if (riidEvent == __uuidof(ISimulatorBreakEvent))
				breaks++;
			else if (riidEvent == __uuidof(ISimulatorSimulateOneEvent))
				simulateOnes++;
			else if (riidEvent == __uuidof(ISimulatorResumeEvent))
				resumes++;
			else
				Assert::Fail();

			return S_OK;
		}
	};

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

		void Write(std::span<const uint8_t> contents)
		{
			wil::unique_hfile file(CreateFileW(_filename, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
			Assert::IsTrue(file.is_valid());

			DWORD bytesWritten = 0;
			Assert::IsTrue(WriteFile(file.get(), contents.data(), contents.size(), &bytesWritten, nullptr) != FALSE);
			Assert::AreEqual<DWORD>(contents.size(), bytesWritten);
		}

		void Write (std::initializer_list<uint8_t> contents)
		{
			Write(std::span<const uint8_t>(contents));
		}

		LPCWSTR Path() const { return _filename; }
	};
}

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

wil::com_ptr_failfast<ISimulator> MakeSimulator (
	LPCWSTR const* romFilenames = nullptr,
	size_t romFilenameCount = 0,
	SpectrumVariant initialVariant = SpectrumVariant48K
);