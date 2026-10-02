
#include <CppUnitTest.h>
#include "Utilities.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

uint8_t ReadMemory(ISimulator* simulator, uint16_t address)
{
	uint8_t value = 0;
	HRESULT hr = simulator->ReadMemoryBus8(address, &value);
	Assert::AreEqual(S_OK, hr);
	return value;
}

void RunUntilTime(ISimulator* simulator, UINT64 targetTime, DWORD timeoutMilliseconds)
{
	DWORD start = GetTickCount();
	for (;;)
	{
		UINT64 time = 0;
		HRESULT hr = simulator->GetTime(&time);
		Assert::AreEqual(S_OK, hr);
		if (time >= targetTime)
			return;
		if (GetTickCount() - start >= timeoutMilliseconds)
			Assert::Fail(L"Simulator did not reach the requested t-state.");
		simulator->SimulateOne();
	}
}

ScreenImage CaptureScreen(ISimulator* simulator, BOOL includeBorder)
{
	unique_safearray image;
	HRESULT hr = simulator->SaveScreen(includeBorder, &image);
	Assert::AreEqual(S_OK, hr);

	void* data = nullptr;
	hr = SafeArrayAccessData(image.get(), &data);
	Assert::AreEqual(S_OK, hr);
	auto unaccess = wil::scope_exit([&image] { SafeArrayUnaccessData(image.get()); });

	const auto* fileHeader = static_cast<const BITMAPFILEHEADER*>(data);
	Assert::AreEqual<WORD>(0x4D42, fileHeader->bfType);
	const auto* header = reinterpret_cast<const BITMAPINFOHEADER*>(static_cast<const uint8_t*>(data) + sizeof(BITMAPFILEHEADER));
	Assert::AreEqual<LONG>(includeBorder ? 352 : 256, header->biWidth);
	LONG imageHeight = header->biHeight < 0 ? -header->biHeight : header->biHeight;
	Assert::AreEqual<LONG>(includeBorder ? 296 : 192, imageHeight);
	Assert::AreEqual<WORD>(32, header->biBitCount);

	LONG lowerBound, upperBound;
	hr = SafeArrayGetLBound(image.get(), 1, &lowerBound);
	Assert::AreEqual(S_OK, hr);
	hr = SafeArrayGetUBound(image.get(), 1, &upperBound);
	Assert::AreEqual(S_OK, hr);
	size_t rowSize = static_cast<size_t>(header->biWidth) * sizeof(uint32_t);
	Assert::IsTrue(fileHeader->bfOffBits + rowSize * imageHeight <= static_cast<size_t>(upperBound - lowerBound + 1));

	ScreenImage result;
	result.width = static_cast<DWORD>(header->biWidth);
	result.height = static_cast<DWORD>(imageHeight);
	result.pixels.reset(static_cast<uint32_t*>(HeapAlloc(GetProcessHeap(), 0, rowSize * result.height)));
	Assert::IsNotNull(result.pixels.get());

	const auto* sourcePixels = static_cast<const uint8_t*>(data) + fileHeader->bfOffBits;
	for (DWORD y = 0; y < result.height; ++y)
	{
		DWORD sourceY = header->biHeight > 0 ? result.height - 1 - y : y;
		memcpy(result.pixels.get() + static_cast<size_t>(y) * result.width, sourcePixels + static_cast<size_t>(sourceY) * rowSize, rowSize);
	}
	return result;
}

unsigned long ScreenImage::GetPixel(DWORD x, DWORD y) const
{
	Assert::IsNotNull(pixels.get());
	Assert::IsTrue(x < width);
	Assert::IsTrue(y < height);
	return pixels[static_cast<size_t>(y) * width + x];
}
