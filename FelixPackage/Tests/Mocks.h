
#pragma once
#include "shared/com.h"
#include "FelixPackage.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

extern wchar_t tempPath[MAX_PATH + 1];

extern com_ptr<IProjectConfig> MakeMockProjectConfig (IVsHierarchy* hier);
extern com_ptr<IVsHierarchy> MakeMockVsHierarchy (const wchar_t* projectDir);
extern com_ptr<IProjectFile> MakeMockSourceFile (IVsHierarchy* hier, VSITEMID itemId, VSITEMID parentItemId,
	BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir);
extern com_ptr<IVsOutputWindowPane2> MakeMockOutputWindowPane (IStream* outputStreamUTF16);
extern com_ptr<IServiceProvider> MakeMockServiceProvider();

extern void WriteFileOnDisk(const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir, const char* fileContent);
extern void DeleteFileOnDisk(const wchar_t* projectDir, const wchar_t* pathRelativeToProjectDir);
