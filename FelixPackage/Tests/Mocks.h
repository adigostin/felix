
#pragma once
#include "shared/com.h"
#include "FelixPackage.h"

extern com_ptr<IProjectConfig> MakeMockProjectConfig (IVsHierarchy* hier);
extern com_ptr<IVsHierarchy> MakeMockVsHierarchy();
extern com_ptr<IProjectFile> MakeMockSourceFile (IVsHierarchy* hier, VSITEMID itemId,
	BuildToolKind buildTool, LPCWSTR pathRelativeToProjectDir, std::string_view fileContent);
extern com_ptr<IVsOutputWindowPane2> MakeMockOutputWindowPane (IStream* outputStreamUTF16);
extern com_ptr<IServiceProvider> MakeMockServiceProvider();
