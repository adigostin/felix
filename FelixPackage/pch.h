
#pragma once

#define STRICT
#define _WIN32_WINNT    _WIN32_WINNT_WIN10
#define NTDDI_VERSION   NTDDI_WIN10_RS5
#define NOMINMAX // Windows Platform min and max macros cause problems for the Standard C++ Library
#define WIN32_LEAN_AND_MEAN	// Exclude rarely-used stuff from the Windows Platform headers
#define _COM_NO_STANDARD_GUIDS_
#define _SILENCE_CXX23_ALIGNED_STORAGE_DEPRECATION_WARNING
#define _HAS_EXCEPTIONS 0 // Visual C++ headers look at this
#undef __EXCEPTIONS // WIL and Intellisense look at this

// Visual C++ headers

// Windows Platform headers
#include <Windows.h>
#undef GetClassName
#undef GetClassInfo
#include <Shlobj.h>
#include <Shlwapi.h>
#include <PathCch.h>
#include <propvarutil.h>
#include <xmllite.h>
#include <OCIdl.h>
#include <Psapi.h>
#undef EnumProcesses

// Visual Studio Platform headers
#include <IVsSccProjectProviderBinding.h>
#include <ocdesign.h>
#include <msdbg.h>
#include <vsdebugguids.h>
#include <vsshell80.h>
#include <vsshell110.h>
#include <vsshell121.h>
#include <vsshell140.h>
#include <vsshell156.h>
#include <vsshell158.h>
#include <vsshell160.h>
#include <vsshell167.h>
#include <VsDbgCmd.h>
#include <fpstfmt.h>
#include <textmgr2.h>
#include <textmgr90.h>
#include <textmgr110.h>
#include <msdbg90.h>
#include <msdbg100.h>
#include <msdbg110.h>
#include <msdbg150.h>
#include <msdbg160.h>
#include <msdbg166.h>
#include <msdbg169.h>
#include <msdbg170.h>
#include <msdbg171.h>
#include <msdbg174.h>
#include <msdbg175.h>
#include <msdbg176.h>
#include <msdbg177.h>
#include <msdbg178.h>
#include <portpriv.h>
#include <vsdebugguids.h>
#include <venuscmddef.h>
#include <sccmnid.h>
#include <vsshlids.h>
#include <sharedids.h>

// WIL
#define RESULT_DIAGNOSTICS_LEVEL 4
#include <wil/result_macros.h>
#include <wil/com.h>
#include <wil/win32_helpers.h>
#include <wil/wistd_type_traits.h>

extern "C" IMAGE_DOS_HEADER __ImageBase;
