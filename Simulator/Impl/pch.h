
#ifndef PCH_H
#define PCH_H

#define STRICT
#define _WIN32_WINNT    _WIN32_WINNT_WIN10
#define NTDDI_VERSION   NTDDI_WIN10_RS5
#define NOMINMAX // Windows Platform min and max macros cause problems for the Standard C++ Library
#define WIN32_LEAN_AND_MEAN	// Exclude rarely-used stuff from the Windows Platform headers
#define _COM_NO_STANDARD_GUIDS_
#define _SILENCE_CXX23_ALIGNED_STORAGE_DEPRECATION_WARNING
#define _HAS_EXCEPTIONS 0 // Visual C++ headers look at this
#undef __EXCEPTIONS // WIL and Intellisense look at this

// Windows Platform headers
#include <Shlwapi.h>
#include <PathCch.h>
#include <OCIdl.h>


// WIL
#define RESULT_DIAGNOSTICS_LEVEL 4
#include <wil/result_macros.h>
#include <wil/com.h>
#include <wil/win32_helpers.h>
#include <wil/wistd_type_traits.h>

extern "C" IMAGE_DOS_HEADER __ImageBase;

#endif //PCH_H
