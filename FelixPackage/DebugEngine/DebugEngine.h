
#pragma once
#include "Simulator.h"
#include "../FelixPackage.h"

// TODO: rename these to E_Z80_XXX
#define E_UNKNOWN_FILE_EXTENSION             MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x201)
#define E_UNRECOGNIZED_DEBUG_FILE_EXTENSION  MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x203)
#define E_NO_MODULE_AT_THIS_ADDRESS          MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x207)
#define E_NO_EXE_FILENAME                    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x20B)

struct __declspec(novtable) __declspec(uuid("{E6B46DE8-63A5-4F11-BDD8-E61559A5525E}")) IZ80DebugPort : IDebugPort2
{
	virtual HRESULT STDMETHODCALLTYPE SendProgramCreateEventToSinks  (IDebugProgram2 *pProgram) = 0;
	virtual HRESULT STDMETHODCALLTYPE SendProgramDestroyEventToSinks (IDebugProgram2 *pProgram, DWORD exitCode) = 0;
};

struct __declspec(novtable) __declspec(uuid("EE50873B-DFAC-4F23-826D-495832A8FD6F")) IZ80Module : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE GetSymbols (IFelixSymbols** symbols) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("25071BD1-2108-4D4D-B665-8FE19E31A6CF") IDebugModuleCollection : IUnknown
{
	// Adds a module to the program and sends VS the IDebugModuleLoadEvent2 event.
	// (VS will in turn attempts to bind breakpoints.)
	virtual HRESULT STDMETHODCALLTYPE AddModule (IDebugModule2* module) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{1EFD0353-0389-4325-BA93-5834A6198E57}") IFelixCodeContext : IUnknown
{
	// If 'false', the address is what the CPU sees (always 0 to 64K).
	// If 'true', the address is in the RAM, which can be different than 64K in size
	//   (different ZX Spectrum models have 32K / 64K / 128K).
	virtual bool PhysicalMemorySpace() const = 0;
	virtual UINT64 Address() const = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{7B10BF8A-903E-4FDD-927C-D9827EEF0229}") IFelixLaunchOptionsProvider : IUnknown
{
	virtual HRESULT GetLaunchOptions (IFelixLaunchOptions** ppOptions) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{FCC163D3-2F23-457E-BDDD-EE5A94D9EC83}") IBreakpointManager : IUnknown
{
	virtual HRESULT AddBreakpoint (IDebugBoundBreakpoint2* bp, BreakpointType type, bool physicalMemorySpace, UINT64 address) = 0;
	virtual HRESULT RemoveBreakpoint (IDebugBoundBreakpoint2* bp) = 0;
	virtual HRESULT ContainsBreakpoint (IDebugBoundBreakpoint2* bp) = 0;
};

HRESULT MakeBreakpointManager (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, IBreakpointManager** ppManager);
HRESULT MakeSimplePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program,
	IBreakpointManager* bpman, bool physicalMemorySpace, UINT64 address, IDebugPendingBreakpoint2** to);
HRESULT MakeSourceLinePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program,
	IBreakpointManager* bpman,
	IDebugBreakpointRequest2* bp_request, const wchar_t* file, uint32_t line_index, IDebugPendingBreakpoint2** to);
HRESULT MakeDebugPort (const wchar_t* portName, const GUID& portId, IDebugPort2** port);
HRESULT MakeDebugProgram (IDebugProcess2* process, IDebugEngine2* engine, IDebugEventCallback2* callback, IDebugProgram2** ppProgram);
HRESULT MakeDebugProcess (IDebugPort2* pPort, LPCOLESTR pszExe, IDebugEngine2* engine,
	IDebugEventCallback2* callback, IDebugProcess2** ppProcess);

// If "symbolsFilePath" is non-NULL, it is the full path to the file with symbols; the module implementation
// will not attempt to load symbols from other places.
// If "symbolsFilePath" is NULL, the module implementation may attempt to load symbols from other places,
// such as from a file located next to the module file with the same filename and a different extension.
// The module implementation attempts to load the symbols only once per debug session, when the symbols
// are needed for the first time, for example in the implementation of IZ80Module::GetSymbols.
HRESULT MakeModule (UINT64 address, DWORD size, const wchar_t* path, const wchar_t* symbolsFilePath, bool user_code,
	IDebugEngine2* engine, IDebugProgram2* program, IDebugEventCallback2* callback, IDebugModule2** to);

HRESULT MakeThread (IDebugEngine2* engine, IDebugProgram2* program, IDebugEventCallback2* callback, IDebugThread2** ppThread);
HRESULT MakeEnumDebugFrameInfo (FRAMEINFO_FLAGS dwFieldSpec, UINT nRadix, IDebugThread2* thread, IEnumDebugFrameInfo2** to);
HRESULT MakeDebugContext (bool physicalMemorySpace, UINT64 uCodeLocationId, IDebugProgram2* program, REFIID riid, void** ppContext);
HRESULT MakeDocumentContext (IDebugModule2* module, uint16_t address, IDebugDocumentContext2** to);
HRESULT MakeEnumRegisterGroupsPropertyInfo (IDebugThread2* thread, DEBUGPROP_INFO_FLAGS flags, IEnumDebugPropertyInfo2** to);
HRESULT MakeDisassemblyStream (DISASSEMBLY_STREAM_SCOPE dwScope, IDebugProgram2* program, IDebugCodeContext2* pCodeContext, IDebugDisassemblyStream2** to);
HRESULT GetAddressFromSourceLocation (IDebugProgram2* program, LPCWSTR projectDirOrNull, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress);
HRESULT GetAddressFromSourceLocation (IDebugModule2* module, LPCWSTR projectDirOrNull, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress);
HRESULT GetModuleAtAddress (IDebugProgram2* program, UINT64 address, OUT IDebugModule2** ppModule);
HRESULT MakeExpressionContextNoSource (IDebugThread2* thread, IDebugExpressionContext2** to);
HRESULT MakeRegisterExpression (IDebugThread2* thread, LPCOLESTR originalText, z80_reg16 reg, IDebugExpression2** to);
HRESULT MakeNumberExpression (LPCOLESTR originalText, bool physicalMemorySpace, uint32_t value, IDebugProgram2* program, IDebugExpression2** to);
HRESULT MakeNumberProperty (LPCOLESTR originalText, bool physicalMemorySpace, UINT64 value, IDebugProgram2* program, IDebugProperty2** to);
HRESULT MakeMemoryBytes (IDebugMemoryBytes2** to);
HRESULT GetSymbolFromAddress(
	__RPC__in IDebugProgram2* program,
	__RPC__in uint16_t address,
	__RPC__in SymbolKind searchKind,
	__RPC__deref_out_opt SymbolKind* foundKind,
	__RPC__deref_out_opt BSTR* foundSymbol,
	__RPC__deref_out_opt UINT16* foundOffset,
	__RPC__deref_out_opt IDebugModule2** foundModule);