
#pragma once
#include "Simulator/Simulator.h"
#include "../FelixPackage.h"

// TODO: rename these to E_Z80_XXX
#define E_UNKNOWN_FILE_EXTENSION             MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x201)
#define E_MODULE_HAS_NO_SYMBOLS              MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x202)
#define E_UNRECOGNIZED_DEBUG_FILE_EXTENSION  MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x203)
#define E_UNRECOGNIZED_SLD_VERSION           MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x204)
#define E_INVALID_SLD_LINE                   MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x205)
#define E_ADDRESS_NOT_IN_SYMBOL_FILE         MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x206)
#define E_NO_MODULE_AT_THIS_ADDRESS          MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x207)
#define E_SRC_LOCATION_NOT_IN_SYMBOLS        MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x208)
#define E_UNRECOGNIZED_Z80SYM_VERSION        MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x209)
#define E_INVALID_Z80SYM_LINE                MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x20A)
#define E_NO_EXE_FILENAME                    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x20B)
#define E_SYMBOL_NOT_IN_SYMBOL_FILE          MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x20C)

struct __declspec(novtable) __declspec(uuid("{E6B46DE8-63A5-4F11-BDD8-E61559A5525E}")) IZ80DebugPort : IDebugPort2
{
	virtual HRESULT STDMETHODCALLTYPE SendProgramCreateEventToSinks  (IDebugProgram2 *pProgram) = 0;
	virtual HRESULT STDMETHODCALLTYPE SendProgramDestroyEventToSinks (IDebugProgram2 *pProgram, DWORD exitCode) = 0;
};

enum SymbolKind : DWORD
{
	SK_None = 0,
	SK_Code = 1,
	SK_Data = 2,
	SK_Both = 3
};

struct __declspec(novtable) __declspec(uuid("32BEBBF2-86DF-4D79-88DF-1123548C4D8E")) IZ80Symbols : IUnknown
{
	// Returns S_OK if the symbol file contains mapping between source code lines and instruction addresses;
	// in this case GetSourceLocationFromAddress and GetAddressFromSourceLocation can be called to attempt
	// mapping between lines and addresses.
	// 
	// Returns S_FALSE if the symbol file doesn't contain this mapping;
	// in this case GetSourceLocationFromAddress and GetAddressFromSourceLocation return E_NOTIMPL.
	// 
	// TODO: implement this as an optional interface on the same object.
	virtual HRESULT STDMETHODCALLTYPE HasSourceLocationInformation() = 0;

	virtual HRESULT STDMETHODCALLTYPE GetSourceLocationFromAddress(
		__RPC__in uint16_t address,
		__RPC__deref_out BSTR* srcFilename,
		__RPC__out UINT32* srcLineIndex) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSourceLocation(
		__RPC__in LPCWSTR src_filename,
		__RPC__in uint32_t line_index,
		__RPC__out UINT16* address_out) = 0;

	// Finds the symbol (code and/or data) located at a given address.
	// If "offset" is NULL, looks only for an exact address match.
	virtual HRESULT STDMETHODCALLTYPE GetSymbolAtAddress(
		__RPC__in uint16_t address,
		__RPC__in SymbolKind searchKind,
		__RPC__deref_out_opt SymbolKind* foundKind,
		__RPC__deref_out_opt BSTR* foundSymbol,
		__RPC__deref_out_opt UINT16* foundOffset) = 0;

	// Returns S_OK or an error if not found.
	virtual HRESULT STDMETHODCALLTYPE GetAddressFromSymbol (LPCWSTR symbolName, UINT16* address) = 0;
};

struct __declspec(novtable) __declspec(uuid("EE50873B-DFAC-4F23-826D-495832A8FD6F")) IZ80Module : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE GetSymbols (IZ80Symbols** symbols) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("25071BD1-2108-4D4D-B665-8FE19E31A6CF") IDebugModuleCollection : IUnknown
{
	// Adds a module to the program, sends VS the IDebugModuleLoadEvent2 event, and attempts to bind breakpoints.
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

HRESULT MakeBreakpointManager (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program, ISimulator* simulator, IBreakpointManager** ppManager);
HRESULT MakeSimplePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program,
	IBreakpointManager* bpman, bool physicalMemorySpace, UINT64 address, IDebugPendingBreakpoint2** to);
HRESULT MakeSourceLinePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IDebugProgram2* program,
	IBreakpointManager* bpman,
	IDebugBreakpointRequest2* bp_request, const wchar_t* file, uint32_t line_index, IDebugPendingBreakpoint2** to);
HRESULT MakeDebugPort (const wchar_t* portName, const GUID& portId, IDebugPort2** port);
HRESULT MakeDebugProgramProcess (IDebugPort2* pPort, LPCOLESTR pszExe, IDebugEngine2* engine, ISimulator* simulator, IDebugEventCallback2* callback, IDebugProcess2** ppProcess);
HRESULT MakeModule (UINT64 address, DWORD size, const wchar_t* path, const wchar_t* debug_info_path, bool user_code,
	IDebugEngine2* engine, IDebugProgram2* program, IDebugEventCallback2* callback, IDebugModule2** to);
HRESULT MakeSldSymbols (IDebugModule2* module, IZ80Symbols** to);
HRESULT MakeZ80SymSymbols (IDebugModule2* module, IZ80Symbols** to);
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
