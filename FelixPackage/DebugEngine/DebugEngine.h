
#pragma once
#include "Simulator.h"
#include "../FelixPackage.h"

// TODO: rename these to E_Z80_XXX
#define E_UNKNOWN_FILE_EXTENSION             MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x201)
#define E_UNRECOGNIZED_DEBUG_FILE_EXTENSION  MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x203)
#define E_NO_EXE_FILENAME                    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x20B)

struct __declspec(novtable) __declspec(uuid("EE50873B-DFAC-4F23-826D-495832A8FD6F")) IZ80Module : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE GetSymbols (IFelixSymbols** symbols) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("80F0C048-F2BA-4FC7-B180-1C93AA7462BA") IFelixProgram : IDebugProgram2
{
	// Adds a module to the program and sends VS the IDebugModuleLoadEvent2 event.
	// (VS will in turn attempts to bind breakpoints.)
	virtual HRESULT STDMETHODCALLTYPE AddModule (IDebugModule2* module) = 0;
	virtual HRESULT STDMETHODCALLTYPE UnloadModules() = 0;
	virtual HRESULT STDMETHODCALLTYPE AddBreakpoint (IDebugBoundBreakpoint2* bp, BreakpointType type, bool physicalMemorySpace, UINT64 address) = 0;
	virtual HRESULT STDMETHODCALLTYPE RemoveBreakpoint (IDebugBoundBreakpoint2* bp) = 0;
	virtual HRESULT STDMETHODCALLTYPE ContainsBreakpoint (IDebugBoundBreakpoint2* bp) = 0;
	virtual HRESULT STDMETHODCALLTYPE EngineAttach (IDebugEngine2* pEngine, IDebugEventCallback2* pCallback, ATTACH_REASON dwReason) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("2BF19E5A-6231-47C6-87DB-2AA904DC51CE") IFelixDebugProcess : IDebugProcess2
{
	virtual HRESULT STDMETHODCALLTYPE ReplaceProgram (LPCWSTR exePath) = 0;
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

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{EE258338-F1FA-4FE5-A518-F7A589F43B60}") IFelixLoadCompleteEvent : IUnknown
{
};

HRESULT MakeSimplePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IFelixProgram* program,
	bool physicalMemorySpace, UINT64 address, IDebugPendingBreakpoint2** to);
HRESULT MakeSourceLinePendingBreakpoint (IDebugEventCallback2* callback, IDebugEngine2* engine, IFelixProgram* program,
	IDebugBreakpointRequest2* bp_request, const wchar_t* file, uint32_t line_index, IDebugPendingBreakpoint2** to);
HRESULT MakeDebugPort (const wchar_t* portName, const GUID& portId, IDebugPort2** port);
HRESULT MakeDebugProgram (IDebugProcess2* process, REFGUID programId, IDebugProgram2** ppProgram);
HRESULT MakeDebugProcess (LPCWSTR romPath, IDebugProcess2** ppProcess);

inline constexpr DWORD SimulatorProcessId = 0xFFFFFFFF;
extern com_ptr<IDebugProcess2> simulatorProcess;

// Symbols are loaded when the module is created and retained for the module's lifetime.
HRESULT MakeModule (UINT64 address, DWORD size, LPCWSTR path, bool user_code, LPCWSTR sourceBasePath, IDebugModule2** to);

HRESULT MakeThread (IDebugProgram2* program, IDebugThread2** ppThread);
HRESULT MakeEnumDebugFrameInfo (FRAMEINFO_FLAGS dwFieldSpec, UINT nRadix, IDebugThread2* thread, IEnumDebugFrameInfo2** to);
HRESULT MakeDebugContext (bool physicalMemorySpace, UINT64 uCodeLocationId, IDebugProgram2* program, REFIID riid, void** ppContext);
HRESULT MakeDocumentContext (IDebugProgram2* program, UINT64 address, IDebugDocumentContext2** to);
HRESULT MakeEnumRegisterGroupsPropertyInfo (IDebugThread2* thread, DEBUGPROP_INFO_FLAGS flags, IEnumDebugPropertyInfo2** to);
HRESULT MakeDisassemblyStream (DISASSEMBLY_STREAM_SCOPE dwScope, IDebugProgram2* program, IDebugCodeContext2* pCodeContext, IDebugDisassemblyStream2** to);
HRESULT GetAddressFromSourceLocation (IDebugProgram2* program, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress);
HRESULT GetAddressFromSourceLocation (IDebugModule2* module, LPCWSTR sourceLocationFilename, DWORD sourceLocationLineIndex, OUT UINT64* pAddress);
HRESULT GetSourceLocationFromAddress (IDebugProgram2* program, UINT64 address, BSTR* pbstrSourceFilePath, UINT32* pSourceLineIndex, _Out_opt_ UINT32* sourceRangeLineIndex);
HRESULT GetModuleAtAddress (IDebugProgram2* program, UINT64 address, IDebugModule2** ppModule);
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