
#pragma once
#include "shared/z80_register_set.h"

struct ISimulator;

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{E420CF93-E250-4C0A-A96D-A58D52651D14}") ISimulatorEvent : IUnknown
{
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{A8DC31F4-2BFD-4BCF-BBA0-7DA0DE6C9F71}") ISimulatorEventHandler : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE ProcessSimulatorEvent (ISimulatorEvent* event, REFIID riidEvent) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("F578EBE3-7596-4FC7-A976-56C9B0AAD856") IScreenCompleteEventHandler : IUnknown
{
	// If the implementation returns a SUCCEEDED error code, it takes ownership of "bi" and must release it with CoTaskMemFree when no longer needed.
	// If the implementation returns a FAILED error code, the caller retains ownership of "bi".
	virtual HRESULT STDMETHODCALLTYPE OnScreenComplete (BITMAPINFO* bi, POINT beamLocation) = 0;
};

enum class BreakpointType { Code, Data };

typedef DWORD SIM_BP_COOKIE;

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{33551613-1FE1-4D48-B805-9D5C82ACDA0E}") ISimulatorBreakpointEvent : ISimulatorEvent
{
	virtual BreakpointType GetType() = 0;
	virtual UINT16 GetAddress() = 0;
	virtual ULONG GetBreakpointCount() = 0;
	virtual HRESULT GetBreakpointAt(ULONG i, SIM_BP_COOKIE* ppKey) = 0;
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{D9C9D2E8-E684-451A-B614-CD7ADB3792F5}") ISimulatorSimulateOneEvent : ISimulatorEvent
{
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{0D2531CF-8EE1-4CD9-9A01-38FE17B7CE26}") ISimulatorResumeEvent : ISimulatorEvent
{
};

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{CD782861-2511-4B52-BE2C-0EC6F9D4F0D6}") ISimulatorBreakEvent : ISimulatorEvent
{
};

#define SIM_E_SNAPSHOT_FILE_WRONG_SIZE                MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x202)

struct DECLSPEC_NOVTABLE DECLSPEC_UUID("{56344845-3DDA-4BC0-9645-7EBA3FE94A93}") ISimulator : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE ReadMemoryBus  (uint16_t address, uint16_t size, void* to) = 0;
	virtual HRESULT STDMETHODCALLTYPE WriteMemoryBus (uint16_t address, uint16_t size, const void* from) = 0;
	virtual HRESULT STDMETHODCALLTYPE Break() = 0;
	virtual HRESULT STDMETHODCALLTYPE Resume (BOOL checkBreakpointsAtCurrentPC) = 0;
	virtual HRESULT STDMETHODCALLTYPE Reset (UINT16 startAddress) = 0;
	virtual HRESULT STDMETHODCALLTYPE Running_HR() = 0;
	virtual HRESULT STDMETHODCALLTYPE SimulateOne() = 0;
	virtual HRESULT STDMETHODCALLTYPE AdviseDebugEvents (ISimulatorEventHandler* handler) = 0;
	virtual HRESULT STDMETHODCALLTYPE UnadviseDebugEvents (ISimulatorEventHandler* handler) = 0;
	virtual HRESULT STDMETHODCALLTYPE AdviseScreenComplete (IScreenCompleteEventHandler* handler) = 0;
	virtual HRESULT STDMETHODCALLTYPE UnadviseScreenComplete (IScreenCompleteEventHandler* handler) = 0;
	virtual HRESULT STDMETHODCALLTYPE ProcessKeyDown (uint32_t vkey, uint32_t modifiers) = 0;
	virtual HRESULT STDMETHODCALLTYPE ProcessKeyUp   (uint32_t vkey, uint32_t modifiers) = 0;
	virtual HRESULT STDMETHODCALLTYPE AddBreakpoint (BreakpointType type, bool physicalMemorySpace, UINT64 address, SIM_BP_COOKIE* pCookie) = 0;
	virtual HRESULT STDMETHODCALLTYPE RemoveBreakpoint (SIM_BP_COOKIE cookie) = 0;
	virtual HRESULT STDMETHODCALLTYPE HasBreakpoints_HR() = 0;
	virtual HRESULT STDMETHODCALLTYPE LoadFile (LPCWSTR pFileName) = 0;
	virtual HRESULT STDMETHODCALLTYPE LoadBinary (IStream* stream, DWORD address) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetPC (uint16_t* pc) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetPC (uint16_t pc) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetStackStartAddress (UINT16* stackStartAddress) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetRegisters (z80_register_set* buffer, uint32_t size) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetRegisters (const z80_register_set* buffer, uint32_t size) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetShowCRTSnapshot() = 0; // returns S_OK or S_FALSE
	virtual HRESULT STDMETHODCALLTYPE SetShowCRTSnapshot(BOOL val) = 0;
};

HRESULT MakeSimulator (LPCWSTR dir, LPCWSTR romFilename, ISimulator** to);
