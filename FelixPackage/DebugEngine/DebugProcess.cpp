
#include "pch.h"
#include "shared/com.h"
#include "shared/OtherGuids.h"
#include "DebugEngine.h"

static const wchar_t SingleDebugProcessName[] = L"Z80 Process";

class DebugProcessImpl : public IDebugProcess2
{
	ULONG _refCount = 0;
	wil::unique_hlocal_string _exe_path;
	AD_PROCESS_ID _pid;
	GUID _processGuid;
	com_ptr<IDebugPort2> _parentPort;
	FILETIME _creationTime;
	com_ptr<IDebugProgram2> _program;

public:
	HRESULT InitInstance (IDebugPort2* pPort, LPCOLESTR pszExe, IDebugEngine2* engine, IDebugEventCallback2* callback)
	{
		_pid.ProcessIdType = AD_PROCESS_ID_GUID;
		if (pszExe)
		{
			_exe_path = wil::make_hlocal_string_nothrow(pszExe); RETURN_IF_NULL_ALLOC(_exe_path);
		}

		auto hr = CoCreateGuid (&_pid.ProcessId.guidProcessId); RETURN_IF_FAILED(hr);
		hr = CoCreateGuid (&_processGuid); RETURN_IF_FAILED(hr);
		_parentPort = pPort;
		::GetSystemTimeAsFileTime (&_creationTime);

		hr = MakeDebugProgram (this, engine, callback, &_program); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	#pragma region IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDebugProcess2*>(this), riid, ppvObject)
			|| TryQI<IDebugProcess2>(this, riid, ppvObject)
		)
			return S_OK;
		
		#ifdef _DEBUG
		if (   riid == IID_IClientSecurity
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == GUID{0xDE34E4B4, 0x500B, 0x487F, { 0xB6, 0x43, 0xCE, 0xE1, 0x43, 0xF4, 0x23, 0xFF } }
			|| riid == IID_IDebugProcessEx2
			|| riid == IID_IAgileObject
			|| riid == IID_IdentityUnmarshal
			|| riid == IID_IStdMarshalInfo
			|| riid == IID_INotARealInterface
			|| riid == IID_IFastRundown
			|| riid == IID_INoIdea7
			|| riid == IID_INoIdea9
			|| riid == IID_INoIdea10
			|| riid == IID_INoIdea11
			|| riid == IID_INoIdea14
			|| riid == IID_INoIdea15
			|| riid == IID_IExternalConnection
		)
			return E_NOINTERFACE;

		for (auto& i : HardcodedRundownIFsOfInterest)
			if (*i == riid)
				return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	#pragma region IDebugProcess2
	HRESULT STDMETHODCALLTYPE GetInfo(PROCESS_INFO_FIELDS Fields, PROCESS_INFO* pProcessInfo) override
	{
		pProcessInfo->Fields = 0;

		if (Fields & PIF_FILE_NAME)
		{
			if (_exe_path)
			{
				pProcessInfo->bstrFileName = SysAllocString(_exe_path.get()); RETURN_IF_NULL_ALLOC(pProcessInfo->bstrFileName);
			}
			else
				pProcessInfo->bstrFileName = nullptr;
			pProcessInfo->Fields |= PIF_FILE_NAME;
		}

		if (Fields & PIF_BASE_NAME)
		{
			pProcessInfo->bstrBaseName = SysAllocString(SingleDebugProcessName);
			pProcessInfo->Fields |= PIF_BASE_NAME;
		}

		if (Fields & PIF_TITLE)
		{
			pProcessInfo->bstrTitle = SysAllocString(L"<TITLE>");
			pProcessInfo->Fields |= PIF_TITLE;
		}

		if (Fields & PIF_PROCESS_ID)
		{
			pProcessInfo->ProcessId = _pid;
			pProcessInfo->Fields |= PIF_PROCESS_ID;
		}

		if (Fields & PIF_SESSION_ID)
		{
			pProcessInfo->dwSessionId = 0;
			pProcessInfo->Fields |= PIF_SESSION_ID;
		}

		if (Fields & PIF_ATTACHED_SESSION_NAME)
		{
			pProcessInfo->bstrAttachedSessionName = SysAllocString(L"<session name>");
			pProcessInfo->Fields |= PIF_ATTACHED_SESSION_NAME;
		}

		if (Fields & PIF_CREATION_TIME)
		{
			pProcessInfo->CreationTime = _creationTime;
			pProcessInfo->Fields |= PIF_CREATION_TIME;
		}

		if (Fields & PIF_FLAGS)
		{
			pProcessInfo->Flags = 0;
			//if (_engineAttached)
			{
				pProcessInfo->Flags |= PIFLAG_DEBUGGER_ATTACHED;
				//if (_running)
				pProcessInfo->Flags |= PIFLAG_PROCESS_RUNNING;
				//else
				//	pProcessInfo->Flags |= PIFLAG_PROCESS_STOPPED;
			}
			pProcessInfo->Fields |= PIF_FLAGS;
		}

		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE EnumPrograms(IEnumDebugPrograms2** ppEnum) override
	{
		return make_single_entry_enumerator<IEnumDebugPrograms2, IDebugProgram2>(_program, ppEnum);
	}

	HRESULT STDMETHODCALLTYPE GetName(GETNAME_TYPE gnType, BSTR* pbstrName) override
	{
		switch (gnType)
		{
		case GN_NAME:
		case GN_BASENAME:
			*pbstrName = SysAllocString(SingleDebugProcessName); RETURN_IF_NULL_ALLOC(*pbstrName);
			return S_OK;

		case GN_FILENAME:
			if (_exe_path)
			{
				*pbstrName = SysAllocString(_exe_path.get()); RETURN_IF_NULL_ALLOC(*pbstrName);
			}
			else
				*pbstrName = nullptr;
			return S_OK;

		default:
			*pbstrName = nullptr;
			RETURN_HR(E_NOTIMPL);
		}
	}

	HRESULT STDMETHODCALLTYPE GetServer(IDebugCoreServer2** ppServer) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	HRESULT STDMETHODCALLTYPE Terminate() override
	{
		// called for example when we return an error code from ResumeProcess()
		_program->Terminate();
		_program = nullptr;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE Attach(IDebugEventCallback2* pCallback, GUID* rgguidSpecificEngines, DWORD celtSpecificEngines, HRESULT* rghrEngineAttach) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE CanDetach() override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE Detach() override
	{
		RETURN_HR(E_NOTIMPL);
	}

	HRESULT STDMETHODCALLTYPE GetPhysicalProcessId(AD_PROCESS_ID* pProcessId) override
	{
		// https://blogs.msdn.microsoft.com/jacdavis/2008/05/01/what-to-do-if-your-debug-engine-doesnt-create-real-processes/
		*pProcessId = _pid;
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE GetProcessId(GUID* pguidProcessId) override
	{
		*pguidProcessId = _processGuid;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetAttachedSessionName(BSTR* pbstrSessionName) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE EnumThreads (IEnumDebugThreads2 **ppEnum) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE CauseBreak() override
	{
		RETURN_HR(E_NOTIMPL);
	}

	HRESULT STDMETHODCALLTYPE GetPort(IDebugPort2** ppPort) override
	{
		*ppPort = _parentPort.get();
		(*ppPort)->AddRef();
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeDebugProcess (IDebugPort2* pPort, LPCOLESTR pszExe, IDebugEngine2* engine,
	IDebugEventCallback2* callback, IDebugProcess2** ppProcess)
{
	auto p = com_ptr(new (std::nothrow) DebugProcessImpl()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(pPort, pszExe, engine, callback); RETURN_IF_FAILED(hr);
	*ppProcess = p.detach();
	return S_OK;
}
