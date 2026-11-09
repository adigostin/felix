
#include "pch.h"
#include "shared/com.h"
#include "shared/OtherGuids.h"
#include "DebugEngine.h"

static const wchar_t SingleDebugProcessName[] = L"Z80 Process";

class DebugProcessImpl : public IFelixDebugProcess
{
	ULONG _refCount = 0;
	WeakRefToThis _weakRefToThis;
	wil::unique_process_heap_string _exe_path;
	AD_PROCESS_ID _pid;
	GUID _processGuid;
	com_ptr<IWeakRef> _parentPort;
	FILETIME _creationTime;
	com_ptr<IDebugProgram2> _program;
	bool _attached = false;

public:
	HRESULT InitInstance (IDebugPort2* pPort)
	{
		HRESULT hr;
		hr = _weakRefToThis.InitInstance(static_cast<IDebugProcess2*>(this)); RETURN_IF_FAILED(hr);
		hr = pPort->QueryInterface(&_parentPort); RETURN_IF_FAILED(hr);
		_pid.ProcessIdType = AD_PROCESS_ID_SYSTEM;
		_pid.ProcessId.dwProcessId = SimulatorProcessId;
		hr = CoCreateGuid (&_processGuid); RETURN_IF_FAILED(hr);
		::GetSystemTimeAsFileTime (&_creationTime);

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE InitDebugProgram (SpectrumVariant platform) override
	{
		HRESULT hr;
		RETURN_HR_IF(E_UNEXPECTED, _attached);

		com_ptr<IDebugPortNotify2> portNotify;
		hr = _parentPort->QueryInterface(&portNotify); RETURN_IF_FAILED(hr);
		RETURN_HR_IF_NULL(E_UNEXPECTED, portNotify);

		if (_program)
		{
			com_ptr<IFelixProgram> modules;
			hr = _program->QueryInterface(&modules); RETURN_IF_FAILED(hr);
			hr = modules->UnloadModules(); RETURN_IF_FAILED(hr);
			com_ptr<IDebugProgramNode2> node;
			hr = _program->QueryInterface(&node); RETURN_IF_FAILED(hr);
			hr = portNotify->RemoveProgramNode(node.get()); RETURN_IF_FAILED(hr);
			_exe_path.reset();
		}

		hr = MakeDebugProgram(this, &_program); RETURN_IF_FAILED(hr);
		wil::unique_process_heap_string dll;
		hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, dll); RETURN_IF_FAILED(hr);
		PathFindFileName(dll.get())[0] = 0; // remove file name
		const wchar_t* rom = romRelativePaths[platform];
		hr = str_concat_nothrow(_exe_path, dll, rom); RETURN_IF_NULL_ALLOC(_exe_path);

		// Create a debug module for the ROM.
		wil::com_ptr_nothrow<IDebugModule2> romModule;
		hr = MakeModule (0, 0x5CCB, _exe_path.get(), false, nullptr, &romModule); RETURN_IF_FAILED(hr);
		com_ptr<IFelixProgram> mcoll;
		hr = _program->QueryInterface(&mcoll); RETURN_IF_FAILED(hr);
		hr = mcoll->AddModule(romModule.get()); RETURN_IF_FAILED(hr);

		com_ptr<IDebugProgramNode2> node;
		hr = _program->QueryInterface(&node); RETURN_IF_FAILED(hr);
		hr = portNotify->AddProgramNode(node.get()); RETURN_IF_FAILED(hr); // VS will call IDebugEngine2::Attach from this

		return S_OK;
	}

	#pragma region IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IDebugProcess2*>(this), riid, ppvObject)
			|| TryQI<IDebugProcess2>(this, riid, ppvObject)
			|| TryQI<IFelixDebugProcess>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

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
			if (simulator)
			{
				if (_attached)
					pProcessInfo->Flags |= PIFLAG_DEBUGGER_ATTACHED;
				pProcessInfo->Flags |= simulator->Running_HR() == S_OK
					? PIFLAG_PROCESS_RUNNING
					: PIFLAG_PROCESS_STOPPED;
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
		auto hr = _program->Terminate(); RETURN_IF_FAILED(hr);
		_attached = false;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE Attach(IDebugEventCallback2* pCallback, GUID* rgguidSpecificEngines, DWORD celtSpecificEngines, HRESULT* rghrEngineAttach) override
	{
		RETURN_HR_IF(E_ATTACH_DEBUGGER_ALREADY_ATTACHED, _attached);
		RETURN_HR_IF(E_POINTER, celtSpecificEngines && rghrEngineAttach == nullptr);

		bool supportedEngine = celtSpecificEngines == 0;
		for (DWORD i = 0; i < celtSpecificEngines; i++)
		{
			bool supported = !rgguidSpecificEngines
				|| rgguidSpecificEngines[i] == GUID_NULL
				|| rgguidSpecificEngines[i] == Engine_Id;
			rghrEngineAttach[i] = supported ? S_OK : E_NOINTERFACE;
			supportedEngine |= supported;
		}

		RETURN_HR_IF(E_NOINTERFACE, !supportedEngine);
		_attached = true;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE CanDetach() override
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE Detach() override
	{
		_attached = false;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetPhysicalProcessId(AD_PROCESS_ID* pProcessId) override
	{
		// https://blogs.msdn.microsoft.com/jacdavis/2008/05/01/what-to-do-if-your-debug-engine-doesnt-create-real-processes/
		// Later edit: VS2022 doesn't show GUID processes in the Debug -> Attach dialog.
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
		RETURN_HR_IF(E_UNEXPECTED, !simulator);
		return simulator->Break();
	}

	HRESULT STDMETHODCALLTYPE GetPort(IDebugPort2** ppPort) override
	{
		RETURN_HR_IF_NULL(E_POINTER, ppPort);
		auto hr = _parentPort->QueryInterface(ppPort); RETURN_IF_FAILED(hr);
		RETURN_HR_IF_NULL(E_UNEXPECTED, *ppPort);
		return S_OK;
	}
	#pragma endregion
};

HRESULT MakeDebugProcess (IDebugPort2* pPort, IDebugProcess2** ppProcess)
{
	auto p = com_ptr(new (std::nothrow) DebugProcessImpl()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(pPort); RETURN_IF_FAILED(hr);
	*ppProcess = p.detach();
	return S_OK;
}
