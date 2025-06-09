
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/inplace_function.h"

struct IBuildStep;

MIDL_INTERFACE("72484332-D17F-41D8-8AF0-48C1EE7D01BF")
IBuildStepCallback : IUnknown
{
	virtual void OnStepComplete (bool success) = 0;
};

struct IBuildStep : IUnknown
{
	// S_OK    - Build step was run synchronously and was successful, callback will not be called.
	// S_FALSE - Build step was successfully started and is running asynchronously, callback will be called with the outcome.
	// error code - Build step could not be started, callback will not be called.
	virtual HRESULT RunStep (IBuildStepCallback* callback) = 0;
	
	// Cancels a build step that's running asynchronously - one that returned S_FALSE from Run and
	// hasn't had its callback called yet. The implementation of this function calls the callback
	// with success=false before returning.
	virtual HRESULT CancelStep() = 0;
};

class BuildStepMessage : public IBuildStep
{
	ULONG _refCount = 0;
	wil::unique_bstr _projectName;
	wil::unique_bstr _message;
	com_ptr<IVsOutputWindowPane2> _outputWindow;

public:
	static HRESULT CreateInstance (PCWSTR projectName, wil::unique_bstr message,
		IVsOutputWindowPane2* outputWindow, IBuildStep** ppStep)
	{
		auto p = com_ptr(new (std::nothrow) BuildStepMessage()); RETURN_IF_NULL_ALLOC(p);
		p->_projectName = wil::make_bstr_nothrow(projectName);
		p->_message = std::move(message);
		p->_outputWindow = outputWindow;
		*ppStep = p.detach();
		return S_OK;
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { RETURN_HR(E_NOTIMPL); }

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	virtual HRESULT RunStep (IBuildStepCallback* callback) override
	{
		auto hr = _outputWindow->OutputTaskItemStringEx2 (_message.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
			nullptr, 0, nullptr, 0, 0, _projectName.get(), nullptr, nullptr); RETURN_IF_FAILED(hr);
		uint32_t len = SysStringLen(_message.get());
		if (len < 2 || _message.get()[len - 2] != 0x0D || _message.get()[len - 1] != 0x0A)
		{
			hr = _outputWindow->OutputTaskItemStringEx2 (L"\r\n", (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
				nullptr, 0, nullptr, 0, 0, _projectName.get(), nullptr, nullptr); RETURN_IF_FAILED(hr);
		}

		return S_OK;
	}

	virtual HRESULT CancelStep() override
	{
		RETURN_HR(E_NOTIMPL);
	}
};

// "pszLine" may contain CR/LF or not
using LineParser = HRESULT(*)(const char* pszLine, const char* lineEnd,
	VSTASKPRIORITY* pnPriority,
	BSTR* pbstrFilename,
	ULONG* pnLineNum,
	ULONG* pnColumn,
	BSTR* pbstrTaskItemText);

class BuildStepRunProcess : public IBuildStep
{
	ULONG _refCount = 0;
	wil::unique_bstr _cmdLine;
	wil::unique_bstr _workDir;
	wil::unique_bstr _projectName;
	bool _printExitCode;
	LineParser _lineParser;
	com_ptr<IVsOutputWindowPane2> _outputWindow;
	//com_ptr<IBuildStepCallback> _callback;
	com_ptr<IWeakRef> _stepCompleteCallback;
	char buffer[250];
	wil::unique_hfile _pipeHandle;
	DWORD _lastOutputStringTickCount;
	OVERLAPPED _overlapped;
	wil::unique_process_information _processInfo;
	vector_nothrow<char> _lineBuffer;
	uint32_t _cursor;
	
	// We create this (with SetTimer) when we run the first step in pendingSteps.
	// We destroy it when we're done running the last step in pendingSteps.
	static UINT_PTR timerId;
	static vector_nothrow<BuildStepRunProcess*> pendingSteps;

public:
	static HRESULT CreateInstance (PCWSTR cmdLine, PCWSTR workDir, PCWSTR projectName, bool printExitCode,
		LineParser lineParser, IVsOutputWindowPane2* outputWindow, IBuildStep** ppStep)
	{
		auto p = com_ptr(new (std::nothrow) BuildStepRunProcess()); RETURN_IF_NULL_ALLOC(p);
		p->_cmdLine = wil::make_bstr_nothrow(cmdLine); RETURN_IF_NULL_ALLOC(p->_cmdLine);
		p->_workDir = wil::make_bstr_nothrow(workDir); RETURN_IF_NULL_ALLOC(p->_workDir);
		p->_projectName = wil::make_bstr_nothrow(projectName); RETURN_IF_NULL_ALLOC(p->_projectName);
		p->_printExitCode = printExitCode;
		p->_lineParser = std::move(lineParser);
		p->_outputWindow = outputWindow;
		bool reserved = p->_lineBuffer.try_reserve(1000); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);
		*ppStep = p.detach();
		return S_OK;
	}

	BuildStepRunProcess()
	{
	}

	~BuildStepRunProcess()
	{
		if (_pipeHandle.is_valid())
		{
			// The user of this object has called RunStep on us, has not yet called CancelStep on us,
			// and has released the last reference to us. Not a meaningful use case, but we should
			// at least not crash and not leak memory.
			// Additionally, as an implementation detail, we choose not to call the callback.
			TerminateStep();
			WI_ASSERT(!_pipeHandle.is_valid());
		}
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { RETURN_HR(E_NOTIMPL); }
	 
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		return ReleaseST(this, _refCount);
	}
	#pragma endregion

	virtual HRESULT RunStep (IBuildStepCallback* callback) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !!_stepCompleteCallback);

		com_ptr<IWeakRef> callbackAsWeak;
		auto hr = callback->QueryInterface(&callbackAsWeak); RETURN_IF_FAILED(hr);

		if (pendingSteps.size() == pendingSteps.capacity())
		{
			bool reserved = pendingSteps.try_reserve(pendingSteps.size() + 10); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);
		}

		// Create a pipe for the child process's STDOUT.
		// Generate a unique name in case we're building multiple projects at the same time.
		static int id;
		static const wchar_t PipeNameFormat[] = L"\\\\.\\pipe\\BuildStepRunProcess%u";
		size_t nameLen =  _countof(PipeNameFormat) + 10;
		auto pipeName = wil::make_hlocal_string_nothrow (nullptr, nameLen); RETURN_IF_NULL_ALLOC(pipeName);
		swprintf_s(pipeName.get(), nameLen + 1, PipeNameFormat, id);
		id++;
		DWORD openMode = PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED;
		DWORD pipeMode = PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_REJECT_REMOTE_CLIENTS;
		wil::unique_hfile pipeHandle (CreateNamedPipe (pipeName.get(), openMode, pipeMode,
			1, 1024, 1024, 0, NULL)); RETURN_LAST_ERROR_IF(!pipeHandle.is_valid());

		SECURITY_ATTRIBUTES secu = { .nLength = sizeof(SECURITY_ATTRIBUTES), .bInheritHandle = TRUE };
		wil::unique_hfile stdoutWriteHandle (CreateFileW (pipeName.get(), GENERIC_WRITE, 0, &secu, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL)); RETURN_LAST_ERROR_IF(!stdoutWriteHandle.is_valid());

		_overlapped = { };
		BOOL bres = ReadFile (pipeHandle.get(), buffer, sizeof(buffer), nullptr, &_overlapped);
		WI_ASSERT(!bres && GetLastError() == ERROR_IO_PENDING);

		STARTUPINFO startupInfo = { .cb = sizeof(startupInfo) };
		startupInfo.hStdError = stdoutWriteHandle.get();
		startupInfo.hStdOutput = stdoutWriteHandle.get();
		startupInfo.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
		startupInfo.dwFlags |= STARTF_USESTDHANDLES;

		bres = CreateProcess (NULL, _cmdLine.get(), NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL,
			_workDir.get(), &startupInfo, &_processInfo);
		if(!bres)
		{
			auto hr = HRESULT_FROM_WIN32(GetLastError());
			return SetErrorInfo(hr, L"Cannot create process with command line: %s\r\n", _cmdLine.get());
		}

		if (pendingSteps.empty())
		{
			// Inefficient but good enough for virtually all ASM projects.
			timerId = ::SetTimer (NULL, 0, 1, TimerProcStatic); RETURN_LAST_ERROR_IF(timerId == 0);
		}
		pendingSteps.try_push_back(this);

		_pipeHandle = std::move(pipeHandle);
		_cursor = 0;
		_stepCompleteCallback = callbackAsWeak;
		_lastOutputStringTickCount = GetTickCount();
		return E_PENDING;
	}
	
	static void CALLBACK TimerProcStatic (HWND hwnd, UINT uMsg, UINT_PTR timerId, DWORD dwTime)
	{
		WI_ASSERT (timerId == BuildStepRunProcess::timerId);
		WI_ASSERT (!pendingSteps.empty());
		for (uint32_t i = pendingSteps.size() - 1; i != -1; i--)
		{
			auto keepAlive = com_ptr(pendingSteps[i]);

			bool stepComplete;
			pendingSteps[i]->TimerProc(&stepComplete);
			if (stepComplete)
			{
				pendingSteps.remove(pendingSteps.begin() + i);
				if (pendingSteps.empty())
				{
					BOOL bres = KillTimer (NULL, BuildStepRunProcess::timerId); WI_ASSERT(bres);
					BuildStepRunProcess::timerId = 0;
				}
			}
		}
	}

	void TimerProc (bool* stepComplete)
	{
		DWORD exitCode;
		HRESULT hr = ReadProcessOutput(&exitCode);
		if (hr == E_PENDING)
		{
			*stepComplete = false;
			return;
		}

		*stepComplete = true;
		_pipeHandle.reset();

		bool callbackSuccess = false;
		if (SUCCEEDED(hr))
		{
			callbackSuccess = (exitCode == 0);
			if (exitCode && _printExitCode)
			{
				const wchar_t Format[] = L"Process finished with exit code %d. Command line was: %s\r\n";
				size_t bufferSize = _countof(Format) + 10 + wcslen(_cmdLine.get());
				if (auto message = wil::make_cotaskmem_string_nothrow(nullptr, bufferSize))
				{
					swprintf_s(message.get(), bufferSize + 1, Format, exitCode, _cmdLine.get());
					hr = _outputWindow->OutputTaskItemStringEx2 (message.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
						nullptr, 0, nullptr, 0, 0, _projectName.get(), nullptr, nullptr);
				}
			}
		}

		if (_lineBuffer.size())
		{
			// There's some output for which we were waiting a CR/LF. Let's send the output now.
			wil::unique_bstr str;
			hr = MakeBstrFromString(_lineBuffer.begin(), _lineBuffer.end(), &str);
			if (SUCCEEDED(hr))
			{
				_outputWindow->OutputTaskItemStringEx2 (str.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
					nullptr, 0, nullptr, 0, 0, _projectName.get(), nullptr, nullptr);
			}
			_lineBuffer.clear();
		}

		com_ptr<IBuildStepCallback> callback;
		hr = _stepCompleteCallback->QueryInterface(&callback); WI_ASSERT(SUCCEEDED(hr));
		if (callback)
			callback->OnStepComplete(callbackSuccess);
		_stepCompleteCallback = nullptr;
	}

	// Returns:
	//  S_OK      - process exited, exit code written at *pExitCode.
	//  E_PENDING - process still running
	//  other error HR - unhandled error
	HRESULT ReadProcessOutput (DWORD* pExitCode)
	{
		HRESULT hr;
		BOOL bres;

		while(true)
		{
			DWORD bytesRead;
			bres = GetOverlappedResult (_pipeHandle.get(), &_overlapped, &bytesRead, FALSE);
			if (!bres)
			{
				DWORD gle = GetLastError();
				if (gle == ERROR_IO_INCOMPLETE) // 996
				{
					// If the read is pending for too long and we have text generated by the process,
					// it may be a prompt waiting for user input. Happens, for example, with a custom
					// build tool that launches xcopy, which generates this prompt without CR/LF:
					// "Overwrite file.ext (Yes/No/All)?"
					// We want to user to see what's going on, so let's send it to the Output Window.
					if (_lineBuffer.size() && GetTickCount() >= _lastOutputStringTickCount + 250)
					{
						wil::unique_bstr bufferU;
						hr = MakeBstrFromString (_lineBuffer.begin(), _lineBuffer.end(), &bufferU); RETURN_IF_FAILED(hr);
						hr = _outputWindow->OutputTaskItemStringEx2 (bufferU.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0, 
							nullptr, 0, nullptr, 0, 0, _projectName.get(), nullptr, nullptr); RETURN_IF_FAILED(hr);
						_lastOutputStringTickCount = GetTickCount();
						_lineBuffer.clear();
					}

					return E_PENDING;
				}
				
				if (gle == ERROR_BROKEN_PIPE) // 109
					// Process exited
					break;

				WI_ASSERT(false);
				RETURN_HR(HRESULT_FROM_WIN32(gle));
			}

			for (uint32_t i = 0; i < bytesRead; i++)
			{
				hr = ProcessReadByte(buffer[i]); RETURN_IF_FAILED(hr);
			}

			// Restart reading right away. If we don't do this, the process might generate some text
			// and then exit before we have a chance to start reading it, and we will lose that text.
			bres = ReadFile (_pipeHandle.get(), buffer, sizeof(buffer), nullptr, &_overlapped);
			if (!bres)
			{
				DWORD gle = GetLastError();
				if (gle == ERROR_BROKEN_PIPE) // 109
					// Process exited
					break;

				if (gle == ERROR_IO_PENDING) // 997
					return E_PENDING;

				WI_ASSERT(false);
				RETURN_HR(HRESULT_FROM_WIN32(gle));
			}

			// ReadFile succeeded again, keep processing the output of the process
		}

		DWORD exitCode;
		bres = GetExitCodeProcess (_processInfo.hProcess, &exitCode);
		if (!bres)
			RETURN_LAST_ERROR();

		*pExitCode = exitCode;
		return S_OK;
	}

	HRESULT ProcessReadByte (BYTE b)
	{
		// TODO: handle non-ANSI output
		if (b == 0x0D)
		{
			// CR
			_cursor = 0;
			//_outputWindowCount = 0;
		}
		else if (b == 0x0A)
		{
			// LF
			// The Output Window wants lines terminated with CR/LF.
			bool pushed = _lineBuffer.try_push_back({ 0x0D, 0x0A }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

			wil::unique_bstr bufferU;
			auto hr = MakeBstrFromString (_lineBuffer.begin(), _lineBuffer.end(), &bufferU); RETURN_IF_FAILED(hr);
			if (_lineParser)
			{
				VSTASKPRIORITY nPriority = { };
				wil::unique_bstr bstrFilename;
				ULONG nLineNum = 0;
				ULONG nColumn = 0;
				wil::unique_bstr bstrTaskItemText;
				hr = _lineParser (_lineBuffer.begin(), _lineBuffer.end(), &nPriority, &bstrFilename, &nLineNum, &nColumn, &bstrTaskItemText); RETURN_IF_FAILED(hr);
				WI_ASSERT(hr == S_OK || hr == S_FALSE);
				hr = _outputWindow->OutputTaskItemStringEx2(bufferU.get(), nPriority, CAT_BUILDCOMPILE, L"", BMP_COMPILE,
					bstrFilename.get(), nLineNum, nColumn, _projectName.get(), bstrTaskItemText.get(), nullptr); RETURN_IF_FAILED(hr);
			}
			else
			{
				hr = _outputWindow->OutputTaskItemStringEx2 (bufferU.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0, 
					nullptr, 0, nullptr, 0, 0, _projectName.get(), nullptr, nullptr); RETURN_IF_FAILED(hr);
			}

			_lastOutputStringTickCount = GetTickCount();
			_lineBuffer.clear();
		}
		else if (b >= 0x20)
		{
			while (_lineBuffer.size() <= _cursor)
			{
				bool pushed = _lineBuffer.try_push_back(' '); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
			}

			_lineBuffer[_cursor] = b;
			_cursor++;
		}
		else
		{
			// Some ASCII character below 0x20. Let's not bother.
		}

		return S_OK;
	}

	void TerminateStep()
	{
		WI_ASSERT(_processInfo.hProcess != nullptr);
		WI_ASSERT(_pipeHandle.is_valid());

		// Terminate the process now, if it hasn't exited already.
		TerminateProcess(_processInfo.hProcess, ERROR_PROCESS_ABORTED);

		// Give it some time to exit.
		DWORD waitResult = WaitForSingleObject(_processInfo.hProcess, 5000);
		WI_ASSERT(waitResult == WAIT_OBJECT_0);

		_pipeHandle.reset();

		pendingSteps.remove(this);
		if (pendingSteps.empty())
		{
			BOOL bres = KillTimer (NULL, timerId); WI_ASSERT(bres);
			timerId = 0;
		}
	}

	virtual HRESULT CancelStep() override
	{
		RETURN_HR_IF (E_UNEXPECTED, !_processInfo.hProcess);
		RETURN_HR_IF (E_UNEXPECTED, !_pipeHandle.is_valid());

		TerminateStep();

		com_ptr<IBuildStepCallback> callback;
		auto hr = _stepCompleteCallback->QueryInterface(&callback); WI_ASSERT(SUCCEEDED(hr));
		if (callback)
			callback->OnStepComplete(false);
		_stepCompleteCallback = nullptr;

		return S_OK;
	}
};

UINT_PTR BuildStepRunProcess::timerId;
vector_nothrow<BuildStepRunProcess*> BuildStepRunProcess::pendingSteps;

struct ProjectConfigBuilder : IProjectConfigBuilder, IBuildStepCallback
{
	ULONG _refCount = 0;
	com_ptr<IVsHierarchy> _hier;
	com_ptr<IProjectConfig> _config;
	wil::unique_bstr _projName;
	com_ptr<IVsOutputWindowPane2> _outputWindow2;
	com_ptr<IProjectConfigBuilderCallback> _callback;
	vector_nothrow<com_ptr<IBuildStep>> _steps;
	uint32_t _currentStep = UINT32_MAX; // UINT32_MAX means we haven't started yet, or have finished already
	WeakRefToThis _weakRefToThis;

public:
	// The presence of an explicit constructor keeps the compiler from memset-ing
	// the object with zero after operator new and before the implicit constructor.
	// (This helped me catch an uninitialized var.)
	ProjectConfigBuilder() { }

	HRESULT InitInstance (IVsHierarchy* hier, IProjectConfig* config, IVsOutputWindowPane2* outputWindowPane)
	{
		HRESULT hr;

		_hier = hier;
		_config = config;
		_outputWindow2 = outputWindowPane;
		hr = _weakRefToThis.InitInstance(static_cast<IProjectConfigBuilder*>(this)); RETURN_IF_FAILED(hr);

		wil::unique_variant projectName;
		hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &projectName); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, projectName.vt != VT_BSTR);
		_projName = wil::unique_bstr(projectName.release().bstrVal);
		return S_OK;
	}

	~ProjectConfigBuilder()
	{
		if (_currentStep < _steps.size())
		{
			// The user of this object has called StartBuild on us, has not yet called
			// CancelBuild on us, and has released the last reference to us.
			WI_ASSERT(_callback);
			ULONG remainingRefCount = _steps[_currentStep].detach()->Release();
			WI_ASSERT(remainingRefCount == 0);
			_callback = nullptr;
		}
		else
			WI_ASSERT(_callback == nullptr);
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		if (   TryQI<IUnknown>(static_cast<IProjectConfigBuilder*>(this), riid, ppvObject)
			|| TryQI<IBuildStepCallback>(this, riid, ppvObject)
		)
			return S_OK;

		if (riid == __uuidof(IWeakRef))
			return _weakRefToThis.QueryIWeakRef(ppvObject);

		RETURN_HR(E_NOINTERFACE);
	}

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	// This function extracts from the CommandLine property an array of lines separated by 0x0D/0x0A.
	// It launches them one by one, stopping in case of a HRESULT error, or in case of a non-zero exit code.
	// This function writes pExitCode only when it returns S_OK. When it writes pExitCode to non-zero,
	// it additionally writes pbstrCmdLine with the command line whose execution returned that exit code.
	HRESULT ParseCommandLines (ICommandLineList* list, const wchar_t* workDir, vector_nothrow<com_ptr<IBuildStep>>& steps)
	{
		wil::unique_bstr desc;
		auto hr = list->get_Description(&desc); RETURN_IF_FAILED(hr);
		if (desc && desc.get()[0])
		{
			com_ptr<IBuildStep> step;
			hr = BuildStepMessage::CreateInstance(_projName.get(), std::move(desc), _outputWindow2, step.addressof()); RETURN_IF_FAILED(hr);
			bool pushed = steps.try_push_back(std::move(step)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}

		wil::unique_bstr cmdLines;
		hr = list->get_CommandLine(&cmdLines); RETURN_IF_FAILED(hr);
		if (!cmdLines)
			return S_OK;
		
		for (PCWSTR p = cmdLines.get(); *p; )
		{
			while (*p && (*p == ' ' || *p == '\t'))
				p++;
			auto from = p;
			while (*p && *p != 0x0D)
				p++;
			while (p > from && (p[-1] == ' ' || p[-1] == '\t'))
				p--;
			if (p > from)
			{
				auto line = SysAllocStringLen(from, (UINT)(p - from)); RETURN_IF_NULL_ALLOC(line);
				com_ptr<IBuildStep> step;
				hr = BuildStepRunProcess::CreateInstance (line, workDir, _projName.get(), true, nullptr, _outputWindow2, step.addressof()); RETURN_IF_FAILED(hr);
				bool pushed = steps.try_push_back(std::move(step)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
			}

			if (p[0] == 0)
				break;

			p++;
			if (*p == 0x0A)
				p++;
		}
		
		return S_OK;
	}

	static HRESULT EnumDescendants (IVsHierarchy* hier, VSITEMID from, const stdext::inplace_function<HRESULT(IChildNode*)>& filter)
	{
		// TODO: search in depth too when we'll support directories

		wil::unique_variant childItemId;
		auto hr = hier->GetProperty(from, VSHPROPID_FirstChild, &childItemId);
		while(SUCCEEDED(hr) && (childItemId.vt == VT_VSITEMID) && (V_VSITEMID(&childItemId) != VSITEMID_NIL))
		{
			wil::unique_variant obj;
			hr = hier->GetProperty(V_VSITEMID(&childItemId), VSHPROPID_BrowseObject, &obj);
			if (SUCCEEDED(hr) && (obj.vt == VT_DISPATCH) && obj.pdispVal)
			{
				com_ptr<IChildNode> item;
				hr = obj.pdispVal->QueryInterface(&item);
				if (SUCCEEDED(hr))
				{
					hr = filter(item);
					if (hr == S_OK || FAILED(hr))
						return hr;
				}
			}

			hr = hier->GetProperty (V_VSITEMID(&childItemId), VSHPROPID_NextSibling, &childItemId);
		}

		return S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE StartBuild (IProjectConfigBuilderCallback* callback) override
	{
		RETURN_HR_IF(E_UNEXPECTED, !_steps.empty());

		vector_nothrow<com_ptr<IBuildStep>> steps;

		wil::unique_variant projectDir;
		auto hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, projectDir.addressof()); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);

		wil::unique_bstr output_dir;
		hr = _config->GetOutputDirectory(&output_dir); RETURN_IF_FAILED(hr);

		// Pre-Build Event
		com_ptr<IProjectConfigPrePostBuildProperties> preBuildProps;
		hr = _config->AsProjectConfigProperties()->get_PreBuildProperties(&preBuildProps); RETURN_IF_FAILED(hr);
		hr = ParseCommandLines (preBuildProps, projectDir.bstrVal, steps);RETURN_IF_FAILED(hr);
		
		// First build the files with a custom build tool. This is similar to what VS does.
		hr = EnumDescendants (_hier, VSITEMID_ROOT, [this, &steps, workDir = projectDir.bstrVal](IChildNode* item)
			{
				com_ptr<IFileNodeProperties> file;
				if (SUCCEEDED(item->QueryInterface(&file)))
				{
					BuildToolKind tool;
					auto hr = file->get_BuildTool(&tool); RETURN_IF_FAILED(hr);
					if (tool == BuildToolKind::CustomBuildTool)
					{
						com_ptr<ICustomBuildToolProperties> props;
						hr = file->get_CustomBuildToolProperties(&props); RETURN_IF_FAILED(hr);
						//static const WCHAR Format[] = L"Custom Build Tool for file \"%s\"";
						//wil::unique_bstr path;
						//hr = f->get_Path(&path); RETURN_IF_FAILED(hr);
						//size_t eventNameLen = _countof(Format) + SysStringLen(path.get());
						//auto eventName = wil::make_hlocal_string_nothrow(nullptr, eventNameLen); RETURN_IF_NULL_ALLOC(eventName);
						//swprintf_s (eventName.get(), eventNameLen, Format, path.get());
						hr = ParseCommandLines (props, workDir, steps); RETURN_IF_FAILED(hr);
					}
				}
				return S_FALSE;
			});

		// Second launch sjasm to build all asm files.
		com_ptr<IProjectConfigAssemblerProperties> asmProps;
		hr = _config->AsProjectConfigProperties()->get_AssemblerProperties(&asmProps); RETURN_IF_FAILED(hr);
		wil::unique_bstr cmdLine;
		hr = MakeSjasmCommandLine (_hier, _config, asmProps, &cmdLine); RETURN_IF_FAILED(hr);
		if (SysStringLen(cmdLine.get()) > 0)
		{
			com_ptr<IBuildStep> buildStep;
			hr = BuildStepRunProcess::CreateInstance (cmdLine.get(), projectDir.bstrVal, _projName.get(), false, ParseSjasmOutput, _outputWindow2, &buildStep); RETURN_IF_FAILED(hr);
			bool pushed = steps.try_push_back(std::move(buildStep)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}

		// Post-Build Event
		com_ptr<IProjectConfigPrePostBuildProperties> postBuildProps;
		hr = _config->AsProjectConfigProperties()->get_PostBuildProperties(&postBuildProps); RETURN_IF_FAILED(hr);
		hr = ParseCommandLines (postBuildProps, projectDir.bstrVal, steps);RETURN_IF_FAILED(hr);

		if (steps.empty())
		{
			hr = _outputWindow2->OutputTaskItemStringEx2 (L"Nothing to build.\r\n", (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
				nullptr, 0, nullptr, 0, 0, _projName.get(), nullptr, nullptr); RETURN_IF_FAILED(hr);
			return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
		}

		int win32err = SHCreateDirectoryExW (nullptr, output_dir.get(), nullptr);
		if (win32err != ERROR_SUCCESS && win32err != ERROR_ALREADY_EXISTS)
			RETURN_WIN32(win32err);

		// Try to run them synchronously.
		uint32_t i;
		for (i = 0; i < steps.size(); i++)
		{
			hr = steps[i]->RunStep(this);
			if (hr == E_PENDING)
				break;
			if (FAILED(hr))
				return hr;
		}

		if (i == steps.size())
			return S_OK;

		// Step i is running asynchronously.
		_callback = callback;
		_steps = std::move(steps);
		_currentStep = i;
		return S_FALSE;
	}

	virtual HRESULT STDMETHODCALLTYPE CancelBuild() override
	{
		RETURN_HR_IF(E_UNEXPECTED, _currentStep >= _steps.size());
		RETURN_HR_IF(E_UNEXPECTED, !_callback);
		auto hr = _steps[_currentStep]->CancelStep(); RETURN_IF_FAILED(hr);
		WI_ASSERT(_callback == nullptr);
		_steps.clear();
		_currentStep = UINT32_MAX;
		return S_OK;
	}

	#pragma region IBuildStepCallback
	virtual void OnStepComplete (bool success) override
	{
		_currentStep++;

		if (!success)
		{
			auto callback = std::move(_callback);
			callback->OnBuildComplete(false);
			while (_steps.size() > _currentStep)
				_steps.remove_back();
			return;
		}

		while (_currentStep < _steps.size())
		{
			auto hr = _steps[_currentStep]->RunStep(this);
			if (hr == E_PENDING)
				return;
			if (FAILED(hr))
			{
				auto callback = std::move(_callback);
				callback->OnBuildComplete(false);
				return;
			}
			_currentStep++;
		}

		auto callback = std::move(_callback);
		callback->OnBuildComplete(true);
	}
	#pragma endregion

	static bool TryParsePrio (const char*& p, VSTASKPRIORITY& prio)
	{
		if (!strncmp(p, "error: ", 7))
		{
			p += 7;
			prio = TP_HIGH;
			return true;
		}

		if (!_strnicmp(p, "warning: ", 9))
		{
			p += 9;
			prio = TP_NORMAL;
			return true;
		}

		if (!_strnicmp(p, "warning[", 8))
		{
			auto pp = p + 8;
			while(pp[0] && pp[0] != ']')
				pp++;
			if (pp[0] != ']' || pp[1] != ':' || pp[2] != ' ')
				return false;
			p = pp + 3;
			prio = TP_NORMAL;
			return true;
		}

		return false;
	}

	static HRESULT ParseGenericOutput (wchar_t* pszLine, IVsOutputWindowPane2* pane, LPCOLESTR projName)
	{
		return pane->OutputTaskItemStringEx2 (pszLine, (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
			nullptr, 0, nullptr, 0, 0, projName, nullptr, nullptr);
	}

	static const char* strchr (const char* ptr, const char* end, char ch)
	{
		while (ptr < end)
		{
			if (*ptr == ch)
				return ptr;
			ptr++;
		}

		return nullptr;
	}

	static HRESULT ParseSjasmOutput (const char* line, const char* lineEnd,
		VSTASKPRIORITY* pnPriority,
		BSTR* pbstrFilename,
		ULONG* pnLineNum,
		ULONG* pnColumn,
		BSTR* pbstrTaskItemText)
	{
		if (auto p = strchr(line, lineEnd, ':'); p
			&& (p - line >= 2) && p[-1] == ')' && isdigit(p[-2]) 
			&& (p + 1 < lineEnd) && (p[1] == ' '))
		{
			const char* pp = p - 2;
			while (pp > line && isdigit(pp[-1]))
				--pp;
			if (pp > line && pp[-1] == '(')
			{
				const char* filenameTo = pp - 1;
				p += 2; // jump over ": "
				VSTASKPRIORITY taskPrio;
				if (TryParsePrio(p, taskPrio))
				{
					const char* taskItemText = p;
					auto tend = p;
					while (*tend && *tend != 0x0D)
						tend++;
					wil::unique_bstr text;
					auto hr = MakeBstrFromString (taskItemText, tend - taskItemText, &text); RETURN_IF_FAILED(hr);
					wil::unique_bstr fn;
					hr = MakeBstrFromString (line, filenameTo, &fn); RETURN_IF_FAILED(hr);
					*pnPriority = taskPrio;
					*pbstrFilename = fn.release();
					*pnLineNum = strtoul(filenameTo + 1, nullptr, 10) - 1;
					*pnColumn = 0;
					*pbstrTaskItemText = text.release();
					return S_OK;
				}
			}
		}

		return S_FALSE;
	}
};

FELIX_API HRESULT MakeProjectConfigBuilder (IVsHierarchy* hier, IProjectConfig* config,
	IVsOutputWindowPane2* outputWindowPane, IProjectConfigBuilder** to)
{
	auto p = com_ptr(new (std::nothrow) ProjectConfigBuilder()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(hier, config, outputWindowPane); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
