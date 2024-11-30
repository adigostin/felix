
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"

struct DECLSPEC_NOINITALL ProjectConfigBuilder : IProjectConfigBuilder
{
	ULONG _refCount = 0;
	com_ptr<IVsUIHierarchy> _hier;
	com_ptr<IProjectConfig> _config;
	wil::unique_bstr _projName;
	com_ptr<IVsOutputWindowPane2> _outputWindow2;
	HWND _hwnd;
	com_ptr<IProjectConfigBuilderCallback> _callback;

	HRESULT InitInstance (IVsUIHierarchy* hier, IProjectConfig* config, IVsOutputWindowPane* outputWindowPane)
	{
		_hier = hier;
		_config = config;
		auto hr = outputWindowPane->QueryInterface(&_outputWindow2); RETURN_IF_FAILED(hr);

		wil::unique_variant projectName;
		hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &projectName); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, projectName.vt != VT_BSTR);
		_projName = wil::unique_bstr(projectName.release().bstrVal);

		return S_OK;
	}

	~ProjectConfigBuilder()
	{
		if (_hwnd)
		{
			::SetWindowLongPtr (_hwnd, GWLP_USERDATA, 0);
			::DestroyWindow(_hwnd);
		}
	}

	#pragma region IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override { RETURN_HR(E_NOTIMPL); }

	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
	#pragma endregion

	HRESULT OutputMessage (LPCOLESTR message)
	{
		return _outputWindow2->OutputTaskItemStringEx2 (message, (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
			nullptr, 0, nullptr, 0, 0, _projName.get(), nullptr, nullptr);
	}

	struct InputFiles
	{
		vector_nothrow<com_ptr<IProjectFile>> CustomBuild;
		vector_nothrow<com_ptr<IProjectFile>> Asm;
	};

	HRESULT GetInputFiles (InputFiles& files)
	{
		wil::unique_variant childItemId;
		auto hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &childItemId); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, childItemId.vt != VT_VSITEMID);
		while (V_VSITEMID(&childItemId) != VSITEMID_NIL)
		{
			wil::unique_variant browseObjectVariant;
			hr = _hier->GetProperty(V_VSITEMID(&childItemId), VSHPROPID_BrowseObject, &browseObjectVariant);
			if (SUCCEEDED(hr) && (browseObjectVariant.vt == VT_DISPATCH))
			{
				com_ptr<IProjectFile> file;
				if (SUCCEEDED(browseObjectVariant.pdispVal->QueryInterface(&file)))
				{
					BuildToolKind tool;
					hr = file->get_BuildTool(&tool); RETURN_IF_FAILED(hr);
					bool pushed;
					switch(tool)
					{
					case BuildToolKind::None:
						break;

					case BuildToolKind::Assembler:
						pushed = files.Asm.try_push_back (std::move(file)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
						break;

					case BuildToolKind::CustomBuildTool:
						pushed = files.CustomBuild.try_push_back(std::move(file)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
						break;

					default:
						RETURN_HR(E_NOTIMPL);
					}
				}
			}

			hr = _hier->GetProperty(V_VSITEMID(&childItemId), VSHPROPID_NextSibling, &childItemId); RETURN_IF_FAILED(hr);
			RETURN_HR_IF(E_FAIL, childItemId.vt != VT_VSITEMID);
		}

		return S_OK;
	}

	static HRESULT Write (ISequentialStream* stream, wchar_t c)
	{
		return stream->Write(&c, sizeof(c), nullptr);
	}

	static HRESULT Write (ISequentialStream* stream, const wchar_t* psz)
	{
		return stream->Write(psz, (ULONG)wcslen(psz) * sizeof(wchar_t), nullptr);
	}

	static HRESULT Write (ISequentialStream* stream, const wchar_t* from, const wchar_t* to)
	{
		return stream->Write(from, (ULONG)(to - from) * sizeof(wchar_t), nullptr);
	}

	HRESULT MakeSjasmCommandLine (IProjectFile* const* files, uint32_t fileCount, LPCWSTR project_dir,
		LPCWSTR output_dir, BSTR* ppCmdLine)
	{
		HRESULT hr;

		com_ptr<IStream> cmdLine;
		hr = CreateStreamOnHGlobal (nullptr, TRUE, &cmdLine); RETURN_IF_FAILED(hr);

		wil::unique_hlocal_string moduleFilename;
		hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, moduleFilename); RETURN_IF_FAILED(hr);
		auto fnres = PathFindFileName(moduleFilename.get()); RETURN_HR_IF(CO_E_BAD_PATH, fnres == moduleFilename.get());
		bool hasSpaces = !!wcschr(moduleFilename.get(), L' ');
		if (hasSpaces)
		{
			hr = Write(cmdLine, L'\"'); RETURN_IF_FAILED(hr);
		}
		hr = Write(cmdLine, moduleFilename.get(), fnres); RETURN_IF_FAILED(hr);
		hr = Write(cmdLine, L"sjasmplus.exe"); RETURN_IF_FAILED(hr);
		if (hasSpaces)
		{
			hr = Write(cmdLine, L"\""); RETURN_IF_FAILED(hr);
		}
		hr = Write(cmdLine, L" --fullpath"); RETURN_IF_FAILED(hr);

		auto addOutputPathParam = [&cmdLine, output_dir, project_dir](const wchar_t* paramName, const wchar_t* output_filename) -> HRESULT
			{
				wil::unique_hlocal_string outputFilePath;
				const DWORD PathFlags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
				auto hr = PathAllocCombine (output_dir, output_filename, PathFlags, &outputFilePath); RETURN_IF_FAILED(hr);
				auto outputFilePathRelativeUgly = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePathRelativeUgly);
				BOOL bRes = PathRelativePathToW (outputFilePathRelativeUgly.get(), project_dir, FILE_ATTRIBUTE_DIRECTORY, outputFilePath.get(), 0); RETURN_HR_IF(CS_E_INVALID_PATH, !bRes);
				size_t len = wcslen(outputFilePathRelativeUgly.get());
				auto outputFilePathRelative = wil::make_hlocal_string_nothrow(nullptr, len); RETURN_IF_NULL_ALLOC(outputFilePathRelative);
				hr = PathCchCanonicalizeEx (outputFilePathRelative.get(), len + 1, outputFilePathRelativeUgly.get(), PathFlags); RETURN_IF_FAILED(hr);
				hr = Write(cmdLine, paramName); RETURN_IF_FAILED(hr);
				hr = Write(cmdLine, outputFilePathRelative.get()); RETURN_IF_FAILED(hr);
				return S_OK;
			};

		// --raw=...
		wil::unique_bstr output_filename;
		hr = _config->GetOutputFileName(&output_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (L" --raw=", output_filename.get()); RETURN_IF_FAILED(hr);

		// --sld=...
		wil::unique_bstr sld_filename;
		hr = _config->GetSldFileName (&sld_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (L" --sld=", sld_filename.get()); RETURN_IF_FAILED(hr);

		// --outprefix
		hr = addOutputPathParam (L" --outprefix=", L""); RETURN_IF_FAILED(hr);

		// --lst
		com_ptr<IProjectConfigAssemblerProperties> asmProps;
		hr = _config->get_AssemblerProperties(&asmProps); RETURN_IF_FAILED(hr);
		VARIANT_BOOL saveListing;
		hr = asmProps->get_SaveListing(&saveListing); RETURN_IF_FAILED(hr);
		if (saveListing)
		{
			hr = Write(cmdLine, L" --lst"); RETURN_IF_FAILED(hr);
			wil::unique_bstr listingFilename;
			hr = asmProps->get_SaveListingFilename(&listingFilename); RETURN_IF_FAILED(hr);
			if (listingFilename && listingFilename.get()[0])
			{
				hr = Write(cmdLine, L"="); RETURN_IF_FAILED(hr);
				hr = Write(cmdLine, listingFilename.get()); RETURN_IF_FAILED(hr);
			}
		}

		// input files
		for (uint32_t i = 0; i < fileCount; i++)
		{
			IProjectFile* file = files[i];
			wil::unique_bstr fileRelativePath;
			hr = file->get_Path(&fileRelativePath);
			if (SUCCEEDED(hr))
			{
				hr = Write(cmdLine, L" "); RETURN_IF_FAILED(hr);
				hr = Write(cmdLine, fileRelativePath.get()); RETURN_IF_FAILED(hr);
			}
		}

		ULARGE_INTEGER size;
		hr = cmdLine->Seek({ .QuadPart = 0 }, STREAM_SEEK_CUR, &size); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(ERROR_FILE_TOO_LARGE, !!size.HighPart);

		HGLOBAL hg;
		hr = GetHGlobalFromStream(cmdLine, &hg); RETURN_IF_FAILED(hr);
		auto buffer = GlobalLock(hg); WI_ASSERT(buffer);
		*ppCmdLine = SysAllocStringLen((OLECHAR*)buffer, size.LowPart / 2);
		GlobalUnlock (hg);
		RETURN_IF_NULL_ALLOC(*ppCmdLine);

		return S_OK;
	}

	// This function extracts from the CommandLine property an array of lines separated by 0x0D/0x0A.
	// It launches them one by one, stopping in case of a HRESULT error, or in case of a non-zero exit code.
	// This function writes pExitCode only when it returns S_OK. When it writes pExitCode to non-zero,
	// it additionally writes pbstrCmdLine with the command line whose execution returned that exit code.
	HRESULT RunCommandLines (ICommandLineList* list, PCWSTR eventName, PCWSTR workDir, DWORD* pdwExitCode)
	{
		wil::unique_bstr cmdLines;
		auto hr = list->get_CommandLine(&cmdLines); RETURN_IF_FAILED(hr);
		if (!cmdLines)
			return S_OK;

		vector_nothrow<wil::unique_bstr> cls;
		for (PCWSTR p = cmdLines.get(); *p; )
		{
			auto from = p;
			while (*p && *p != 0x0D)
				p++;
			if (p > from)
			{
				auto line = SysAllocStringLen(from, (UINT)(p - from)); RETURN_IF_NULL_ALLOC(line);
				bool pushed = cls.try_push_back(wil::unique_bstr(line)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
			}

			if (p[0] == 0)
				break;

			p++;
			if (*p == 0x0A)
				p++;
		}

		if (cls.size())
		{
			wil::unique_bstr desc;
			hr = list->get_Description(&desc); RETURN_IF_FAILED(hr);
			if (desc)
			{
				hr = OutputMessage(desc.get()); RETURN_IF_FAILED(hr);
			}

			for (auto& cl : cls)
			{
				wil::unique_bstr outputStr;
				hr = RunTool(cl.get(), workDir, pdwExitCode, outputStr.addressof()); RETURN_IF_FAILED_EXPECTED(hr);
				hr = OutputMessage(outputStr.get()); RETURN_IF_FAILED(hr);
				if (*pdwExitCode)
				{
					static const wchar_t Format[] = L"%s returned exit code %u: %s";
					size_t bufferLen = wcslen(eventName) + _countof(Format) + 10 + SysStringLen(cl.get());
					auto buffer = wil::make_hlocal_string_nothrow(nullptr, bufferLen); RETURN_IF_NULL_ALLOC(buffer);
					swprintf_s (buffer.get(), bufferLen, Format, eventName, *pdwExitCode, cl.get());
					OutputMessage(buffer.get());
					return S_OK;
				}
			}
		}

		*pdwExitCode = 0;
		return S_OK;
	}

	virtual HRESULT StartBuild (IProjectConfigBuilderCallback* callback) override
	{
		HRESULT hr;
		DWORD exitCode;

		wil::unique_variant projectDir;
		hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);

		wil::unique_bstr output_dir;
		hr = _config->GetOutputDirectory(&output_dir); RETURN_IF_FAILED(hr);

		int win32err = SHCreateDirectoryExW (nullptr, output_dir.get(), nullptr);
		if (win32err != ERROR_SUCCESS && win32err != ERROR_ALREADY_EXISTS)
			RETURN_WIN32(win32err);

		// Pre-Build Event
		com_ptr<IProjectConfigPrePostBuildProperties> preBuildProps;
		hr = _config->get_PreBuildProperties(&preBuildProps); RETURN_IF_FAILED(hr);
		hr = RunCommandLines(preBuildProps, L"Pre-Build Event", projectDir.bstrVal, &exitCode); RETURN_IF_FAILED(hr);
		if (exitCode)
			return callback->OnBuildComplete(FALSE), S_OK;

		InputFiles inputFiles;
		hr = GetInputFiles(inputFiles); RETURN_IF_FAILED(hr);

		// First build the files with a custom build tool. This is similar to what VS does.
		for (auto& f : inputFiles.CustomBuild)
		{
			com_ptr<ICustomBuildToolProperties> props;
			hr = f->get_CustomBuildToolProperties(&props); RETURN_IF_FAILED(hr);
			static const WCHAR Format[] = L"Custom Build Tool for file \"%s\"";
			wil::unique_bstr path;
			hr = f->get_Path(&path); RETURN_IF_FAILED(hr);
			size_t eventNameLen = _countof(Format) + SysStringLen(path.get());
			auto eventName = wil::make_hlocal_string_nothrow(nullptr, eventNameLen); RETURN_IF_NULL_ALLOC(eventName);
			swprintf_s (eventName.get(), eventNameLen, Format, path.get());
			hr = RunCommandLines(props, eventName.get(), projectDir.bstrVal, &exitCode); RETURN_IF_FAILED_EXPECTED(hr);
			if (exitCode)
				return callback->OnBuildComplete(FALSE), S_OK;
		}

		// Second launch sjasm to build all asm files.
		if (inputFiles.Asm.size())
		{
			wil::unique_bstr cmdLine;
			hr = MakeSjasmCommandLine (inputFiles.Asm[0].addressof(), inputFiles.Asm.size(),
				projectDir.bstrVal, output_dir.get(), &cmdLine); RETURN_IF_FAILED(hr);

			wil::unique_bstr outputStr;
			hr = RunTool(cmdLine.get(), projectDir.bstrVal, &exitCode, outputStr.addressof()); RETURN_IF_FAILED(hr);

			hr = ParseSjasmOutput (outputStr.get(), outputStr.get() + SysStringLen(outputStr.get())); RETURN_IF_FAILED(hr);
			if (exitCode)
				return callback->OnBuildComplete(FALSE), S_OK;
		}

		// Post-Build Event
		com_ptr<IProjectConfigPrePostBuildProperties> postBuildProps;
		hr = _config->get_PostBuildProperties(&postBuildProps); RETURN_IF_FAILED(hr);
		hr = RunCommandLines(postBuildProps, L"Post-Build Event", projectDir.bstrVal, &exitCode); RETURN_IF_FAILED_EXPECTED(hr);
		if (exitCode)
			return callback->OnBuildComplete(FALSE), S_OK;

		callback->OnBuildComplete(TRUE);
		return S_OK;
	}

	// This function writes pExitCode and pbstrOutput only when it returns S_OK.
	HRESULT RunTool (PWSTR cmdLine, PCWSTR workDir, DWORD* pExitCode, BSTR* pbstrOutput)
	{
		// Create a pipe for the child process's STDOUT.
		wil::unique_handle stdoutReadHandle;
		wil::unique_handle stdoutWriteHandle;
		SECURITY_ATTRIBUTES saAttr = { .nLength = sizeof(SECURITY_ATTRIBUTES), .bInheritHandle = TRUE, };
		BOOL bres = CreatePipe(&stdoutReadHandle, &stdoutWriteHandle, &saAttr, 0); RETURN_IF_WIN32_BOOL_FALSE(bres);

		// Ensure the read handle to the pipe for STDOUT is not inherited.
		bres = SetHandleInformation(stdoutReadHandle.get(), HANDLE_FLAG_INHERIT, 0); RETURN_IF_WIN32_BOOL_FALSE(bres);

		STARTUPINFO startupInfo = { .cb = sizeof(startupInfo) };
		startupInfo.hStdError = stdoutWriteHandle.get();
		startupInfo.hStdOutput = stdoutWriteHandle.get();
		//startupInfo.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
		startupInfo.dwFlags |= STARTF_USESTDHANDLES;

		wil::unique_process_information processInfo;
		bres = CreateProcess (NULL, cmdLine, NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL,
			workDir, &startupInfo, &processInfo);
		if(!bres)
			return SetErrorInfo(HRESULT_FROM_WIN32(GetLastError()), L"Cannot create process with command line: %s\r\n", cmdLine);

		WaitForSingleObject(processInfo.hProcess, IsDebuggerPresent() ? INFINITE : 5000);
		DWORD exitCode = E_FAIL;
		GetExitCodeProcess (processInfo.hProcess, &exitCode);
		stdoutWriteHandle = nullptr;

		const DWORD ReadSize = 100;
		DWORD outputSize = 0;
		wil::unique_cotaskmem_ansistring output;
		while (true)
		{
			auto newBlock = (PSTR)CoTaskMemRealloc(output.get(), outputSize + ReadSize); RETURN_IF_NULL_ALLOC(newBlock);
			output.release();
			output.reset(newBlock);
			DWORD bytesRead;
			bres = ReadFile (stdoutReadHandle.get(), output.get() + outputSize, ReadSize, &bytesRead, nullptr);
			if (!bres)
			{
				DWORD gle = GetLastError();
				if (gle != ERROR_BROKEN_PIPE)
					RETURN_WIN32(gle);
				break;
			}
			outputSize += bytesRead;
		}

		// Let's assume UTF-8. Looking at its source code, sjasmplus seems to be ASCII-only.
		auto hr = MakeBstrFromString (output.get(), output.get() + outputSize, pbstrOutput); RETURN_IF_FAILED(hr);
		*pExitCode = exitCode;
		return S_OK;
	}

	static bool TryParsePrio (wchar_t*& p, VSTASKPRIORITY& prio)
	{
		if (!wcsncmp(p, L"error: ", 7))
		{
			p += 7;
			prio = TP_HIGH;
			return true;
		}

		if (!_wcsnicmp(p, L"warning: ", 9))
		{
			p += 9;
			prio = TP_NORMAL;
			return true;
		}

		if (!_wcsnicmp(p, L"warning[", 8))
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

	HRESULT ParseSjasmOutput (wchar_t* outputText, wchar_t* outputTextEnd)
	{
		HRESULT hr;

		for (wchar_t* line = outputText; line < outputTextEnd; )
		{
			bool addedTaskItem = false;
			if (auto p = wcspbrk(line, L":\x0D"); p
				&& (p[0] == ':') 
				&& (p - line >= 2) 
				&& p[-1] == ')' 
				&& isdigit(p[-2]) 
				&& p[1] == ' ')
			{
				wchar_t* pp = p - 2;
				while (pp > line && isdigit(pp[-1]))
					--pp;
				if (pp > line && pp[-1] == '(')
				{
					wchar_t* filenameTo = pp - 1;
					p += 2; // jump over ": "
					VSTASKPRIORITY taskPrio;
					if (TryParsePrio(p, taskPrio))
					{
						const wchar_t* taskItemText = p;
						while (*p && *p != 0x0D)
							p++;
						if (*p == 0x0D)
						{
							*p++ = 0;
							if (*p == 0x0A)
								p++;
						}

						wil::unique_bstr taskFilename (SysAllocStringLen(line, (UINT)(filenameTo - line)));
						ULONG lineNum = wcstoul(filenameTo + 1, nullptr, 10) - 1;
						hr = _outputWindow2->OutputTaskItemStringEx2(line, taskPrio, CAT_BUILDCOMPILE, L"", BMP_COMPILE,
							taskFilename.get(), lineNum, 0, _projName.get(), taskItemText, L""); RETURN_IF_FAILED(hr);
						addedTaskItem = true;
						line = p;
					}
				}
			}

			if (!addedTaskItem)
			{
				auto p = line;
				while (*p && *p != 0x0D)
					p++;
				if (*p == 0x0D)
				{
					*p++ = 0;
					if (*p == 0x0A)
						p++;
				}

				hr = _outputWindow2->OutputTaskItemStringEx2 (line, (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
					nullptr, 0, nullptr, 0, 0, _projName.get(), nullptr, nullptr); RETURN_IF_FAILED(hr);
				line = p;
			}
		}

		return S_OK;
	}
};

HRESULT MakeProjectConfigBuilder (IVsUIHierarchy* hier, IProjectConfig* config,
	IVsOutputWindowPane* outputWindowPane, IProjectConfigBuilder** to)
{
	auto p = com_ptr(new (std::nothrow) ProjectConfigBuilder()); RETURN_IF_NULL_ALLOC(p);
	auto hr = p->InitInstance(hier, config, outputWindowPane); RETURN_IF_FAILED(hr);
	*to = p.detach();
	return S_OK;
}
