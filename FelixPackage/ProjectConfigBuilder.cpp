
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/string_builder.h"

static constexpr wchar_t wnd_class_name[] = L"ProjectConfig-{9838F078-469A-4F89-B08E-881AF33AE76D}";
using unique_atom = wil::unique_any<ATOM, void(ATOM), [](ATOM a) { UnregisterClassW((LPCWSTR)a, (HINSTANCE)&__ImageBase); }>;
static inline unique_atom atom;
static constexpr UINT WM_BUILD_COMPLETE = WM_APP + 1; // lParam = ThreadData*

struct DECLSPEC_NOINITALL ProjectConfigBuilder : IProjectConfigBuilder
{
	ULONG _refCount = 0;
	com_ptr<IVsUIHierarchy> _hier;
	com_ptr<IProjectConfig> _config;
	wil::unique_bstr _projName;
	com_ptr<IVsOutputWindowPane> _outputWindowPane;
	com_ptr<IVsOutputWindowPane2> outputWindow2;
	com_ptr<IVsLaunchPadFactory> launchPadFactory;
	HWND _hwnd;
	com_ptr<IProjectConfigBuilderCallback> _callback;

	struct ThreadData
	{
		wil::unique_bstr cmdLine;
		wil::unique_bstr workDir;
		com_ptr<IVsLaunchPadFactory> launchPadFactory;
		HWND hwnd;
		com_ptr<IVsOutputWindowPane> outputWindow;
		wil::unique_handle thread_handle;
		wil::slim_event_auto_reset thread_started_event;
		wil::unique_bstr outputText;
		DWORD exitCode;
	};

	vector_nothrow<wistd::unique_ptr<ThreadData>> _threads;

	HRESULT InitInstance (IVsUIHierarchy* hier, IProjectConfig* config, IVsOutputWindowPane* outputWindowPane)
	{
		HRESULT hr;

		_hier = hier;
		_config = config;
		_outputWindowPane = outputWindowPane;

		hr = outputWindowPane->QueryInterface(&outputWindow2); RETURN_IF_FAILED(hr);

		wil::unique_variant projectName;
		hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &projectName); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, projectName.vt != VT_BSTR);
		_projName = wil::unique_bstr(projectName.release().bstrVal);

		if (!atom)
		{
			WNDCLASS wc = { };
			wc.lpfnWndProc = ProjectConfigBuilder::window_proc;
			wc.hInstance = (HINSTANCE)&__ImageBase;
			wc.lpszClassName = wnd_class_name;
			atom.reset (RegisterClassW(&wc)); RETURN_LAST_ERROR_IF(!atom);
		}

		_hwnd = CreateWindowExW (0, wnd_class_name, L"", WS_CHILD, 0, 0, 0, 0, HWND_MESSAGE, 0, (HINSTANCE)&__ImageBase, this); RETURN_LAST_ERROR_IF(!_hwnd);
		SetWindowLongPtr (_hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
		hr = serviceProvider->QueryService(SID_SVsLaunchPadFactory, &launchPadFactory); RETURN_IF_FAILED(hr);
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

	HRESULT MakeSjasmCommandLine (IProjectFile* const* files, uint32_t fileCount, LPCWSTR project_dir,
		LPCWSTR output_dir, BSTR* ppCmdLine)
	{
		HRESULT hr;

		wstring_builder cmdLine;

		wil::unique_hlocal_string moduleFilename;
		hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, moduleFilename); RETURN_IF_FAILED(hr);
		auto fnres = PathFindFileName(moduleFilename.get()); RETURN_HR_IF(CO_E_BAD_PATH, fnres == moduleFilename.get());
		cmdLine << L'\"';
		cmdLine.append(moduleFilename.get(), fnres - moduleFilename.get());
		cmdLine << "sjasmplus.exe" << L'\"';

		cmdLine << " --fullpath";

		auto addOutputPathParam = [&cmdLine, output_dir, project_dir](const char* paramName, const wchar_t* output_filename) -> HRESULT
			{
				wil::unique_hlocal_string outputFilePath;
				const DWORD PathFlags = PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS;
				auto hr = PathAllocCombine (output_dir, output_filename, PathFlags, &outputFilePath); RETURN_IF_FAILED(hr);
				auto outputFilePathRelativeUgly = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePathRelativeUgly);
				BOOL bRes = PathRelativePathToW (outputFilePathRelativeUgly.get(), project_dir, FILE_ATTRIBUTE_DIRECTORY, outputFilePath.get(), 0); RETURN_HR_IF(CS_E_INVALID_PATH, !bRes);
				size_t len = wcslen(outputFilePathRelativeUgly.get());
				auto outputFilePathRelative = wil::make_hlocal_string_nothrow(nullptr, len); RETURN_IF_NULL_ALLOC(outputFilePathRelative);
				hr = PathCchCanonicalizeEx (outputFilePathRelative.get(), len + 1, outputFilePathRelativeUgly.get(), PathFlags); RETURN_IF_FAILED(hr);
				cmdLine << paramName << outputFilePathRelative.get();
				return S_OK;
			};

		// --raw=...
		wil::unique_bstr output_filename;
		hr = _config->GetOutputFileName(&output_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (" --raw=", output_filename.get()); RETURN_IF_FAILED(hr);

		// --sld=...
		wil::unique_bstr sld_filename;
		hr = _config->GetSldFileName (&sld_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (" --sld=", sld_filename.get()); RETURN_IF_FAILED(hr);

		// --outprefix
		hr = addOutputPathParam (" --outprefix=", L""); RETURN_IF_FAILED(hr);

		// --lst
		com_ptr<IProjectConfigAssemblerProperties> asmProps;
		hr = _config->get_AssemblerProperties(&asmProps); RETURN_IF_FAILED(hr);
		VARIANT_BOOL saveListing;
		hr = asmProps->get_SaveListing(&saveListing); RETURN_IF_FAILED(hr);
		if (saveListing)
		{
			wil::unique_bstr _listingFilename;
			hr = asmProps->get_SaveListingFilename(&_listingFilename); RETURN_IF_FAILED(hr);
			cmdLine << " --lst";
			if (_listingFilename && _listingFilename.get()[0])
				cmdLine << "=" << _listingFilename.get();
		}

		// input files
		for (uint32_t i = 0; i < fileCount; i++)
		{
			IProjectFile* file = files[i];
			wil::unique_bstr fileRelativePath;
			hr = file->get_Path(&fileRelativePath);
			if (SUCCEEDED(hr))
				cmdLine << ' ' << fileRelativePath.get();
		}

		cmdLine << '\0';
		RETURN_HR_IF(E_OUTOFMEMORY, cmdLine.out_of_memory());

		*ppCmdLine = SysAllocStringLen (cmdLine.data(), cmdLine.size()); RETURN_IF_NULL_ALLOC(*ppCmdLine);
		return S_OK;
	}

	virtual HRESULT StartBuild (IProjectConfigBuilderCallback* callback) override
	{
		HRESULT hr;

		_callback = callback;

		wil::unique_variant project_dir;
		hr = _hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &project_dir); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_FAIL, project_dir.vt != VT_BSTR);

		wil::unique_bstr output_dir;
		hr = _config->GetOutputDirectory(&output_dir); RETURN_IF_FAILED(hr);

		int win32err = SHCreateDirectoryExW (nullptr, output_dir.get(), nullptr);
		if (win32err != ERROR_SUCCESS && win32err != ERROR_ALREADY_EXISTS)
			RETURN_WIN32(win32err);

		InputFiles inputFiles;
		hr = GetInputFiles(inputFiles); RETURN_IF_FAILED(hr);

		// First build the files with a custom build tool. This is similar to what VS does.
		//for (auto& f : inputFiles.CustomBuild)
		//{
		//	wstring_builder cmdLine;
		//}

		// Second launch sjasm to build all asm files.
		wil::unique_bstr cmdLine;
		hr = MakeSjasmCommandLine (inputFiles.Asm[0].addressof(), inputFiles.Asm.size(),
			project_dir.bstrVal, output_dir.get(), &cmdLine); RETURN_IF_FAILED(hr);

		auto td = wistd::unique_ptr<ThreadData>(new (std::nothrow) ThreadData()); RETURN_IF_NULL_ALLOC(td);
		td->cmdLine = std::move(cmdLine);
		td->workDir = wil::unique_bstr(SysAllocString(project_dir.bstrVal));
		td->launchPadFactory = launchPadFactory;
		td->hwnd = _hwnd;
		td->outputWindow = _outputWindowPane;
		td->thread_handle = wil::unique_handle (CreateThread (nullptr, 0, ThreadProc, td.get(), 0, nullptr)); RETURN_LAST_ERROR_IF_NULL(td->thread_handle);
		td->thread_started_event.wait();
		_threads.try_push_back(std::move(td));
		return S_OK;
	}

	static DWORD WINAPI ThreadProc (LPVOID arg)
	{
		ThreadData* data = static_cast<ThreadData*>(arg);
		data->thread_started_event.SetEvent();
		com_ptr<IVsLaunchPad> lp;
		auto hr = data->launchPadFactory->CreateLaunchPad(&lp);
		if (FAILED(hr))
		{
			data->exitCode = hr;
			PostMessageW (data->hwnd, WM_BUILD_COMPLETE, 0, reinterpret_cast<LPARAM>(data));
			return hr;
		}
		
		// LPF_PipeStdoutToOutputWindow, and possibly LPF_PipeStdoutToTaskList too,
		// don't actually do anything. I debugged through the VS code and execution eventually
		// gets to a function related to the Output Window that throws if not on GUI code.
		// The same functionality does work in the VS2005 samples. I suspect this functionality
		// was abandoned since then and nobody noticed cause nobody used it with a worker thread.
		// So let's pass 0 here, and we parse the output ourselves.
		LAUNCHPAD_FLAGS flags = 0;

		wil::unique_bstr output;
		hr = lp->ExecCommand(nullptr, data->cmdLine.get(), data->workDir.get(), flags, data->outputWindow,
			CAT_BUILDCOMPILE, BMP_COMPILE, L"", nullptr, &data->exitCode, &output);
		if (FAILED(hr))
		{
			data->exitCode = hr;
			PostMessageW (data->hwnd, WM_BUILD_COMPLETE, 0, reinterpret_cast<LPARAM>(data));
			return hr;
		}

		data->outputText = std::move(output);
		PostMessageW (data->hwnd, WM_BUILD_COMPLETE, 0, reinterpret_cast<LPARAM>(data));
		return 0;
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
						hr = outputWindow2->OutputTaskItemStringEx2(line, taskPrio, CAT_BUILDCOMPILE, L"", BMP_COMPILE,
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

				hr = _outputWindowPane->OutputString(line); RETURN_IF_FAILED(hr);
				line = p;
			}
		}

		return S_OK;
	}

	static LRESULT CALLBACK window_proc (HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
	{
		if (msg == WM_BUILD_COMPLETE)
		{
			auto p = reinterpret_cast<ProjectConfigBuilder*>(GetWindowLongPtr (hwnd, GWLP_USERDATA)); WI_ASSERT(p);
			auto* td = reinterpret_cast<ThreadData*>(lparam);
			wchar_t* outputText = td->outputText.get();
			wchar_t* outputTextEnd = outputText + SysStringLen(td->outputText.get());
			auto hr = p->ParseSjasmOutput(outputText, outputTextEnd); LOG_IF_FAILED(hr);
			DWORD exitCode = td->exitCode;
			p->_threads.remove([td](auto& u) { return u.get() == td; });
			auto callback = std::move(p->_callback);
			callback->OnBuildComplete(exitCode == 0);
			return 0;
		}

		return DefWindowProc (hwnd, msg, wparam, lparam);
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
