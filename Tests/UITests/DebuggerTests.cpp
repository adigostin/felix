
#include "pch.h"
#include "shared/com.h"
#include "UITests.h"
#include "dispids.h"
#include <string_view>
#include <vector>

namespace UITests
{
	static const char TemplateProjectXML[] = ""
		"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
		"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
		"  <Configurations>\r\n"
		"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\">\r\n"
		"      <GeneralProperties OutputFileType=\"Binary\" />\r\n"
		"    </Configuration>\r\n"
		"  </Configurations>\r\n"
		"  <Items>\r\n"
		"    <File Path=\"file.asm\" BuildTool=\"Assembler\" />\r\n"
		"  </Items>\r\n"
		"</Z80Project>\r\n";

	static const char EmptyTemplateProjectXML[] = ""
		"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
		"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">\r\n"
		"  <Configurations>\r\n"
		"    <Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\" />\r\n"
		"  </Configurations>\r\n"
		"</Z80Project>\r\n";

	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (VxDTE::DTE2* dte, PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path);

	struct DebuggerTD : TD
	{
		wil::com_ptr_failfast<VxDTE::Debugger> debugger;

		DebuggerTD(PCWSTR testClassPath, PCWSTR testName, PCWSTR projectTemplatePath)
			: TD(testClassPath, testName, projectTemplatePath)
		{
			auto hr = dte->get_Debugger(&debugger);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			long breakpointCount;
			hr = breakpoints->get_Count(&breakpointCount);
			Assert::AreEqual(S_OK, hr);
			while (breakpointCount > 0)
			{
				VARIANT index;
				VariantInit(&index);
				V_VT(&index) = VT_I4;
				V_I4(&index) = breakpointCount;
				wil::com_ptr_failfast<VxDTE::Breakpoint> breakpoint;
				hr = breakpoints->Item(index, &breakpoint);
				Assert::AreEqual(S_OK, hr);
				hr = breakpoint->Delete();
				Assert::AreEqual(S_OK, hr);
				--breakpointCount;
			}

			VxDTE::dbgDebugMode mode;
			hr = debugger->get_CurrentMode(&mode);
			Assert::AreEqual(S_OK, hr);
			if (mode != VxDTE::dbgDesignMode)
			{
				hr = debugger->Stop();
				Assert::AreEqual(S_OK, hr);
				HRESULT lastHr = S_OK;
				VxDTE::dbgDebugMode currentMode = (VxDTE::dbgDebugMode)0;
				bool reachedMode = WaitWithMessageLoop([&]
					{
						lastHr = debugger->get_CurrentMode(&currentMode);
						Assert::IsTrue(SUCCEEDED(lastHr) || lastHr == RPC_E_CALL_REJECTED,
							str_printf(L"Reading the debugger mode failed: 0x%08x", lastHr).get());
						return SUCCEEDED(lastHr) && currentMode == VxDTE::dbgDesignMode;
					}, 5000);
				Assert::IsTrue(reachedMode,
					str_printf(L"Timed out waiting for debugger mode %u (last HRESULT 0x%08x, mode %u)",
						(unsigned)VxDTE::dbgDesignMode, lastHr, (unsigned)currentMode).get());
			}
		}

		~DebuggerTD()
		{
			if (debugger)
			{
				VxDTE::dbgDebugMode mode;
				if (SUCCEEDED(debugger->get_CurrentMode(&mode)) && mode != VxDTE::dbgDesignMode)
					debugger->Stop();
			}
		}
	};

	TEST_CLASS(DebuggerTests)
	{
		static inline wil::unique_process_heap_string TemplateProjectPath;
		static inline wil::unique_process_heap_string EmptyTemplateProjectPath;
		static inline wil::unique_process_heap_string testClassPath;

		TEST_CLASS_INITIALIZE(ClassInit)
		{
			HRESULT hr;

			testClassPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"DebuggerTests\\");
			if (PathFileExists(testClassPath.get()))
				ClearDirectoryContents(testClassPath.get());
			else
				Assert::IsTrue(CreateDirectory(testClassPath.get(), nullptr));

			auto templateDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testClassPath, L"TemplateProject\\");
			TemplateProjectPath = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"proj.flx");
			WriteFileCreateDirs(TemplateProjectPath, TemplateProjectXML);
			auto file = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"file.asm");
			WriteFileCreateDirs(file, "start:\tret");

			auto emptyTemplateDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testClassPath, L"EmptyTemplateProject\\");
			EmptyTemplateProjectPath = wil::str_concat_failfast<wil::unique_process_heap_string>(emptyTemplateDir, L"proj.flx");
			WriteFileCreateDirs(EmptyTemplateProjectPath, EmptyTemplateProjectXML);
		}

		TEST_CLASS_CLEANUP(ClassCleanup)
		{
			if (testClassPath)
			{
				TryRemoveDirectoryTree(testClassPath.get());
				testClassPath.reset();
			}
		}

		static void WaitIDEDebugMode(VxDTE::DTE2* dte, DWORD timeoutMillis = 5000)
		{
			HRESULT lastHr = S_OK;
			VxDTE::vsIDEMode currentMode = (VxDTE::vsIDEMode)0;
			bool reachedMode = WaitWithMessageLoop([&]
			{
				lastHr = dte->get_Mode(&currentMode);
				Assert::IsTrue(SUCCEEDED(lastHr) || lastHr == RPC_E_CALL_REJECTED,
					str_printf(L"Reading the IDE mode failed: 0x%08x", lastHr).get());
				return SUCCEEDED(lastHr) && currentMode == VxDTE::vsIDEMode::vsIDEModeDebug;
			}, IsDebuggerPresent() ? INFINITE : timeoutMillis);
			Assert::IsTrue(reachedMode,
				str_printf(L"Timed out waiting for vsIDEModeDebug (last HRESULT 0x%08x, mode %u)", lastHr, (unsigned)currentMode).get());
		}

		static wil::com_ptr_failfast<IProjectConfigAssemblerProperties> GetAsmProps(TD& testData)
		{
			wil::com_ptr_failfast<IVsCfg> cfg;
			ULONG actual;
			VSCFGFLAGS flags;
			auto hr = testData.proj.query<IVsCfgProvider>()->GetCfgs(1, cfg.addressof(), &actual, &flags);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<IProjectConfigAssemblerProperties> asmProps;
			hr = cfg.query<IProjectConfigProperties>()->get_AssemblerProperties(&asmProps);
			Assert::AreEqual(S_OK, hr);

			return asmProps;
		}

		TEST_METHOD(DebugBinaryWithBreakpoint)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"DebugBinaryWithBreakpoint", TemplateProjectPath.get());

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = td.debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add (nullptr, wil::make_bstr_failfast(L"file.asm").get(), 1, 1, nullptr,
				VxDTE::dbgBreakpointConditionTypeWhenTrue, nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(added.get());

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode (td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr);
			WaitDebugMode (td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}

		TEST_METHOD(DebugBinary_EntryPointIsNumber)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"DebugBinary_EntryPointIsNumber", TemplateProjectPath.get());

			auto asmProps = GetAsmProps(td);

			DWORD baseAddress;
			hr = asmProps->get_BaseAddress(&baseAddress);
			Assert::AreEqual(S_OK, hr);

			wchar_t buffer[10];
			swprintf_s(buffer, L"0x%x", baseAddress);
			hr = asmProps->put_EntryPointAddress(wil::make_bstr_failfast(buffer).get());
			Assert::AreEqual(S_OK, hr);

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.StepInto").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode (td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr);
			WaitDebugMode (td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}

		TEST_METHOD(BreakpointInSourceFileInProjectSubdir)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"BreakpointInSourceFileInProjectSubdir", EmptyTemplateProjectPath.get());

			auto sourceFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\subdir\\file.asm");
			WriteFileCreateDirs(sourceFilePath, "start:\tret");
			LPCOLESTR files[] = { sourceFilePath.get() };
			VSADDRESULT addResult;
			auto operation = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, operation, L"", 1, files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			td.proj->Save();

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = td.debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			auto filePathFromSLD = wil::make_bstr_failfast(L"subdir\\file.asm");
			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add(nullptr, filePathFromSLD.get(), 1, 1, nullptr, VxDTE::dbgBreakpointConditionTypeWhenTrue,
				nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(added.get());

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}

		TEST_METHOD(BreakpointInSourceFileOutsideProjectDir)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"BreakpointInSourceFileOutsideProjectDir", EmptyTemplateProjectPath.get());

			// We put it outside the solution dir because, if our debug engine can't bind it,
			// VS will attempt to bind it with all kinds of "smart" paths under the solution dir.
			auto sourceFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(testClassPath, L"file.asm");
			WriteFileCreateDirs(sourceFilePath, "start:\tret");
			LPCOLESTR files[] = { sourceFilePath.get() };
			VSADDRESULT addResult;
			auto operation = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, operation, L"", 1, files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			td.proj->Save();

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = td.debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			auto filePathFromSLD = wil::make_bstr_failfast(sourceFilePath.get());
			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add(nullptr, filePathFromSLD.get(), 1, 1, nullptr, VxDTE::dbgBreakpointConditionTypeWhenTrue,
				nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(added.get());

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}

		TEST_METHOD(BreakpointInSourceFileOnOtherDrive)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"BreakpointInSourceFileOnOtherDrive", EmptyTemplateProjectPath.get());

			// We put it outside the solution dir because, if our debug engine can't bind it,
			// VS will attempt to bind it with all kinds of "smart" paths under the solution dir.
			BOOL bres = DefineDosDeviceW (DDD_NO_BROADCAST_SYSTEM, L"Q:", testClassPath.get());
			Assert::IsTrue(bres);
			auto cleanup = wil::scope_exit([&] {
				DefineDosDeviceW (DDD_REMOVE_DEFINITION | DDD_EXACT_MATCH_ON_REMOVE | DDD_NO_BROADCAST_SYSTEM, L"Q:", testClassPath.get());
			});

			const wchar_t* sourceFilePath = L"Q:\\file.asm";
			WriteFileCreateDirs(sourceFilePath, "start:\tret");
			LPCOLESTR files[] = { sourceFilePath };
			VSADDRESULT addResult;
			auto operation = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, operation, L"", 1, files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			td.proj->Save();

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = td.debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			auto filePathFromSLD = wil::make_bstr_failfast(sourceFilePath);
			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add(nullptr, filePathFromSLD.get(), 1, 1, nullptr, VxDTE::dbgBreakpointConditionTypeWhenTrue,
				nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(added.get());

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr); // it sometimes fails on here, don't know why
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}

		struct DisassemblyLines
		{
			wil::unique_bstr text;
			std::vector<std::wstring_view> lines;
		};

		static DisassemblyLines GetDisassembly(VxDTE::DTE2* dte)
		{
			HRESULT hr;

			hr = dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Disassembly").get());
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Windows> windows;
			hr = dte->get_Windows(&windows);
			Assert::AreEqual(S_OK, hr);

			auto disassemblyGuid = wil::make_bstr_failfast(L"{CF577B8C-4134-11D2-83E5-00C04F9902C1}");
			VARIANT index{};
			index.vt = VT_BSTR;
			index.bstrVal = disassemblyGuid.get();
			wil::com_ptr_failfast<VxDTE::Window> disassemblyWindow;
			hr = windows->Item(index, &disassemblyWindow);
			Assert::AreEqual(S_OK, hr);
			hr = disassemblyWindow->Activate();
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<IUIAutomation> uiAutomation;
			hr = CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&uiAutomation));
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Window> mainWindow;
			hr = dte->get_MainWindow(&mainWindow);
			Assert::AreEqual(S_OK, hr);
			long mainWindowHWnd;
			hr = mainWindow->get_HWnd(&mainWindowHWnd);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<IUIAutomationElement> mainWindowElement;
			hr = uiAutomation->ElementFromHandle((HWND)(LONG_PTR)mainWindowHWnd, &mainWindowElement);
			Assert::AreEqual(S_OK, hr);
			wil::unique_bstr disassemblyCaption;
			hr = disassemblyWindow->get_Caption(disassemblyCaption.put());
			Assert::AreEqual(S_OK, hr);
			wil::com_ptr_failfast<IUIAutomationCondition> disassemblyCondition;
			hr = uiAutomation->CreatePropertyCondition(UIA_NamePropertyId, wil::make_variant_bstr_failfast(disassemblyCaption.get()), &disassemblyCondition);
			Assert::AreEqual(S_OK, hr);
			wil::com_ptr_failfast<IUIAutomationElement> disassemblyElement;
			hr = mainWindowElement->FindFirst(TreeScope_Descendants, disassemblyCondition.get(), &disassemblyElement);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(disassemblyElement.get());

			wil::com_ptr_failfast<IUIAutomationCondition> viewingOptionsCondition;
			hr = uiAutomation->CreatePropertyCondition(UIA_NamePropertyId, wil::make_variant_bstr_failfast(L"Viewing Options"), &viewingOptionsCondition);
			Assert::AreEqual(S_OK, hr);
			wil::com_ptr_failfast<IUIAutomationElement> viewingOptionsElement;
			hr = disassemblyElement->FindFirst(TreeScope_Descendants, viewingOptionsCondition.get(), &viewingOptionsElement);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(viewingOptionsElement.get());

			wil::com_ptr_failfast<IUnknown> viewingOptionsPatternUnknown;
			hr = viewingOptionsElement->GetCurrentPattern(UIA_ExpandCollapsePatternId, &viewingOptionsPatternUnknown);
			Assert::AreEqual(S_OK, hr);
			auto viewingOptionsPattern = viewingOptionsPatternUnknown.query<IUIAutomationExpandCollapsePattern>();
			ExpandCollapseState viewingOptionsState;
			hr = viewingOptionsPattern->get_CurrentExpandCollapseState(&viewingOptionsState);
			Assert::AreEqual(S_OK, hr);
			if (viewingOptionsState == ExpandCollapseState_Collapsed)
			{
				hr = viewingOptionsPattern->Expand();
				Assert::AreEqual(S_OK, hr);
			}

			wil::com_ptr_failfast<IUIAutomationCondition> showSourceCodeCondition;
			hr = uiAutomation->CreatePropertyCondition(UIA_NamePropertyId, wil::make_variant_bstr_failfast(L"Show source code"), &showSourceCodeCondition);
			Assert::AreEqual(S_OK, hr);
			wil::com_ptr_failfast<IUIAutomationElement> showSourceCodeElement;
			hr = disassemblyElement->FindFirst(TreeScope_Descendants, showSourceCodeCondition.get(), &showSourceCodeElement);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(showSourceCodeElement.get());

			wil::com_ptr_failfast<IUnknown> showSourceCodePatternUnknown;
			hr = showSourceCodeElement->GetCurrentPattern(UIA_TogglePatternId, &showSourceCodePatternUnknown);
			Assert::AreEqual(S_OK, hr);
			auto showSourceCodePattern = showSourceCodePatternUnknown.query<IUIAutomationTogglePattern>();
			ToggleState showSourceCodeState;
			hr = showSourceCodePattern->get_CurrentToggleState(&showSourceCodeState);
			Assert::AreEqual(S_OK, hr);
			if (showSourceCodeState != ToggleState_On)
			{
				hr = showSourceCodePattern->Toggle();
				Assert::AreEqual(S_OK, hr);
			}
			hr = showSourceCodePattern->get_CurrentToggleState(&showSourceCodeState);
			Assert::AreEqual(S_OK, hr);
			Assert::AreEqual((int)ToggleState_On, (int)showSourceCodeState);

			VARIANT textPatternAvailable{};
			textPatternAvailable.vt = VT_BOOL;
			textPatternAvailable.boolVal = VARIANT_TRUE;
			wil::com_ptr_failfast<IUIAutomationCondition> textPatternCondition;
			hr = uiAutomation->CreatePropertyCondition(UIA_IsTextPatternAvailablePropertyId, textPatternAvailable, &textPatternCondition);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<IUIAutomationElement> textElement;
			hr = disassemblyElement->FindFirst(TreeScope_Descendants, textPatternCondition.get(), &textElement);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(textElement.get());

			wil::com_ptr_failfast<IUnknown> textPatternUnknown;
			hr = textElement->GetCurrentPattern(UIA_TextPatternId, &textPatternUnknown);
			Assert::AreEqual(S_OK, hr);
			auto textPattern = textPatternUnknown.query<IUIAutomationTextPattern>();

			wil::com_ptr_failfast<IUIAutomationTextRange> textRange;
			hr = textPattern->get_DocumentRange(&textRange);
			Assert::AreEqual(S_OK, hr);
			wil::unique_bstr disassemblyText;
			hr = textRange->GetText(-1, disassemblyText.put());
			Assert::AreEqual(S_OK, hr);
			//return wil::make_process_heap_string_failfast(disassemblyText.get(), SysStringLen(disassemblyText.get()));

			DisassemblyLines lines = { std::move(disassemblyText) };

			const wchar_t* lineStart = lines.text.get();
			while (true)
			{
				const wchar_t* p = lineStart;
				while (*p != L'\n' && *p != L'\0')
					p++;
				auto line = std::wstring_view(lineStart, p - lineStart);
				if (!line.empty() && line.back() == L'\r')
					line.remove_suffix(1);
				lines.lines.push_back(line);
				if (*p == L'\0')
					break;
				lineStart = p + 1;
			}

			return lines;
		}

		TEST_METHOD(DisassemblyIncludesSource)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"DisassemblyIncludesSource", EmptyTemplateProjectPath.get());

			auto sourceFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\file.asm");
			static const char source[] =
				"start:\r\n"
				"\tld a, 1\r\n"
				"; disassembly comment\r\n"
				"\tinc a\r\n"
				"\r\n"
				"\tld b, 2\r\n"
				"\tinc b\r\n"
				"\tnop\r\n"
				"\tdec a\r\n"
				"\tret\r\n";
			WriteFileCreateDirs(sourceFilePath, source);
			LPCOLESTR files[] = { sourceFilePath.get() };
			VSADDRESULT addResult;
			auto operation = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, operation, L"", 1, files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			td.proj->Save();

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = td.debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add(nullptr, wil::make_bstr_failfast(L"file.asm").get(), 1, 1, nullptr,
				VxDTE::dbgBreakpointConditionTypeWhenTrue, nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(added.get());

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			auto disasm = GetDisassembly(td.dte);

			auto assertAdjacentLines = [&disasm](const wchar_t* target, const wchar_t* previous, const wchar_t* next)
			{
				size_t lineIndex = 0;
				while (lineIndex < disasm.lines.size() && disasm.lines[lineIndex].find(target) == std::wstring_view::npos)
					++lineIndex;
				Assert::IsTrue(lineIndex > 0);
				Assert::IsTrue(lineIndex + 1 < disasm.lines.size());
				Assert::IsTrue(disasm.lines[lineIndex - 1].find(previous) != std::wstring_view::npos);
				Assert::IsTrue(disasm.lines[lineIndex + 1].find(next) != std::wstring_view::npos);
			};
			assertAdjacentLines(L"ld a, 1", L"start:", L"LD A, 1");
			assertAdjacentLines(L"inc a", L"; disassembly comment", L"INC A");

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}

		TEST_METHOD(DisassemblyCorrectOnFileTransitionInSLD)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"DisassemblyIncludesCommentsAtStartOfSourceFile", EmptyTemplateProjectPath.get());

			auto libFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\lib.asm");
			WriteFileCreateDirs(libFilePath, "e0:\tret\r\n");
			auto startFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\start.asm");
			static const char source[] =
				"; start comment 1\r\n"
				"; start comment 2\r\n"
				"; start comment 3\r\n"
				"start:\r\n"
				"\tcall e0\r\n"
				"\tret\r\n";
			WriteFileCreateDirs(startFilePath, source);

			LPCOLESTR files[] = { libFilePath.get(), startFilePath.get() };
			VSADDRESULT addResult;
			auto operation = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, operation, L"", 2, files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			td.proj->Save();

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = td.debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add(nullptr, wil::make_bstr_failfast(L"start.asm").get(), 5, 1, nullptr,
				VxDTE::dbgBreakpointConditionTypeWhenTrue, nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(added.get());

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			auto disasm = GetDisassembly(td.dte);
			size_t i = 0;
			while (i < disasm.lines.size() && disasm.lines[i].find(L"; start comment 1") == std::wstring_view::npos)
				++i;
			Assert::IsTrue(i < disasm.lines.size());
			++i;
			while (i < disasm.lines.size() && disasm.lines[i].find(L"; start comment 2") == std::wstring_view::npos)
				++i;
			Assert::IsTrue(i < disasm.lines.size());
			++i;
			while (i < disasm.lines.size() && disasm.lines[i].find(L"; start comment 3") == std::wstring_view::npos)
				++i;
			Assert::IsTrue(i < disasm.lines.size());
			++i;
			while (i < disasm.lines.size() && disasm.lines[i].find(L"start:") == std::wstring_view::npos)
				++i;
			Assert::IsTrue(i < disasm.lines.size());

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}

		TEST_METHOD(DisassemblyIncludesSourceForRepeatedBlock)
		{
			HRESULT hr;
			DebuggerTD td(testClassPath.get(), L"DisassemblyIncludesSourceForRepeatedBlock", EmptyTemplateProjectPath.get());

			auto sourceFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(td.projDir, L"\\repeat.asm");
			static const char source[] =
				"start:\r\n"
				"\tDUP 2\r\n"
				"\tld a, 1\r\n"
				"\tinc a\r\n"
				"\tld h, d : ld l, e\r\n"
				"\tEDUP\r\n"
				"\tret\r\n";
			WriteFileCreateDirs(sourceFilePath, source);
			LPCOLESTR files[] = { sourceFilePath.get() };
			VSADDRESULT addResult;
			auto operation = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, operation, L"", 1, files, nullptr, &addResult);
			Assert::AreEqual(S_OK, hr);
			hr = td.proj->Save();
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = td.debugger->get_Breakpoints(&breakpoints);
			Assert::AreEqual(S_OK, hr);

			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add(nullptr, wil::make_bstr_failfast(L"repeat.asm").get(), 3, 1, nullptr,
				VxDTE::dbgBreakpointConditionTypeWhenTrue, nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);
			Assert::IsNotNull(added.get());

			hr = td.dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			WaitIDEDebugMode(td.dte);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgBreakMode);

			auto disasm = GetDisassembly(td.dte);
			size_t lineIndex = 0;
			for (unsigned iteration = 0; iteration < 2; iteration++)
			{
				for (auto [opcode, previousLine] : {
					std::pair { L"LD A, 1", L"ld a, 1" },
					std::pair { L"INC A", L"inc a" },
					std::pair { L"LD H, D", L"ld h, d : ld l, e" },
					std::pair { L"LD L, E", L"LD H, D" } })
				{
					while (lineIndex < disasm.lines.size() && disasm.lines[lineIndex].find(opcode) == std::wstring_view::npos)
						++lineIndex;
					Assert::IsTrue(lineIndex < disasm.lines.size(),
						str_printf(L"Missing %s in repetition %u.\r\n%s", opcode, iteration + 1, disasm.text.get()).get());
					Assert::IsTrue(lineIndex > 0);
					Assert::IsTrue(disasm.lines[lineIndex - 1].find(previousLine) != std::wstring_view::npos,
						str_printf(L"Missing %s before %s in repetition %u.\r\n%s", previousLine, opcode, iteration + 1, disasm.text.get()).get());
					++lineIndex;
				}
			}

			hr = td.debugger->Stop();
			Assert::AreEqual(S_OK, hr);
			WaitDebugMode(td.debugger, VxDTE::dbgDebugMode::dbgDesignMode);
		}
	};
}
