
#include "pch.h"
#include "shared/com.h"
#include "UITests.h"
#include "dispids.h"

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
		}

		TEST_METHOD(BreakpointInSourceFileOnOtherDrive)
		{
		}
	};
}
