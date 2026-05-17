
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

	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (VxDTE::DTE2* dte, PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path);


	TEST_CLASS(DebuggerTests)
	{
		static inline wil::com_ptr_failfast<VxDTE::DTE2> dte;
		static inline wil::unique_process_heap_string TemplateProjectPath;
		static inline wil::unique_process_heap_string testClassPath;

		TEST_CLASS_INITIALIZE(ClassInit)
		{
			dte = GetDefaultVSInstance();

			testClassPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"DebuggerTests\\");
			Assert::IsTrue(CreateDirectory(testClassPath.get(), nullptr));

			auto templateDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testClassPath, L"TemplateProject\\");
			TemplateProjectPath = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"proj.flx");
			WriteFileCreateDirs(TemplateProjectPath, TemplateProjectXML);
			auto file = wil::str_concat_failfast<wil::unique_process_heap_string>(templateDir, L"file.asm");
			WriteFileCreateDirs(file, "start:\tret");
		}

		TEST_CLASS_CLEANUP(ClassCleanup)
		{
			dte.reset();
		}

		wil::unique_process_heap_string testDir;
		wil::unique_process_heap_string projPath;
		wil::com_ptr_failfast<VxDTE::_Solution> sln;
		wil::com_ptr_failfast<VxDTE::Project> proj;

		TEST_METHOD_INITIALIZE(MethodInit)
		{
			HRESULT hr;

			testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testClassPath, L"Test\\");
			Assert::IsTrue(CreateDirectory(testDir.get(), nullptr));

			wil::com_ptr_failfast<IUnknown> solution;
			hr = dte->get_Solution((VxDTE::Solution**)solution.addressof());
			Assert::AreEqual(S_OK, hr);
			sln = solution.query<VxDTE::_Solution>();
			hr = sln->Create(wil::make_bstr_failfast(testDir.get()).get(), wil::make_bstr_failfast(L"test").get());
			Assert::AreEqual(S_OK, hr);

			hr = sln->AddFromTemplate (
				wil::make_bstr_failfast(TemplateProjectPath.get()).get(),
				wil::make_bstr_failfast(testDir.get()).get(),
				wil::make_bstr_failfast(L"test.flx").get(), VARIANT_FALSE, &proj);
			Assert::AreEqual(S_OK, hr);
			hr = sln->SaveAs(wil::make_bstr_failfast(L"test").get());
			Assert::AreEqual(S_OK, hr);
		}

		TEST_METHOD_CLEANUP(MethodCleanup)
		{
			if (sln)
			{
//				sln->Close();
				sln.reset();
				proj.reset();
			}

			if (testDir)
			{
//				RemoveDirectoryTree(testDir);
				testDir.reset();
			}
		}

		TEST_METHOD(DebugBinaryWithBreakpoint)
		{
			HRESULT hr;

			wil::com_ptr_failfast<VxDTE::Debugger> debugger;
			hr = dte->get_Debugger(&debugger);

			wil::com_ptr_failfast<VxDTE::Breakpoints> breakpoints;
			hr = debugger->get_Breakpoints(&breakpoints);

			wil::com_ptr_failfast<VxDTE::Breakpoints> added;
			hr = breakpoints->Add (nullptr, wil::make_bstr_failfast(L"file.asm").get(), 1, 1, nullptr,
				VxDTE::dbgBreakpointConditionTypeWhenTrue, nullptr, nullptr, 0, nullptr, 0, VxDTE::dbgHitCountTypeNone, &added);
			Assert::AreEqual(S_OK, hr);

			hr = dte->ExecuteCommand(wil::make_bstr_failfast(L"Debug.Start").get());
			Assert::AreEqual(S_OK, hr);

			DWORD tickStart = GetTickCount();
			while(true)
			{
				VxDTE::vsIDEMode mode = (VxDTE::vsIDEMode)0;
				hr = dte->get_Mode(&mode); // This sometimes returns RPC_E_CALL_REJECTED
				if (mode == VxDTE::vsIDEMode::vsIDEModeDebug)
					break;
				Sleep(50);
				Assert::IsTrue(IsDebuggerPresent() || GetTickCount() - tickStart < 2000);
			}

			tickStart = GetTickCount();
			while(true)
			{
				VxDTE::dbgDebugMode mode = (VxDTE::dbgDebugMode)0;
				debugger->get_CurrentMode(&mode);
				if (mode == VxDTE::dbgDebugMode::dbgBreakMode)
					break;
				Sleep(50);
				Assert::IsTrue(IsDebuggerPresent() || GetTickCount() - tickStart < 2000);
			}

			hr = debugger->Stop();
			Assert::AreEqual(S_OK, hr);
		}
	};
}
