
#include "pch.h"
#include "shared/com.h"
#include "../TestsCommon.h"
#include "dispids.h"
#include "FelixPackage_h.h"

#define FORCE_EXPLICIT_DTE_NAMESPACE
#include <dte.h>
namespace VxDTE
{
	#include <dte80.h>
	#include <dte90.h>
}

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace UITests
{
	extern std::pair<wil::com_ptr_failfast<VxDTE::_Solution>, wil::com_ptr_failfast<VxDTE::Project>>
		CreateSolutionAndProject (PCWSTR testDir, PCWSTR solutionName, PCWSTR projectName);
	extern wil::unique_process_heap_string MakeVolumeGuidPath (const wchar_t* path);

	TEST_CLASS(FolderTests)
	{
		struct TD
		{
			wil::unique_process_heap_string testPath;
			wil::unique_process_heap_string slnFilePath;
			wil::unique_process_heap_string projPath;
			wil::unique_process_heap_string projFilePath;
			wil::com_ptr_failfast<VxDTE::_Solution> sln;
			wil::com_ptr_failfast<VxDTE::Project> proj;

			TD()
			{
				testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"FolderTest");
				Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
				std::tie(sln, proj) = CreateSolutionAndProject(testPath.get(), L"test", L"proj");
				slnFilePath = CombinePath(testPath.get(), L"test.sln");
				projPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj");
				projFilePath = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\proj.flx");
			}

			~TD()
			{
				//sln->Close();
				//RemoveDirectoryTree(testPath.get());
			}
		};

		TEST_METHOD(CreateFolderNameExists)
		{
			TD td;
			HRESULT hr;

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_CLONEFILE | 0x1000);
			hr = td.proj.query<IVsProject>()->AddItem (VSITEMID_ROOT, oper, L"Name", 1, const_cast<LPCOLESTR*>(TemplatePath_EmptyFile.addressof()), NULL, &addResult);
			Assert::AreEqual(S_OK, hr);

			VSITEMID fileItemId;
			hr = td.proj.query<IVsHierarchy>()->ParseCanonicalName(L"Name", &fileItemId);
			Assert::AreEqual(S_OK, hr);
			wil::unique_bstr fileMk;
			hr = td.proj.query<IVsProject>()->GetMkDocument(fileItemId, &fileMk);
			Assert::AreEqual(S_OK, hr);

			VSITEMID nextVSItemID;
			hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextVSItemID);
			Assert::AreEqual(S_OK, hr);

			auto tryAddFolder = [&td, nextVSItemID]
				{
					// Try to add a folder with the same name as the existing file, with same casing and then different casing.
					// We should get ERROR_ALREADY_EXISTS, and the hierarchy should not be changed.
					wil::unique_variant folderItemId;
					auto hr = td.proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
						OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"Name").addressof(), &folderItemId);
					Assert::IsFalse(SUCCEEDED(hr));
					VSITEMID nextNow;
					hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextNow);
					Assert::AreEqual(nextVSItemID, nextNow);

					// Try to add it again, with the same name and different casing. We should get ERROR_ALREADY_EXISTS.
					hr = td.proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
						OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"NAME").addressof(), &folderItemId);
					Assert::IsFalse(SUCCEEDED(hr));
					hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextNow);
					Assert::AreEqual(nextVSItemID, nextNow);

					// Try to add it again, with the same name and different casing. We should get ERROR_ALREADY_EXISTS.
					hr = td.proj.query<IVsUIHierarchy>()->ExecCommand(VSITEMID_ROOT, &CMDSETID_StandardCommandSet97, cmdidNewFolder,
						OLECMDEXECOPT_DONTPROMPTUSER, wil::make_variant_bstr_failfast(L"name").addressof(), &folderItemId);
					Assert::IsFalse(SUCCEEDED(hr));
					hr = td.proj.query<IProjectNodeProperties>()->get_NextVSItemIdDebug(&nextNow);
					Assert::AreEqual(nextVSItemID, nextNow);
				};

			tryAddFolder();

			// Now delete the file from disk (keep it in the hierarchy), and try again to add the folder.
			BOOL bres = DeleteFile(fileMk.get());
			Assert::IsTrue(bres);
			tryAddFolder();
			RemoveDirectory(fileMk.get()); //remove any leftover directory (we're ok with the leftover for this test)

			// Now put the file back on disk and remove it from the hierarchy, and try again to add the folder.
			WriteFileOnDisk(fileMk.get(), "");
			hr = td.proj.query<IVsHierarchyDeleteHandler3>()->DeleteItems(1, DELITEMOP_RemoveFromProject, &fileItemId, DHO_SUPPRESS_UI);
			Assert::AreEqual(S_OK, hr);
			tryAddFolder();

			// TODO: do the same with a directory in a directory. With the outer directory present and then missing.
		}
	};
}
