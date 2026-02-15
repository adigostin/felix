
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

	TEST_CLASS(FileTests)
	{
		wil::unique_process_heap_string testPath;
		wil::unique_process_heap_string projPath;
		wil::com_ptr_failfast<VxDTE::_Solution> sln;
		wil::com_ptr_failfast<VxDTE::Project> proj;

		TEST_METHOD_INITIALIZE(FileTestInit)
		{
			testPath = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"FileTests");
			Assert::IsTrue(CreateDirectory(testPath.get(), nullptr));
			std::tie(sln, proj) = CreateSolutionAndProject(testPath.get(), L"test", L"proj");
			projPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\proj");
		}

		TEST_METHOD_CLEANUP(FileTestCleanup)
		{
			if (sln)
			{
				sln->Close();
				sln.reset();
				proj.reset();
			}

			if (testPath)
			{
				RemoveDirectoryTree(testPath.get());
				testPath.reset();
			}
		}

		TEST_METHOD(put_Items_EmptyProject)
		{
		}

		TEST_METHOD(put_Items_NonEmptyProject)
		{
		}

		TEST_METHOD(RenameFileInProjectDir)
		{
			VSITEMID id;
			auto hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &id);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj.query<IVsHierarchy>()->SetProperty(id, VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			auto newFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\new.asm");
			Assert::IsTrue(PathFileExists(newFullPath.get()));
		}
		
		TEST_METHOD(RenameFileInFolder)
		{
			HRESULT hr;
			auto fullPathOld = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\folder/test.asm");
			WriteFileOnDisk(fullPathOld.get(), nullptr);

			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000u);
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fullPathOld.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			VSITEMID fileItemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"folder\\test.asm", &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));

			// Now rename it.
			hr = proj.query<IVsHierarchy>()->SetProperty(fileItemId, VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant newName;
			proj.query<IVsHierarchy>()->GetProperty(fileItemId, VSHPROPID_SaveName, &newName);
			Assert::AreEqual(L"new.asm", newName.bstrVal);

			auto fullPathNew = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\folder\\new.asm");
			Assert::IsTrue(PathFileExists(fullPathNew.get()));
		}
		
		TEST_METHOD(RenameFileOutsideProjectDirSameDrive)
		{
			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\file.asm");
			VSADDRESULT addResult;
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000u);
			auto hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fileFullPath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
			VSITEMID fileItemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"..\\file.asm", &fileItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant fileVar;
			hr = proj.query<IVsHierarchy>()->GetProperty(fileItemId, VSHPROPID_BrowseObject, &fileVar);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<int>(VT_DISPATCH, fileVar.vt);
			auto fileProps = wil::com_query_failfast<IFileNodeProperties>(fileVar.pdispVal);
			wil::unique_bstr pathProp;
			fileProps->get_Path(&pathProp);
			Assert::AreEqual(L"..\\file.asm", pathProp.get());

			wil::unique_hfile handle (CreateFile(fileFullPath.get(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			Assert::IsTrue(handle.is_valid());
			handle.reset();

			hr = proj.query<IVsHierarchy>()->SetProperty(fileItemId, VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));
			fileProps->get_Path(&pathProp);
			Assert::AreEqual(L"..\\new.asm", pathProp.get());

			auto newFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\new.asm");
			Assert::IsTrue(PathFileExists(newFullPath.get()));
		}
		
		TEST_METHOD(RenameFileOutsideProjectDirOtherDrive)
		{
			HRESULT hr;
			auto testPathOtherDrive = MakeVolumeGuidPath(testPath.get());
			Assert::IsTrue(PathFileExists(testPathOtherDrive.get()));
			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPathOtherDrive, L"\\file.asm");
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000u);
			VSADDRESULT addResult;
			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fileFullPath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));
			WriteFileOnDisk(fileFullPath.get(), nullptr);

			VSITEMID id;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(fileFullPath.get(), &id);
			Assert::IsTrue(SUCCEEDED(hr));

			auto newFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPathOtherDrive, L"\\new.asm");
			Assert::IsFalse(PathFileExists(newFullPath.get()));
			hr = proj.query<IVsHierarchy>()->SetProperty(id, VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(PathFileExists(newFullPath.get()));
		}
		
		TEST_METHOD(AddFileAlreadyExists)
		{
			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\folder/test.asm");
			WriteFileOnDisk(fileFullPath.get(), "");

			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000u);
			VSADDRESULT addResult;
			auto hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fileFullPath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fileFullPath.addressof()), nullptr, &addResult);
			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), hr);
		}
		
		TEST_METHOD(GetMkDocument_FileInProjectDir)
		{
			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(projPath, L"\\file.asm");
			VSITEMID childItemId;
			auto hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &childItemId);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_bstr mk;
			hr = proj.query<IVsProject>()->GetMkDocument(childItemId, &mk);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual((PCWSTR)fileFullPath.get(), (PCWSTR)mk.get());
		}
		
		TEST_METHOD(GetMkDocument_FileNotInProjectDir_SameDrive)
		{
			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testPath, L"\\file.asm");
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000u);
			VSADDRESULT addResult;
			auto hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fileFullPath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			VSITEMID itemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"..\\file.asm", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_bstr mk;
			hr = proj.query<IVsProject>()->GetMkDocument(itemId, &mk);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual((PCWSTR)fileFullPath.get(), (PCWSTR)mk.get());
		}
		
		TEST_METHOD(GetMkDocument_FileNotInProjectDir_OtherDrive)
		{
			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(MakeVolumeGuidPath(testPath.get()), L"\\file.asm");
			auto oper = (VSADDITEMOPERATION)(VSADDITEMOP_OPENFILE | 0x1000u);
			VSADDRESULT addResult;
			auto hr = proj.query<IVsProject>()->AddItem(VSITEMID_ROOT, oper, L"", 1, const_cast<LPCOLESTR*>(fileFullPath.addressof()), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			VSITEMID itemId;
			hr = proj.query<IVsHierarchy>()->ParseCanonicalName(fileFullPath.get(), &itemId);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_bstr mk;
			hr = proj.query<IVsProject>()->GetMkDocument(itemId, &mk);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual((PCWSTR)fileFullPath.get(), (PCWSTR)mk.get());
		}
		
		TEST_METHOD(NotifyPropertyChanged_BuildToolKind)
		{
			VSITEMID itemId;
			auto hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant fileBO;
			hr = proj.query<IVsHierarchy>()->GetProperty(itemId, VSHPROPID_BrowseObject, &fileBO);
			Assert::IsTrue(SUCCEEDED(hr));

			auto sink = MakeTestPropertyNotifySink();
			AdviseSinkToken token;
			hr = AdviseSink<IPropertyNotifySink>(fileBO.pdispVal, sink, &token);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = wil::com_query_failfast<IFileNodeProperties>(fileBO.pdispVal)->put_BuildTool(BuildToolKind::CustomBuildTool);
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsTrue(sink->Called({ dispidBuildToolKind }));
		}
		
		TEST_METHOD(ProjectDirtyOnFilePropertyChange)
		{
			VSITEMID itemId;
			auto hr = proj.query<IVsHierarchy>()->ParseCanonicalName(L"file.asm", &itemId);
			Assert::IsTrue(SUCCEEDED(hr));

			auto pff = proj.query<IPersistFileFormat>();

			hr = pff->Save(L"", TRUE, 0);
			Assert::IsTrue(SUCCEEDED(hr));

			BOOL dirty;
			hr = pff->IsDirty(&dirty); Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsFalse(dirty);

			wil::unique_variant fileBO;
			hr = proj.query<IVsHierarchy>()->GetProperty(itemId, VSHPROPID_BrowseObject, &fileBO);
			Assert::IsTrue(SUCCEEDED(hr));

			auto fileProps = wil::com_query<IFileNodeProperties>(fileBO.pdispVal);
			fileProps->put_BuildTool(BuildToolKind::None);

			hr = pff->IsDirty(&dirty); Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(dirty);
		}
	};
}
