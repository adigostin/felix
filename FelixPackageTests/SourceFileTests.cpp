
#include "pch.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(FileTests)
	{
		TEST_METHOD(put_Items_EmptyProject)
		{
		}

		TEST_METHOD(put_Items_NonEmptyProject)
		{
		}

		TEST_METHOD(RenameFileInProjectDir)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			LPCOLESTR templateasm[] = { TemplatePath_EmptyFile.get() };
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"file.asm", 1, templateasm, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_variant id;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &id);
			Assert::IsTrue(SUCCEEDED(hr));

			wchar_t newFullPath[MAX_PATH];
			PathCombine(newFullPath, tempPath, L"new.asm");
			BOOL bres = DeleteFile(newFullPath);
			Assert::IsTrue(bres || GetLastError() == ERROR_FILE_NOT_FOUND);

			hr = hier->SetProperty(V_VSITEMID(&id), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsTrue(PathFileExists(newFullPath));
			DeleteFile(newFullPath);
		}

		TEST_METHOD(RenameFileInFolder)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			const wchar_t* relativePathOld = L"folder/test.asm";
			auto fullPathOld = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine(fullPathOld.get(), tempPath, relativePathOld);
			wchar_t folderFullPath[MAX_PATH];
			PathCombine(folderFullPath, tempPath, L"folder");
			BOOL bres = CreateDirectory(folderFullPath, nullptr);
			Assert::IsTrue(bres || GetLastError() == ERROR_ALREADY_EXISTS);
			WriteFileOnDisk(tempPath, relativePathOld);

			wchar_t fullPathNew[MAX_PATH];
			PathCombine(fullPathNew, tempPath, L"folder/new.asm");
			bres = DeleteFile(fullPathNew);
			Assert::IsTrue(bres || GetLastError() == ERROR_FILE_NOT_FOUND);

			VSADDRESULT addResult;
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fullPathOld.addressof(), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			auto firstChild = hier.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(firstChild);
			auto folder = wil::try_com_query_nothrow<IFolderNode>(firstChild);
			Assert::IsNotNull(folder.get());
			Assert::AreEqual(L"folder", GetProperty_String(hier, folder->GetItemId(), VSHPROPID_SaveName).get());

			firstChild = folder.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(firstChild);
			auto file = wil::try_com_query_nothrow<IFileNode>(firstChild);
			Assert::IsNotNull(file.get());

			// Now rename it.

			hr = hier->SetProperty(file->GetItemId(), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::AreEqual(L"new.asm", GetProperty_String(hier, file->GetItemId(), VSHPROPID_SaveName).get());

			Assert::IsTrue(PathFileExists(fullPathNew));
			DeleteFile(fullPathNew);
		}

		TEST_METHOD(RenameFileOutsideProjectDirSameDrive)
		{
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\RenameFileOutsideProjectDirSameDrive");
			CreateDirectory (testDir.get(), nullptr);
			auto projDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\projdir");
			CreateDirectory (projDir.get(), nullptr);

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (TemplatePath_EmptyProject.get(), projDir.get(), L"proj.flx", CPF_CLONEFILE, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\file.asm");
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fileFullPath.addressof(), nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			auto fileProps = wil::try_com_query_nothrow<IFileNodeProperties>(hier.try_query<IParentNode>()->FirstChild());
			wil::unique_bstr pathProp;
			fileProps->get_Path(&pathProp);
			Assert::AreEqual(L"..\\file.asm", pathProp.get());

			wil::unique_hfile handle (CreateFile(fileFullPath.get(), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			Assert::IsTrue(handle.is_valid());
			handle.reset();

			wil::unique_variant id;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &id);
			Assert::IsTrue(SUCCEEDED(hr));

			auto newFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\new.asm");
			BOOL bres = DeleteFile(newFullPath.get());
			Assert::IsTrue(bres || GetLastError() == ERROR_FILE_NOT_FOUND);

			hr = hier->SetProperty(V_VSITEMID(&id), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			fileProps->get_Path(&pathProp);
			Assert::AreEqual(L"..\\new.asm", pathProp.get());

			Assert::IsTrue(PathFileExists(newFullPath.get()));
			DeleteFile(newFullPath.get());
		}

		TEST_METHOD(RenameFileOutsideProjectDirOtherDrive)
		{
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\RenameFileOutsideProjectDirSameDrive");
			CreateDirectory(testDir.get(), nullptr);

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, testDir.get(), nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsTrue(SUCCEEDED(hr));
			const wchar_t* fileFullPath = L"D:\\FelixTest\\RenameFileOutsideProjectDirSameDrive\\file.asm";
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, &fileFullPath, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));
			wil::unique_hfile handle (CreateFile(fileFullPath, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			Assert::IsTrue(handle.is_valid());
			handle.reset();

			wil::unique_variant id;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &id);
			Assert::IsTrue(SUCCEEDED(hr));

			const wchar_t* newFullPath = L"D:\\FelixTest\\RenameFileOutsideProjectDirSameDrive\\new.asm";
			BOOL bres = DeleteFile(newFullPath);
			Assert::IsTrue(bres || GetLastError() == ERROR_FILE_NOT_FOUND);

			hr = hier->SetProperty(V_VSITEMID(&id), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::IsTrue(PathFileExists(newFullPath));
			DeleteFile(newFullPath);
		}

		TEST_METHOD(AddFileAlreadyExists)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			wchar_t dir[MAX_PATH];
			PathCombine(dir, tempPath, L"folder");
			CreateDirectory(dir, nullptr);

			auto fileFullPath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine(fileFullPath.get(), tempPath, L"folder/test.asm");

			WriteFileOnDisk(tempPath, L"folder/test.asm");

			VSADDRESULT addResult;
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fileFullPath.addressof(), nullptr, &addResult);
			Assert::IsTrue(SUCCEEDED(hr));

			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fileFullPath.addressof(), nullptr, &addResult);
			Assert::AreEqual(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), hr);

			DeleteFile(fileFullPath.get());
		}

		TEST_METHOD(GetMkDocument_FileInProjectDir)
		{
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\GetMkDocument_FileInProjectDir");

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, testDir.get(), nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto project = hier.try_query<IVsProject>();

			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\file.asm");
			hr = project->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fileFullPath.addressof(), nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, child.vt);
			wil::unique_bstr mk;
			hr = project->GetMkDocument(V_VSITEMID(&child), &mk);
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::AreEqual((PCWSTR)fileFullPath.get(), (PCWSTR)mk.get());
		}

		TEST_METHOD(GetMkDocument_FileNotInProjectDir_SameDrive)
		{
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\GetMkDocument_FileNotInProjectDir_SameDrive");
			auto projDir = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\subdir");

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, projDir.get(), nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto project = hier.try_query<IVsProject>();

			auto fileFullPath = wil::str_concat_failfast<wil::unique_process_heap_string>(testDir, L"\\file.asm");
			hr = project->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fileFullPath.addressof(), nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, child.vt);
			wil::unique_bstr mk;
			hr = project->GetMkDocument(V_VSITEMID(&child), &mk);
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::AreEqual((PCWSTR)fileFullPath.get(), (PCWSTR)mk.get());
		}

		TEST_METHOD(GetMkDocument_FileNotInProjectDir_OtherDrive)
		{
			auto testDir = wil::str_concat_failfast<wil::unique_process_heap_string>(tempPath, L"\\GetMkDocument_FileNotInProjectDir_OtherDrive");
			
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, testDir.get(), nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto project = hier.try_query<IVsProject>();

			const wchar_t* fileFullPath = L"D:\\FelixTest\\GetMkDocument_FileNotInProjectDir_OtherDrive\\file.asm";
			hr = project->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, &fileFullPath, nullptr, nullptr);
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant child;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_FirstChild, &child);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<VARTYPE>(VT_VSITEMID, child.vt);
			wil::unique_bstr mk;
			hr = project->GetMkDocument(V_VSITEMID(&child), &mk);
			Assert::IsTrue(SUCCEEDED(hr));

			Assert::AreEqual((PCWSTR)fileFullPath, (PCWSTR)mk.get());
		}
	};
}
