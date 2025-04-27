
#include "pch.h"
#include "Mocks.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(SourceFileTests)
	{
		TEST_METHOD(SourceFileUnnamedTest)
		{
			HRESULT hr;

			com_ptr<IVsHierarchy> hier;
			hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));
			auto config = MakeProjectConfig(hier);
			auto pane = MakeMockOutputWindowPane(nullptr);

			com_ptr<IFileNode> file;
			hr = MakeFileNode(&file);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = AddFileToParent(file, hier.try_query<IParentNode>());
			Assert::IsTrue(SUCCEEDED(hr));

			wil::unique_variant value;
			hr = hier->GetProperty(1000, VSHPROPID_SaveName, &value);
			Assert::IsTrue(FAILED(hr));
			hr = hier->GetProperty(1000, VSHPROPID_Caption, &value);
			Assert::IsTrue(FAILED(hr));
		}

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

			com_ptr<IFileNode> file;
			hr = MakeFileNode(&file);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = file.try_query<IFileNodeProperties>()->put_Path(wil::make_bstr_nothrow(L"file.asm").get());
			Assert::IsTrue(SUCCEEDED(hr));
			WriteFileOnDisk(tempPath, L"file.asm");
			hr = AddFileToParent(file, hier.try_query<IParentNode>());
			Assert::IsTrue(SUCCEEDED(hr));

			wchar_t newFullPath[MAX_PATH];
			PathCombine(newFullPath, tempPath, L"new.asm");
			BOOL bres = DeleteFile(newFullPath);
			Assert::IsTrue(bres || GetLastError() == ERROR_FILE_NOT_FOUND);

			hr = hier->SetProperty(file->GetItemId(), VSHPROPID_EditLabel, wil::make_variant_bstr_nothrow(L"new.asm"));
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
	};
}
