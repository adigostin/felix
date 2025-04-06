
#include "pch.h"
#include "CppUnitTest.h"
#include "Mocks.h"
#include "Z80Xml.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

wchar_t tempPath[MAX_PATH + 1];

#pragma warning (push)
#pragma warning (disable: 4995)

TEST_MODULE_INITIALIZE(ModuleInit)
{
	serviceProvider = MakeMockServiceProvider();

	GetTempPathW (MAX_PATH + 1, tempPath);
	wchar_t* end;
	StringCchCatExW (tempPath, _countof(tempPath), L"FelixTest", &end, NULL, 0);
	Assert::IsTrue (end < tempPath + _countof(tempPath) - 1);
	end[1] = 0;
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = tempPath, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	SHFileOperation(&file_op);

	BOOL bres = CreateDirectory(tempPath, 0);
	Assert::IsTrue(bres);
}

TEST_MODULE_CLEANUP(ModuleCleanup)
{
	SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = tempPath, .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
	SHFileOperation(&file_op);

	serviceProvider = nullptr;
}

namespace FelixTests
{
	TEST_CLASS(ProjectTests)
	{
		TEST_METHOD(PutItemsOneFile)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file = MakeMockFileNode(BuildToolKind::Assembler, L"test.asm");
			Assert::AreEqual<VSITEMID>(VSITEMID_NIL, file->GetItemId());

			{
				SAFEARRAYBOUND bound = { .cElements = 1, .lLbound = 0 };
				auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
				Assert::IsTrue(SUCCEEDED(hr));
				LONG i = 0;
				SafeArrayPutElement (sa.get(), &i, file.try_query<IDispatch>());
				hr = hier.try_query<IProjectNodeProperties>()->put_Items(sa.get());
				Assert::IsTrue(SUCCEEDED(hr));

				Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file->GetItemId());

				auto pip = hier.try_query<IParentNode>();
				IUnknown* firstChild = wil::try_com_query_nothrow<IUnknown>(pip->FirstChild());
				Assert::AreEqual<void*>(file.try_query<IUnknown>().get(), firstChild);

				wil::unique_variant parentID;
				hr = file->GetProperty(VSHPROPID_Parent, &parentID);
				Assert::IsTrue(SUCCEEDED(hr));
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, V_VSITEMID(&parentID));
			}

			ULONG refCount = hier.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);

			refCount = file.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);
		}

		TEST_METHOD(PutItemsTwoFilesOneInFolder)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file1 = MakeMockFileNode(BuildToolKind::Assembler, L"file1.asm");
			auto file2 = MakeMockFileNode(BuildToolKind::Assembler, L"testfolder/file2.asm");

			com_ptr<IFolderNode> folder;

			{
				Assert::AreEqual<VSITEMID>(VSITEMID_NIL, file1->GetItemId());
				Assert::AreEqual<VSITEMID>(VSITEMID_NIL, file2->GetItemId());

				SAFEARRAYBOUND bound = { .cElements = 2, .lLbound = 0 };
				auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
				LONG i = 0;
				SafeArrayPutElement (sa.get(), &i, file1.try_query<IDispatch>());
				i++;
				SafeArrayPutElement (sa.get(), &i, file2.try_query<IDispatch>());
				hr = hier.try_query<IProjectNodeProperties>()->put_Items(sa.get());
				Assert::IsTrue(SUCCEEDED(hr));

				// The project should have assigned them item ids
				Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file1->GetItemId());
				Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, file2->GetItemId());

				// First project child should be a folder.
				auto folderItemId = GetProperty_VSITEMID(hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
				auto folderDisp = GetProperty_Dispatch(hier, folderItemId, VSHPROPID_BrowseObject);
				folder = folderDisp.try_query<IFolderNode>();
				Assert::IsNotNull(folder.get());
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, folderItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"testfolder", GetProperty_String(hier, folderItemId, VSHPROPID_SaveName).get());

				// Next project child should be our first file.
				auto file1ItemId = GetProperty_VSITEMID(hier, folderItemId, VSHPROPID_NextSibling);
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"file1.asm", GetProperty_String(hier, file1ItemId, VSHPROPID_SaveName).get());

				// There should be no more files after that.
				Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_NextSibling));

				// First child in the folder should be our second file.
				auto file2ItemId = GetProperty_VSITEMID (hier, folderItemId, VSHPROPID_FirstChild);
				Assert::AreEqual<VSITEMID>(folderItemId, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"file2.asm", GetProperty_String(hier, file2ItemId, VSHPROPID_SaveName).get());

				// There should be no more files after that.
				Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_NextSibling));
			}

			Assert::AreEqual<ULONG>(0, hier.detach()->Release());
			Assert::AreEqual<ULONG>(0, folder.detach()->Release());
			Assert::AreEqual<ULONG>(0, file1.detach()->Release());
			Assert::AreEqual<ULONG>(0, file2.detach()->Release());
		}

		TEST_METHOD(PutItemsTwoFilesInTwoFolders)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file1 = MakeMockFileNode(BuildToolKind::Assembler, L"testfolder1/file1.asm");
			auto file2 = MakeMockFileNode(BuildToolKind::Assembler, L"testfolder2/file2.asm");
			com_ptr<IFolderNode> folder1;
			com_ptr<IFolderNode> folder2;

			{
				SAFEARRAYBOUND bound = { .cElements = 2, .lLbound = 0 };
				auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
				LONG i = 0;
				SafeArrayPutElement (sa.get(), &i, file1.try_query<IDispatch>());
				i++;
				SafeArrayPutElement (sa.get(), &i, file2.try_query<IDispatch>());
				hr = hier.try_query<IProjectNodeProperties>()->put_Items(sa.get());
				Assert::IsTrue(SUCCEEDED(hr));

				// First project child should be a folder.
				auto folder1ItemId = GetProperty_VSITEMID(hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
				auto folder1Disp = GetProperty_Dispatch(hier, folder1ItemId, VSHPROPID_BrowseObject);
				folder1 = folder1Disp.try_query<IFolderNode>();
				Assert::IsNotNull(folder1.get());
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, folder1ItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"testfolder1", GetProperty_String(hier, folder1ItemId, VSHPROPID_SaveName).get());

				// Next project child should be the other folder.
				auto folder2ItemId = GetProperty_VSITEMID(hier, folder1ItemId, VSHPROPID_NextSibling);
				auto folder2Disp = GetProperty_Dispatch(hier, folder2ItemId, VSHPROPID_BrowseObject);
				folder2 = folder2Disp.try_query<IFolderNode>();
				Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, folder2ItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"testfolder2", GetProperty_String(hier, folder2ItemId, VSHPROPID_SaveName).get());

				// There should be no more nodes after that.
				Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, folder2ItemId, VSHPROPID_NextSibling));

				// First child in first folder should be our first file.
				auto file1ItemId = GetProperty_VSITEMID (hier, folder1ItemId, VSHPROPID_FirstChild);
				Assert::AreEqual<VSITEMID>(folder1ItemId, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"file1.asm", GetProperty_String(hier, file1ItemId, VSHPROPID_SaveName).get());
				// and then no more nodes
				Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_NextSibling));

				// First child in second folder should be our second file.
				auto file2ItemId = GetProperty_VSITEMID (hier, folder2ItemId, VSHPROPID_FirstChild);
				Assert::AreEqual<VSITEMID>(folder2ItemId, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_Parent));
				Assert::AreEqual(L"file2.asm", GetProperty_String(hier, file2ItemId, VSHPROPID_SaveName).get());
				// and then no more nodes
				Assert::AreEqual<VSITEMID>(VSITEMID_NIL, GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_NextSibling));
			}

			Assert::AreEqual<ULONG>(0, hier.detach()->Release());
			Assert::AreEqual<ULONG>(0, folder1.detach()->Release());
			Assert::AreEqual<ULONG>(0, folder2.detach()->Release());
			Assert::AreEqual<ULONG>(0, file1.detach()->Release());
			Assert::AreEqual<ULONG>(0, file2.detach()->Release());
		}

		TEST_METHOD(PutItemsFourFilesUnsorted)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			auto file1 = MakeMockFileNode(BuildToolKind::Assembler, L"file1.asm");
			auto file2 = MakeMockFileNode(BuildToolKind::Assembler, L"file2.asm");
			auto file3 = MakeMockFileNode(BuildToolKind::Assembler, L"file3.asm");
			auto file4 = MakeMockFileNode(BuildToolKind::Assembler, L"file4.asm");

			{
				SAFEARRAYBOUND bound = { .cElements = 4, .lLbound = 0 };
				auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound));
				LONG i = 0;
				SafeArrayPutElement (sa.get(), &i, file3.try_query<IDispatch>()); // Add to empty parent
				i++;
				SafeArrayPutElement (sa.get(), &i, file1.try_query<IDispatch>()); // Insert in first pos
				i++;
				SafeArrayPutElement (sa.get(), &i, file2.try_query<IDispatch>()); // Insert between the two above
				i++;
				SafeArrayPutElement (sa.get(), &i, file4.try_query<IDispatch>()); // Add at the end
				hr = hier.try_query<IProjectNodeProperties>()->put_Items(sa.get());
				Assert::IsTrue(SUCCEEDED(hr));

				// First project child should be file1.
				auto file1ItemId = GetProperty_VSITEMID(hier, VSITEMID_ROOT, VSHPROPID_FirstChild);
				Assert::AreEqual(L"file1.asm", GetProperty_String(hier, file1ItemId, VSHPROPID_SaveName).get());

				// Next should be file2.
				auto file2ItemId = GetProperty_VSITEMID(hier, file1ItemId, VSHPROPID_NextSibling);
				Assert::AreEqual(L"file2.asm", GetProperty_String(hier, file2ItemId, VSHPROPID_SaveName).get());

				// Then file3.
				auto file3ItemId = GetProperty_VSITEMID(hier, file2ItemId, VSHPROPID_NextSibling);
				Assert::AreEqual(L"file3.asm", GetProperty_String(hier, file3ItemId, VSHPROPID_SaveName).get());

				// And then file4.
				auto file4ItemId = GetProperty_VSITEMID(hier, file3ItemId, VSHPROPID_NextSibling);
				Assert::AreEqual(L"file4.asm", GetProperty_String(hier, file4ItemId, VSHPROPID_SaveName).get());
			}

			Assert::AreEqual<ULONG>(0, hier.detach()->Release());
			Assert::AreEqual<ULONG>(0, file1.detach()->Release());
			Assert::AreEqual<ULONG>(0, file2.detach()->Release());
			Assert::AreEqual<ULONG>(0, file3.detach()->Release());
			Assert::AreEqual<ULONG>(0, file4.detach()->Release());
		}

		TEST_METHOD(PutItemsTwoDirsUnsorted)
		{
		}
		TEST_METHOD(PutItemsOneFileInOneSubfolder)
		{
		}

		TEST_METHOD(PutItemsGetItemsSameWithFolders)
		{
		}

		TEST_METHOD(PutItemsOutsideOfProjectDir)
		{
		}

		TEST_METHOD(RemoveItemsFromRootNode)
		{
		}

		TEST_METHOD(RemoveItemsFromFolderNode)
		{
		}

		TEST_METHOD(AddItemNew)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsProject2> proj;
			hr = hier->QueryInterface(&proj);
			Assert::IsTrue(SUCCEEDED(hr));

			wchar_t templateFullPath[MAX_PATH];
			PathCombine(templateFullPath, tempPath, L"template.asm");
			HANDLE h = CreateFile(templateFullPath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
			Assert::AreNotEqual(INVALID_HANDLE_VALUE, h);
			CloseHandle(h);

			const wchar_t* templateName = templateFullPath;
			VSADDRESULT result;
			hr = proj->AddItem(VSITEMID_ROOT, VSADDITEMOP_CLONEFILE, L"test.asm", 1, &templateName, nullptr, &result);
			Assert::IsTrue(SUCCEEDED(hr));

			auto pip = hier.try_query<IParentNode>();
			com_ptr<IChildNode> firstChild = pip->FirstChild();
			Assert::IsNotNull(firstChild.get());
			Assert::AreNotEqual<VSITEMID>(VSITEMID_NIL, firstChild->GetItemId());
			Assert::AreEqual<VSITEMID>(VSITEMID_ROOT, GetProperty_VSITEMID(hier, firstChild->GetItemId(), VSHPROPID_Parent));

			pip.reset();
			proj.reset();

			ULONG refCount = hier.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);

			refCount = firstChild.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);
		}

		TEST_METHOD(AddItemNewToExistingFolder)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IFolderNode> folder;
			hr = MakeFolderNode(&folder);
			Assert::IsTrue(SUCCEEDED(hr));
			hr = folder.try_query<IFolderNodeProperties>()->put_FolderName(L"folder");
			Assert::IsTrue(SUCCEEDED(hr));

			hr = AddFolderToParent (folder, hier.try_query<IParentNode>());
			Assert::IsTrue(SUCCEEDED(hr));

			{
				com_ptr<IVsProject2> proj;
				hr = hier->QueryInterface(&proj);
				Assert::IsTrue(SUCCEEDED(hr));

				wchar_t templateFullPath[MAX_PATH];
				PathCombine(templateFullPath, tempPath, L"template.asm");
				HANDLE h = CreateFile(templateFullPath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
				Assert::AreNotEqual(INVALID_HANDLE_VALUE, h);
				CloseHandle(h);

				const wchar_t* templateName = templateFullPath;
				VSADDRESULT result;
				hr = proj->AddItem (folder->GetItemId(), VSADDITEMOP_CLONEFILE, L"folder/subfolder/file.asm", 1, &templateName, nullptr, &result);
				Assert::IsTrue(SUCCEEDED(hr));

				IChildNode* folderFirstChild = folder.try_query<IParentNode>()->FirstChild();
				Assert::IsNotNull(folderFirstChild);
				com_ptr<IFolderNode> subfolder = wil::try_com_query_nothrow<IFolderNode>(folderFirstChild);
				Assert::IsNotNull(subfolder.get());
				Assert::AreEqual(L"subfolder", GetProperty_String(hier, subfolder->GetItemId(), VSHPROPID_SaveName).get());
				Assert::AreNotEqual(VSITEMID_NIL, subfolder->GetItemId());

				IChildNode* subfolderFirstChild = subfolder.try_query<IParentNode>()->FirstChild();
				Assert::IsNotNull(subfolderFirstChild);
				com_ptr<IFileNode> file = wil::try_com_query_nothrow<IFileNode>(subfolderFirstChild);
				Assert::IsNotNull(file.get());
				Assert::AreEqual(L"file.asm", GetProperty_String(hier, file->GetItemId(), VSHPROPID_SaveName).get());
				Assert::AreNotEqual(VSITEMID_NIL, file->GetItemId());
			}

			ULONG refCount = hier.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);

			refCount = folder.detach()->Release();
			Assert::AreEqual<ULONG>(0, refCount);
		}

		TEST_METHOD(AddExistingItemWithHierarchyEventSinks)
		{
			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			auto sink = MakeMockHierarchyEventSink();
			VSCOOKIE cookie;
			hr = hier->AdviseHierarchyEvents(sink, &cookie);
			Assert::IsTrue(SUCCEEDED(hr));
			auto unadvise = wil::scope_exit([cookie, hier=hier.get()] { hier->UnadviseHierarchyEvents(cookie); });

			auto fullPathSource = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH);
			PathCombine(fullPathSource.get(), tempPath, L"file.asm");
			VSADDRESULT result;
			hr = hier.try_query<IVsProject>()->AddItem(VSITEMID_ROOT, VSADDITEMOP_OPENFILE, nullptr, 1, (LPCOLESTR*)fullPathSource.addressof(), nullptr, &result);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual<DWORD>(ADDRESULT_Success, result);
		}

		TEST_METHOD(AddItemOpen)
		{
		}

		TEST_METHOD(AddItemNewWithSubfolder_TestHierarchyEvents)
		{
			// test that OnItemAdded and OnPropertyChanged are called
		}

		TEST_METHOD(AddItemOutsideOfProjectDir)
		{
		}

		TEST_METHOD(DirtyAfterAddItem)
		{
		}

		TEST_METHOD(GetItemsPutItems_WithFolders)
		{
			com_ptr<IVsHierarchy> hier1;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier1));
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IFolderNode> folder1;
			hr = MakeFolderNode(&folder1);
			hr = folder1.try_query<IFolderNodeProperties>()->put_FolderName(L"folder");
			Assert::IsTrue(SUCCEEDED(hr));
			hr = AddFolderToParent(folder1, hier1.try_query<IParentNode>());
			Assert::IsTrue(SUCCEEDED(hr));

			auto file1 = MakeMockFileNode(BuildToolKind::Assembler, L"folder/test.asm");
			hr = AddFileToParent(file1, folder1.try_query<IParentNode>());

			// ------------------------------------------------

			auto stream = com_ptr(SHCreateMemStream(nullptr, 0));
			hr = SaveToXml(hier1.try_query<IProjectNodeProperties>(), L"Temp", 0, stream);
			Assert::IsTrue(SUCCEEDED(hr));

			com_ptr<IVsHierarchy> hier2;
			hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier2));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = stream->Seek({ 0 }, STREAM_SEEK_SET, nullptr);
			hr = LoadFromXml(hier2.try_query<IProjectNodeProperties>(), L"Temp", stream);
			Assert::IsTrue(SUCCEEDED(hr));

			// ---------------------

			auto pip2 = hier2.try_query<IParentNode>();
			Assert::IsNotNull(pip2->FirstChild());
			auto folder2 = wil::try_com_query_nothrow<IFolderNode>(pip2->FirstChild());
			Assert::IsNotNull(pip2->FirstChild());
			wil::unique_bstr folder2Name;
			hr = folder2.try_query<IFolderNodeProperties>()->get_FolderName(&folder2Name);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"folder", folder2Name.get());

			auto file2 = folder2.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(file2);
			wil::unique_bstr file2Path;
			hr = wil::try_com_query_nothrow<IFileNodeProperties>(file2)->get_Path(&file2Path);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::AreEqual(L"folder/test.asm", file2Path.get());
		}

		TEST_METHOD(TestLoadFromXml1)
		{
			static const char xml[] = ""
				"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
				"<Z80Project Guid=\"{2839FDD7-4C8F-4772-90E6-222C702D045E}\">"
					"<Configurations>"
						"<Configuration ConfigName=\"Debug\" PlatformName=\"ZX Spectrum 48K\">"
							"<AssemblerProperties SaveListingFilename=\"output.lst\" />"
						"</Configuration>"
					"</Configurations>"
					"<Items>"
						"<File Path=\"start.asm\" BuildTool=\"Assembler\" />"
						"<File Path=\"lib.asm\" BuildTool=\"Assembler\" />"
						"<File Path=\"More/EvenMore/file.inc\" BuildTool=\"None\" />"
						"<File Path=\"GeneratedFiles/preinclude.inc\" BuildTool=\"None\" />"
						"<File Path=\"GeneratedFiles/postinclude.inc\" BuildTool=\"None\" />"
					"</Items>"
				"</Z80Project>";

			// After loading, the hierarchy is supposed to look like this:
			// Project
			//   GeneratedFiles
			//     postinclude.inc
			//     preinclude.inc
			//   More
			//     EvenMore
			//       file.inc
			//   lib.asm
			//   start.asm

			auto s = SHCreateMemStream((BYTE*)xml, sizeof(xml) - 1);
			com_ptr<IStream> stream;
			stream.attach(s);

			com_ptr<IVsHierarchy> hier;
			auto hr = MakeProjectNode (nullptr, tempPath, nullptr, 0, IID_PPV_ARGS(&hier));
			Assert::IsTrue(SUCCEEDED(hr));

			hr = LoadFromXml(hier.try_query<IProjectNodeProperties>(), ProjectElementName, stream);
			Assert::IsTrue(SUCCEEDED(hr));

			auto c = hier.try_query<IParentNode>()->FirstChild();
			Assert::IsNotNull(c);
			auto genFilesFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(genFilesFolder.get());
			Assert::AreEqual(L"GeneratedFiles", GetProperty_String(hier, genFilesFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = genFilesFolder.try_query<IParentNode>()->FirstChild();
			auto postinc = wil::try_com_query_nothrow<IFileNode>(c);
			Assert::IsNotNull(postinc.get());
			Assert::AreEqual(L"postinclude.inc", GetProperty_String(hier, postinc->GetItemId(), VSHPROPID_SaveName).get());

			auto preinc = wil::try_com_query_nothrow<IFileNode>(postinc->Next());
			Assert::IsNotNull(preinc.get());
			Assert::AreEqual(L"preinclude.inc", GetProperty_String(hier, preinc->GetItemId(), VSHPROPID_SaveName).get());

			c = genFilesFolder->Next();
			Assert::IsNotNull(c);
			auto moreFolder = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(moreFolder.get());
			Assert::AreEqual(L"More", GetProperty_String(hier, moreFolder->GetItemId(), VSHPROPID_SaveName).get());

			c = moreFolder.try_query<IParentNode>()->FirstChild();
			auto evenMore = wil::try_com_query_nothrow<IFolderNode>(c);
			Assert::IsNotNull(evenMore.get());
			Assert::AreEqual(L"EvenMore", GetProperty_String(hier, evenMore->GetItemId(), VSHPROPID_SaveName).get());

			c = evenMore.try_query<IParentNode>()->FirstChild();
			auto fileinc = wil::try_com_query_nothrow<IFileNode>(c);
			Assert::IsNotNull(fileinc.get());
			Assert::AreEqual(L"file.inc", GetProperty_String(hier, fileinc->GetItemId(), VSHPROPID_SaveName).get());

			auto libasm = wil::try_com_query_nothrow<IFileNode>(moreFolder->Next());
			Assert::IsNotNull(libasm.get());
			Assert::AreEqual(L"lib.asm", GetProperty_String(hier, libasm->GetItemId(), VSHPROPID_SaveName).get());

			auto startasm = wil::try_com_query_nothrow<IFileNode>(libasm->Next());
			Assert::IsNotNull(startasm.get());
			Assert::AreEqual(L"start.asm", GetProperty_String(hier, startasm->GetItemId(), VSHPROPID_SaveName).get());

			Assert::IsNull(startasm->Next());
		}
	};
}

#pragma warning (pop)
