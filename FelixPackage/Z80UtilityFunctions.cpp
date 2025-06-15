
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/inplace_function.h"
#include "../FelixPackageUi/resource.h"
#include "guids.h"

#define __dte_h__
#include <VSShell174.h>

const char* PropIDToString (VSHPROPID propid)
{
	switch (propid)
	{
		case VSHPROPID_NIL:                   return "NIL";
		case VSHPROPID_Parent:                return "Parent";
		case VSHPROPID_FirstChild:            return "FirstChild";
		case VSHPROPID_NextSibling:           return "NextSibling";
		case VSHPROPID_Root:                  return "Root";
		case VSHPROPID_TypeGuid:              return "TypeGuid";
		case VSHPROPID_SaveName:              return "SaveName";
		case VSHPROPID_Caption:               return "Caption";
		case VSHPROPID_IconImgList:           return "IconImgList";
		case VSHPROPID_IconIndex:             return "IconIndex";
		case VSHPROPID_Expandable:            return "Expandable";
		case VSHPROPID_ExpandByDefault:       return "ExpandByDefault";
		case VSHPROPID_ProjectName:           return "ProjectName / Name";
		case VSHPROPID_IconHandle:            return "IconHandle";
		case VSHPROPID_OpenFolderIconHandle:  return "OpenFolderIconHandle";
		case VSHPROPID_OpenFolderIconIndex:   return "OpenFolderIconIndex";
		case VSHPROPID_CmdUIGuid:             return "CmdUIGuid";
		case VSHPROPID_SelContainer:          return "SelContainer";
		case VSHPROPID_BrowseObject:          return "BrowseObject";
		case VSHPROPID_AltHierarchy:          return "AltHierarchy";
		case VSHPROPID_AltItemid:             return "AltItemid";
		case VSHPROPID_ProjectDir:            return "ProjectDir";
		case VSHPROPID_SortPriority:          return "SortPriority";
		case VSHPROPID_UserContext:           return "UserContext";
		case VSHPROPID_EditLabel:             return "EditLabel";
		case VSHPROPID_ExtObject:             return "ExtObject";
		case VSHPROPID_ExtSelectedItem:       return "ExtSelectedItem";
		case VSHPROPID_StateIconIndex:        return "StateIconIndex";
		case VSHPROPID_ProjectType:           return "ProjectType / TypeName";
		case VSHPROPID_ReloadableProjectFile: return "ReloadableProjectFile / HandlesOwnReload";
		case VSHPROPID_ParentHierarchy:       return "ParentHierarchy";
		case VSHPROPID_ParentHierarchyItemid: return "ParentHierarchyItemid";
		case VSHPROPID_ItemDocCookie:         return "ItemDocCookie";
		case VSHPROPID_Expanded:              return "Expanded";
		case VSHPROPID_ConfigurationProvider: return "ConfigurationProvider";
		case VSHPROPID_ImplantHierarchy:      return "ImplantHierarchy";
		case VSHPROPID_OwnerKey:              return "OwnerKey";
		case VSHPROPID_StartupServices:       return "StartupServices";
		case VSHPROPID_FirstVisibleChild:     return "FirstVisibleChild";
		case VSHPROPID_NextVisibleSibling:    return "NextVisibleSibling";
		case VSHPROPID_IsHiddenItem:          return "IsHiddenItem";
		case VSHPROPID_IsNonMemberItem:       return "IsNonMemberItem";
		case VSHPROPID_IsNonLocalStorage:     return "IsNonLocalStorage";
		case VSHPROPID_StorageType:           return "StorageType";
		case VSHPROPID_ItemSubType:           return "ItemSubType";
		case VSHPROPID_OverlayIconIndex:      return "OverlayIconIndex";
		case VSHPROPID_DefaultNamespace:      return "DefaultNamespace";
		case VSHPROPID_IsNonSearchable:       return "IsNonSearchable";
		case VSHPROPID_IsFindInFilesForegroundOnly:   return "IsFindInFilesForegroundOnly";
		case VSHPROPID_CanBuildFromMemory:            return "CanBuildFromMemory";
		case VSHPROPID_PreferredLanguageSID:          return "PreferredLanguageSID";
		case VSHPROPID_ShowProjInSolutionPage:        return "ShowProjInSolutionPage";
		case VSHPROPID_AllowEditInRunMode:            return "AllowEditInRunMode";
		case VSHPROPID_IsNewUnsavedItem:              return "IsNewUnsavedItem";
		case VSHPROPID_ShowOnlyItemCaption:           return "ShowOnlyItemCaption";
		case VSHPROPID_ProjectIDGuid:                 return "ProjectIDGuid";
		case VSHPROPID_DesignerVariableNaming:        return "DesignerVariableNaming";
		case VSHPROPID_DesignerFunctionVisibility:    return "DesignerFunctionVisibility";
		case VSHPROPID_HasEnumerationSideEffects:     return "HasEnumerationSideEffects";
		case VSHPROPID_DefaultEnableBuildProjectCfg:  return "DefaultEnableBuildProjectCfg";
		case VSHPROPID_DefaultEnableDeployProjectCfg: return "DefaultEnableDeployProjectCfg";
		// VSHPROPID2
		case VSHPROPID_PropertyPagesCLSIDList:          return "PropertyPagesCLSIDList";
		case VSHPROPID_CfgPropertyPagesCLSIDList:       return "CfgPropertyPagesCLSIDList";
		case VSHPROPID_ExtObjectCATID:                  return "ExtObjectCATID";
		case VSHPROPID_BrowseObjectCATID:               return "BrowseObjectCATID";
		case VSHPROPID_CfgBrowseObjectCATID:            return "CfgBrowseObjectCATID";
		case VSHPROPID_AddItemTemplatesGuid:            return "AddItemTemplatesGuid";
		case VSHPROPID_ChildrenEnumerated:              return "ChildrenEnumerated";
		case VSHPROPID_StatusBarClientText:             return "StatusBarClientText";
		case VSHPROPID_DebuggeeProcessId:               return "DebuggeeProcessId";
		case VSHPROPID_IsLinkFile:                      return "IsLinkFile";
		case VSHPROPID_KeepAliveDocument:               return "KeepAliveDocument";
		case VSHPROPID_SupportsProjectDesigner:         return "SupportsProjectDesigner";
		case VSHPROPID_IntellisenseUnknown:             return "IntellisenseUnknown";
		case VSHPROPID_IsUpgradeRequired:               return "IsUpgradeRequired";
		case VSHPROPID_DesignerHiddenCodeGeneration:    return "DesignerHiddenCodeGeneration";
		case VSHPROPID_SuppressOutOfDateMessageOnBuild: return "SuppressOutOfDateMessageOnBuild";
		case VSHPROPID_Container:                       return "Container";
		case VSHPROPID_UseInnerHierarchyIconList:       return "UseInnerHierarchyIconList";
		case VSHPROPID_EnableDataSourceWindow:          return "EnableDataSourceWindow";
		case VSHPROPID_AppTitleBarTopHierarchyName:     return "AppTitleBarTopHierarchyName";
		case VSHPROPID_DebuggerSourcePaths:             return "DebuggerSourcePaths";
		case VSHPROPID_CategoryGuid:                    return "CategoryGuid";
		case VSHPROPID_DisableApplicationSettings:      return "DisableApplicationSettings";
		case VSHPROPID_ProjectDesignerEditor:           return "ProjectDesignerEditor";
		case VSHPROPID_PriorityPropertyPagesCLSIDList:  return "PriorityPropertyPagesCLSIDList";
		case VSHPROPID_NoDefaultNestedHierSorting:      return "NoDefaultNestedHierSorting";
		case VSHPROPID_ExcludeFromExportItemTemplate:   return "ExcludeFromExportItemTemplate";
		case VSHPROPID_SupportedMyApplicationTypes:     return "SupportedMyApplicationTypes";
		// VSHPROPID3
		case VSHPROPID_TargetFrameworkVersion          : return "TargetFrameworkVersion";           // -2093
		case VSHPROPID_WebReferenceSupported           : return "WebReferenceSupported";            // -2094
		case VSHPROPID_ServiceReferenceSupported       : return "ServiceReferenceSupported";        // -2095
		case VSHPROPID_SupportsHierarchicalUpdate      : return "SupportsHierarchicalUpdate";       // -2096
		case VSHPROPID_SupportsNTierDesigner           : return "SupportsNTierDesigner";            // -2097
		case VSHPROPID_SupportsLinqOverDataSet         : return "SupportsLinqOverDataSet";          // -2098
		case VSHPROPID_ProductBrandName                : return "ProductBrandName";                 // -2099
		case VSHPROPID_RefactorExtensions              : return "RefactorExtensions";               // -2100
		case VSHPROPID_IsDefaultNamespaceRefactorNotify: return "IsDefaultNamespaceRefactorNotify"; // -2101
		// VSHPROPID4
		case VSHPROPID_TargetFrameworkMoniker:          return "TargetFrameworkMoniker";
		case VSHPROPID_ExternalItem:                    return "ExternalItem";
		case VSHPROPID_SupportsAspNetIntegration:       return "SupportsAspNetIntegration";
		case VSHPROPID_DesignTimeDependencies:          return "DesignTimeDependencies";
		case VSHPROPID_BuildDependencies:               return "BuildDependencies";
		case VSHPROPID_BuildAction:                     return "BuildAction";
		case VSHPROPID_DescriptiveName:                 return "DescriptiveName";
		case VSHPROPID_AlwaysBuildOnDebugLaunch:        return "AlwaysBuildOnDebugLaunch";
		// VSHPROPID5
		case VSHPROPID_MinimumDesignTimeCompatVersion:  return "MinimumDesignTimeCompatVersion";
		case VSHPROPID_ProvisionalViewingStatus:        return "ProvisionalViewingStatus";
		case VSHPROPID_SupportedOutputTypes:            return "SupportedOutputTypes";
		case VSHPROPID_TargetPlatformIdentifier:        return "TargetPlatformIdentifier";
		case VSHPROPID_TargetPlatformVersion:           return "TargetPlatformVersion";
		case VSHPROPID_TargetRuntime:                   return "TargetRuntime";
		case VSHPROPID_AppContainer:                    return "AppContainer";
		case VSHPROPID_OutputType:                      return "OutputType";
		case VSHPROPID_ReferenceManagerUser:            return "ReferenceManagerUser";
		case VSHPROPID_ProjectUnloadStatus:             return "ProjectUnloadStatus";
		case VSHPROPID_DemandLoadDependencies:          return "DemandLoadDependencies";
		case VSHPROPID_IsFaulted:                       return "IsFaulted";
		case VSHPROPID_FaultMessage:                    return "FaultMessage";
		case VSHPROPID_ProjectCapabilities:             return "ProjectCapabilities";
		case VSHPROPID_RequiresReloadForExternalFileChange: return "RequiresReloadForExternalFileChange";
		case VSHPROPID_ForceFrameworkRetarget:          return "ForceFrameworkRetarget";
		case VSHPROPID_IsProjectProvisioned:            return "IsProjectProvisioned";
		case VSHPROPID_SupportsCrossRuntimeReferences:  return "SupportsCrossRuntimeReferences";
		case VSHPROPID_WinMDAssembly:                   return "WinMDAssembly";
		case VSHPROPID_MonikerSameAsPersistFile:        return "MonikerSameAsPersistFile";
		case VSHPROPID_IsPackagingProject:              return "IsPackagingProject";
		case VSHPROPID_ProjectPropertiesDebugPageArg:   return "ProjectPropertiesDebugPageArg";
		// VSHPROPID6
		case VSHPROPID_ConnectedServicesPersistence:    return "ConnectedServicesPersistence";
		case VSHPROPID_ProjectRetargeting:              return "ProjectRetargeting";
		case VSHPROPID_ShowAllProjectFilesInProjectView:return "ShowAllProjectFilesInProjectView";
		case VSHPROPID_Subcaption:                      return "Subcaption";
		case VSHPROPID_ScriptJmcProjectControl:         return "ScriptJmcProjectControl";
		case VSHPROPID_NuGetPackageProjectTypeContext:  return "NuGetPackageProjectTypeContext";
		case VSHPROPID_RequiresLegacyManagedDebugEngine:return "RequiresLegacyManagedDebugEngine";
		case VSHPROPID_CurrentTargetId:                 return "CurrentTargetId";
		case VSHPROPID_NewTargetId:                     return "NewTargetId";
		// VSHPROPID7
		case VSHPROPID_IsSharedItem:                   return "IsSharedItem";
		case VSHPROPID_SharedItemContextHierarchy:     return "SharedItemContextHierarchy";
		case VSHPROPID_ShortSubcaption:                return "ShortSubcaption";
		case VSHPROPID_SharedItemsImportFullPaths:     return "SharedItemsImportFullPaths";
		case VSHPROPID_ProjectTreeCapabilities:        return "ProjectTreeCapabilities";
		case VSHPROPID_DeploymentRelativePath:         return "DeploymentRelativePath";
		case VSHPROPID_IsSharedFolder:                 return "IsSharedFolder";
		case VSHPROPID_OneAppCapabilities:             return "OneAppCapabilities";
		case VSHPROPID_MSBuildImportsStorage:          return "MSBuildImportsStorage";
		case VSHPROPID_SharedProjectHierarchy:         return "SharedProjectHierarchy";
		case VSHPROPID_SharedAssetsProject:            return "SharedAssetsProject";
		case VSHPROPID_IsSharedItemsImportFile:        return "IsSharedItemsImportFile";
		case VSHPROPID_ExcludeFromMoveFileToProjectUI: return "ExcludeFromMoveFileToProjectUI";
		case VSHPROPID_CanBuildQuickCheck:             return "CanBuildQuickCheck";
		case VSHPROPID_CanDebugLaunchQuickCheck:       return "CanDebugLaunchQuickCheck";
		case VSHPROPID_CanDeployQuickCheck:            return "CanDeployQuickCheck";
		// VSHPROPID8
		case VSHPROPID_SupportsIconMonikers:            return "SupportsIconMonikers";
		case VSHPROPID_IconMonikerGuid:                 return "IconMonikerGuid";
		case VSHPROPID_IconMonikerId:                   return "IconMonikerId";
		case VSHPROPID_OpenFolderIconMonikerGuid:       return "OpenFolderIconMonikerGuid";
		case VSHPROPID_OpenFolderIconMonikerId:         return "OpenFolderIconMonikerId";
		case VSHPROPID_IconMonikerImageList:            return "IconMonikerImageList";
		case VSHPROPID_SharedProjectReference:          return "SharedProjectReference";
		case VSHPROPID_DiagHubPlatform:                 return "DiagHubPlatform";
		case VSHPROPID_DiagHubPlatformVersion:          return "DiagHubPlatformVersion";
		case VSHPROPID_DiagHubLanguage:                 return "DiagHubLanguage";
		case VSHPROPID_DiagHubProjectTargetFactory:     return "DiagHubProjectTargetFactory";
		case VSHPROPID_DiagHubProjectTarget:            return "DiagHubProjectTarget";
		case VSHPROPID_SolutionGuid:                    return "SolutionGuid";
		case VSHPROPID_ActiveIntellisenseProjectContext:return "ActiveIntellisenseProjectContext";
		case VSHPROPID_ProjectCapabilitiesChecker:      return "ProjectCapabilitiesChecker";
		case VSHPROPID_ContainsStartupTask:             return "ContainsStartupTask";

		// VSSPROPID13
		case VSSPROPID_EnableEnhancedTooltips:          return "VSSPROPID_EnableEnhancedTooltips"; // -9088
		case VSHPROPID_SlowEnumeration:                 return "VSHPROPID_SlowEnumeration"; // -9089
	}

	return nullptr;
}

void PrintProperty (const char* prefix, VSHPROPID propid, const VARIANT* pvar)
{
	OutputDebugStringA (prefix);
	OutputDebugStringA ("propid=");
	auto s = PropIDToString(propid);
	if (s != nullptr)
		OutputDebugStringA (s);
	else
	{
		char buffer[20];
		sprintf_s(buffer, "%d", propid);
		OutputDebugStringA(buffer);
	}

	if (pvar->vt != VT_EMPTY)
	{
		OutputDebugStringA (", pvar=");

		VARIANT v;
		VariantInit(&v);
		VariantChangeType (&v, pvar, 0, VT_BSTR);
		OutputDebugString (v.bstrVal);
		VariantClear(&v);
	}

	OutputDebugStringA("\r\n");
}

HRESULT MakeBstrFromString (const char* name, BSTR* bstr)
{
	return MakeBstrFromString (name, name + strlen(name), bstr);
}

HRESULT MakeBstrFromString (const char* name, size_t len, BSTR* bstr)
{
	return MakeBstrFromString (name, name + len, bstr);
}

HRESULT MakeBstrFromString (const char* sl_name_from, const char* sl_name_to, BSTR* to)
{
	if (sl_name_from == sl_name_to)
		return (*to = nullptr), S_OK;
	int ires = MultiByteToWideChar(CP_UTF8, 0, sl_name_from, (int)(sl_name_to - sl_name_from), nullptr, 0); RETURN_LAST_ERROR_IF(!ires);
	auto wname = wil::make_hlocal_string_nothrow(nullptr, ires); RETURN_IF_NULL_ALLOC(wname);
	int ires1 = MultiByteToWideChar (CP_UTF8, 0, sl_name_from, (int)(sl_name_to - sl_name_from), wname.get(), ires + 1); RETURN_LAST_ERROR_IF(!ires); RETURN_HR_IF(E_FAIL, ires1 != ires);
	*to = SysAllocStringLen(wname.get(), ires); RETURN_IF_NULL_ALLOC(*to);
	return S_OK;
}

FELIX_API HRESULT MakeBstrFromStreamOnHGlobal (IStream* stream, BSTR* pBstr)
{
	STATSTG stat;
	auto hr = stream->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(ERROR_FILE_TOO_LARGE, !!stat.cbSize.HighPart);

	HGLOBAL hg;
	hr = GetHGlobalFromStream(stream, &hg); RETURN_IF_FAILED(hr);
	auto buffer = GlobalLock(hg); RETURN_LAST_ERROR_IF(!buffer);
	*pBstr = SysAllocStringLen((OLECHAR*)buffer, stat.cbSize.LowPart / 2);
	GlobalUnlock (hg);
	RETURN_IF_NULL_ALLOC(*pBstr);

	return S_OK;
}

static HRESULT Write (ISequentialStream* stream, const wchar_t* psz)
{
	return stream->Write(psz, (ULONG)wcslen(psz) * sizeof(wchar_t), nullptr);
}

static HRESULT Write (ISequentialStream* stream, const wchar_t* from, const wchar_t* to)
{
	return stream->Write(from, (ULONG)(to - from) * sizeof(wchar_t), nullptr);
}

static HRESULT GeneratePrePostIncludeFilesInner (IProjectNode* project, IProjectMacroResolver* macroResolver)
{
	HRESULT hr;

	com_ptr<IVsShell> shell;
	hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);

	wil::unique_hlocal_string packageDir;
	hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, packageDir); RETURN_IF_FAILED(hr);
	auto fnres = PathFindFileName(packageDir.get()); RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME), fnres == packageDir.get());
	*fnres = 0;

	wil::unique_bstr genFilesStr;
	hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERATED_FILES, &genFilesStr); RETURN_IF_FAILED(hr);
	com_ptr<IFolderNode> folder;
	hr = GetOrCreateChildFolder(project->AsParentNode(), genFilesStr.get(), true, &folder); RETURN_IF_FAILED(hr);

	for (ULONG resID : { IDS_PREINCLUDE, IDS_POSTINCLUDE })
	{
		wil::unique_bstr fileName;
		hr = shell->LoadPackageString(CLSID_FelixPackage, resID, &fileName); RETURN_IF_FAILED(hr);

		com_ptr<IFileNode> file = FindChildFileByName(folder->AsParentNode(), fileName.get());
		if (!file)
		{
			hr = MakeFileNodeForExistingFile (fileName.get(), &file); RETURN_IF_FAILED(hr);
			file.try_query<IFileNodeProperties>()->put_IsGenerated(TRUE);
			hr = AddFileToParent(file, folder->AsParentNode()); RETURN_IF_FAILED(hr);
		}

		wil::unique_process_heap_string templatePath;
		hr = wil::str_concat_nothrow(templatePath, packageDir, L"Templates\\", fileName); RETURN_IF_FAILED(hr);

		wil::unique_process_heap_string includePath;
		hr = GetPathOf(file, includePath); RETURN_IF_FAILED(hr);
		hr = CreateFileFromTemplate(templatePath.get(), includePath.get(), macroResolver); RETURN_IF_FAILED(hr);
	}

	wil::unique_bstr projectMk;
	hr = project->AsVsProject()->GetMkDocument(VSITEMID_ROOT, &projectMk); RETURN_IF_FAILED(hr);
	com_ptr<IVsFileChangeEx> fileChange;
	hr = serviceProvider->QueryService(SID_SVsFileChangeEx, IID_PPV_ARGS(&fileChange)); RETURN_IF_FAILED(hr);
	hr = fileChange->IgnoreFile(VSCOOKIE_NIL, projectMk.get(), TRUE); RETURN_IF_FAILED(hr);
	auto unignore = wil::scope_exit([&fileChange, &projectMk] { fileChange->IgnoreFile(0, projectMk.get(), FALSE); });
	hr = wil::try_com_query_nothrow<IPersistFileFormat>(project)->Save(nullptr, 0, 0); RETURN_IF_FAILED(hr);
	hr = fileChange->SyncFile(projectMk.get()); (void)hr;
	unignore.reset();

	return S_OK;
}

HRESULT GeneratePrePostIncludeFiles (IProjectNode* project, const wchar_t* configName, IProjectMacroResolver* macroResolver)
{
	HRESULT hr;

	com_ptr<IVsShell> shell;
	hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);

	com_ptr<IVsOutputWindowPane> op;
	hr = serviceProvider->QueryService(SID_SVsGeneralOutputWindowPane, IID_PPV_ARGS(&op)); RETURN_IF_FAILED(hr);
	hr = op->Activate(); RETURN_IF_FAILED(hr);
	com_ptr<IVsOutputWindowPane2> op2;
	hr = op->QueryInterface(&op2); RETURN_IF_FAILED(hr);

	wil::unique_variant projectName;
	hr = project->AsHierarchy()->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &projectName); RETURN_IF_FAILED(hr);
	wil::unique_bstr str;
	if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_GEN_PRE_POST_MESSAGE, &str)))
	{
		wil::unique_process_heap_string message;
		if (SUCCEEDED(wil::str_printf_nothrow(message, str.get(), configName)))
			op2->OutputTaskItemStringEx2(message.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
				nullptr, 0, nullptr, 0, 0, projectName.bstrVal, nullptr, nullptr);
	}
	
	hr = GeneratePrePostIncludeFilesInner (project, macroResolver);
	if (FAILED(hr))
	{
		if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_GEN_PRE_POST_MESSAGE_FAIL, &str)))
			op2->OutputTaskItemStringEx2(str.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
				nullptr, 0, nullptr, 0, 0, projectName.bstrVal, nullptr, nullptr);
		return hr;
	}

	if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_GEN_PRE_POST_MESSAGE_DONE, &str)))
		op2->OutputTaskItemStringEx2(str.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
			nullptr, 0, nullptr, 0, 0, projectName.bstrVal, nullptr, nullptr);
	return S_OK;
};

FELIX_API HRESULT MakeSjasmCommandLine (IVsHierarchy* hier, IProjectConfig* config, IProjectConfigAssemblerProperties* asmPropsOverride, BSTR* ppCmdLine)
{
	HRESULT hr;

	com_ptr<IVsShell> shell;
	hr = serviceProvider->QueryService(SID_SVsShell, IID_PPV_ARGS(&shell)); RETURN_IF_FAILED(hr);

	wil::unique_variant projectDir;
	hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, projectDir.addressof()); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);

	wil::unique_bstr output_dir;
	hr = config->GetOutputDirectory(&output_dir); RETURN_IF_FAILED(hr);

	vector_nothrow<com_ptr<IFileNodeProperties>> asmFiles;

	wil::unique_bstr generatedFilesName;
	hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_GENERATED_FILES, &generatedFilesName); RETURN_IF_FAILED(hr);
	com_ptr<IFolderNode> genFilesFolder;
	com_ptr<IFileNodeProperties> preIncludeFile, postIncludeFile;
	for (auto c = wil::try_com_query_nothrow<IParentNode>(hier)->FirstChild(); c; c = c->Next())
	{
		com_ptr<IFolderNode> folder;
		wil::unique_variant folderName;
		if (SUCCEEDED(c->QueryInterface(IID_PPV_ARGS(&folder)))
			&& SUCCEEDED(folder->GetProperty(VSHPROPID_SaveName, &folderName))
			&& folderName.vt == VT_BSTR && folderName.bstrVal
			&& !wcscmp(folderName.bstrVal, generatedFilesName.get()))
		{
			wil::unique_bstr preincludeName, postincludeName;
			hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_PREINCLUDE, &preincludeName); RETURN_IF_FAILED(hr);
			hr = shell->LoadPackageString(CLSID_FelixPackage, IDS_POSTINCLUDE, &postincludeName); RETURN_IF_FAILED(hr);

			if (auto file = FindChildFileByName(folder->AsParentNode(), preincludeName.get()))
				preIncludeFile = wil::try_com_query_nothrow<IFileNodeProperties>(file);

			if (auto file = FindChildFileByName(folder->AsParentNode(), postincludeName.get()))
				postIncludeFile = wil::try_com_query_nothrow<IFileNodeProperties>(file);

			genFilesFolder = std::move(folder);
			break;
		}
	}

	stdext::inplace_function<HRESULT(IParentNode*)> enumDescendants;

	enumDescendants = [&enumDescendants, &asmFiles, &genFilesFolder](IParentNode* parent) -> HRESULT
		{
			for (auto c = parent->FirstChild(); c; c = c->Next())
			{
				if (c == genFilesFolder)
					continue;

				if (auto file = wil::try_com_query_nothrow<IFileNodeProperties>(c))
				{
					BuildToolKind tool;
					auto hr = file->get_BuildTool(&tool); RETURN_IF_FAILED(hr);
					if (tool == BuildToolKind::Assembler)
					{
						bool pushed = asmFiles.try_push_back(std::move(file)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
					}
				}
				else if (auto cAsParent = wil::try_com_query_nothrow<IParentNode>(c))
				{
					auto hr = enumDescendants(cAsParent); RETURN_IF_FAILED(hr);
				}
			}

			return S_OK;
		};
	if (preIncludeFile)
		asmFiles.try_push_back(std::move(preIncludeFile));
	hr = enumDescendants(wil::try_com_query_nothrow<IParentNode>(hier)); RETURN_IF_FAILED(hr);
	if (postIncludeFile)
		asmFiles.try_push_back(std::move(postIncludeFile));

	if (asmFiles.empty())
		return (*ppCmdLine = nullptr), S_OK;

	com_ptr<IStream> cmdLine;
	hr = CreateStreamOnHGlobal (nullptr, TRUE, &cmdLine); RETURN_IF_FAILED(hr);

	wil::unique_hlocal_string moduleFilename;
	hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, moduleFilename); RETURN_IF_FAILED(hr);
	auto fnres = PathFindFileName(moduleFilename.get()); RETURN_HR_IF(CO_E_BAD_PATH, fnres == moduleFilename.get());
	bool hasSpaces = !!wcschr(moduleFilename.get(), L' ');
	if (hasSpaces)
	{
		hr = Write(cmdLine, L"\""); RETURN_IF_FAILED(hr);
	}
	hr = Write(cmdLine, moduleFilename.get(), fnres); RETURN_IF_FAILED(hr);
	hr = Write(cmdLine, L"sjasmplus.exe"); RETURN_IF_FAILED(hr);
	if (hasSpaces)
	{
		hr = Write(cmdLine, L"\""); RETURN_IF_FAILED(hr);
	}
	hr = Write(cmdLine, L" --fullpath"); RETURN_IF_FAILED(hr);

	auto addOutputPathParam = [&cmdLine, output_dir=output_dir.get(), project_dir=projectDir.bstrVal](const wchar_t* paramName, const wchar_t* output_filename) -> HRESULT
		{
			auto outputFilePath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePath);
			auto pres = PathCombine (outputFilePath.get(), output_dir, output_filename); RETURN_HR_IF(CO_E_BAD_PATH, !pres);
			auto outputFilePathRelativeUgly = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePathRelativeUgly);
			BOOL bRes = PathRelativePathToW (outputFilePathRelativeUgly.get(), project_dir, FILE_ATTRIBUTE_DIRECTORY, outputFilePath.get(), 0); RETURN_HR_IF(CS_E_INVALID_PATH, !bRes);
			auto outputFilePathRelative = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePathRelative);
			BOOL bres = PathCanonicalize (outputFilePathRelative.get(), outputFilePathRelativeUgly.get()); RETURN_IF_WIN32_BOOL_FALSE(bres);
			auto hr = Write(cmdLine, paramName); RETURN_IF_FAILED(hr);
			hr = Write(cmdLine, outputFilePathRelative.get()); RETURN_IF_FAILED(hr);
			return S_OK;
		};

	com_ptr<IProjectConfigAssemblerProperties> asmProps;
	if (asmPropsOverride)
		asmProps = asmPropsOverride;
	else
	{
		hr = config->AsProjectConfigProperties()->get_AssemblerProperties(&asmProps); RETURN_IF_FAILED(hr);
	}

	OutputFileType outputFileType;
	hr = asmProps->get_OutputFileType(&outputFileType); RETURN_IF_FAILED(hr);
	if (outputFileType == OutputFileType::Binary)
	{
		// --raw=...
		wil::unique_bstr output_filename;
		hr = asmProps->GetOutputFileName(&output_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (L" --raw=", output_filename.get()); RETURN_IF_FAILED(hr);
	}

	// --sld=...
	wil::unique_bstr sld_filename;
	hr = asmProps->GetSldFileName (&sld_filename); RETURN_IF_FAILED(hr);
	hr = addOutputPathParam (L" --sld=", sld_filename.get()); RETURN_IF_FAILED(hr);

	// --outprefix
	hr = addOutputPathParam (L" --outprefix=", L""); RETURN_IF_FAILED(hr);

	// --lst
	VARIANT_BOOL saveListing;
	hr = asmProps->get_SaveListing(&saveListing); RETURN_IF_FAILED(hr);
	if (saveListing)
	{
		wil::unique_bstr listingFilename;
		hr = asmProps->get_SaveListingFilename(&listingFilename); RETURN_IF_FAILED(hr);
		if (listingFilename && listingFilename.get()[0])
		{
			hr = addOutputPathParam (L" --lst=", listingFilename.get()); RETURN_IF_FAILED(hr);
		}
	}

	// input files
	for (auto& asmFile : asmFiles)
	{
		wil::unique_bstr path;
		hr = asmFile->get_Path(&path); RETURN_IF_FAILED(hr);
		hr = Write(cmdLine, L" "); RETURN_IF_FAILED(hr);
		if (PathIsFileSpec(path.get()))
		{
			// (1)
			com_ptr<IFileNode> fn;
			hr = asmFile->QueryInterface(IID_PPV_ARGS(&fn)); RETURN_IF_FAILED(hr);
			wil::unique_process_heap_string relative;
			hr = GetPathOf (fn, relative, true); RETURN_IF_FAILED(hr);
			hr = Write(cmdLine, relative.get()); RETURN_IF_FAILED(hr);
		}
		else
		{
			// (2) or (3)
			hr = Write(cmdLine, path.get()); RETURN_IF_FAILED(hr);
		}
	}

	return MakeBstrFromStreamOnHGlobal (cmdLine, ppCmdLine);
}

BOOL LUtilFixFilename (wchar_t* strName)
{
	// The shell removes leading spaces, and trailing dots and spaces.
	// Let's remove them too, cause we don't want to create a file that the shell doesn't understand.
	BOOL bFixupDone = FALSE;

	// Trailing dots and spaces.
	wchar_t* p = strName + wcslen(strName);
	if (p[-1] == ' ' || p[-1] == '.')
	{
		while (p > strName && (p[-1] == ' ' || p[-1] == '.'))
			p--;
		p[0] = 0;
		bFixupDone = TRUE;
	}

	// Leading spaces.
	if (strName[0] == ' ')
	{
		wchar_t* first = strName;
		while(first[0] == ' ')
			first++;
		memmove(strName, first, wcslen(first) * sizeof(wchar_t));
		bFixupDone = TRUE;
	}

	return bFixupDone;
}

HRESULT QueryEditProjectFile (IVsHierarchy* hier)
{
	HRESULT hr;

	com_ptr<IPersistFileFormat> pff;
	hr = hier->QueryInterface(&pff); RETURN_IF_FAILED(hr);

	BOOL dirty = FALSE;
	if (SUCCEEDED(pff->IsDirty(&dirty)) && dirty)
		return S_OK;

	VSQueryEditResult fEditVerdict;
	com_ptr<IVsQueryEditQuerySave2> queryEdit;
	hr = serviceProvider->QueryService (SID_SVsQueryEditQuerySave, &queryEdit);
	if (SUCCEEDED(hr))
	{
		wil::unique_cotaskmem_string fullPathName;
		DWORD unused;
		hr = pff->GetCurFile (&fullPathName, &unused); RETURN_IF_FAILED(hr);
		hr = queryEdit->QueryEditFiles (QEF_DisallowInMemoryEdits, 1, fullPathName.addressof(), nullptr, nullptr, &fEditVerdict, nullptr); LOG_IF_FAILED(hr);
		if (FAILED(hr) || (fEditVerdict != QER_EditOK))
			return OLE_E_PROMPTSAVECANCELLED;
	}
	
	return S_OK;
}

HRESULT GetHierarchyWindow (IVsUIHierarchyWindow** ppHierWindow)
{
	com_ptr<IVsUIShell> shell;
	auto hr = serviceProvider->QueryService(SID_SVsUIShell, IID_PPV_ARGS(shell.addressof()));
	if(FAILED(hr))
		return hr;

	com_ptr<IVsWindowFrame> frame;
	hr = shell->FindToolWindow(0, GUID_SolutionExplorer, frame.addressof());
	if(FAILED(hr))
		return hr;

	wil::unique_variant docViewVar;
	hr = frame->GetProperty(VSFPROPID_DocView, &docViewVar);
	if(FAILED(hr))
		return hr;
	if (docViewVar.vt != VT_UNKNOWN)
		return E_UNEXPECTED;

	return docViewVar.punkVal->QueryInterface(ppHierWindow);
}

HRESULT GetPathTo (IChildNode* node, wil::unique_process_heap_string& dir, bool relativeToProjectDir)
{
	HRESULT hr;

	com_ptr<IParentNode> parent;
	hr = node->GetParent(&parent); RETURN_IF_FAILED(hr);
	if (auto hier = parent.try_query<IVsHierarchy>())
	{
		if (!relativeToProjectDir)
		{
			wil::unique_variant projDir;
			hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projDir); RETURN_IF_FAILED(hr);
			dir = wil::make_process_heap_string_nothrow(projDir.bstrVal); RETURN_IF_NULL_ALLOC(dir);
		}
		else
		{
			dir = wil::make_process_heap_string_nothrow(L""); RETURN_IF_NULL_ALLOC(dir);
		}
	}
	else
	{
		com_ptr<IChildNode> parentAsChild;
		hr = parent->QueryInterface(IID_PPV_ARGS(&parentAsChild)); RETURN_IF_FAILED(hr);
		hr = GetPathTo (parentAsChild, dir, relativeToProjectDir);
		wil::unique_variant parentName;
		hr = parentAsChild->GetProperty(VSHPROPID_SaveName, &parentName); RETURN_IF_FAILED(hr);
		hr = wil::str_concat_nothrow(dir, parentName.bstrVal, L"\\"); RETURN_IF_FAILED(hr);
	}
		
	return S_OK;
}

HRESULT GetPathOf (IChildNode* node, wil::unique_process_heap_string& path, bool relativeToProjectDir)
{
	com_ptr<IVsHierarchy> hier;
	auto hr = FindHier(node, IID_PPV_ARGS(hier.addressof())); RETURN_IF_FAILED(hr);
	hr = GetPathTo (node, path, relativeToProjectDir); RETURN_IF_FAILED(hr);
	if (!relativeToProjectDir)
		WI_ASSERT(path && path.get()[0] && wcschr(path.get(), 0)[-1] == L'\\');
	else
		WI_ASSERT(path && path.get()[0] != L'\\');
	wil::unique_variant name;
	hr = node->GetProperty(VSHPROPID_SaveName, &name); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(E_UNEXPECTED, name.vt != VT_BSTR);
	hr = wil::str_concat_nothrow(path, name.bstrVal); RETURN_IF_FAILED(hr);
	return S_OK;
}

HRESULT FindHier (IChildNode* from, REFIID riid, void** ppvHier)
{
	com_ptr<IParentNode> parent;
	auto hr = from->GetParent(&parent); RETURN_IF_FAILED(hr);
	return FindHier(parent, riid, ppvHier);
}

HRESULT FindHier (IParentNode* from, REFIID riid, void** ppvHier)
{
	com_ptr<IVsHierarchy> hier;
	auto hr = from->QueryInterface(IID_PPV_ARGS(hier.addressof()));
	if (hr == S_OK)
		return hier->QueryInterface(riid, ppvHier);
	if (hr != E_NOINTERFACE)
		RETURN_HR(hr);

	com_ptr<IChildNode> pc;
	hr = from->QueryInterface(IID_PPV_ARGS(&pc)); RETURN_IF_FAILED(hr);
	return FindHier(pc, riid, ppvHier);
}

// Enum depth-first (just because it's simpler) pre-order mode (so that parents get their ItemId before children).
static HRESULT SetItemIdsTree (IChildNode* child, IChildNode* childPrevSibling, IParentNode* addTo)
{
	com_ptr<IProjectNode> root;
	auto hr = FindHier(addTo, IID_PPV_ARGS(&root)); RETURN_IF_FAILED(hr);

	stdext::inplace_function<HRESULT(IChildNode*, IChildNode*, IParentNode*)> enumNodeAndChildren;

	enumNodeAndChildren = [root, &enumNodeAndChildren](IChildNode* node, IChildNode* nodePrevSibling, IParentNode* nodeParent) -> HRESULT
		{
			auto hr = node->SetItemId(nodeParent, root->MakeItemId()); RETURN_IF_FAILED(hr);

			com_ptr<IParentNode> nodeAsParent;
			if (SUCCEEDED(node->QueryInterface(&nodeAsParent)))
			{
				IChildNode* childPrevSibling = nullptr;
				for (auto c = nodeAsParent->FirstChild(); c; c = c->Next())
				{
					hr = enumNodeAndChildren(c, childPrevSibling, nodeAsParent); RETURN_IF_FAILED(hr);
					childPrevSibling = c;
				}
			}

			com_ptr<IEnumHierarchyEvents> eventSinks;
			if (SUCCEEDED(root->EnumHierarchyEventSinks(&eventSinks)) && eventSinks)
			{
				com_ptr<IVsHierarchyEvents> sink;
				ULONG fetched;
				while (SUCCEEDED(eventSinks->Next(1, &sink, &fetched)) && fetched)
				{
					VSITEMID itemidSiblingPrev = nodePrevSibling ? nodePrevSibling->GetItemId() : VSITEMID_NIL;
					sink->OnItemAdded (nodeParent->GetItemId(), itemidSiblingPrev, node->GetItemId());

					// Since our expandable status may have changed, we need to refresh it in the UI.
					sink->OnPropertyChanged (nodeParent->GetItemId(), VSHPROPID_Expandable, 0);
				}
			}

			return S_OK;
		};

	return enumNodeAndChildren(child, childPrevSibling, addTo);
}

HRESULT AddFileToParent (IFileNode* child, IParentNode* addTo)
{
	HRESULT hr;
	RETURN_HR_IF(E_UNEXPECTED, child->GetItemId() != VSITEMID_NIL);

	IChildNode* prevChild = nullptr;
	if (!addTo->FirstChild())
		addTo->SetFirstChild(child);
	else
	{
		com_ptr<IFileNodeProperties> childProps;
		hr = child->QueryInterface(IID_PPV_ARGS(childProps.addressof())); RETURN_IF_FAILED(hr);
		wil::unique_bstr childPath;
		hr = childProps->get_Path(&childPath); RETURN_IF_FAILED(hr);
		const wchar_t* childName = PathFindFileName(childPath.get());

		// Do we need to insert it in the first position?
		wil::unique_variant name;
		if (!wil::try_com_query_nothrow<IFolderNode>(addTo->FirstChild())
			&& SUCCEEDED(addTo->FirstChild()->GetProperty(VSHPROPID_SaveName, &name))
			&& _wcsicmp(childName, name.bstrVal) < 0)
		{
			// Yes
			child->SetNext(addTo->FirstChild());
			addTo->SetFirstChild(child);
		}
		else
		{
			// No, insert it after some existing node
			IChildNode* insertAfter = addTo->FirstChild();

			// Skip any existing folder nodes.
			while (insertAfter->Next() && wil::try_com_query_nothrow<IFolderNode>(insertAfter->Next()))
				insertAfter = insertAfter->Next();

			if (insertAfter->Next())
			{
				while (insertAfter->Next()
					&& SUCCEEDED(insertAfter->Next()->GetProperty(VSHPROPID_SaveName, &name))
					&& _wcsicmp(childName, name.bstrVal) > 0)
				{
					insertAfter = insertAfter->Next();
				}
			}

			child->SetNext(insertAfter->Next());
			insertAfter->SetNext(child);

			prevChild = insertAfter;
		}
	}

	if (addTo->GetItemId() != VSITEMID_NIL)
	{
		// Adding it to a hierarchy.
		hr = SetItemIdsTree(child, prevChild, addTo); RETURN_IF_FAILED(hr);
	}

	return S_OK;
}

HRESULT GetOrCreateChildFolder (IParentNode* parent, std::wstring_view folderName, bool createDirectoryOnFileSystem, IFolderNode** ppFolder)
{
	HRESULT hr;

	auto createDir = [](IFolderNode* f) -> HRESULT
		{
			wil::unique_process_heap_string path;
			auto hr = GetPathOf (f, path); RETURN_IF_FAILED(hr);
			BOOL bres = ::CreateDirectoryW (path.get(), nullptr);
			if (bres)
				return S_OK;
			DWORD le = GetLastError();
			if (le == ERROR_ALREADY_EXISTS)
				return S_OK;
			RETURN_WIN32(le);
		};

	com_ptr<IChildNode> insertAfter;
	com_ptr<IChildNode> insertBefore = parent->FirstChild();
	while (insertBefore)
	{
		auto insertBeforeAsFolder = wil::try_com_query_nothrow<IFolderNode>(insertBefore);
		if (!insertBeforeAsFolder)
			break;

		wil::unique_variant name;
		if (SUCCEEDED(insertBeforeAsFolder->GetProperty(VSHPROPID_SaveName, &name)) && (V_VT(&name) == VT_BSTR))
		{
			int cmpRes = folderName.compare(V_BSTR(&name));
			if (cmpRes == 0)
			{
				if (createDirectoryOnFileSystem)
				{
					hr = createDir(insertBeforeAsFolder); RETURN_IF_FAILED(hr);
				}

				*ppFolder = insertBeforeAsFolder.detach();
				return S_OK;
			}
			
			if (cmpRes < 0)
				break;
		}

		insertAfter = insertBefore;
		insertBefore = insertBefore->Next();
	}

	com_ptr<IFolderNode> newFolder;
	hr = MakeFolderNode (&newFolder); RETURN_IF_FAILED(hr);
	auto name = wil::unique_bstr(SysAllocStringLen(folderName.data(), (UINT)folderName.size())); RETURN_IF_NULL_ALLOC(name);
	hr = newFolder.try_query<IFolderNodeProperties>()->put_Name(name.get()); RETURN_IF_FAILED(hr);
	newFolder->SetNext(insertBefore);
	if (!insertAfter)
		parent->SetFirstChild(newFolder);
	else
		insertAfter->SetNext(newFolder);
	hr = SetItemIdsTree(newFolder, insertAfter, parent); RETURN_IF_FAILED(hr);
	if (createDirectoryOnFileSystem)
	{
		hr = createDir(newFolder); RETURN_IF_FAILED(hr);
	}
	*ppFolder = newFolder.detach();
	return S_OK;
}

// Enum depth-first (just because it's simpler) post-order mode (so that children clear their ItemId before parent).
static HRESULT ClearItemIdsTree (IProjectNode* root, IChildNode* child)
{
	stdext::inplace_function<HRESULT(IChildNode*)> enumNodeAndChildren;

	enumNodeAndChildren = [root, &enumNodeAndChildren](IChildNode* node) -> HRESULT
		{
			HRESULT hr;

			com_ptr<IParentNode> nodeAsParent;
			if (SUCCEEDED(node->QueryInterface(&nodeAsParent)))
			{
				for (auto c = nodeAsParent->FirstChild(); c; c = c->Next())
				{
					hr = enumNodeAndChildren(c); RETURN_IF_FAILED(hr);
				}
			}

			VSITEMID itemId = node->GetItemId();
			hr = node->ClearItemId(); RETURN_IF_FAILED(hr);

			com_ptr<IEnumHierarchyEvents> eventSinks;
			if (SUCCEEDED(root->EnumHierarchyEventSinks(&eventSinks)) && eventSinks)
			{
				com_ptr<IVsHierarchyEvents> sink;
				ULONG fetched;
				while (SUCCEEDED(eventSinks->Next(1, &sink, &fetched)) && fetched)
					sink->OnItemDeleted(itemId);
			}

			return S_OK;
		};

	return enumNodeAndChildren(child);
}

HRESULT RemoveChildFromParent (IProjectNode* root, IChildNode* node)
{
	HRESULT hr;

	com_ptr<IParentNode> parent;
	hr = node->GetParent(&parent); RETURN_IF_FAILED(hr);

	hr = ClearItemIdsTree(root, node); RETURN_IF_FAILED(hr);

	if (parent->FirstChild() == node)
		parent->SetFirstChild(node->Next());
	else
	{
		auto prev = parent->FirstChild();
		while(prev->Next() && (prev->Next() != node))
			prev = prev->Next();
		RETURN_HR_IF(E_UNEXPECTED, !prev->Next());
		prev->SetNext(node->Next());
	}

	node->SetNext(nullptr);

	return S_OK;
}

HRESULT CreatePathOfNode (IParentNode* node, wil::unique_process_heap_string& path)
{
	stdext::inplace_function<HRESULT(IParentNode* pn)> createDirRecursively;
	createDirRecursively = [&createDirRecursively, &path](IParentNode* node) -> HRESULT
		{
			if (node->GetItemId() == VSITEMID_ROOT)
			{
				com_ptr<IVsHierarchy> hier;
				auto hr = node->QueryInterface(IID_PPV_ARGS(&hier)); RETURN_IF_FAILED(hr);
				wil::unique_variant projectDir;
				hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr);
				if (!PathFileExists(projectDir.bstrVal))
					return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
				path = wil::make_process_heap_string_nothrow(projectDir.bstrVal); RETURN_IF_NULL_ALLOC(path);
				return S_OK;
			}
			else
			{
				com_ptr<IChildNode> cn;
				auto hr = node->QueryInterface(IID_PPV_ARGS(&cn)); RETURN_IF_FAILED(hr);
				com_ptr<IParentNode> parent;
				hr = cn->GetParent(&parent); RETURN_IF_FAILED(hr);
				hr = createDirRecursively(parent);
				if (FAILED(hr))
					return hr;
				wil::unique_variant saveName;
				hr = cn->GetProperty(VSHPROPID_SaveName, &saveName); RETURN_IF_FAILED(hr);
				hr = wil::str_concat_nothrow(path, saveName.bstrVal, L"\\"); RETURN_IF_FAILED(hr);
				DWORD attrs = GetFileAttributes(path.get());
				if (attrs == INVALID_FILE_ATTRIBUTES)
				{
					DWORD gle = GetLastError();
					if (gle != ERROR_FILE_NOT_FOUND)
						return HRESULT_FROM_WIN32(gle);

					BOOL bres = CreateDirectory(path.get(), nullptr);
					if (!bres)
						return HRESULT_FROM_WIN32(GetLastError());
					return S_OK;
				}
				else
				{
					if ((attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
						return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
					return S_OK;
				}
			}
		};

	return createDirRecursively(node);
}

HRESULT GetItems (IParentNode* parent, SAFEARRAY** itemsOut)
{
	HRESULT hr;
	*itemsOut = nullptr;

	vector_nothrow<com_ptr<IDispatch>> nodes;

	for (auto c = parent->FirstChild(); c; c = c->Next())
	{
		if (auto file = wil::try_com_query_nothrow<IFileNodeProperties>(c))
		{
			bool pushed = nodes.try_push_back(std::move(file)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}
		else if (auto folder = wil::try_com_query_nothrow<IFolderNodeProperties>(c))
		{
			bool pushed = nodes.try_push_back(std::move(folder)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	SAFEARRAYBOUND bound;
	bound.cElements = nodes.size();
	bound.lLbound = 0;
	auto sa = unique_safearray(SafeArrayCreate(VT_DISPATCH, 1, &bound)); RETURN_HR_IF(E_OUTOFMEMORY, !sa);
	for (LONG i = 0; i < (LONG)nodes.size(); i++)
	{
		hr = SafeArrayPutElement(sa.get(), &i, nodes[i].get()); RETURN_IF_FAILED(hr);
	}

	*itemsOut = sa.release();
	return S_OK;
}

HRESULT PutItems (SAFEARRAY* sa, IParentNode* parent)
{
	HRESULT hr;

	VARTYPE vt;
	hr = SafeArrayGetVartype(sa, &vt); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(E_NOTIMPL, vt != VT_DISPATCH);
	UINT dim = SafeArrayGetDim(sa);
	RETURN_HR_IF(E_NOTIMPL, dim != 1);
	LONG lbound;
	hr = SafeArrayGetLBound(sa, 1, &lbound); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(E_NOTIMPL, lbound != 0);
	LONG ubound;
	hr = SafeArrayGetUBound(sa, 1, &ubound); RETURN_IF_FAILED(hr);

	// We don't support replacing items with this function, we only support adding them once.
	RETURN_HR_IF(E_UNEXPECTED, parent->FirstChild() != nullptr);

	for (LONG i = 0; i <= ubound; i++)
	{
		com_ptr<IDispatch> child;
		hr = SafeArrayGetElement (sa, &i, child.addressof()); RETURN_IF_FAILED(hr);
		if (auto node = child.try_query<IChildNode>())
		{
			IChildNode* insertAfter = parent->FirstChild();
			if (!insertAfter)
			{
				parent->SetFirstChild(node);
			}
			else
			{
				while (insertAfter->Next())
					insertAfter = insertAfter->Next();
				node->SetNext(insertAfter->Next());
				insertAfter->SetNext(node);
			}

			if (parent->GetItemId() != VSITEMID_NIL)
				hr = SetItemIdsTree(node, insertAfter, parent); RETURN_IF_FAILED(hr);
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	// This is meant to be called only from LoadXml, no need to set dirty flag or send notifications.

	return S_OK;
}

HRESULT CreateFileFromTemplate (LPCWSTR fromPath, LPCWSTR toPath, IProjectMacroResolver* macroResolver)
{
	HRESULT hr;

	com_ptr<IStream> fromStream;
	hr = SHCreateStreamOnFileEx (fromPath, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &fromStream); RETURN_IF_FAILED(hr);

	com_ptr<IStream> toStream;
	hr = SHCreateStreamOnFileEx (toPath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &toStream); RETURN_IF_FAILED(hr);

	char ch;
	ULONG cbRead;
	while(true)
	{
		hr = fromStream->Read(&ch, 1, &cbRead); RETURN_IF_FAILED(hr);
		if (!cbRead)
			break;

		if (ch == '%')
		{
			char macro[32];
			uint32_t i = 0;
			while(true)
			{
				hr = fromStream->Read(&ch, 1, &cbRead); RETURN_IF_FAILED(hr);
				RETURN_HR_IF(E_UNEXPECTED, cbRead == 0);
				if (ch != '%')
				{
					RETURN_HR_IF(E_UNEXPECTED, i == _countof(macro) - 1);
					macro[i++] = ch;
				}
				else
				{
					macro[i] = 0;
					break;
				}
			}

			wil::unique_cotaskmem_ansistring value;
			hr = macroResolver->ResolveMacro(macro, macro + i, &value); RETURN_IF_FAILED(hr);
			hr = toStream->Write (value.get(), strlen(value.get()), &cbRead); RETURN_IF_FAILED(hr);
		}
		else
		{
			ULONG cbWritten;
			hr = toStream->Write(&ch, 1, &cbWritten); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_UNEXPECTED, cbWritten != 1);
		}
	}

	return S_OK;
}

IFileNode* FindChildFileByName (IParentNode* parent, const wchar_t* fileName)
{
	// Skip all folders.
	auto child = parent->FirstChild();
	while (child && !wil::try_com_query_nothrow<IFileNode>(child))
		child = child->Next();
	if (!child)
		return nullptr;

	while(child)
	{
		auto fn = wil::try_com_query_nothrow<IFileNode>(child);
		if (!fn)
			return nullptr;

		wil::unique_variant n;
		if (SUCCEEDED(fn->GetProperty(VSHPROPID_SaveName, &n)) && V_VT(&n) == VT_BSTR && !_wcsicmp(fileName, V_BSTR(&n)))
			return fn;

		child = child->Next();
	}

	return nullptr;
}

HRESULT MakeFileNodeForExistingFile (LPCWSTR path, IFileNode** ppFile)
{
	com_ptr<IFileNode> file;
	auto hr = MakeFileNode(&file); RETURN_IF_FAILED(hr);
	com_ptr<IFileNodeProperties> fileProps;
	hr = file->QueryInterface(&fileProps); RETURN_IF_FAILED(hr);
	hr = fileProps->put_Path(wil::make_bstr_nothrow(path).get()); RETURN_IF_FAILED(hr);
	auto buildTool = _wcsicmp(PathFindExtension(path), L".asm") ? BuildToolKind::None : BuildToolKind::Assembler;
	hr = fileProps->put_BuildTool(buildTool); RETURN_IF_FAILED(hr);
	*ppFile = file.detach();
	return S_OK;
}

