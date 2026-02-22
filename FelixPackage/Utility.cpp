
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/inplace_function.h"
#include "../FelixPackageUi/resource.h"
#include <string_view>

const wchar_t MacroOutputName[] = L"OUTPUT_NAME";
const wchar_t MacroProjectName[] = L"PROJECT_NAME";
const wchar_t MacroProjectDir[] = L"PROJECT_DIR";
const wchar_t MacroConfigName[] = L"CONFIG_NAME";
const wchar_t MacroOutputDir[] = L"OUTPUT_DIR";
const wchar_t MacroOutputFilename[] = L"OUTPUT_FILENAME";

static HRESULT SetItemIdsTree (IProjectNode* root, IChildNode* child, IChildNode* childPrevSibling, IParentNode* addTo);
HRESULT InsertFolderNode (IProjectNode* proj, IParentNode* parent, IChildNode* insertBefore, IChildNode* insertAfter, IFolderNode* newFolder);

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

static HRESULT Write (ISequentialStream* stream, const wchar_t* str1, const wchar_t* str2)
{
	auto hr = stream->Write(str1, (ULONG)wcslen(str1) * sizeof(wchar_t), nullptr);
	if (SUCCEEDED(hr))
		hr = stream->Write(str2, (ULONG)wcslen(str2) * sizeof(wchar_t), nullptr);
	return hr;
}

HRESULT GetCountOfBuildToolAssemblerFiles(IProjectNode* project, UINT* pCount)
{
	HRESULT hr;

	*pCount = 0;
	stdext::inplace_function<HRESULT(IParentNode*)> searchDescendants;
	searchDescendants = [&searchDescendants, pCount](IParentNode* parent) -> HRESULT
		{
			for (auto c = parent->FirstChild(); c; c = c->Next())
			{
				if (auto file = wil::try_com_query_nothrow<IFileNodeProperties>(c))
				{
					BuildToolKind tool;
					auto hr = file->get_BuildTool(&tool); RETURN_IF_FAILED(hr);
					if (tool == BuildToolKind::Assembler)
						(*pCount)++;
				}
				else if (auto cAsParent = wil::try_com_query_nothrow<IParentNode>(c))
				{
					if (c->GetItemId() != VSITEMID_GENFILES)
					{
						auto hr = searchDescendants(cAsParent); RETURN_IF_FAILED(hr);
					}
				}
			}

			return S_FALSE;
		};
	return searchDescendants(project);
}

HRESULT GetActiveCfgGeneratePrePostIncludeFiles (IVsHierarchy* hier)
{
	com_ptr<IVsSolutionBuildManager> buildManager;
	auto hr = serviceProvider->QueryService(SID_SVsSolutionBuildManager, IID_PPV_ARGS(&buildManager)); RETURN_IF_FAILED(hr);
	com_ptr<IVsProjectCfg> activeConfig;
	hr = buildManager->FindActiveProjectCfg (nullptr, nullptr, hier, &activeConfig); RETURN_IF_FAILED_EXPECTED(hr);
	// The call to FindActiveProjectCfg will fail if we get here while loading a project from XML.
	// See explanation in comment in ProjectNode::put_Configurations().
	com_ptr<IProjectConfig> activeCfg;
	hr = activeConfig->QueryInterface(&activeCfg); RETURN_IF_FAILED(hr);
	VARIANT_BOOL generate;
	hr = activeCfg->AsmProps()->get_GeneratePrePostIncludeFiles(&generate); RETURN_IF_FAILED(hr);
	return generate ? S_OK : S_FALSE;
}

static HRESULT GeneratePrePostIncludeFilesInner (IProjectNode* proj, IProjectConfig* macroResolver)
{
	wil::unique_variant projDir;
	auto hr = proj->AsHierarchy()->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projDir); RETURN_IF_FAILED(hr);
	wil::unique_process_heap_string genFilesDir;
	hr = wil::str_concat_nothrow(genFilesDir, V_BSTR(&projDir), genFilesStr); RETURN_IF_FAILED(hr);
	hr = EnsureDirectoryExists(genFilesDir.get()); RETURN_IF_FAILED_EXPECTED(hr);

	com_ptr<IFolderNode> folder;
	com_ptr<IChildNode> insertBefore, insertAfter;
	hr = FindFolderNodeOrInsertLocation (proj, proj, genFilesStr.get(), &folder, insertBefore, insertAfter); RETURN_IF_FAILED(hr);
	if (hr == S_FALSE)
	{
		hr = MakeFolderNodeGenerated(&folder); RETURN_IF_FAILED(hr);
		InsertFolderNode(proj, proj, insertBefore, insertAfter, folder); RETURN_IF_FAILED(hr);
	}

	wil::unique_process_heap_string packageDir;
	hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, packageDir); RETURN_IF_FAILED(hr);
	*PathFindFileName(packageDir.get()) = 0;

	struct Info { wil::unique_bstr& filename; const wchar_t* templateFilename; bool post; };
	static const Info info[2] = { { preincludeFilename, L"preinclude.asm", false }, { postincludeFilename, L"postinclude.asm", true } };
	for (auto& i : info)
	{
		com_ptr<IFileNode> file = FindChildFileByName(proj, folder->AsParentNode(), i.filename.get());
		if (!file)
		{
			hr = MakeFileNodePrePostInc(i.post, &file); RETURN_IF_FAILED(hr);
			hr = AddFileToParent(proj, file, folder->AsParentNode()); RETURN_IF_FAILED(hr);
		}

		wil::unique_process_heap_string templatePath;
		hr = wil::str_concat_nothrow(templatePath, packageDir, L"Templates\\", i.templateFilename); RETURN_IF_FAILED(hr);

		wil::unique_process_heap_string includePath;
		hr = GetPathOf(proj, file, includePath); RETURN_IF_FAILED(hr);
		hr = CreateFileFromTemplate(templatePath.get(), includePath.get(), macroResolver); RETURN_IF_FAILED(hr);
	}

	return S_OK;
}

HRESULT GeneratePrePostIncludeFiles (IProjectNode* project)
{
	HRESULT hr;

	com_ptr<IVsOutputWindowPane> op;
	hr = serviceProvider->QueryService(SID_SVsGeneralOutputWindowPane, IID_PPV_ARGS(&op)); RETURN_IF_FAILED(hr);
	hr = op->Activate(); RETURN_IF_FAILED(hr);
	com_ptr<IVsOutputWindowPane2> op2;
	hr = op->QueryInterface(&op2); RETURN_IF_FAILED(hr);

	wil::unique_variant projectName;
	hr = project->AsHierarchy()->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &projectName); RETURN_IF_FAILED(hr);

	com_ptr<IVsSolutionBuildManager> buildManager;
	hr = serviceProvider->QueryService(SID_SVsSolutionBuildManager, IID_PPV_ARGS(&buildManager)); RETURN_IF_FAILED(hr);
	com_ptr<IVsProjectCfg> projectConfig;
	hr = buildManager->FindActiveProjectCfg (nullptr, nullptr, project->AsHierarchy(), &projectConfig); RETURN_IF_FAILED(hr);
	com_ptr<IProjectConfig> config;
	hr = projectConfig->QueryInterface(IID_PPV_ARGS(&config)); RETURN_IF_FAILED(hr);

	wil::unique_bstr str;
	if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_GEN_PRE_POST_MESSAGE, &str)))
	{
		wil::unique_bstr configName;
		hr = config->AsVsProjectConfig()->get_DisplayName(&configName); RETURN_IF_FAILED(hr);

		wil::unique_process_heap_string message;
		if (SUCCEEDED(wil::str_printf_nothrow(message, str.get(), configName)))
			op2->OutputTaskItemStringEx2(message.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
				nullptr, 0, nullptr, 0, 0, projectName.bstrVal, nullptr, nullptr);
	}
	
	hr = GeneratePrePostIncludeFilesInner (project, config);
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

HRESULT DeletePrePostIncludeFiles (IProjectNode* project)
{
	HRESULT hr;

	com_ptr<IVsOutputWindowPane> op;
	hr = serviceProvider->QueryService(SID_SVsGeneralOutputWindowPane, IID_PPV_ARGS(&op)); RETURN_IF_FAILED(hr);
	hr = op->Activate(); RETURN_IF_FAILED(hr);
	com_ptr<IVsOutputWindowPane2> op2;
	hr = op->QueryInterface(&op2); RETURN_IF_FAILED(hr);

	wil::unique_variant projectName;
	hr = project->AsHierarchy()->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &projectName); RETURN_IF_FAILED(hr);

	com_ptr<IFolderNode> genFilesFolder;
	for (auto c = project->FirstChild(); c; c = c->Next())
	{
		if (c->GetItemId() == VSITEMID_GENFILES)
		{
			hr = c->QueryInterface(&genFilesFolder); RETURN_IF_FAILED(hr);
			break;
		}
	}

	if (genFilesFolder)
	{
		wil::unique_bstr str;
		if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_REMOVING_GEN_FILES, &str)))
			op2->OutputTaskItemStringEx2 (str.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
				nullptr, 0, nullptr, 0, 0, projectName.bstrVal, nullptr, nullptr);

		com_ptr<IVsRunningDocumentTable> rdt;
		if (SUCCEEDED(serviceProvider->QueryService(SID_SVsRunningDocumentTable, IID_PPV_ARGS(&rdt))))
		{
			com_ptr<IVsSolution> solution;
			hr = serviceProvider->QueryService(SID_SVsSolution, &solution); RETURN_IF_FAILED(hr);

			for (auto c = genFilesFolder->AsParentNode()->FirstChild(); c; c = c->Next())
			{
				wil::unique_process_heap_string path;
				hr = GetPathOf (project, c, path); RETURN_IF_FAILED(hr);
				VSDOCCOOKIE docCookie;
				hr = rdt->FindAndLockDocument(RDT_NoLock, path.get(), nullptr, nullptr, nullptr, &docCookie);
				if (SUCCEEDED(hr) && docCookie != VSDOCCOOKIE_NIL)
				{
					hr = solution->CloseSolutionElement (SLNSAVEOPT_NoSave, project->AsHierarchy(), docCookie); LOG_IF_FAILED(hr);
				}
			}
		}

		wil::unique_process_heap_string genDirPath;
		hr = GetPathOf(project, genFilesFolder, genDirPath); RETURN_IF_FAILED(hr);

		hr = RemoveChildFromParent(project, genFilesFolder); RETURN_IF_FAILED(hr);

		auto buffer = wil::str_printf_failfast<wil::unique_process_heap_string>(L"%s%c", genDirPath.get(), L'\0');
		SHFILEOPSTRUCT file_op = { .wFunc = FO_DELETE, .pFrom = buffer.get(), .fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT };
		int ires = SHFileOperation(&file_op); WI_ASSERT(!ires);

		if (SUCCEEDED(shell->LoadPackageString(CLSID_FelixPackage, IDS_GEN_PRE_POST_MESSAGE_DONE, &str)))
			op2->OutputTaskItemStringEx2(str.get(), (VSTASKPRIORITY)0, (VSTASKCATEGORY)0,
				nullptr, 0, nullptr, 0, 0, projectName.bstrVal, nullptr, nullptr);
	}

	return S_OK;
}

// Returns S_FALSE when there are no files with BuildTool=Assembler.
FELIX_API HRESULT MakeSjasmCommandLine (IProjectNode* project, IProjectConfig* config, IProjectConfigAssemblerProperties* asmPropsOverride, BSTR* ppCmdLine)
{
	HRESULT hr;

	wil::unique_variant projectDir;
	hr = project->AsHierarchy()->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, projectDir.addressof()); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);

	wil::unique_process_heap_string packageDir;
	hr = wil::GetModuleFileNameW((HMODULE)&__ImageBase, packageDir); RETURN_IF_FAILED(hr);
	*PathFindFileName(packageDir.get()) = 0;

	wil::unique_bstr outputDirUnresolved;
	hr = config->GeneralProps()->get_OutputDirectory(&outputDirUnresolved); RETURN_IF_FAILED(hr);
	wil::unique_process_heap_string output_dir;
	hr = ResolveMacros(outputDirUnresolved.get(), config, output_dir); RETURN_IF_FAILED(hr);
	hr = EnsureDirHasBackslash (output_dir.get(), output_dir); RETURN_IF_FAILED(hr);

	vector_nothrow<com_ptr<IFileNode>> asmFiles;
	com_ptr<IFileNode> preIncludeFile, postIncludeFile;

	stdext::inplace_function<HRESULT(IParentNode*)> enumDescendants;

	enumDescendants = [&enumDescendants, &asmFiles, &preIncludeFile, &postIncludeFile](IParentNode* parent) -> HRESULT
		{
			for (auto c = parent->FirstChild(); c; c = c->Next())
			{
				if (c->GetItemId() == VSITEMID_GENFILES)
				{
					com_ptr<IParentNode> genFilesFolder;
					auto hr = c->QueryInterface(&genFilesFolder); RETURN_IF_FAILED(hr);
					for (auto cc = genFilesFolder->FirstChild(); cc; cc = cc->Next())
					{
						if (cc->GetItemId() == VSITEMID_PREINCLUDE)
							RETURN_IF_FAILED(cc->QueryInterface(&preIncludeFile));
						else if (cc->GetItemId() == VSITEMID_POSTINCLUDE)
							RETURN_IF_FAILED(cc->QueryInterface(&postIncludeFile));
					}
					RETURN_HR_IF(E_UNEXPECTED, !preIncludeFile);
					RETURN_HR_IF(E_UNEXPECTED, !postIncludeFile);
				}
				else if (auto file = wil::try_com_query_nothrow<IFileNodeProperties>(c))
				{
					BuildToolKind tool;
					auto hr = file->get_BuildTool(&tool); RETURN_IF_FAILED(hr);
					if (tool == BuildToolKind::Assembler)
					{
						com_ptr<IFileNode> fileNode;
						hr = file->QueryInterface(&fileNode); RETURN_IF_FAILED(hr);
						bool pushed = asmFiles.try_push_back(std::move(fileNode)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
					}
				}
				else if (auto cAsParent = wil::try_com_query_nothrow<IParentNode>(c))
				{
					auto hr = enumDescendants(cAsParent); RETURN_IF_FAILED(hr);
				}
			}

			return S_OK;
		};
	hr = enumDescendants(project); RETURN_IF_FAILED(hr);

	if (asmFiles.empty())
		return (*ppCmdLine = nullptr), S_FALSE;

	com_ptr<IStream> cmdLine;
	hr = CreateStreamOnHGlobal (nullptr, TRUE, &cmdLine); RETURN_IF_FAILED(hr);

	bool hasSpaces = !!wcschr(packageDir.get(), L' ');
	if (hasSpaces)
	{
		hr = Write(cmdLine, L"\""); RETURN_IF_FAILED(hr);
	}
	hr = Write(cmdLine, packageDir.get()); RETURN_IF_FAILED(hr);
	hr = Write(cmdLine, L"sjasmplus.exe"); RETURN_IF_FAILED(hr);
	if (hasSpaces)
	{
		hr = Write(cmdLine, L"\""); RETURN_IF_FAILED(hr);
	}
	hr = Write(cmdLine, L" --fullpath"); RETURN_IF_FAILED(hr);

	// We launch sjasmplus in the project directory, so that's the CWD of the sjasmplus.exe process.
	// It looks for its input files, and generates its output files, in directories relative to the CWD.
	// But we gave our user a way to specify a different output directory - the property "OutputDirectory".
	// Thus for output files we must first construct a full path of the output files, then try to make it relative to the project directory;
	// If it's not possible to make it relative to the project directory, then we pass it to sjasmplus as an absolute path.
	auto addOutputPathParam = [&cmdLine, output_dir=output_dir.get(), project_dir=projectDir.bstrVal](const wchar_t* paramName, const wchar_t* output_filename) -> HRESULT
		{
			RETURN_HR_IF(ERROR_BAD_PATHNAME, output_dir[wcslen(output_dir) - 1] != L'\\');
			if (!PathIsRelative(output_dir) && PathIsSameRootW(output_dir, project_dir))
			{
				// Output dir is an absolute path on same drive as the project dir. We can generate an output path relative to the project dir.
				auto outputFilePath = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePath);
				auto pres = PathCombine (outputFilePath.get(), output_dir, output_filename); RETURN_HR_IF(ERROR_BAD_PATHNAME, !pres);
				auto outputFilePathRelative = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePathRelative);
				BOOL bRes = PathRelativePathToW (outputFilePathRelative.get(), project_dir, FILE_ATTRIBUTE_DIRECTORY, outputFilePath.get(), 0);
				if (!bRes)
					return SetFelixErrorInfo (E_INVALIDARG, IDS_CANNOT_MAKE_RELATIVE_PATH_S_S, outputFilePath.get(), project_dir);
				auto hr = Write(cmdLine, paramName); RETURN_IF_FAILED(hr);
				hr = Write(cmdLine, outputFilePathRelative.get()); RETURN_IF_FAILED(hr);
			}
			else
			{
				// Output dir is a relative path (which we consider to be relative to the project dir), or it is a path on a different drive.
				auto hr = Write(cmdLine, paramName); RETURN_IF_FAILED(hr);
				hr = Write(cmdLine, output_dir); RETURN_IF_FAILED(hr);
				hr = Write(cmdLine, output_filename); RETURN_IF_FAILED(hr);
			}
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
	hr = config->GeneralProps()->get_OutputFileType(&outputFileType); RETURN_IF_FAILED(hr);
	if (outputFileType == OutputFileType::Binary)
	{
		// --raw=...
		wil::unique_bstr output_filename;
		hr = config->GeneralProps()->get_OutputFilename(&output_filename); RETURN_IF_FAILED(hr);
		hr = addOutputPathParam (L" --raw=", output_filename.get()); RETURN_IF_FAILED(hr);
	}

	// --sld=...
	wil::unique_process_heap_string sld_filename;
	wil::unique_bstr outputName;
	hr = config->GeneralProps()->get_OutputName(&outputName); RETURN_IF_FAILED(hr);
	hr = wil::str_printf_nothrow(sld_filename, L"%s.sld", outputName.get()); RETURN_IF_FAILED(hr);
	hr = ResolveMacros (sld_filename.get(), config, sld_filename); RETURN_IF_FAILED(hr);

	hr = addOutputPathParam (L" --sld=", sld_filename.get()); RETURN_IF_FAILED(hr);

	// --outprefix
	hr = addOutputPathParam (L" --outprefix=", L""); RETURN_IF_FAILED(hr);

	// --lst
	VARIANT_BOOL saveListing;
	hr = asmProps->get_SaveListing(&saveListing); RETURN_IF_FAILED(hr);
	if (saveListing)
	{
		wil::unique_bstr listingFilename;
		hr = asmProps->get_ListingFilename(&listingFilename); RETURN_IF_FAILED(hr);
		if (listingFilename && listingFilename.get()[0])
		{
			hr = addOutputPathParam (L" --lst=", listingFilename.get()); RETURN_IF_FAILED(hr);
		}
	}

	wil::unique_bstr additionalOpts;
	hr = asmProps->get_AdditionalAssemblerOptions(&additionalOpts); RETURN_IF_FAILED(hr);
	if (additionalOpts && additionalOpts.get()[0])
	{
		hr = Write(cmdLine, L" ", additionalOpts.get()); RETURN_IF_FAILED(hr);
	}

	// input files
	wil::unique_bstr filename;
	if (preIncludeFile)
	{
		hr = preIncludeFile->GetCanonicalName(project, &filename); RETURN_IF_FAILED(hr);
		hr = Write(cmdLine, L" ", filename.get()); RETURN_IF_FAILED(hr);
	}	
	for (auto& asmFile : asmFiles)
	{
		hr = asmFile->GetCanonicalName(project, &filename); RETURN_IF_FAILED(hr);
		hr = Write(cmdLine, L" ", filename.get()); RETURN_IF_FAILED(hr);
	}
	if (postIncludeFile)
	{
		hr = postIncludeFile->GetCanonicalName(project, &filename); RETURN_IF_FAILED(hr);
		hr = Write(cmdLine, L" ", filename.get()); RETURN_IF_FAILED(hr);
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
	HRESULT hr;
	com_ptr<IVsWindowFrame> frame;
	hr = uiShell->FindToolWindow(0, GUID_SolutionExplorer, frame.addressof());
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

HRESULT GetPathTo (IProjectNode* proj, IChildNode* node, wil::unique_process_heap_string& dir, bool relativeToProjectDir)
{
	HRESULT hr;

	wil::unique_bstr filePath;
	if (auto fileNode = wil::try_com_query_nothrow<IFileNodeProperties>(node); fileNode && SUCCEEDED(fileNode->get_Path(&filePath)))
		WI_ASSERT(PathIsFileSpec(filePath.get())); // Only case (1) supported in this function

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
		hr = GetPathTo (proj, parentAsChild, dir, relativeToProjectDir);
		wil::unique_variant parentName;
		hr = parentAsChild->GetProperty(proj, VSHPROPID_SaveName, &parentName); RETURN_IF_FAILED(hr);
		hr = wil::str_concat_nothrow(dir, parentName.bstrVal, L"\\"); RETURN_IF_FAILED(hr);
	}
		
	return S_OK;
}

HRESULT GetPathOf (IProjectNode* proj, IChildNode* node, wil::unique_process_heap_string& path, bool relativeToProjectDir)
{
	HRESULT hr;
	hr = GetPathTo (proj, node, path, relativeToProjectDir); RETURN_IF_FAILED(hr);
	if (!relativeToProjectDir)
		WI_ASSERT(path && path.get()[0] && wcschr(path.get(), 0)[-1] == L'\\');
	else
		WI_ASSERT(path && path.get()[0] != L'\\');
	wil::unique_variant name;
	hr = node->GetProperty(proj, VSHPROPID_SaveName, &name); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(E_UNEXPECTED, name.vt != VT_BSTR);
	hr = wil::str_concat_nothrow(path, name.bstrVal); RETURN_IF_FAILED(hr);
	return S_OK;
}

// Enum depth-first (just because it's simpler) pre-order mode (so that parents get their ItemId before children).
static HRESULT SetItemIdsTree (IProjectNode* root, IChildNode* child, IChildNode* childPrevSibling, IParentNode* addTo)
{
	stdext::inplace_function<HRESULT(IChildNode*, IChildNode*, IParentNode*)> enumNodeAndChildren;

	enumNodeAndChildren = [root, &enumNodeAndChildren](IChildNode* node, IChildNode* nodePrevSibling, IParentNode* nodeParent) -> HRESULT
		{
			root->NotifyNodeInsertingIntoHier(node);

			auto hr = node->SetItemId(root, nodeParent); RETURN_IF_FAILED(hr);

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

			root->NotifyNodeInsertedIntoHier (nodeParent, nodePrevSibling, node);
			return S_OK;
		};

	return enumNodeAndChildren(child, childPrevSibling, addTo);
}

HRESULT AddFileToParent (IProjectNode* proj, IFileNode* child, IParentNode* addTo)
{
	HRESULT hr;
	RETURN_HR_IF(E_UNEXPECTED, child->GetItemId() != VSITEMID_NIL);

	IChildNode* prevChild = nullptr;
	if (!addTo->FirstChild())
		addTo->SetFirstChild(child);
	else
	{
		wil::unique_bstr childPath;
		hr = child->GetPath(&childPath); RETURN_IF_FAILED(hr);
		const wchar_t* childName = PathFindFileName(childPath.get());

		// Do we need to insert it in the first position?
		wil::unique_variant name;
		if (!wil::try_com_query_nothrow<IFolderNode>(addTo->FirstChild())
			&& SUCCEEDED(addTo->FirstChild()->GetProperty(proj, VSHPROPID_SaveName, &name))
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
					&& SUCCEEDED(insertAfter->Next()->GetProperty(proj, VSHPROPID_SaveName, &name))
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
		hr = SetItemIdsTree (proj, child, prevChild, addTo); RETURN_IF_FAILED(hr);

		// Since our expandable status may have changed, we need to refresh it in the UI.
		proj->NotifyPropertyChangedHierNode (addTo->GetItemId(), VSHPROPID_Expandable);
	}

	return S_OK;
}

HRESULT EnsureDirectoryExists (const wchar_t* path)
{
	if (PathFileExists(path))
	{
		DWORD attrs = GetFileAttributes(path); RETURN_LAST_ERROR_IF_EXPECTED(attrs == INVALID_FILE_ATTRIBUTES);
		if ((attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
			return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
	}
	else
	{
		int ires = SHCreateDirectoryEx(nullptr, path, nullptr); RETURN_HR_IF_EXPECTED(HRESULT_FROM_WIN32(ires), ires);
	}

	return S_OK;
}

// Returns S_OK if found, result in insertBefore.
// Returns S_FALSE if not found, location to insert given by insertBefore and insertAfter.
HRESULT FindFolderNodeOrInsertLocation (IProjectNode* proj, IParentNode* parent, const wchar_t* folderName,
	IFolderNode** ppFound, com_ptr<IChildNode>& insertBefore, com_ptr<IChildNode>& insertAfter)
{
	insertAfter = nullptr;
	insertBefore = parent->FirstChild();
	while (insertBefore)
	{
		auto insertBeforeAsFolder = wil::try_com_query_nothrow<IFolderNode>(insertBefore);
		if (!insertBeforeAsFolder)
			break; // No more folders, only files from now on

		wil::unique_variant name;
		auto hr = insertBeforeAsFolder->GetProperty(proj, VSHPROPID_SaveName, &name); RETURN_IF_FAILED(hr);
		RETURN_HR_IF(E_UNEXPECTED, V_VT(&name) != VT_BSTR);

		if (!_wcsicmp(folderName, V_BSTR(&name)))
		{
			if (ppFound)
				*ppFound = insertBeforeAsFolder.detach();
			return S_OK;
		}

		int cmpRes = wcscmp(folderName, V_BSTR(&name));
		if (cmpRes < 0)
			break;

		insertAfter = insertBefore;
		insertBefore = insertBefore->Next();
	}

	// We found the place to insert the new folder. Let's make sure there aren't any other nodes with the same name.
	if (insertBefore)
	{
		for (auto c = insertBefore->Next(); c != nullptr; c = c->Next())
		{
			wil::unique_variant name;
			if (SUCCEEDED(c->GetProperty(proj, VSHPROPID_SaveName, &name)) && !_wcsicmp(name.bstrVal, folderName))
				return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
		}
	}

	return S_FALSE;
}

HRESULT InsertFolderNode (IProjectNode* proj, IParentNode* parent, IChildNode* insertBefore, IChildNode* insertAfter, IFolderNode* newFolder)
{
	HRESULT hr;
	newFolder->SetNext(insertBefore);
	if (!insertAfter)
		parent->SetFirstChild(newFolder);
	else
		insertAfter->SetNext(newFolder);
	hr = SetItemIdsTree (proj, newFolder, insertAfter, parent); RETURN_IF_FAILED(hr);
	proj->NotifyPropertyChangedHierNode (parent->GetItemId(), VSHPROPID_Expandable);
	return S_OK;
}

// Enum depth-first (just because it's simpler) post-order mode (so that children clear their ItemId before parent).
static HRESULT ClearItemIdsTree (IProjectNode* root, IChildNode* child)
{
	stdext::inplace_function<HRESULT(IChildNode*)> enumNodeAndChildren;

	enumNodeAndChildren = [root, &enumNodeAndChildren](IChildNode* node) -> HRESULT
		{
			HRESULT hr;
			
			VSITEMID oldItemID = node->GetItemId();
			root->NotifyNodeRemovingFromHier(node);

			com_ptr<IParentNode> nodeAsParent;
			if (SUCCEEDED(node->QueryInterface(&nodeAsParent)))
			{
				for (auto c = nodeAsParent->FirstChild(); c; c = c->Next())
				{
					hr = enumNodeAndChildren(c); RETURN_IF_FAILED(hr);
				}
			}

			hr = node->ClearItemId(); RETURN_IF_FAILED(hr);

			root->NotifyNodeRemovedFromHier(node, oldItemID);

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

	auto keepAlive = com_ptr(node);

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

HRESULT GetItems (IParentNode* parent, HRESULT(*filter)(IChildNode*), SAFEARRAY** itemsOut)
{
	HRESULT hr;
	*itemsOut = nullptr;

	vector_nothrow<com_ptr<IDispatch>> nodes;

	for (auto c = parent->FirstChild(); c; c = c->Next())
	{
		if (filter)
		{
			hr = filter(c); RETURN_IF_FAILED(hr);
			if (hr == S_FALSE)
				continue;
		}

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

	// To keep things simple and the loading code fast, this function expects to put items either
	// directly in the project (in which case it calls SetItemIdsTree), or to some other parent item
	// that's _not_yet_ added to a hierarchy.
	auto proj = wil::try_com_query_nothrow<IProjectNode>(parent);
	RETURN_HR_IF(E_UNEXPECTED, !proj && parent->GetItemId() != VSITEMID_NIL);

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

			if (proj)
			{
				hr = SetItemIdsTree (proj, node, insertAfter, parent); RETURN_IF_FAILED(hr);
			}
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	// This is meant to be called only from LoadXml, no need to set dirty flag or send notifications.

	return S_OK;
}

static bool IsMacroName (const wchar_t* from, const wchar_t* to)
{
	while(from != to)
	{
		if (!iswupper(*from) && !iswdigit(*from) && (*from != L'_'))
			return false;
		from++;
	}

	return true;
}

static HRESULT ResolveTemplateFileMacro (const wchar_t* macroFrom, const wchar_t* macroTo, IProjectConfig* config, wil::unique_process_heap_string& valueOut)
{
	HRESULT hr;

	std::wstring_view macro = { macroFrom, macroTo };

	if (macro == L"BASE_ADDR")
	{
		unsigned long baseAddress;
		hr = config->AsmProps()->get_BaseAddress(&baseAddress); RETURN_IF_FAILED(hr);
		hr = wil::str_printf_nothrow(valueOut, L"%u", baseAddress); RETURN_IF_FAILED(hr);
		return S_OK;
	}
	else if (macro == L"DEVICE")
	{
		valueOut = wil::make_process_heap_string_nothrow(L"ZXSPECTRUM48"); RETURN_IF_NULL_ALLOC(valueOut);
		return S_OK;
	}
	else if (macro == L"ENTRY_POINT_ADDR")
	{
		wil::unique_bstr addr;
		hr = config->AsmProps()->get_EntryPointAddress(&addr); RETURN_IF_FAILED(hr);
		valueOut = wil::make_process_heap_string_nothrow(addr.get()); RETURN_IF_NULL_ALLOC(valueOut);
		return S_OK;
	}
	else if (macro == L"SAVESNA")
	{
		OutputFileType ft;
		hr = config->GeneralProps()->get_OutputFileType(&ft); RETURN_IF_FAILED(hr);
		if (ft == OutputFileType::Binary)
		{
			static const wchar_t str[] = L"; (Binary selected. Filename is passed as command-line parameter)";
			valueOut = wil::make_process_heap_string_nothrow(str, _countof(str) - 1); RETURN_IF_NULL_ALLOC(valueOut);
			return S_OK;
		}
		else if (ft == OutputFileType::Sna)
		{
			wil::unique_bstr outputName;
			hr = config->GeneralProps()->get_OutputName(&outputName); RETURN_IF_FAILED(hr);
			wil::unique_process_heap_string on;
			hr = ResolveMacros(outputName.get(), config, on); RETURN_IF_FAILED(hr);
			hr = wil::str_printf_nothrow (valueOut, L"SAVESNA %s%s", on.get(), GetOutputExtensionFromOutputType(OutputFileType::Sna)); RETURN_IF_FAILED(hr);
			return S_OK;
		}
		else
			RETURN_HR(E_NOTIMPL);
	}

	RETURN_HR(E_NOTIMPL);
}

HRESULT CreateFileFromTemplate (LPCWSTR fromPath, LPCWSTR toPath, IProjectConfig* config)
{
	HRESULT hr;

	// Read template UTF8 to memory.
	wil::unique_hfile fromFile (CreateFile (fromPath, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr)); RETURN_LAST_ERROR_IF(!fromFile.is_valid());
	DWORD fileSize = GetFileSize (fromFile.get(), nullptr); RETURN_LAST_ERROR_IF(fileSize == INVALID_FILE_SIZE);
	auto fromBufferUtf8 = wil::make_hlocal_ansistring_nothrow(nullptr, fileSize); RETURN_IF_NULL_ALLOC(fromBufferUtf8);
	DWORD bytesRead;
	BOOL bres = ReadFile (fromFile.get(), fromBufferUtf8.get(), fileSize, &bytesRead, nullptr); RETURN_IF_WIN32_BOOL_FALSE(bres);
	fromBufferUtf8.get()[fileSize] = 0;
	fromFile.reset();

	// Convert template UTF8 to UTF16
	wil::unique_process_heap_string fromBuffer;
	hr = UTF8ToUTF16(fromBufferUtf8, fromBuffer); RETURN_IF_FAILED(hr);

	// Generate UTF16 content.
	com_ptr<IStream> toStreamUtf16;
	hr = CreateStreamOnHGlobal (nullptr, TRUE, &toStreamUtf16); RETURN_IF_FAILED(hr);
	for (const wchar_t* p = fromBuffer.get(); *p; )
	{
		const wchar_t* to;
		if (p[0] == L'%' && (to = wcschr(p + 1, L'%')) && IsMacroName(p + 1, to))
		{
			wil::unique_process_heap_string value;
			hr = ResolveTemplateFileMacro(p + 1, to, config, value); RETURN_IF_FAILED(hr);
			hr = toStreamUtf16->Write (value.get(), sizeof(wchar_t) * wcslen(value.get()), nullptr); RETURN_IF_FAILED(hr);
			p = to + 1;
		}
		else
		{
			hr = toStreamUtf16->Write (p, sizeof(wchar_t), nullptr); RETURN_IF_FAILED(hr);
			p++;
		}
	}

	// Convert UTF16 content to UTF8 content;
	STATSTG stat;
	hr = toStreamUtf16->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
	HGLOBAL hg;
	hr = GetHGlobalFromStream(toStreamUtf16, &hg); RETURN_IF_FAILED(hr);
	auto buffer = wil::unique_hglobal_locked(hg); RETURN_LAST_ERROR_IF(!buffer);
	wil::unique_hlocal_ansistring content;
	size_t contentLen;
	hr = UTF16ToUTF8(static_cast<wchar_t*>(buffer.get()), content, &contentLen); RETURN_IF_FAILED(hr);

	// Write UTF8 content to disk.
	com_ptr<IStream> toStream;
	hr = SHCreateStreamOnFileEx (toPath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &toStream); RETURN_IF_FAILED(hr);
	hr = toStream->Write(content.get(), contentLen, nullptr); RETURN_IF_FAILED(hr);
	
	return S_OK;
}

IFileNode* FindChildFileByName (IProjectNode* proj, IParentNode* parent, const wchar_t* fileName)
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
		if (SUCCEEDED(fn->GetProperty(proj, VSHPROPID_SaveName, &n)) && V_VT(&n) == VT_BSTR && !_wcsicmp(fileName, V_BSTR(&n)))
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

// returns S_OK or S_FALSE
HRESULT ParseNumber (LPCWSTR str, DWORD* value)
{
	size_t len = wcslen(str);
	if (!len)
		return S_FALSE;
	if (!isdigit(str[0]))
		return S_FALSE;

	uint32_t val;
	wchar_t* endPtr;
	if ((str[1] & 0xDF) == 'X')
	{
		// C-style hex
		val = wcstoul(&str[2], &endPtr, 16);
		if (endPtr != &str[len])
			return S_FALSE;
	}
	else if ((str[len - 1] & 0xDF) == 'H')
	{
		// ASM-style hex
		val = wcstoul(str, &endPtr, 16);
		if (endPtr != &str[len - 1])
			return S_FALSE;
	}
	else
	{
		// Decimal (we don't support octal)
		val = wcstoul(str, &endPtr, 10);
		if (endPtr != &str[len])
			return S_FALSE;
	}

	*value = val;
	return S_OK;
}

const wchar_t* GetOutputExtensionFromOutputType (OutputFileType type)
{
	if (type == OutputFileType::Binary)
		return L".bin";
	if (type == OutputFileType::Sna)
		return L".sna";
	WI_ASSERT(false); return L"";
}

static HRESULT ResolveMacro (IProjectNode* project, IProjectConfig* config, const wchar_t* macroFrom, const wchar_t* macroTo, wil::unique_process_heap_string& valueOut, const wchar_t*& macroUniqueNameOut)
{
	HRESULT hr;

	std::wstring_view macro = { macroFrom, macroTo };
	if (macro == MacroOutputName)
	{
		wil::unique_bstr outputName;
		hr = config->GeneralProps()->get_OutputName(&outputName); RETURN_IF_FAILED(hr);
		valueOut = wil::make_process_heap_string_nothrow(outputName.get()); RETURN_IF_NULL_ALLOC(valueOut);
		macroUniqueNameOut = MacroOutputName;
		return S_OK;
	}
	else if (macro == MacroProjectName)
	{
		wil::unique_variant name;
		hr = project->AsHierarchy()->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &name); RETURN_IF_FAILED(hr); RETURN_HR_IF(E_FAIL, name.vt != VT_BSTR);
		valueOut = wil::make_process_heap_string_nothrow(name.bstrVal); RETURN_IF_NULL_ALLOC(valueOut);
		macroUniqueNameOut = MacroProjectName;
		return S_OK;
	}
	else if (macro == MacroProjectDir)
	{
		wil::unique_variant projectDir;
		hr = project->AsHierarchy()->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, &projectDir); RETURN_IF_FAILED(hr);
		valueOut = wil::make_process_heap_string_nothrow(projectDir.bstrVal); RETURN_IF_NULL_ALLOC(valueOut);
		macroUniqueNameOut = MacroProjectDir;
		return S_OK;
	}
	else if (macro == MacroConfigName)
	{
		wil::unique_bstr name;
		hr = config->AsProjectConfigProperties()->get_ConfigName(&name); RETURN_IF_FAILED(hr);
		valueOut = wil::make_process_heap_string_nothrow(name.get()); RETURN_IF_NULL_ALLOC(valueOut);
		macroUniqueNameOut = MacroConfigName;
		return S_OK;
	}
	else if (macro == MacroOutputDir)
	{
		wil::unique_bstr outputDir;
		hr = config->GeneralProps()->get_OutputDirectory(&outputDir); RETURN_IF_FAILED(hr);
		valueOut = wil::make_process_heap_string_nothrow(outputDir.get()); RETURN_IF_NULL_ALLOC(valueOut);
		macroUniqueNameOut = MacroOutputDir;
		return S_OK;
	}
	else if (macro == MacroOutputFilename)
	{
		wil::unique_bstr fn;
		hr = config->GeneralProps()->get_OutputFilename(&fn); RETURN_IF_FAILED(hr);
		valueOut = wil::make_process_heap_string_nothrow(fn.get()); RETURN_IF_NULL_ALLOC(valueOut);
		macroUniqueNameOut = MacroOutputFilename;
		return S_OK;
	}

	RETURN_HR(E_NOTIMPL);
}

static HRESULT ResolveMacros (IProjectNode* project, IProjectConfig* config, vector_nothrow<const wchar_t*>& resolving, const wchar_t* pszIn, IStream* pOut)
{
	HRESULT hr;

	for (const wchar_t* p = pszIn; *p; )
	{
		const wchar_t* to;
		if (p[0] == L'%' && (to = wcschr(p + 1, L'%')) && IsMacroName(p + 1, to))
		{
			wil::unique_process_heap_string value;
			const wchar_t* macroUniqueName;
			hr = ResolveMacro (project, config, p + 1, to, value, macroUniqueName); RETURN_IF_FAILED(hr);
			auto it = resolving.find(macroUniqueName);
			if (it == resolving.end())
			{
				// Only resolve if no circular dependency
				bool pushed = resolving.try_push_back(macroUniqueName); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
				auto remove = wil::scope_exit([&resolving]{ resolving.remove_back(); });
				hr = ResolveMacros (project, config, resolving, value.get(), pOut); RETURN_IF_FAILED(hr);
			}
			else
			{
				// Circular dependency. We leave the macro as it. Maybe later we'll tell the user about it somehow.
				hr = pOut->Write (p, 2 * (to + 1 - p), nullptr); RETURN_IF_FAILED(hr);
			}

			p = to + 1;
		}
		else
		{
			hr = pOut->Write(p, 2, nullptr); RETURN_IF_FAILED(hr);
			p++;
		}
	}

	return S_OK;
}

HRESULT ResolveMacros (const wchar_t* pszIn, IProjectConfig* config, wil::unique_process_heap_string& out)
{
	HRESULT hr;
	com_ptr<IProjectNode> project;
	hr = config->GetSite(IID_PPV_ARGS(&project)); RETURN_IF_FAILED(hr);
	IStream* sraw = SHCreateMemStream(nullptr, 0); RETURN_IF_NULL_ALLOC(sraw);
	com_ptr<IStream> s;
	s.attach(sraw);
	vector_nothrow<const wchar_t*> resolving;
	hr = ResolveMacros (project, config, resolving, pszIn, sraw); RETURN_IF_FAILED_EXPECTED(hr);
	STATSTG stat;
	hr = sraw->Stat(&stat, STATFLAG_NONAME); RETURN_IF_FAILED(hr);
	out = wil::make_process_heap_string_nothrow(nullptr, stat.cbSize.LowPart); RETURN_IF_NULL_ALLOC(out);
	hr = sraw->Seek({ 0 }, SEEK_SET, nullptr); RETURN_IF_FAILED(hr);
	hr = sraw->Read(out.get(), stat.cbSize.LowPart, nullptr); RETURN_IF_FAILED(hr);
	out.get()[stat.cbSize.LowPart] = 0;
	return S_OK;
}

HRESULT IsDescendantOf (IParentNode* possibleAncestor, IChildNode* node)
{
	com_ptr<IParentNode> parent;
	HRESULT hr = node->GetParent(&parent); RETURN_IF_FAILED(hr);
	while (true)
	{
		if (parent == possibleAncestor)
			return S_OK;
		if (parent->GetItemId() == VSITEMID_ROOT)
			return S_FALSE;
		com_ptr<IChildNode> parentAsChild;
		hr = parent->QueryInterface(IID_PPV_ARGS(&parentAsChild)); RETURN_IF_FAILED(hr);
		hr = parentAsChild->GetParent(&parent); RETURN_IF_FAILED(hr);
	}
}

HRESULT EnsureDirHasBackslash (LPCOLESTR pszLocation, wil::unique_process_heap_string& dir)
{
	size_t len = wcslen(pszLocation);
	if (pszLocation[len - 1] == '\\')
	{
		dir = wil::make_process_heap_string_nothrow(pszLocation, len); RETURN_IF_NULL_ALLOC(dir);
	}
	else if (pszLocation[len - 1] == '/')
	{
		dir = wil::make_process_heap_string_nothrow(pszLocation, len); RETURN_IF_NULL_ALLOC(dir);
		dir.get()[len - 1] = '\\';
	}
	else
	{
		dir = wil::make_process_heap_string_nothrow(pszLocation, len + 1); RETURN_IF_NULL_ALLOC(dir);
		dir.get()[len] = '\\';
		dir.get()[len + 1] = 0;
	}

	return S_OK;
}

HRESULT NotifyPropertyChanging (ConnectionPointImpl<IPropertyChangeSink>* cp, IDispatch* pDisp, DISPID dispid)
{
	return cp->Notify([pDisp,dispid](IPropertyChangeSink* sink)
		{
			auto hr = sink->OnPropertyChanging(pDisp, dispid, { }); RETURN_HR(hr);
		});
}

HRESULT NotifyPropertyChanged (ConnectionPointImpl<IPropertyChangeSink>* cp, IDispatch* pDisp, DISPID dispid)
{
	return cp->Notify([pDisp, dispid](IPropertyChangeSink* sink)
		{
			auto hr = sink->OnPropertyChanged(pDisp, dispid, { }); RETURN_HR(hr);
		});
}

HRESULT NotifyPropertyChanging (ConnectionPointImpl<IPropertyChangeSink>* cp, IDispatch* pDisp, std::initializer_list<DISPID> dispids)
{
	return cp->Notify([pDisp,&dispids](IPropertyChangeSink* sink)
		{
			for (auto it = std::begin(dispids); it != std::end(dispids); it++)
			{
				auto hr = sink->OnPropertyChanging(pDisp, *it, { }); RETURN_IF_FAILED(hr);
			}
			return S_OK;
		});
}

HRESULT NotifyPropertyChanged (ConnectionPointImpl<IPropertyChangeSink>* cp, IDispatch* pDisp, std::initializer_list<DISPID> dispids)
{
	return cp->Notify([pDisp,&dispids](IPropertyChangeSink* sink)
		{
			for (auto it = std::rbegin(dispids); it != std::rend(dispids); it++)
			{
				auto hr = sink->OnPropertyChanged(pDisp, *it, { }); RETURN_IF_FAILED(hr);
			}
			return S_OK;
		});
}

HRESULT NotifyPropertyChanged (ConnectionPointImpl<IPropertyNotifySink>* cp, DISPID dispid)
{
	return cp->Notify([dispid](IPropertyNotifySink* sink)
		{ 
			auto hr = sink->OnChanged(dispid); RETURN_HR(hr);
		});
}

HRESULT NotifyPropertyChanged (ConnectionPointImpl<IPropertyNotifySink>* cp, std::initializer_list<DISPID> dispids)
{
	return cp->Notify([&dispids](IPropertyNotifySink* sink)
		{
			for (auto dispid : dispids)
			{
				auto hr = sink->OnChanged(dispid); RETURN_IF_FAILED(hr);
			}
			return S_OK;
		});
}
