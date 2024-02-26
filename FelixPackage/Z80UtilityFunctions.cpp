
#include "pch.h"
#include "FelixPackage.h"

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
	int ires = MultiByteToWideChar(CP_UTF8, 0, sl_name_from, (int)(sl_name_to - sl_name_from), nullptr, 0); RETURN_LAST_ERROR_IF(!ires);
	auto wname = wil::make_hlocal_string_nothrow(nullptr, ires); RETURN_IF_NULL_ALLOC(wname);
	int ires1 = MultiByteToWideChar (CP_UTF8, 0, sl_name_from, (int)(sl_name_to - sl_name_from), wname.get(), ires + 1); RETURN_LAST_ERROR_IF(!ires); RETURN_HR_IF(E_FAIL, ires1 != ires);
	*to = SysAllocStringLen(wname.get(), ires); RETURN_IF_NULL_ALLOC(*to);
	return S_OK;
}

