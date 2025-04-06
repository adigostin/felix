
#include "pch.h"
#include "FelixPackage.h"
#include "shared/com.h"
#include "shared/inplace_function.h"

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

HRESULT MakeBstrFromStreamOnHGlobal (IStream* stream, BSTR* pBstr)
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

HRESULT MakeSjasmCommandLine (IVsHierarchy* hier, IProjectConfig* config, IProjectConfigAssemblerProperties* asmProps, BSTR* ppCmdLine)
{
	HRESULT hr;

	wil::unique_variant projectDir;
	hr = hier->GetProperty(VSITEMID_ROOT, VSHPROPID_ProjectDir, projectDir.addressof()); RETURN_IF_FAILED(hr);
	RETURN_HR_IF(E_FAIL, projectDir.vt != VT_BSTR);

	wil::unique_bstr output_dir;
	hr = config->GetOutputDirectory(&output_dir); RETURN_IF_FAILED(hr);

	vector_nothrow<com_ptr<IProjectFileProperties>> asmFiles;

	stdext::inplace_function<HRESULT(IProjectItemParent*)> enumDescendants;

	enumDescendants = [&enumDescendants, &asmFiles](IProjectItemParent* parent) -> HRESULT
		{
			for (auto c = parent->FirstChild(); c; c = c->Next())
			{
				if (auto file = wil::try_com_query_nothrow<IProjectFileProperties>(c))
				{
					BuildToolKind tool;
					auto hr = file->get_BuildTool(&tool); RETURN_IF_FAILED(hr);
					if (tool == BuildToolKind::Assembler)
					{
						bool pushed = asmFiles.try_push_back(std::move(file)); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
					}
				}
				else if (auto cAsParent = wil::try_com_query_nothrow<IProjectItemParent>(c))
				{
					auto hr = enumDescendants(cAsParent); RETURN_IF_FAILED(hr);
				}
			}

			return S_OK;
		};
	enumDescendants(wil::try_com_query_nothrow<IProjectItemParent>(hier));

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
			wil::unique_hlocal_string outputFilePath;
			auto hr = PathAllocCombine (output_dir, output_filename, PathFlags, &outputFilePath); RETURN_IF_FAILED(hr);
			auto outputFilePathRelativeUgly = wil::make_hlocal_string_nothrow(nullptr, MAX_PATH); RETURN_IF_NULL_ALLOC(outputFilePathRelativeUgly);
			BOOL bRes = PathRelativePathToW (outputFilePathRelativeUgly.get(), project_dir, FILE_ATTRIBUTE_DIRECTORY, outputFilePath.get(), 0); RETURN_HR_IF(CS_E_INVALID_PATH, !bRes);
			size_t len = wcslen(outputFilePathRelativeUgly.get());
			auto outputFilePathRelative = wil::make_hlocal_string_nothrow(nullptr, len); RETURN_IF_NULL_ALLOC(outputFilePathRelative);
			hr = PathCchCanonicalizeEx (outputFilePathRelative.get(), len + 1, outputFilePathRelativeUgly.get(), PathFlags); RETURN_IF_FAILED(hr);
			hr = Write(cmdLine, paramName); RETURN_IF_FAILED(hr);
			hr = Write(cmdLine, outputFilePathRelative.get()); RETURN_IF_FAILED(hr);
			return S_OK;
		};

	// --raw=...
	wil::unique_bstr output_filename;
	hr = asmProps->GetOutputFileName(&output_filename); RETURN_IF_FAILED(hr);
	hr = addOutputPathParam (L" --raw=", output_filename.get()); RETURN_IF_FAILED(hr);

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
	for (uint32_t i = 0; i < asmFiles.size(); i++)
	{
		IProjectFileProperties* file = asmFiles[i];
		wil::unique_bstr fileRelativePath;
		hr = file->get_Path(&fileRelativePath);
		if (SUCCEEDED(hr))
		{
			hr = Write(cmdLine, L" "); RETURN_IF_FAILED(hr);
			hr = Write(cmdLine, fileRelativePath.get()); RETURN_IF_FAILED(hr);
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

	wil::unique_cotaskmem_string fullPathName;
	DWORD unused;
	hr = pff->GetCurFile (&fullPathName, &unused); RETURN_IF_FAILED(hr);

	VSQueryEditResult fEditVerdict;
	com_ptr<IVsQueryEditQuerySave2> queryEdit;
	hr = serviceProvider->QueryService (SID_SVsQueryEditQuerySave, &queryEdit); RETURN_IF_FAILED(hr);
	hr = queryEdit->QueryEditFiles (QEF_DisallowInMemoryEdits, 1, fullPathName.addressof(), nullptr, nullptr, &fEditVerdict, nullptr); LOG_IF_FAILED(hr);
	if (FAILED(hr) || (fEditVerdict != QER_EditOK))
		return OLE_E_PROMPTSAVECANCELLED;
	
	return S_OK;
}
