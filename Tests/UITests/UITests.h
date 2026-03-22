
#pragma once
#include "..\TestsCommon.h"
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
	wil::com_ptr_failfast<VxDTE::DTE2> GetDefaultVSInstance();
	wil::com_ptr_failfast<VxDTE::DTE2> LaunchVS();
	void CloseVS (VxDTE::DTE2* dte, bool hard = false);
}


