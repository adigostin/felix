
#include "pch.h"
#include "FelixPackage.h"
#include "guids.h"
#include "shared/com.h"
#include "../FelixPackageUi/resource.h"
#include <optional>

struct S
{
	PCWSTR valueBefore;
	wil::unique_bstr valueAfter;
};

static std::optional<LRESULT> CommandLineDialogProc (S* params, HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (uMsg == WM_INITDIALOG)
	{
		#pragma region Center in parent
		RECT rcOwner;
		GetWindowRect (GetParent(hwnd), &rcOwner);
		RECT rcDlg;
		GetWindowRect (hwnd, &rcDlg);
		RECT rc = rcOwner;
		// Offset the owner and dialog box rectangles so that right and bottom  values represent the width
		// and height, and then offset the owner again to discard space taken up by the dialog box.
		OffsetRect(&rcDlg, -rcDlg.left, -rcDlg.top);
		OffsetRect(&rc, -rc.left, -rc.top);
		OffsetRect(&rc, -rcDlg.right, -rcDlg.bottom);
		// The new position is the sum of half the remaining space and the owner's original position.
		SetWindowPos (hwnd, HWND_TOP, rcOwner.left + (rc.right / 2), rcOwner.top + (rc.bottom / 2), rcDlg.right - rcDlg.left, rcDlg.bottom - rcDlg.top, 0);
		#pragma endregion
		if (params->valueBefore)
			SetWindowText(GetDlgItem(hwnd, IDC_EDIT_COMMANDS), params->valueBefore);
		return TRUE;
	}

	if (uMsg == WM_COMMAND)
	{
		if (wParam == IDOK)
		{
			HWND edit = GetDlgItem(hwnd, IDC_EDIT_COMMANDS);
			int len = GetWindowTextLength(edit);
			if (len == 0)
				return EndDialog(hwnd, HRESULT_FROM_WIN32(GetLastError())), 0;

			auto valueAfter = wil::make_hlocal_string_nothrow(nullptr, len);
			if (!valueAfter)
				return EndDialog(hwnd, E_OUTOFMEMORY), 0;
			GetWindowText(edit, valueAfter.get(), len + 1);

			params->valueAfter.reset(SysAllocStringLen(valueAfter.get(), (UINT)len));
			if (!params->valueAfter)
				return EndDialog(hwnd, E_OUTOFMEMORY), 0;

			return EndDialog(hwnd, S_OK), 0;
		}

		if (wParam == IDCANCEL)
			return EndDialog(hwnd, HRESULT_FROM_WIN32(ERROR_CANCELLED)), 0;

		return std::nullopt;
	}

	return std::nullopt;
}

static INT_PTR CALLBACK CommandLineDialogProcStatic (HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (uMsg == WM_INITDIALOG)
		SetWindowLongPtr (hwnd, DWLP_USER, lParam);

	S* params = reinterpret_cast<S*>(GetWindowLongPtr(hwnd, DWLP_USER));
	if (!params)
		return FALSE; // It's one of those messages sent before WM_INITDIALOG.

	auto res = CommandLineDialogProc (params, hwnd, uMsg, wParam, lParam);
	if (res)
	{
		SetWindowLongPtr (hwnd, DWLP_MSGRESULT, res.value());
		return TRUE;
	}

	return FALSE;
}

HRESULT ShowCommandLinePropertyBuilder (HWND hwndParent, BSTR valueBefore, BSTR* valueAfter)
{
	com_ptr<IVsShell> shell;
	auto hr = serviceProvider->QueryService(SID_SVsShell, &shell);
	HINSTANCE uiLibrary;
	hr = shell->LoadUILibrary(CLSID_FelixPackage, 0, (DWORD_PTR*)&uiLibrary); RETURN_IF_FAILED(hr);
	auto params = wistd::unique_ptr<S>(new (std::nothrow) S()); RETURN_IF_NULL_ALLOC(params);
	params->valueBefore = valueBefore;
	INT_PTR dbres = DialogBoxParamW(uiLibrary, MAKEINTRESOURCE(IDD_DIALOG_COMMAND_LINE_EDITOR), hwndParent,
		CommandLineDialogProcStatic, reinterpret_cast<LPARAM>(params.get()));
	RETURN_LAST_ERROR_IF(dbres == -1);
	hr = (HRESULT)dbres;
	if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
		return hr;
	RETURN_IF_FAILED(hr);
	*valueAfter = params->valueAfter.release();
	return S_OK;
}
