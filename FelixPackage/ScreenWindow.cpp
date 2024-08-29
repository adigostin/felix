
#include "pch.h"
#include "shared/com.h"
#include "shared/OtherGuids.h"
#include "FelixPackage.h"
#include "../FelixPackageUi/CommandIds.h"
#include "guids.h"
#include <queue>

using unique_cotaskmem_bitmapinfo = wil::unique_any<BITMAPINFO*, decltype(&::CoTaskMemFree), ::CoTaskMemFree>;

class ScreenWindowImpl 
	: public IVsWindowPane
	, public IVsDpiAware
	, public IVsDebuggerEvents
	, public IVsWindowFrameNotify4
	, public IVsWindowFrameNotify3
	, public IOleCommandTarget
	, public IScreenCompleteEventHandler
	, public IVsBroadcastMessageEvents
{
	ULONG _refCount = 0;
	static const WNDCLASS wndClass;
	static ATOM wndClassAtom;
	HWND _hwnd = nullptr;
	VS_RGBA _windowColor;
	VSCOOKIE _broadcastCookie = VSCOOKIE_NIL;
	com_ptr<ISimulator> _simulator;
	wil::unique_event_nothrow _design_mode_event;
	com_ptr<IServiceProvider> _sp;
	DWORD _debugger_events_cookie = 0;
	bool _advisingScreenCompleteEvents = false;
	struct Rational { LONG numerator; LONG denominator; };
	Rational _zoom;
	wil::unique_hfont _debugFont;
	wil::unique_hfont _beamFont;
	wil::unique_hgdiobj _beamPen;
	LONG _debugFontLineHeight;
	bool _debugFlagFpsAndDuration = false;

	struct render_perf_info
	{
		LARGE_INTEGER start_time;
		float duration;
	};

	LARGE_INTEGER _performance_counter_frequency;
	std::deque<render_perf_info> perf_info_queue;
	unique_cotaskmem_bitmapinfo _bitmap;
	POINT beamLocation;

public:
	HRESULT InitInstance()
	{
		QueryPerformanceFrequency(&_performance_counter_frequency);

		if (!wndClassAtom)
		{
			wndClassAtom = RegisterClass(&wndClass);
			RETURN_LAST_ERROR_IF(!wndClassAtom);
		}
		
		return S_OK;
	}

	#pragma region IUnknown
	virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

	virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }

	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
	{
		RETURN_HR_IF(E_POINTER, !ppvObject);
		*ppvObject = nullptr;

		if (   TryQI<IUnknown>(static_cast<IVsWindowPane*>(this), riid, ppvObject)
			|| TryQI<IVsWindowPane>(this, riid, ppvObject)
			|| TryQI<IVsDpiAware>(this, riid, ppvObject)
			|| TryQI<IVsDebuggerEvents>(this, riid, ppvObject)
			|| TryQI<IVsWindowFrameNotify4>(this, riid, ppvObject)
			|| TryQI<IVsWindowFrameNotify3>(this, riid, ppvObject)
			|| TryQI<IOleCommandTarget>(this, riid, ppvObject)
			|| TryQI<IScreenCompleteEventHandler>(this, riid, ppvObject)
			|| TryQI<IVsBroadcastMessageEvents>(this, riid, ppvObject)
		)
			return S_OK;

		#ifdef _DEBUG
		if (riid == IID_IManagedObject
			|| riid == IID_IProvideClassInfo
			|| riid == IID_IProvideClassInfo2
			|| riid == IID_IInspectable
			|| riid == IID_IMarshal
			|| riid == IID_INoMarshal
			|| riid == IID_IAgileObject
			|| riid == IID_IRpcOptions
			|| riid == IID_IReplaceablePane
			|| riid == IID_IVsWindowSearch
			|| riid == IID_IVsUIElementPane
			|| riid == IID_ISupportErrorInfo
			|| riid == IID_IVsWindowFrameNotify
			|| riid == IID_IVsStatusbarUser
			|| riid == IID_IVsWindowPaneCommitFilter
			|| riid == IID_IVsWindowPaneCommit
			|| riid == IID_IConvertible
			|| riid == IID_IDispatch
			|| riid == IID_IVsWindowView
			|| riid == IID_IVsUIHierarchyWindow
			|| riid == IID_IVsFindTarget
			|| riid == IID_IVsCodeWindow
			|| riid == IID_IServiceProvider
			|| riid == IID_IVsTextView
		)
			return E_NOINTERFACE;
		#endif

		return E_NOINTERFACE;
	}
	#pragma endregion

	static LRESULT CALLBACK WindowProcStatic (HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
	{
		//if (!assert_function_running)
		{
			if (auto w = reinterpret_cast<ScreenWindowImpl*>(GetWindowLongPtr(hwnd, GWLP_USERDATA)))
			{
				return w->WindowProc(hwnd, msg, wparam, lparam);
			}
		}

		return DefWindowProc (hwnd, msg, wparam, lparam);
	}

	LRESULT WindowProc (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
	{
		if (msg == WM_SIZE)
		{
			if (_bitmap)
				_zoom = GetZoom (&_bitmap.get()->bmiHeader, LOWORD(lParam), HIWORD(lParam));
			return 0;
		}

		if (msg == WM_ERASEBKGND)
			return ProcessWmErase((HDC)wParam);

		if (msg == WM_PAINT)
			return ProcessWmPaint();

		if (msg == WM_DESTROY)
		{
			WI_ASSERT_MSG(!GetWindowLongPtr(hWnd, GWLP_USERDATA), "This window is meant to be destroyed by releasing all references to it.");
			return 0;
		}

		if (msg == WM_KEYDOWN)
		{
			if (wParam == VK_F1)
			{
				_debugFlagFpsAndDuration = !_debugFlagFpsAndDuration;
				InvalidateRect(hWnd, NULL, FALSE);
				return 0;
			}
			
			UINT modifiers = (GetKeyState(VK_SHIFT) >> 15) ? MK_SHIFT : 0;
			_simulator->ProcessKeyDown((UINT)wParam, modifiers);
			return 0;
		}

		if (msg == WM_KEYUP)
		{
			UINT modifiers = (GetKeyState(VK_SHIFT) >> 15) ? MK_SHIFT : 0;
			_simulator->ProcessKeyUp((UINT)wParam, modifiers);
			return 0;
		}

		if ((msg >= WM_LBUTTONDOWN) && (msg <= WM_MBUTTONDBLCLK))
		{
			SetFocus(_hwnd);
			return 0;
		}

		return DefWindowProc (hWnd, msg, wParam, lParam);
	}
	
	static Rational GetZoom (const BITMAPINFOHEADER* bi, LONG clientWidth, LONG clientHeight)
	{
		if (clientWidth >= bi->biWidth && clientHeight >= bi->biHeight)
			// zoom >= 1
			return Rational{ .numerator = std::min(clientWidth / bi->biWidth, clientHeight / bi->biHeight), .denominator = 1 };

		if (clientWidth > 0 && clientHeight > 0)
			// zoom <= 1
			return Rational{ .numerator = 1, .denominator = std::max ((bi->biWidth + clientWidth - 1) / clientWidth, (bi->biHeight + clientHeight - 1) / clientHeight) };
		
		return { 0, 1 };
	}

	LRESULT ProcessWmErase (HDC hdc)
	{
		RECT clientRect;
		::GetClientRect(_hwnd, &clientRect);

		auto b = wil::unique_hbrush(CreateSolidBrush(_windowColor & 0xFFFFFF));

		if (!_bitmap)
			FillRect(hdc, &clientRect, b.get());
		else
		{
			LONG w = _bitmap.get()->bmiHeader.biWidth * _zoom.numerator / _zoom.denominator;
			LONG h = _bitmap.get()->bmiHeader.biHeight * _zoom.numerator / _zoom.denominator;
			LONG xDest = (clientRect.right - w) / 2;
			LONG yDest = (clientRect.bottom - h) / 2;
			RECT rc = { 0, 0, clientRect.right, yDest };
			FillRect (hdc, &rc, b.get());
			rc = { 0, yDest, xDest, yDest + h };
			FillRect (hdc, &rc, b.get());
			rc = { xDest + w, yDest, clientRect.right, yDest + h };
			FillRect (hdc, &rc, b.get());
			rc = { 0, yDest + h, clientRect.right, clientRect.bottom };
			FillRect (hdc, &rc, b.get());
		}

		return 1;
	}

	LRESULT ProcessWmPaint()
	{
		RECT clientRect;
		::GetClientRect(_hwnd, &clientRect);

		LARGE_INTEGER start_time;
		BOOL bRes = QueryPerformanceCounter(&start_time); WI_ASSERT(bRes);

		RECT frameDurationAndFpsRect;
		if (_debugFlagFpsAndDuration)
		{
			frameDurationAndFpsRect.left = clientRect.right - 100;
			frameDurationAndFpsRect.top = clientRect.bottom - 2 * _debugFontLineHeight - 2 * 5;
			frameDurationAndFpsRect.right = clientRect.right;
			frameDurationAndFpsRect.bottom = clientRect.bottom;
			InvalidateRect (_hwnd, &frameDurationAndFpsRect, FALSE);
		}

		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(_hwnd, &ps);
		
		if (auto bitmapInfo = _bitmap.get())
		{
			int w = bitmapInfo->bmiHeader.biWidth;
			int h = bitmapInfo->bmiHeader.biHeight;
			int wscaled = w * _zoom.numerator / _zoom.denominator;
			int hscaled = h * _zoom.numerator / _zoom.denominator;
			int xDest = (clientRect.right - wscaled) / 2;
			int yDest = (clientRect.bottom - hscaled) / 2;
			int ires = StretchDIBits (hdc, xDest, yDest, wscaled, hscaled,
				0, 0, w, h, bitmapInfo->bmiColors, bitmapInfo, DIB_RGB_COLORS, SRCCOPY);
		
			if (_simulator->Running_HR() == S_FALSE)
			{
				SetBkColor(ps.hdc, 0xFFFF00);
				auto oldFont = wil::SelectObject(ps.hdc, _beamFont.get());
				auto oldPen = wil::SelectObject(ps.hdc, _beamPen.get());

				// ----------------------------------------------------------------
				// horz line that shows the Y value

				wchar_t buffer[16];
				int cc = swprintf_s (buffer, L"Y=%d", beamLocation.y);
				RECT rc = { };
				DrawTextW (ps.hdc, buffer, cc, &rc, DT_CALCRECT);
				LONG margin = rc.bottom / 4;

				POINT from = { margin + rc.right + margin, yDest + beamLocation.y * _zoom.numerator / _zoom.denominator };
				POINT to   = { clientRect.right, from.y };
				::MoveToEx (ps.hdc, from.x, from.y, nullptr);
				::LineTo (ps.hdc, to.x, to.y);
				LONG textY = std::max (margin, from.y - rc.bottom / 2);
				textY = std::min (textY, clientRect.bottom - rc.bottom - margin);
				rc.top = textY;
				rc.bottom += textY;
				DrawTextW (ps.hdc, buffer, cc, &rc, 0);

				// ----------------------------------------------------------------
				// vert line that shows the X value

				cc = swprintf_s (buffer, L"X=%d", beamLocation.x);
				rc = { };
				DrawTextW (ps.hdc, buffer, cc, &rc, DT_CALCRECT);

				from = { xDest + beamLocation.x * _zoom.numerator / _zoom.denominator, margin + rc.bottom + margin };
				to   = { from.x, clientRect.bottom };
				::MoveToEx (ps.hdc, from.x, from.y, nullptr);
				::LineTo (ps.hdc, to.x, to.y);
				LONG textX = std::max (margin, from.x - rc.right / 2);
				textX = std::min (textX, clientRect.right - rc.right - margin);
				rc.left = textX;
				rc.right += textX;
				rc.top = margin;
				rc.bottom += margin;
				DrawTextW (ps.hdc, buffer, cc, &rc, 0);
			}
		}

		if (_debugFlagFpsAndDuration)
		{
			GdiFlush();
			#pragma region Calculate performance data.
			LARGE_INTEGER timeNow;
			bRes = QueryPerformanceCounter(&timeNow); WI_ASSERT(bRes);

			render_perf_info perfInfo;
			perfInfo.start_time = start_time;
			perfInfo.duration = (float)(timeNow.QuadPart - start_time.QuadPart) / (float)_performance_counter_frequency.QuadPart * 1000.0f;

			perf_info_queue.push_back(perfInfo);
			if (perf_info_queue.size() > 100)
				perf_info_queue.pop_front();
			#pragma endregion

			#pragma region Draw performance data.
			wchar_t ss[50];
			unsigned dur = (unsigned)round(average_render_duration() * 100);
			int sslen = swprintf_s(ss, L"%4u FPS\r\n%3u.%02u ms", (unsigned)round(fps()), dur / 100, dur % 100);
			SetBkColor(ps.hdc, 0xFFFF00);
			auto oldFont = SelectObject (ps.hdc, _debugFont.get());
			RECT ssrect;
			ssrect.left = frameDurationAndFpsRect.right - 4 - 100;// - tl.width();
			ssrect.top = frameDurationAndFpsRect.top + 5;
			ssrect.right = clientRect.right;
			ssrect.bottom = clientRect.bottom;
			DrawTextW (ps.hdc, ss, sslen, &ssrect, DT_RIGHT | DT_NOCLIP);
			SelectObject (ps.hdc, oldFont);
			ssrect = { 0, 0, 100, 100 };
			HBRUSH hb = CreateSolidBrush ((rand() % 256) << 16 | 0x8080);
			FillRect(ps.hdc, &ssrect, hb);
			DeleteObject(hb);
			#pragma endregion
		}

		EndPaint(_hwnd, &ps);

		return 0;
	}

	float fps() const
	{
		if (perf_info_queue.size() < 2)
			return 0;

		LARGE_INTEGER start_time = perf_info_queue.front().start_time;
		LARGE_INTEGER end_time = perf_info_queue.back().start_time;

		float seconds = (float)(end_time.QuadPart - start_time.QuadPart) / (float)_performance_counter_frequency.QuadPart;

		float fps = ((float)perf_info_queue.size() - 1) / seconds;

		return fps;
	}

	float average_render_duration() const
	{
		if (perf_info_queue.empty())
			return 0;

		float sum = 0;
		for (const auto& entry : perf_info_queue)
		{
			sum += entry.duration;
		}

		float avg = sum / perf_info_queue.size();

		return avg;
	}

	HRESULT InitVSColors()
	{
		wil::com_ptr_nothrow<IVsUIShell5> shell;
		auto hr = _sp->QueryService (SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);

		hr = shell->GetThemedColor(GUID_EnvironmentColorsCategory, L"Window", TCT_Background, &_windowColor); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	#pragma region IVsWindowPane
	virtual HRESULT STDMETHODCALLTYPE SetSite (IServiceProvider *pSP) override
	{
		_sp = pSP;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE CreatePaneWindow (HWND hwndParent, int x, int y, int cx, int cy, HWND *hwnd) override
	{
		HRESULT hr;

		// TODO: if this fails somewhere in the middle, destroy/release/unregister everything created up to that point.

		uint32_t dpi = GetDpiForWindow(hwndParent);

		LOGFONT lf = { };
		lf.lfHeight = 11 * (LONG)dpi / 72;
		lf.lfQuality = NONANTIALIASED_QUALITY;
		wcscpy_s(lf.lfFaceName, L"Consolas");
		_debugFont.reset(CreateFontIndirect(&lf)); WI_ASSERT(_debugFont);
		_beamFont.reset(CreateFontIndirect(&lf)); WI_ASSERT(_beamFont);
		_beamPen.reset(CreatePen(PS_DASH, 1, RGB(255, 0, 255))); WI_ASSERT(_beamPen);
		HDC tempDC = GetDC(hwndParent);
		auto oldFont = SelectObject(tempDC, _debugFont.get());
		SIZE textSize;
		BOOL bRes = GetTextExtentPoint32W (tempDC, L"0", 1, &textSize); WI_ASSERT(bRes);
		_debugFontLineHeight = textSize.cy;
		SelectObject (tempDC, oldFont);
		ReleaseDC(hwndParent, tempDC);

		_hwnd = CreateWindowExW (0, wndClass.lpszClassName, L"", WS_CHILD | WS_VISIBLE, x, y, cx, cy, hwndParent, 0, (HINSTANCE)&__ImageBase, this); RETURN_LAST_ERROR_IF(!_hwnd);
		SetWindowLongPtr (_hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
		auto destroyHwndIfFailed = wil::scope_exit([this]
			{ SetWindowLongPtr(_hwnd, GWLP_USERDATA, 0); DestroyWindow(_hwnd); _hwnd = nullptr; });

		RECT cr;
		GetClientRect(_hwnd, &cr);

		hr = _sp->QueryService (SID_Simulator, &_simulator); RETURN_IF_FAILED(hr);

		hr = _simulator->AdviseScreenComplete(this); RETURN_IF_FAILED(hr);
		_advisingScreenCompleteEvents = true;

		com_ptr<IVsDebugger> debugger;
		hr = _sp->QueryService(SID_SVsShellDebugger, &debugger); RETURN_IF_FAILED(hr);
		hr = debugger->AdviseDebuggerEvents(this, &_debugger_events_cookie); RETURN_IF_FAILED(hr);
		auto unadvise_if_failed = wil::scope_exit([debugger, this]
			{ debugger->UnadviseDebuggerEvents(_debugger_events_cookie); _debugger_events_cookie = 0; });
/*
		DBGMODE mode;
		hr = debugger->GetMode(&mode); RETURN_IF_FAILED(hr);
		if (mode == DBGMODE_Design)
		{
			hr = _simulator->Resume(false); RETURN_IF_FAILED(hr);
		}
*/
		wil::com_ptr_nothrow<IVsWindowFrame> windowFrame;
		hr = _sp->QueryService (SID_SVsWindowFrame, &windowFrame); RETURN_IF_FAILED(hr);
		wil::unique_variant v;
		hr = windowFrame->GetProperty (VSFPROPID_ToolbarHost, &v); RETURN_IF_FAILED(hr);
		if (v.vt != VT_UNKNOWN)
			RETURN_HR(E_INVALIDARG);
		wil::com_ptr_nothrow<IVsToolWindowToolbarHost> toolbarHost;
		hr = v.punkVal->QueryInterface(&toolbarHost); RETURN_IF_FAILED(hr);
		hr = toolbarHost->AddToolbar (VSTWT_TOP, &CLSID_FelixPackageCmdSet, TWToolbar); RETURN_IF_FAILED(hr);

		hr = InitVSColors(); RETURN_IF_FAILED(hr);
		wil::com_ptr_nothrow<IVsShell> shell;
		hr = _sp->QueryService(SID_SVsShell, &shell); RETURN_IF_FAILED(hr);
		hr = shell->AdviseBroadcastMessages (this, &_broadcastCookie); RETURN_IF_FAILED(hr);
		auto unadviseBroadcastIfFailed = wil::scope_exit([this, shell=shell.get()]
			{ shell->UnadviseBroadcastMessages (_broadcastCookie); _broadcastCookie = VSCOOKIE_NIL; });




		unadviseBroadcastIfFailed.release();
		unadvise_if_failed.release();
		destroyHwndIfFailed.release();
		*hwnd = _hwnd;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetDefaultSize (SIZE* psize) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE ClosePane() override
	{
		HRESULT hr;

		if (_broadcastCookie != VSCOOKIE_NIL)
		{
			wil::com_ptr_nothrow<IVsShell> shell;
			hr = _sp->QueryService(SID_SVsShell, &shell);
			WI_ASSERT(SUCCEEDED(hr));
			if (SUCCEEDED(hr))
			{
				hr = shell->UnadviseBroadcastMessages (_broadcastCookie);
				WI_ASSERT(SUCCEEDED(hr));
			}
			_broadcastCookie = VSCOOKIE_NIL;
		}

		com_ptr<IVsDebugger> debugger;
		hr = _sp->QueryService(SID_SVsShellDebugger, &debugger); LOG_IF_FAILED(hr);
		if (SUCCEEDED(hr))
		{
			hr = debugger->UnadviseDebuggerEvents(_debugger_events_cookie); LOG_IF_FAILED(hr);
			_debugger_events_cookie = 0;
		}

		if (_advisingScreenCompleteEvents)
		{
			_simulator->UnadviseScreenComplete(this);
			_advisingScreenCompleteEvents = false;
		}

		_simulator = nullptr;
		if (_hwnd)
		{
			WI_ASSERT(reinterpret_cast<ScreenWindowImpl*>(GetWindowLongPtr(_hwnd, GWLP_USERDATA)) == this);
			::SetWindowLongPtr (_hwnd, GWLP_USERDATA, 0);
			::DestroyWindow(_hwnd);
			_hwnd = nullptr;
		}

		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE LoadViewState (IStream* pstream) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE SaveViewState (IStream *pstream) override
	{
		RETURN_HR(E_NOTIMPL);
	}

	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator (LPMSG lpmsg) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion

	#pragma region IVsBroadcastMessageEvents
	virtual HRESULT STDMETHODCALLTYPE OnBroadcastMessage (UINT msg, WPARAM wParam, LPARAM lParam) override
	{
		switch (msg)
		{
			case WM_SYSCOLORCHANGE:
			case WM_PALETTECHANGED:
				InitVSColors();
				if (_hwnd)
					::InvalidateRect(_hwnd, nullptr, TRUE);
				break;
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IScreenCompleteEventHandler
	virtual HRESULT STDMETHODCALLTYPE OnScreenComplete (BITMAPINFO* bi, POINT beamLocation) override
	{
		if (!_bitmap)
		{
			RECT cr;
			::GetClientRect(_hwnd, &cr);
			_zoom = GetZoom (&bi->bmiHeader, cr.right, cr.bottom);
		}

		_bitmap = unique_cotaskmem_bitmapinfo(bi);
		this->beamLocation = beamLocation;
		BOOL erase = _simulator->Running_HR() == S_FALSE;
		InvalidateRect(_hwnd, 0, erase);
		return S_OK;
	}
	#pragma endregion

	BOOL AtlWaitWithMessageLoop(_In_ HANDLE hEvent)
	{
		DWORD dwRet;
		MSG msg;

		while(1)
		{
			dwRet = MsgWaitForMultipleObjectsEx(1, &hEvent, INFINITE, QS_ALLINPUT, MWMO_INPUTAVAILABLE);

			if (dwRet == WAIT_OBJECT_0)
				return TRUE;    // The event was signaled

			if (dwRet != WAIT_OBJECT_0 + 1)
				break;          // Something else happened

			// There is one or more window message available. Dispatch them
			while(PeekMessage(&msg,0,0,0,PM_NOREMOVE))
			{
				// check for unicode window so we call the appropriate functions
				BOOL bUnicode = ::IsWindowUnicode(msg.hwnd);
				BOOL bRet;

				if (bUnicode)
					bRet = ::GetMessageW(&msg, NULL, 0, 0);
				else
					bRet = ::GetMessageA(&msg, NULL, 0, 0);

				if (bRet > 0)
				{
					::TranslateMessage(&msg);

					if (bUnicode)
						::DispatchMessageW(&msg);
					else
						::DispatchMessageA(&msg);
				}

				if (WaitForSingleObject(hEvent, 0) == WAIT_OBJECT_0)
					return TRUE; // Event is now signaled.
			}
		}
		return FALSE;
	}

	HRESULT LoadFile (bool start_debugging)
	{
		HRESULT hr;

		com_ptr<IVsDebugger> debugger;
		hr = _sp->QueryService(SID_SVsShellDebugger, &debugger); RETURN_IF_FAILED(hr);

		com_ptr<IVsDebugger2> debugger2;
		hr = _sp->QueryService(SID_SVsShellDebugger, &debugger2); RETURN_IF_FAILED(hr);

		DBGMODE mode;
		hr = debugger->GetMode(&mode); RETURN_IF_FAILED(hr);
		if (mode != DBGMODE_Design)
		{
			hr = _design_mode_event.create(); RETURN_IF_FAILED(hr);
			auto destroy_event = wil::scope_exit([this] { _design_mode_event.reset(); });

			hr = debugger2->ConfirmStopDebugging(nullptr); RETURN_IF_FAILED(hr);
			if (hr == S_FALSE)
				return S_OK;

			// It's not a good thing to hijack the message loop, I know.
			// I'll look into this some other time.
			BOOL bres = AtlWaitWithMessageLoop(_design_mode_event.get()); WI_ASSERT(bres);
		}

		wil::unique_bstr initial_directory;
		com_ptr<IVsWritableSettingsStore> settings_store;
		{
			com_ptr<IVsSettingsManager> sm;
			hr = _sp->QueryService(IID_SVsSettingsManager, &sm); LOG_IF_FAILED(hr);
			if (SUCCEEDED(hr))
			{
				hr = sm->GetWritableSettingsStore (SettingsScope_UserSettings, &settings_store);
				if (SUCCEEDED(hr))
					settings_store->GetString (SettingsCollection, SettingLoadSavePath, &initial_directory); // no need to check for errors
			}
		}

		wchar_t filename[MAX_PATH];
		filename[0] = 0;
		com_ptr<IVsUIShell> shell;
		hr = _sp->QueryService (SID_SVsUIShell, &shell); RETURN_IF_FAILED(hr);
		HWND dialogOwner;
		hr = shell->GetDialogOwnerHwnd(&dialogOwner); RETURN_IF_FAILED(hr);
		VSOPENFILENAMEW of = { };
		of.lStructSize = (DWORD)sizeof(of);
		of.hwndOwner = dialogOwner;
		of.pwzDlgTitle = L"ABC";
		of.pwzFileName = filename;
		of.nMaxFileName = (DWORD)ARRAYSIZE(filename);
		of.pwzInitialDir = initial_directory.get();
		of.pwzFilter = L"Snapshot Files (*.sna)\0*.sna\0All Files (*.*)\0*.*\0";
		hr = shell->GetOpenFileNameViaDlg(&of);
		if (hr == OLE_E_PROMPTSAVECANCELLED)
			return S_OK;
		RETURN_IF_FAILED(hr);

		if (settings_store)
		{
			// No error checking, nothing we could do about it anyway.
			if (auto dir = wil::make_hlocal_string_nothrow(of.pwzFileName, of.nFileOffset))
				settings_store->SetString (SettingsCollection, SettingLoadSavePath, dir.get());
		}

		if (!start_debugging)
		{
			com_ptr<IStream> stream;
			hr = SHCreateStreamOnFileEx (filename, STGM_READ | STGM_SHARE_DENY_WRITE, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &stream); RETURN_IF_FAILED(hr);

			hr = _simulator->LoadSnapshot(stream.get());
			if (hr == SIM_E_SNAPSHOT_FILE_WRONG_SIZE)
			{
				LONG result;
				hr = shell->ShowMessageBox (0, CLSID_NULL, (LPOLESTR)L"Wrong file size",
					(LPOLESTR)L"The file has an unrecognized size.\r\n(Only ZX Spectrum 48K snapshot files are supported for now.)",
					nullptr, 0, OLEMSGBUTTON_OK, OLEMSGDEFBUTTON_FIRST, OLEMSGICON_WARNING, FALSE, &result);
				return S_OK;
			}
			RETURN_IF_FAILED(hr);

			if (_simulator->Running_HR() == S_FALSE)
			{
				hr = _simulator->Resume(false); RETURN_IF_FAILED(hr);
			}
			return S_OK;
		}

		VsDebugTargetInfo2 dti = { };
		dti.cbSize = sizeof(dti);
		dti.dlo = DLO_CreateProcess;
		dti.LaunchFlags = DBGLAUNCH_StopAtEntryPoint;
		dti.bstrExe = SysAllocString(of.pwzFileName); RETURN_IF_NULL_ALLOC(dti.bstrExe);
		dti.guidLaunchDebugEngine = Engine_Id;
		dti.guidPortSupplier = PortSupplier_Id;
		dti.bstrPortName = SysAllocString(SingleDebugPortName); RETURN_IF_NULL_ALLOC(dti.bstrPortName);
		dti.fSendToOutputWindow = TRUE;
		hr = debugger2->LaunchDebugTargets2 (1, &dti); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	#pragma region IOleCommandTarget
	virtual HRESULT STDMETHODCALLTYPE QueryStatus (const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT* pCmdText) override
	{
		if (*pguidCmdGroup == CLSID_FelixPackageCmdSet)
		{
			RETURN_HR_IF(E_NOTIMPL, cCmds != 1);
			if (prgCmds[0].cmdID == cmdidScreenWindowDebug)
			{
				prgCmds[0].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED
					| (_debugFlagFpsAndDuration ? OLECMDF_LATCHED : 0);
				return S_OK;
			}

			if (prgCmds[0].cmdID == cmdidShowCRTSnapshot)
			{
				BOOL show = _simulator->GetShowCRTSnapshot() == S_OK;
				prgCmds[0].cmdf = OLECMDF_SUPPORTED | OLECMDF_ENABLED
					| (show ? OLECMDF_LATCHED : 0);
				return S_OK;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		return OLECMDERR_E_UNKNOWNGROUP;
	}

	virtual HRESULT STDMETHODCALLTYPE Exec (const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvaIn, VARIANT* pvaOut) override
	{
		if (*pguidCmdGroup == CMDSETID_StandardCommandSet97)
		{
			// These are the cmdidXxxYyy constants from stdidcmd.h
			return OLECMDERR_E_NOTSUPPORTED;
		}

		if (*pguidCmdGroup == CLSID_FelixPackageCmdSet)
		{
			if (nCmdID == cmdidResetSimulator)
			{
				auto hr = _simulator->Reset(0); RETURN_IF_FAILED(hr);
				return hr;
			}

			if (nCmdID == cmdidOpenZ80File)
				return LoadFile(false);

			if (nCmdID == cmdidDebugZ80File)
				return LoadFile(true);

			if (nCmdID == cmdidScreenWindowDebug)
			{
				_debugFlagFpsAndDuration = !_debugFlagFpsAndDuration;
				InvalidateRect(_hwnd, nullptr, true);
				return S_OK;
			}

			if (nCmdID == cmdidShowCRTSnapshot)
			{
				BOOL show = _simulator->GetShowCRTSnapshot() == S_OK;
				auto hr = _simulator->SetShowCRTSnapshot(!show); RETURN_IF_FAILED(hr);
				return S_OK;
			}

			return OLECMDERR_E_NOTSUPPORTED;
		}

		return OLECMDERR_E_UNKNOWNGROUP;
	}
	#pragma endregion

	#pragma region IVsDpiAware
	virtual HRESULT STDMETHODCALLTYPE get_Mode(VSDPIMODE* dwMode) override
	{
		*dwMode = VSDM_PerMonitor;
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsDebuggerEvents
	virtual HRESULT STDMETHODCALLTYPE OnModeChange (DBGMODE dbgmodeNew) override
	{
		if (dbgmodeNew == DBGMODE_Design)
		{
			if (_design_mode_event)
				::SetEvent(_design_mode_event.get());
		}

		return S_OK;
	}
	#pragma endregion

	#pragma region IVsWindowFrameNotify3
	virtual HRESULT STDMETHODCALLTYPE OnShow (FRAMESHOW2 fShow) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnMove (int x, int y, int w, int h) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnSize (int x, int y, int w, int h) override
	{
		// The width and height here are those of the pane (our window plus the toolbar)
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE OnDockableChange (BOOL fDockable, int x, int y, int w, int h) override
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE OnClose (FRAMECLOSE *pgrfSaveOptions) override
	{
		return S_OK;
	}
	#pragma endregion

	#pragma region IVsWindowFrameNotify4
	virtual HRESULT STDMETHODCALLTYPE OnPropertyChanged(VSFPROPID propid) override
	{
		return E_NOTIMPL;
	}
	#pragma endregion
};

//static
const WNDCLASS ScreenWindowImpl::wndClass = {
	.style = CS_DBLCLKS | CS_HREDRAW | CS_VREDRAW,
	.lpfnWndProc = &ScreenWindowImpl::WindowProcStatic,
	.hInstance = (HINSTANCE)&__ImageBase,
	.hIcon = nullptr,
	.hCursor = ::LoadCursor (nullptr, IDC_ARROW),
	.hbrBackground = (HBRUSH)(1 + COLOR_WINDOW),
	.lpszMenuName = nullptr,
	.lpszClassName = L"ScreenWindowImpl-{24B42526-2970-4B3C-A753-2DABD22C4BB0}",
};

// static
ATOM ScreenWindowImpl::wndClassAtom = 0;

HRESULT SimulatorWindowPane_CreateInstance (IVsWindowPane** to)
{
	com_ptr<ScreenWindowImpl> sw = new (std::nothrow) ScreenWindowImpl(); RETURN_IF_NULL_ALLOC(sw);
	auto hr = sw->InitInstance(); RETURN_IF_FAILED(hr);
	*to = sw.detach();
	return S_OK;
}
