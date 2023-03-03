#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "wininet")

#include <windows.h>
#include <atlbase.h>
#include <atlhost.h>
#include <string>
#include <winternl.h>
#include <wininet.h>
#include <iostream>
#include "json11.hpp"
#include "resource.h"

#define DEFAULT_DPI 96
#define SCALEX(X) MulDiv(X, uDpiX, DEFAULT_DPI)
#define SCALEY(Y) MulDiv(Y, uDpiY, DEFAULT_DPI)
#define POINT2PIXEL(PT) MulDiv(PT, uDpiY, 72)

TCHAR szClassName[] = TEXT("aichat");
WCHAR szAPIKey[] = L"YOUR_API_KEY";

LPWSTR a2w(LPCSTR lpszText)
{
	const DWORD dwSize = MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, 0, 0);
	LPWSTR lpszReturnText = (LPWSTR)GlobalAlloc(0, dwSize * sizeof(WCHAR));
	MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, lpszReturnText, dwSize);
	return lpszReturnText;
}

LPSTR w2a(LPCWSTR lpszText)
{
	const DWORD dwSize = WideCharToMultiByte(CP_UTF8, 0, lpszText, -1, 0, 0, 0, 0);
	LPSTR lpszReturnText = (char*)GlobalAlloc(GPTR, dwSize * sizeof(char));
	WideCharToMultiByte(CP_UTF8, 0, lpszText, -1, lpszReturnText, dwSize, 0, 0);
	return lpszReturnText;
}

BOOL GetScaling(HWND hWnd, UINT* pnX, UINT* pnY)
{
	BOOL bSetScaling = FALSE;
	const HMONITOR hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
	if (hMonitor)
	{
		HMODULE hShcore = LoadLibraryW(L"SHCORE");
		if (hShcore)
		{
			typedef HRESULT __stdcall GetDpiForMonitor(HMONITOR, int, UINT*, UINT*);
			GetDpiForMonitor* fnGetDpiForMonitor = reinterpret_cast<GetDpiForMonitor*>(GetProcAddress(hShcore, "GetDpiForMonitor"));
			if (fnGetDpiForMonitor)
			{
				UINT uDpiX, uDpiY;
				if (SUCCEEDED(fnGetDpiForMonitor(hMonitor, 0, &uDpiX, &uDpiY)) && uDpiX > 0 && uDpiY > 0)
				{
					*pnX = uDpiX;
					*pnY = uDpiY;
					bSetScaling = TRUE;
				}
			}
			FreeLibrary(hShcore);
		}
	}
	if (!bSetScaling)
	{
		HDC hdc = GetDC(NULL);
		if (hdc)
		{
			*pnX = GetDeviceCaps(hdc, LOGPIXELSX);
			*pnY = GetDeviceCaps(hdc, LOGPIXELSY);
			ReleaseDC(NULL, hdc);
			bSetScaling = TRUE;
		}
	}
	if (!bSetScaling)
	{
		*pnX = DEFAULT_DPI;
		*pnY = DEFAULT_DPI;
		bSetScaling = TRUE;
	}
	return bSetScaling;
}

LPWSTR RunCommand(LPCWSTR lpszMessage)
{
	if (!lpszMessage) return 0;
	LPBYTE lpszByte = 0;
	DWORD dwSize = 0;
	const HINTERNET hSession = InternetOpenW(0, INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, INTERNET_FLAG_NO_COOKIES);
	if (hSession)
	{
		const HINTERNET hConnection = InternetConnectW(hSession, L"api.openai.com", INTERNET_DEFAULT_HTTPS_PORT, 0, 0, INTERNET_SERVICE_HTTP, 0, 0);
		if (hConnection)
		{
			const HINTERNET hRequest = HttpOpenRequestW(hConnection, L"POST", L"/v1/chat/completions", 0, 0, 0, INTERNET_FLAG_NO_CACHE_WRITE | INTERNET_FLAG_SECURE, 0);
			if (hRequest)
			{
				WCHAR szHeaders[1024] = L"";
				wsprintf(szHeaders, L"Content-Type: application/json\r\nAuthorization: Bearer %s", szAPIKey);
				std::string body_json = "{\"model\":\"gpt-3.5-turbo\", \"messages\":[{\"role\":\"user\", \"content\":\"";
				LPSTR lpszMessageA = w2a(lpszMessage);
				body_json += lpszMessageA;
				body_json += "\"}]}";
				GlobalFree(lpszMessageA);
				HttpSendRequestW(hRequest, szHeaders, lstrlenW(szHeaders), (LPVOID)(body_json.c_str()), (DWORD)body_json.length());
				lpszByte = (LPBYTE)GlobalAlloc(GMEM_FIXED, 1);
				if (lpszByte) {
					DWORD dwRead;
					BYTE szBuf[1024 * 4];
					for (;;) {
						if (!InternetReadFile(hRequest, szBuf, (DWORD)sizeof(szBuf), &dwRead) || !dwRead) break;
						LPBYTE lpTmp = (LPBYTE)GlobalReAlloc(lpszByte, (SIZE_T)(dwSize + dwRead), GMEM_MOVEABLE);
						if (lpTmp == NULL) break;
						lpszByte = lpTmp;
						CopyMemory(lpszByte + dwSize, szBuf, dwRead);
						dwSize += dwRead;
					}
				}
				InternetCloseHandle(hRequest);
			}
			InternetCloseHandle(hConnection);
		}
		InternetCloseHandle(hSession);
	}
	if (lpszByte)
	{
		std::string src((LPSTR)lpszByte, dwSize);
		GlobalFree(lpszByte);
		lpszByte = 0;
		std::string err;
		json11::Json v = json11::Json::parse(src, err);
		for (auto& item : v["choices"].array_items()) {
			return a2w(item["message"]["content"].string_value().c_str());
		}
	}
	LPCWSTR lpszErrorMessage = L"エラーとなりました。しばらくたってからリトライしてください。";
	LPWSTR lpszReturn = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR) * (lstrlenW(lpszErrorMessage) + 1));
	if (lpszReturn) {
		lstrcpy(lpszReturn, lpszErrorMessage);
	}
	return lpszReturn;
}

LPWSTR HtmlEncode(LPCWSTR lpText)
{
	int i, m;
	LPWSTR ptr;
	const int n = lstrlenW(lpText);
	for (i = 0, m = 0; i < n; i++)
	{
		switch (lpText[i])
		{
		case L'>':
		case L'<':
			m += 4;
			break;
		case L'&':
			m += 5;
			break;
		default:
			m++;
		}
	}
	if (n == m)return 0;
	LPWSTR lpOutText = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR) * (m + 1));
	if (lpOutText) {
		for (i = 0, ptr = lpOutText; i <= n; i++) {
			switch (lpText[i]) {
			case L'>':
				lstrcpyW(ptr, L"&gt;");
				ptr += lstrlenW(ptr);
				break;
			case L'<':
				lstrcpyW(ptr, L"&lt;");
				ptr += lstrlenW(ptr);
				break;
			case L'&':
				lstrcpyW(ptr, L"&amp;");
				ptr += lstrlenW(ptr);
				break;
			default:
				*ptr++ = lpText[i];
			}
		}
	}
	return lpOutText;
}

DWORD WINAPI ThreadFunc(LPVOID p)
{
	HWND hWnd = (HWND)p;
	LPCWSTR lpszMessage = (LPCWSTR)GetWindowLongPtr(hWnd, DWLP_USER);
	PostMessageW(hWnd, WM_APP, 0, (LPARAM)RunCommand(lpszMessage));
	ExitThread(0);
}

VOID execBrowserCommand(IHTMLDocument2* pDocument, LPCWSTR lpszMessage)
{
	VARIANT var = { 0 };
	VARIANT_BOOL varBool = { 0 };
	var.vt = VT_EMPTY;
	BSTR command = SysAllocString(lpszMessage);
	pDocument->execCommand(command, VARIANT_FALSE, var, &varBool);
	SysFreeString(command);
}

VOID ScrollBottom(IHTMLDocument2* pDocument)
{
	IHTMLWindow2* pHtmlWindow2;
	pDocument->get_parentWindow(&pHtmlWindow2);
	if (pHtmlWindow2) {
		pHtmlWindow2->scrollTo(0, SHRT_MAX);
		pHtmlWindow2->Release();
	}
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static CComQIPtr<IWebBrowser2>pWB2;
	static CComQIPtr<IHTMLDocument2>pDocument;
	static HWND hEdit, hBrowser;
	static HBRUSH hBrush;
	static HFONT hFont;
	static UINT uDpiX = DEFAULT_DPI, uDpiY = DEFAULT_DPI;
	static HANDLE hThread;
	static LPWSTR lpszMessage;
	switch (msg)
	{
	case WM_CREATE:
		hBrush = CreateSolidBrush(RGB(200, 219, 249));
		hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", 0, WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessageW(hEdit, EM_SETCUEBANNER, TRUE, (LPARAM)L"メッセージを入力");
		hBrowser = CreateWindowExW(0, L"AtlAxWin140", L"about:blank", WS_CHILD | WS_VISIBLE | WS_VSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		{
			CComPtr<IUnknown> punkIE;
			if (AtlAxGetControl(hBrowser, &punkIE) == S_OK) {
				pWB2 = punkIE;
				punkIE.Release();
				if (pWB2) {
					pWB2->put_Silent(VARIANT_TRUE);
					pWB2->put_RegisterAsDropTarget(VARIANT_FALSE);
					pWB2->get_Document((IDispatch**)&pDocument);
					{
						CComPtr<IUnknown> punkIE;
						ATL::CComPtr<IAxWinAmbientDispatch> ambient;
						AtlAxGetHost(hBrowser, &punkIE);
						ambient = punkIE;
						if (ambient) {
							ambient->put_AllowContextMenu(VARIANT_FALSE);
						}
					}
				}
			}
		}
		if (!pDocument) {
			return -1;
		}
		{
			WCHAR szModuleFilePath[MAX_PATH] = { 0 };
			GetModuleFileNameW(NULL, szModuleFilePath, MAX_PATH);
			WCHAR szURL[MAX_PATH] = { 0 };
			wsprintfW(szURL, L"res://%s/%d", szModuleFilePath, IDR_HTML1);
			BSTR url = SysAllocString(szURL);
			pWB2->Navigate(url, NULL, NULL, NULL, NULL);
			SysFreeString(url);
		}
		SendMessage(hWnd, WM_APP + 1, 0, 0);
		RECT rect;
		GetWindowRect(hWnd, &rect);
		SetWindowPos(hWnd, NULL, 0, 0, (rect.right - rect.left) / 3, (rect.bottom - rect.top) / 2, SWP_NOMOVE | SWP_NOZORDER | SWP_NOSENDCHANGING | SWP_NOREDRAW);
		break;
	case WM_SIZE:
		MoveWindow(hBrowser, 0, 0, LOWORD(lParam), HIWORD(lParam) - POINT2PIXEL(32), TRUE);
		MoveWindow(hEdit, 0, HIWORD(lParam) - POINT2PIXEL(32), LOWORD(lParam), POINT2PIXEL(32), TRUE);
		break;
	case WM_CTLCOLOREDIT:
	case WM_CTLCOLORSTATIC:
		SetBkMode((HDC)wParam, TRANSPARENT);
		return(INT_PTR)hBrush;
	case WM_SETFOCUS:
		SetFocus(hEdit);
		break;
	case WM_APP:
		WaitForSingleObject(hThread, INFINITE);
		CloseHandle(hThread);
		hThread = 0;
		{
			LPWSTR lpszReturnString = (LPWSTR)lParam;
			if (lpszReturnString) {
				StrTrim(lpszReturnString, L"\n");
				LPWSTR lpszReturnString2 = HtmlEncode(lpszReturnString);
				if (lpszReturnString2) {
					GlobalFree((HGLOBAL)lpszReturnString);
					lpszReturnString = lpszReturnString2;
				}
				const int nHtmlLength = lstrlenW(lpszReturnString) + 256;
				LPWSTR lpszHtml = (LPWSTR)GlobalAlloc(0, nHtmlLength * sizeof(WCHAR));
				if (lpszHtml) {
					lstrcpyW(lpszHtml, L"<div class=\"result\"><div class=\"icon\"><img></div><div class=\"output\"><pre>");
					lstrcatW(lpszHtml, lpszReturnString);
					GlobalFree((HGLOBAL)lpszReturnString);
					lstrcatW(lpszHtml, L"</pre></div></div>");
					CComQIPtr<IHTMLElement>pElementBody;
					pDocument->get_body((IHTMLElement**)&pElementBody);
					if (pElementBody) {
						BSTR where = SysAllocString(L"beforeEnd");
						BSTR html = SysAllocString(lpszHtml);
						pElementBody->insertAdjacentHTML(where, html);
						SysFreeString(where);
						SysFreeString(html);
						pElementBody.Release();
						ScrollBottom(pDocument);
					}
					GlobalFree(lpszHtml);
				}
			}
		}
		EnableWindow(hBrowser, TRUE);
		EnableWindow(hEdit, TRUE);
		SetFocus(hEdit);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == ID_RUN)
		{
			if (GetFocus() != hEdit)
				SetFocus(hEdit);
			const int nTextLength = GetWindowTextLength(hEdit);
			if (nTextLength > 0) {
				GlobalFree(lpszMessage);
				lpszMessage = (LPWSTR)GlobalAlloc(0, (nTextLength + 1) * sizeof(WCHAR));
				if (lpszMessage) {
					GetWindowTextW(hEdit, lpszMessage, nTextLength + 1);
					SetWindowLongPtrW(hWnd, DWLP_USER, (LONG_PTR)lpszMessage);
					LPWSTR lpszMessage2 = HtmlEncode(lpszMessage);
					UINT nHtmlLength = (lpszMessage2 ? lstrlenW(lpszMessage2) : nTextLength) + 256;
					LPWSTR lpszHtml = (LPWSTR)GlobalAlloc(0, nHtmlLength * sizeof(WCHAR));
					if (lpszHtml) {
						lstrcpyW(lpszHtml, L"<div class=\"input\"><pre>");
						lstrcatW(lpszHtml, lpszMessage2 ? lpszMessage2 : lpszMessage);
						lstrcatW(lpszHtml, L"</pre></div>");
						CComQIPtr<IHTMLElement>pElementBody;
						pDocument->get_body((IHTMLElement**)&pElementBody);
						if (pElementBody) {
							BSTR where = SysAllocString(L"beforeEnd");
							BSTR html = SysAllocString(lpszHtml);
							pElementBody->insertAdjacentHTML(where, html);
							SysFreeString(where);
							SysFreeString(html);
							pElementBody.Release();
							ScrollBottom(pDocument);
						}
						GlobalFree(lpszHtml);
						SetWindowText(hEdit, 0);
						DWORD dwParam;
						EnableWindow(hBrowser, FALSE);
						EnableWindow(hEdit, FALSE);
						hThread = CreateThread(0, 0, ThreadFunc, (LPVOID)hWnd, 0, &dwParam);
					}
					GlobalFree((HGLOBAL)lpszMessage2);
				}
			}
		}
		else if (LOWORD(wParam) == ID_COPY) {
			if (GetFocus() == hEdit) {
				SendMessage(hEdit, WM_COPY, 0, 0);
			}
			else {
				execBrowserCommand(pDocument, L"Copy");
				execBrowserCommand(pDocument, L"Unselect");
				SetFocus(hEdit);
			}
		}
		else if (LOWORD(wParam) == ID_TAB) {
			execBrowserCommand(pDocument, L"Unselect");
			SendMessage(hEdit, EM_SETSEL, 0, -1);
			SetFocus(hEdit);
		}
		break;
	case WM_NCCREATE:
	{
		const HMODULE hModUser32 = GetModuleHandleW(L"user32.dll");
		if (hModUser32)
		{
			typedef BOOL(WINAPI* fnTypeEnableNCScaling)(HWND);
			const fnTypeEnableNCScaling fnEnableNCScaling = (fnTypeEnableNCScaling)GetProcAddress(hModUser32, "EnableNonClientDpiScaling");
			if (fnEnableNCScaling)
			{
				fnEnableNCScaling(hWnd);
			}
		}
	}
	return DefWindowProc(hWnd, msg, wParam, lParam);
	case WM_DPICHANGED:
	case WM_APP + 1:
		GetScaling(hWnd, &uDpiX, &uDpiY);
		DeleteObject(hFont);
		hFont = CreateFontW(-POINT2PIXEL(12), 0, 0, 0, FW_NORMAL, 0, 0, 0, ANSI_CHARSET, 0, 0, 0, 0, L"Consolas");
		SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, 0);
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		GlobalFree(lpszMessage);
		pDocument.Release();
		pWB2.Release();
		DeleteObject(hBrush);
		DeleteObject(hFont);
		PostQuitMessage(0);
		break;
	default:
		return DefDlgProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	if (FAILED(CoInitialize(NULL)))
	{
		return 0;
	}
	::AtlAxWinInit();
	MSG msg = { 0 };
	{
		CComModule _Module;
		WNDCLASS wndclass = {
			0,
			WndProc,
			0,
			DLGWINDOWEXTRA,
			hInstance,
			LoadIcon(hInstance,(LPCTSTR)IDI_ICON1),
			0,
			0,
			0,
			szClassName
		};
		RegisterClass(&wndclass);
		HWND hWnd = CreateWindow(
			szClassName,
			TEXT("AI chat"),
			WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
			CW_USEDEFAULT,
			0,
			CW_USEDEFAULT,
			0,
			0,
			0,
			hInstance,
			0
		);
		ShowWindow(hWnd, SW_SHOWDEFAULT);
		UpdateWindow(hWnd);
		HACCEL hAccel = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));
		while (GetMessage(&msg, 0, 0, 0))
		{
			if (!TranslateAccelerator(hWnd, hAccel, &msg) && !IsDialogMessage(hWnd, &msg))
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
	}
	::AtlAxWinTerm();
	CoUninitialize();
	return (int)msg.wParam;
}
