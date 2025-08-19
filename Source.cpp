#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>  // DefSubclassProc 用
#include <winhttp.h>
#include <vector>
#include <string>
#include "json11.hpp"  // json11 の使用を前提
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "comctl32.lib")  // 必須

WCHAR szAPIKey[] = L"YOUR_API_KEY";

struct ChatMessage {
    std::wstring text;
    bool isMine;
    RECT rect;
};

struct ThreadData {
    HWND hWnd;
    LPWSTR lpszMessage;
};

std::vector<ChatMessage> g_messages;
int g_scrollPos = 0;      // 縦スクロール位置（ピクセル）
int g_contentHeight = 0;  // 全メッセージの合計高さ
int g_lineHeight = 20;    // 最低行送り

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

LPWSTR RunCommand(LPCWSTR lpszMessage)
{
    if (!lpszMessage) return 0;

    std::string response;
    DWORD dwSize = 0;

    // WinHTTP セッションを作成
    HINTERNET hSession = WinHttpOpen(L"WinHTTP ChatClient/1.0",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return 0;

    // 接続
    HINTERNET hConnect = WinHttpConnect(hSession, L"api.openai.com",
        INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hConnect) {
        WinHttpCloseHandle(hSession);
        return 0;
    }

    // POST リクエスト
    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST",
        L"/v1/chat/completions",
        nullptr, WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES,
        WINHTTP_FLAG_SECURE);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return 0;
    }

    // ヘッダー作成
    WCHAR szHeaders[1024];
    wsprintf(szHeaders, L"Content-Type: application/json\r\nAuthorization: Bearer %s", szAPIKey);

    // JSON ボディ作成
    std::string body_json = "{\"model\":\"gpt-3.5-turbo\", \"messages\":[{\"role\":\"user\", \"content\":\"";
    LPSTR lpszMessageA = w2a(lpszMessage);
    body_json += lpszMessageA;
    body_json += "\"}]}";
    GlobalFree(lpszMessageA);

    // 送信
    BOOL bResult = WinHttpSendRequest(hRequest, szHeaders, (DWORD)wcslen(szHeaders),
        (LPVOID)body_json.c_str(), (DWORD)body_json.length(),
        (DWORD)body_json.length(), 0);
    if (!bResult) {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        goto error_return;
    }

    // レスポンス受信
    if (!WinHttpReceiveResponse(hRequest, nullptr)) {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        goto error_return;
    }

    for (;;) {
        DWORD dwAvailable = 0;
        if (!WinHttpQueryDataAvailable(hRequest, &dwAvailable) || dwAvailable == 0) break;
        std::unique_ptr<char[]> buffer(new char[dwAvailable]);
        DWORD dwRead = 0;
        if (!WinHttpReadData(hRequest, buffer.get(), dwAvailable, &dwRead) || dwRead == 0) break;
        response.append(buffer.get(), dwRead);
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    {
        std::string err;
        json11::Json json = json11::Json::parse(response, err);
        if (!err.empty()) {
            LPWSTR result = (LPWSTR)GlobalAlloc(GMEM_FIXED, 64 * sizeof(wchar_t));
            wcscpy_s(result, 64, L"JSON解析エラー");
            return result;
        }

        auto choices = json["choices"].array_items();
        if (!choices.empty()) {
            std::wstring content = a2w(choices[0]["message"]["content"].string_value().c_str());
            LPWSTR result = (LPWSTR)GlobalAlloc(GMEM_FIXED, (content.size() + 1) * sizeof(wchar_t));
            wcscpy_s(result, content.size() + 1, content.c_str());
            return result;
        }
    }

error_return:
    LPCWSTR lpszErrorMessage = L"エラーとなりました。しばらくたってからリトライしてください。";
    LPWSTR lpszReturn = (LPWSTR)GlobalAlloc(GMEM_FIXED, sizeof(WCHAR) * (lstrlenW(lpszErrorMessage) + 1));
    if (lpszReturn) {
        lstrcpy(lpszReturn, lpszErrorMessage);
    }
    return lpszReturn;
}

DWORD WINAPI ThreadFunc(LPVOID p)
{
    ThreadData* data = (ThreadData*)p;
    PostMessageW(data->hWnd, WM_APP, 0, (LPARAM)RunCommand(data->lpszMessage));
    ExitThread(0);
}

// フォントに基づいて行の高さを更新
void UpdateLineHeight(HWND hWnd) {
    HDC hdc = GetDC(hWnd);
    HFONT hFont = (HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0);
    HFONT hOld = nullptr;
    if (hFont) hOld = (HFONT)SelectObject(hdc, hFont);

    TEXTMETRIC tm;
    GetTextMetrics(hdc, &tm);
    g_lineHeight = tm.tmHeight + tm.tmExternalLeading + 10;

    if (hOld) SelectObject(hdc, hOld);
    ReleaseDC(hWnd, hdc);
}

void UpdateScrollBar(HWND hWnd) {
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);

    const int inputHeight = 32;    // 入力欄の高さ
    const int bottomMargin = 64;   // 入力欄の上の余白
    int visibleHeight = rcClient.bottom - inputHeight - bottomMargin;

    HDC hdc = GetDC(hWnd);
    HFONT hFont = (HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0);
    HFONT hOld = nullptr;
    if (hFont) hOld = (HFONT)SelectObject(hdc, hFont);

    const int bubblePaddingX = 15;
    const int bubblePaddingY = 10;

    g_contentHeight = 5;
    for (auto& msg : g_messages) {
        RECT rcCalc = { 0,0, rcClient.right / 2,0 };
        DrawTextW(hdc, msg.text.c_str(), -1, &rcCalc,
            DT_CALCRECT | DT_WORDBREAK | DT_LEFT | DT_NOPREFIX);

        int bubbleHeight = (rcCalc.bottom - rcCalc.top) + bubblePaddingY * 2;
        int lineHeight = max(bubbleHeight + 10, g_lineHeight);
        g_contentHeight += lineHeight;
    }

    if (hOld) SelectObject(hdc, hOld);
    ReleaseDC(hWnd, hdc);

    SCROLLINFO si = { sizeof(SCROLLINFO), SIF_RANGE | SIF_PAGE | SIF_POS };
    si.nMin = 0;
    si.nMax = max(g_contentHeight - 1, 0);
    si.nPage = visibleHeight;
    if (g_scrollPos > si.nMax - si.nPage + 1)
        g_scrollPos = max(0, si.nMax - si.nPage + 1);
    si.nPos = g_scrollPos;
    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
}

void ScrollToBottom(HWND hWnd) {
    SCROLLINFO si = { sizeof(SCROLLINFO), SIF_ALL };
    GetScrollInfo(hWnd, SB_VERT, &si);

    // 最大値を取得して一番下に設定
    si.nPos = si.nMax - (int)si.nPage + 1;
    if (si.nPos < si.nMin) si.nPos = si.nMin;

    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
    g_scrollPos = si.nPos;

    InvalidateRect(hWnd, NULL, TRUE);
}

void AddMessage(HWND hWnd, const std::wstring& msg, bool mine) {
    ChatMessage m{ msg, mine };
    g_messages.push_back(m);
    UpdateScrollBar(hWnd);
    ScrollToBottom(hWnd);
}

void OnPaint(HWND hWnd, HDC hdc) {
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);

    const int inputHeight = 32;
    const int bottomMargin = 64; // 入力欄上の余白
    const int bottomLimit = rcClient.bottom - inputHeight - bottomMargin;

    const int bubblePaddingX = 15;
    const int bubblePaddingY = 10;

    HBRUSH hBrushMine = CreateSolidBrush(RGB(187, 215, 234));
    HBRUSH hBrushOther = CreateSolidBrush(RGB(244, 244, 244));
    HPEN hPen = CreatePen(PS_SOLID, 1, RGB(250, 250, 250));
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

    int y = 5 - g_scrollPos;
    for (auto& msg : g_messages) {
        RECT rcCalc = { 0,0, rcClient.right / 2,0 };
        DrawTextW(hdc, msg.text.c_str(), -1, &rcCalc,
            DT_CALCRECT | DT_WORDBREAK | DT_LEFT | DT_NOPREFIX);

        int bubbleWidth = (rcCalc.right - rcCalc.left) + bubblePaddingX * 2;
        int bubbleHeight = (rcCalc.bottom - rcCalc.top) + bubblePaddingY * 2;
        int lineHeight = max(bubbleHeight + 10, g_lineHeight);

        if (y + bubbleHeight > bottomLimit)
            break; // 入力欄の上限まで

        HBRUSH hOldBrush;
        if (msg.isMine) {
            msg.rect = { rcClient.right - bubbleWidth - 10, y,
                         rcClient.right - 10, y + bubbleHeight };
            hOldBrush = (HBRUSH)SelectObject(hdc, hBrushMine);
        }
        else {
            msg.rect = { 10, y,
                         10 + bubbleWidth, y + bubbleHeight };
            hOldBrush = (HBRUSH)SelectObject(hdc, hBrushOther);
        }

        RoundRect(hdc, msg.rect.left, msg.rect.top,
            msg.rect.right, msg.rect.bottom, 20, 20);

        RECT rcDraw = msg.rect;
        InflateRect(&rcDraw, -bubblePaddingX, -bubblePaddingY);
        SetBkMode(hdc, TRANSPARENT);

        DrawTextW(hdc, msg.text.c_str(), -1, &rcDraw,
            DT_LEFT | DT_NOPREFIX | DT_WORDBREAK);

        SelectObject(hdc, hOldBrush);
        y += lineHeight;
    }

    SelectObject(hdc, hOldPen);
    DeleteObject(hPen);
    DeleteObject(hBrushMine);
    DeleteObject(hBrushOther);
}

void OnVScroll(HWND hWnd, WPARAM wParam) {
    SCROLLINFO si = { sizeof(SCROLLINFO), SIF_ALL };
    GetScrollInfo(hWnd, SB_VERT, &si);
    int pos = si.nPos;

    switch (LOWORD(wParam)) {
    case SB_LINEUP:   si.nPos -= 20; break;
    case SB_LINEDOWN: si.nPos += 20; break;
    case SB_PAGEUP:   si.nPos -= (int)si.nPage; break;
    case SB_PAGEDOWN: si.nPos += (int)si.nPage; break;
    case SB_THUMBTRACK: si.nPos = si.nTrackPos; break;
    }

    if (si.nPos < si.nMin) si.nPos = si.nMin;
    if (si.nPos > si.nMax - (int)si.nPage + 1) si.nPos = si.nMax - (int)si.nPage + 1;

    si.fMask = SIF_POS;
    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
    g_scrollPos = si.nPos;
    InvalidateRect(hWnd, NULL, TRUE);
}

void CopyToClipboard(HWND hWnd, const std::wstring& text) {
    if (OpenClipboard(hWnd)) {
        EmptyClipboard();
        size_t size = (text.size() + 1) * sizeof(wchar_t);
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, size);
        if (hMem) {
            memcpy(GlobalLock(hMem), text.c_str(), size);
            GlobalUnlock(hMem);
            SetClipboardData(CF_UNICODETEXT, hMem);
        }
        CloseClipboard();
    }
}

LRESULT CALLBACK EditSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    switch (msg)
    {
    case WM_CHAR:
        if (wParam == VK_RETURN) {
            SendMessage(GetParent(hWnd), WM_APP + 1, 0, 0);
            return 0;
        }
        break;
    }
    return DefSubclassProc(hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HWND hEdit;
    static HBRUSH hBrushOther;
    static LPWSTR text;
    static ThreadData data;
    static HANDLE hThread;
    switch (msg) {
    case WM_CREATE:
        UpdateLineHeight(hWnd);
        hBrushOther = CreateSolidBrush(RGB(244, 244, 244));
        hEdit = CreateWindow(L"EDIT", 0, WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
        SetWindowSubclass(hEdit, EditSubclassProc, 0, 0);
        return 0;

    case WM_APP: {
        WaitForSingleObject(hThread, INFINITE);
        CloseHandle(hThread);
        hThread = 0;
        LPWSTR lpszMessage = (LPWSTR)lParam;
        if (lpszMessage) {
            AddMessage(hWnd, lpszMessage, false);
            GlobalFree(lpszMessage);
        }
        EnableWindow(hEdit, TRUE);
        SetFocus(hEdit);
        return 0;
    }

    case WM_APP + 1: {
        int len = GetWindowTextLength(hEdit);
        if (len > 0) {
            if (text) {
                GlobalFree(text);
                text = nullptr;
            }
            text = (LPWSTR)GlobalAlloc(0, (len + 1) * sizeof(WCHAR));
            if (text) {
                GetWindowText(hEdit, text, len + 1);
                AddMessage(hWnd, text, true);
                SetWindowText(hEdit, 0);
                EnableWindow(hEdit, FALSE);
                data.hWnd = hWnd;
                data.lpszMessage = text;
                DWORD dwParam;
                hThread = CreateThread(0, 0, ThreadFunc, (LPVOID)&data, 0, &dwParam);
            }
        }
        return 0;
    }

    case WM_CTLCOLOREDIT:
        if ((HWND)lParam == hEdit) {
            HDC hdc = (HDC)wParam;
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, RGB(0, 0, 0));
            return (LRESULT)hBrushOther;
        }
        break;

    case WM_SETFONT:
        UpdateLineHeight(hWnd);
        UpdateScrollBar(hWnd);
        InvalidateRect(hWnd, NULL, TRUE);
        break;

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        OnPaint(hWnd, hdc);

        {
            RECT rect;
            GetClientRect(hWnd, &rect);
            HPEN hPen = CreatePen(PS_SOLID, 1, RGB(250, 250, 250));
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, hBrushOther);
            RoundRect(hdc, 32, rect.bottom - 64 - 32,
                rect.right - 32, rect.bottom - 32, 32, 32);
            SelectObject(hdc, hOldBrush);
            SelectObject(hdc, hOldPen);
            DeleteObject(hPen);
        }

        EndPaint(hWnd, &ps);
    } return 0;

    case WM_VSCROLL:
        OnVScroll(hWnd, wParam);
        return 0;

    case WM_MOUSEWHEEL: {
        int zDelta = GET_WHEEL_DELTA_WPARAM(wParam);
        int scroll = -(zDelta / WHEEL_DELTA) * 40; // 1ノッチ40pxスクロール
        g_scrollPos += scroll;
        if (g_scrollPos < 0) g_scrollPos = 0;
        if (g_scrollPos > g_contentHeight) g_scrollPos = g_contentHeight;
        UpdateScrollBar(hWnd);
        InvalidateRect(hWnd, NULL, TRUE);
        return 0;
    }

    case WM_SIZE:
        UpdateScrollBar(hWnd);
        MoveWindow(hEdit, 32 + 16, HIWORD(lParam) - 32 - 32 - 16, LOWORD(lParam) - 64 - 32, 32, TRUE);
        return 0;

    case WM_SETFOCUS:
        SetFocus(hEdit);
        return 0;

    case WM_RBUTTONDOWN: {
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        for (auto& msg : g_messages) {
            if (PtInRect(&msg.rect, pt)) {
                CopyToClipboard(hWnd, msg.text);
                MessageBoxW(hWnd, L"コピーしました！", L"情報", MB_OK);
                break;
            }
        }
        return 0;
    }

    case WM_DESTROY:
        if (text) {
            GlobalFree(text);
            text = nullptr;
        }
        DeleteObject(hBrushOther);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hWnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int) {
    INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icex);  // DefSubclassProc を使う前に必須
    MSG msg;
    WNDCLASS wc = {
        CS_HREDRAW | CS_VREDRAW,
        WndProc,0,0,hInstance,
        0, LoadCursor(0,IDC_ARROW),
        (HBRUSH)(COLOR_WINDOW + 1),0,L"aichat"
    };
    RegisterClass(&wc);
    HWND hWnd = CreateWindow(
        L"aichat", L"aichat",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_VSCROLL,
        CW_USEDEFAULT, 0, CW_USEDEFAULT, 0,
        0, 0, hInstance, 0);
    ShowWindow(hWnd, SW_SHOWDEFAULT);
    UpdateWindow(hWnd);
    while (GetMessage(&msg, 0, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}