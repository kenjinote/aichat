#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "d2d1")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "winhttp")
#pragma comment(lib, "comctl32")

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <winhttp.h>
#include <d2d1.h>
#include <dwrite.h>
#include <vector>
#include <string>
#include "json11.hpp"
#include "resource.h"

WCHAR szAPIKey[] = L"YOUR_API_KEY";

struct ChatMessage
{
    std::wstring text;
    bool isMine;
    RECT rect;
};

std::vector<ChatMessage> g_messages;
int g_scrollPos = 0;      // 縦スクロール位置（ピクセル）
int g_contentHeight = 0;  // 全メッセージの合計高さ
int g_lineHeight = 20;    // 最低行送り
ID2D1Factory* g_pD2DFactory = nullptr;
ID2D1HwndRenderTarget* g_pRenderTarget = nullptr;
ID2D1SolidColorBrush* g_pBrushMine = nullptr;
ID2D1SolidColorBrush* g_pBrushOther = nullptr;
ID2D1SolidColorBrush* g_pBrushText = nullptr;
IDWriteFactory* g_pDWriteFactory = nullptr;
IDWriteTextFormat* g_pTextFormat = nullptr;

// 初期化
HRESULT InitD2D(HWND hWnd)
{
    HRESULT hr;
    // Direct2D ファクトリ
    if (!g_pD2DFactory) {
        hr = D2D1CreateFactory(
            D2D1_FACTORY_TYPE_SINGLE_THREADED,
            &g_pD2DFactory
        );
        if (FAILED(hr)) return hr;
    }
    // クライアント領域取得
    RECT rc;
    GetClientRect(hWnd, &rc);
    D2D1_SIZE_U size = D2D1::SizeU(rc.right - rc.left, rc.bottom - rc.top);
    // RenderTarget 作成
    if (!g_pRenderTarget) {
        hr = g_pD2DFactory->CreateHwndRenderTarget(
            D2D1::RenderTargetProperties(),
            D2D1::HwndRenderTargetProperties(hWnd, size),
            &g_pRenderTarget
        );
        if (FAILED(hr)) return hr;
    }
    // ブラシ作成
    if (!g_pBrushMine) {
        hr = g_pRenderTarget->CreateSolidColorBrush(
            D2D1::ColorF(0xBB / 255.0f, 0xD7 / 255.0f, 0xEA / 255.0f), // RGB(187,215,234)
            &g_pBrushMine
        );
        if (FAILED(hr)) return hr;
    }
    if (!g_pBrushOther) {
        hr = g_pRenderTarget->CreateSolidColorBrush(
            D2D1::ColorF(0xF4 / 255.0f, 0xF4 / 255.0f, 0xF4 / 255.0f), // RGB(244,244,244)
            &g_pBrushOther
        );
        if (FAILED(hr)) return hr;
    }
    if (!g_pBrushText) {
        hr = g_pRenderTarget->CreateSolidColorBrush(
            D2D1::ColorF(D2D1::ColorF::Black),
            &g_pBrushText
        );
        if (FAILED(hr)) return hr;
    }
    // DirectWrite 初期化
    if (!g_pDWriteFactory) {
        hr = DWriteCreateFactory(
            DWRITE_FACTORY_TYPE_SHARED,
            __uuidof(IDWriteFactory),
            reinterpret_cast<IUnknown**>(&g_pDWriteFactory)
        );
        if (FAILED(hr)) return hr;
    }
    if (!g_pTextFormat) {
        hr = g_pDWriteFactory->CreateTextFormat(
            L"Segoe UI",                // フォント
            nullptr,
            DWRITE_FONT_WEIGHT_REGULAR,
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            18.0f,
            L"ja-jp",
            &g_pTextFormat
        );
        if (FAILED(hr)) return hr;
    }
    return S_OK;
}

void CleanupD2D()
{
    if (g_pBrushMine) { g_pBrushMine->Release();   g_pBrushMine = nullptr; }
    if (g_pBrushOther) { g_pBrushOther->Release();  g_pBrushOther = nullptr; }
    if (g_pBrushText) { g_pBrushText->Release();   g_pBrushText = nullptr; }
    if (g_pTextFormat) { g_pTextFormat->Release();  g_pTextFormat = nullptr; }
    if (g_pDWriteFactory) { g_pDWriteFactory->Release(); g_pDWriteFactory = nullptr; }
    if (g_pRenderTarget) { g_pRenderTarget->Release(); g_pRenderTarget = nullptr; }
    if (g_pD2DFactory) { g_pD2DFactory->Release();   g_pD2DFactory = nullptr; }
}

void DiscardDeviceResources()
{
    if (g_pBrushMine) { g_pBrushMine->Release();   g_pBrushMine = nullptr; }
    if (g_pBrushOther) { g_pBrushOther->Release();  g_pBrushOther = nullptr; }
    if (g_pBrushText) { g_pBrushText->Release();   g_pBrushText = nullptr; }
    if (g_pRenderTarget) { g_pRenderTarget->Release(); g_pRenderTarget = nullptr; }
}

// 先頭付近のグローバルと合わせる前提
int GetVisibleHeight(HWND hWnd)
{
    RECT rcClient; GetClientRect(hWnd, &rcClient);
    const int inputHeight = 32;   // 入力欄の高さ
    const int bottomMargin = 64;  // 入力欄の上の余白
    const int topMargin = 5;      // 上の余白（描画開始 y と合わせる）
    int bottomLimit = rcClient.bottom - inputHeight - bottomMargin;
    int visibleHeight = bottomLimit - topMargin;
    return max(0, visibleHeight);
}

void RecalcContentHeight(HWND hWnd)
{
    // DirectWrite で OnPaint と同じ条件で高さを測る
    RECT rcClient; GetClientRect(hWnd, &rcClient);
    const float layoutWidth = rcClient.right / 2.0f;
    const float bubblePaddingY = 10.0f;
    g_contentHeight = 5; // 上余白ぶん
    if (!g_pDWriteFactory || !g_pTextFormat) return;
    for (auto& msg : g_messages) {
        IDWriteTextLayout* pTextLayout = nullptr;
        HRESULT hr = g_pDWriteFactory->CreateTextLayout(
            msg.text.c_str(),
            (UINT32)msg.text.length(),
            g_pTextFormat,
            layoutWidth,
            FLT_MAX,
            &pTextLayout
        );
        if (FAILED(hr) || !pTextLayout) continue;
        DWRITE_TEXT_METRICS metrics{};
        pTextLayout->GetMetrics(&metrics);
        int bubbleHeight = (int)ceilf(metrics.height + bubblePaddingY * 2.0f);
        int lineHeight = max(bubbleHeight + 10, g_lineHeight);
        g_contentHeight += lineHeight;
        pTextLayout->Release();
    }
}

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

// JSON 文字列用エスケープ（インプレース）
inline char _hex(unsigned v)
{
    return (v < 10) ? char('0' + v) : char('A' + (v - 10));
}

void escape_json_inplace(std::string& s)
{
    std::string out;
    out.reserve(s.size() + 16); // 少し余裕を持って確保
    for (size_t i = 0; i < s.size(); ++i) {
        unsigned char c = static_cast<unsigned char>(s[i]);
        switch (c) {
        case '\"': out += "\\\""; break;
        case '\\': out += "\\\\"; break;
        case '\b': out += "\\b";  break;
        case '\f': out += "\\f";  break;
        case '\n': out += "\\n";  break;
        case '\r': out += "\\r";  break;
        case '\t': out += "\\t";  break;
        default:
            if (c < 0x20) {
                // 0x00〜0x1F は \u00XX でエスケープ
                out += "\\u00";
                out += _hex((c >> 4) & 0x0F);
                out += _hex(c & 0x0F);
            }
            else {
                // UTF-8 バイトはそのまま
                out.push_back(static_cast<char>(c));
            }
            break;
        }
    }
    s.swap(out);
}

LPWSTR RunCommand(HWND hWnd)
{
    std::string response;
    DWORD dwSize = 0;
    // WinHTTP セッションを作成
    HINTERNET hSession = WinHttpOpen(L"WinHTTP ChatClient/1.0",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return 0;
    HINTERNET hConnect = WinHttpConnect(hSession, L"api.openai.com",
        INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hConnect) {
        WinHttpCloseHandle(hSession);
        return 0;
    }
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
    std::string body_json = "{\"model\":\"gpt-3.5-turbo\",\"messages\":[";
    for (size_t i = 0; i < g_messages.size(); i++) {
        const ChatMessage& msg = g_messages[i];
        std::string role = msg.isMine ? "user" : "assistant";
        std::string content = w2a(msg.text.c_str());
        escape_json_inplace(content);
        body_json += "{\"role\":\"" + role + "\",\"content\":\"" + content + "\"}";
        if (i + 1 < g_messages.size()) body_json += ",";
    }
    body_json += "]}";
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
    {
        LPCWSTR lpszErrorMessage = L"エラーとなりました。しばらくたってからリトライしてください。";
        LPWSTR lpszReturn = (LPWSTR)GlobalAlloc(GMEM_FIXED, sizeof(WCHAR) * (lstrlenW(lpszErrorMessage) + 1));
        if (lpszReturn) {
            lstrcpy(lpszReturn, lpszErrorMessage);
        }
        return lpszReturn;
    }
}

DWORD WINAPI ThreadFunc(LPVOID p)
{
    PostMessageW((HWND)p, WM_APP, 0, (LPARAM)RunCommand((HWND)p));
    ExitThread(0);
}

// フォントに基づいて行の高さを更新
void UpdateLineHeight(HWND hWnd)
{
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

void UpdateScrollBar(HWND hWnd)
{
    // DirectWrite 計測で正しい内容高さを得る
    RecalcContentHeight(hWnd);
    // ページ（可視領域）高さ
    int visibleHeight = GetVisibleHeight(hWnd);
    if (visibleHeight <= 0) visibleHeight = 1;
    SCROLLINFO si = { sizeof(SCROLLINFO) };
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = max(g_contentHeight - 1, 0);
    si.nPage = (UINT)visibleHeight;
    // 現在位置を最大スクロール位置にクランプ
    int maxPos = max(0, si.nMax - (int)si.nPage + 1);
    if (g_scrollPos > maxPos) g_scrollPos = maxPos;
    if (g_scrollPos < si.nMin) g_scrollPos = si.nMin;
    si.nPos = g_scrollPos;
    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
}

void ScrollToBottom(HWND hWnd)
{
    SCROLLINFO si = { sizeof(SCROLLINFO), SIF_ALL };
    GetScrollInfo(hWnd, SB_VERT, &si);
    // 最大値を取得して一番下に設定
    si.nPos = si.nMax - (int)si.nPage + 1;
    if (si.nPos < si.nMin) si.nPos = si.nMin;
    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
    g_scrollPos = si.nPos;
    InvalidateRect(hWnd, NULL, TRUE);
}

void AddMessage(HWND hWnd, const std::wstring& msg, bool mine)
{
    ChatMessage m{ msg, mine };
    g_messages.push_back(m);
    UpdateScrollBar(hWnd);
    ScrollToBottom(hWnd);
}

void OnPaint(HWND hWnd)
{
    if (!g_pRenderTarget) return;
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);
    const int inputHeight = 32;
    const int bottomMargin = 64; // 入力欄上の余白
    //const int bottomLimit = rcClient.bottom - inputHeight - bottomMargin;
    const int bottomLimit = rcClient.bottom;
    const float bubblePaddingX = 15.0f;
    const float bubblePaddingY = 10.0f;
    g_pRenderTarget->BeginDraw();
    // 背景塗りつぶし
    g_pRenderTarget->Clear(D2D1::ColorF(D2D1::ColorF::White));
    float y = 5.0f - (float)g_scrollPos;
    for (auto& msg : g_messages) {
        // 文字列サイズを計測
        D2D1_SIZE_F layoutSize = D2D1::SizeF((rcClient.right / 2.0f), FLT_MAX);
        IDWriteTextLayout* pTextLayout = nullptr;
        HRESULT hr = g_pDWriteFactory->CreateTextLayout(
            msg.text.c_str(),
            (UINT32)msg.text.length(),
            g_pTextFormat,
            layoutSize.width,
            layoutSize.height,
            &pTextLayout
        );
        if (FAILED(hr)) continue;
        DWRITE_TEXT_METRICS metrics;
        pTextLayout->GetMetrics(&metrics);
        float bubbleWidth = metrics.width + bubblePaddingX * 2.0f;
        float bubbleHeight = metrics.height + bubblePaddingY * 2.0f;
        float lineHeight = (float)max((int)(bubbleHeight + 10.0f), g_lineHeight);
        if (y > bottomLimit) {
            pTextLayout->Release();
            break; // 入力欄の上限まで
        }
        // 吹き出し矩形
        if (msg.isMine) {
            msg.rect = {
                rcClient.right - (LONG)bubbleWidth - 10,
                (LONG)y,
                rcClient.right - 10,
                (LONG)(y + bubbleHeight)
            };
        }
        else {
            msg.rect = {
                10,
                (LONG)y,
                10 + (LONG)bubbleWidth,
                (LONG)(y + bubbleHeight)
            };
        }
        // 吹き出し背景
        D2D1_ROUNDED_RECT roundedRect = D2D1::RoundedRect(
            D2D1::RectF(
                (FLOAT)msg.rect.left,
                (FLOAT)msg.rect.top,
                (FLOAT)msg.rect.right,
                (FLOAT)msg.rect.bottom
            ),
            20.0f, 20.0f
        );
        g_pRenderTarget->FillRoundedRectangle(
            &roundedRect,
            msg.isMine ? g_pBrushMine : g_pBrushOther
        );
        // テキスト描画領域
        D2D1_POINT_2F origin = {
            (FLOAT)msg.rect.left + bubblePaddingX,
            (FLOAT)msg.rect.top + bubblePaddingY
        };
        g_pRenderTarget->DrawTextLayout(
            origin,
            pTextLayout,
            g_pBrushText,
            D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT
        );
        pTextLayout->Release();
        y += lineHeight;
    }
    {
        D2D1_ROUNDED_RECT roundedRect = D2D1::RoundedRect(
            D2D1::RectF(
                32.0f,
                (FLOAT)(rcClient.bottom - 64 - 32),
                (FLOAT)(rcClient.right - 32),
                (FLOAT)(rcClient.bottom - 32)
            ),
            32.0f, // 角丸半径X
            32.0f  // 角丸半径Y
        );
        // 塗りつぶし
        g_pRenderTarget->FillRoundedRectangle(
            &roundedRect,
            g_pBrushOther // 事前に作ってある RGB(250,250,250) 相当のブラシ
        );
        // 枠線を描きたいなら DrawRoundedRectangle を追加
        ID2D1SolidColorBrush* pStrokeBrush = nullptr;
        g_pRenderTarget->CreateSolidColorBrush(
            D2D1::ColorF(D2D1::ColorF::Black), &pStrokeBrush);
        if (pStrokeBrush) {
            g_pRenderTarget->DrawRoundedRectangle(
                &roundedRect, pStrokeBrush, 2.0f // 線幅1
            );
            pStrokeBrush->Release();
        }
    }
    HRESULT hr = g_pRenderTarget->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET) {
        DiscardDeviceResources();
        InvalidateRect(hWnd, nullptr, FALSE); // 再描画要求
    }
}

void OnVScroll(HWND hWnd, WPARAM wParam)
{
    SCROLLINFO si = { sizeof(SCROLLINFO), SIF_ALL };
    GetScrollInfo(hWnd, SB_VERT, &si);
    int pos = si.nPos;
    switch (LOWORD(wParam)) {
    case SB_LINEUP:     pos -= 20; break;
    case SB_LINEDOWN:   pos += 20; break;
    case SB_PAGEUP:     pos -= (int)si.nPage; break;
    case SB_PAGEDOWN:   pos += (int)si.nPage; break;
    case SB_THUMBTRACK: pos = si.nTrackPos; break;
    default: break;
    }
    int maxPos = max(0, (int)si.nMax - (int)si.nPage + 1);
    if (pos < (int)si.nMin) pos = (int)si.nMin;
    if (pos > maxPos)        pos = maxPos;
    si.fMask = SIF_POS;
    si.nPos = pos;
    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
    g_scrollPos = pos;
    InvalidateRect(hWnd, NULL, TRUE);
}

void CopyToClipboard(HWND hWnd, const std::wstring& text)
{
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

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static HWND hEdit;
    static HBRUSH hBrushOther;
    static HANDLE hThread;
    switch (msg) {
    case WM_CREATE:
		InitD2D(hWnd);
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
            LPWSTR text = (LPWSTR)GlobalAlloc(0, (len + 1) * sizeof(WCHAR));
            if (text) {
                GetWindowText(hEdit, text, len + 1);
                AddMessage(hWnd, text, true);
				GlobalFree(text);

                SetWindowText(hEdit, 0);
                EnableWindow(hEdit, FALSE);

                DWORD dwParam;
                hThread = CreateThread(0, 0, ThreadFunc, (LPVOID)hWnd, 0, &dwParam);
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
    case WM_ERASEBKGND:
        return 1;
    case WM_PAINT: {
        if (!g_pRenderTarget) {
            InitD2D(hWnd);
        }
        {
            RECT rcUpdate;
            if (GetUpdateRect(hWnd, &rcUpdate, FALSE)) {
                OnPaint(hWnd);
            }
            ValidateRect(hWnd, NULL);
        }
    } return 0;
    case WM_VSCROLL:
        OnVScroll(hWnd, wParam);
        return 0;
    case WM_MOUSEWHEEL: {
        SCROLLINFO si = { sizeof(SCROLLINFO), SIF_ALL };
        GetScrollInfo(hWnd, SB_VERT, &si);
        int delta = -(GET_WHEEL_DELTA_WPARAM(wParam) / WHEEL_DELTA) * 40; // 1ノッチ40px
        int pos = (int)si.nPos + delta;
        int maxPos = max(0, (int)si.nMax - (int)si.nPage + 1);
        if (pos < (int)si.nMin) pos = (int)si.nMin;
        if (pos > maxPos)        pos = maxPos;
        si.fMask = SIF_POS;
        si.nPos = pos;
        SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
        g_scrollPos = pos;
        InvalidateRect(hWnd, NULL, TRUE);
        return 0;
    }
    case WM_SIZE:
        if (g_pRenderTarget) {
            g_pRenderTarget->Resize(D2D1::SizeU(LOWORD(lParam), HIWORD(lParam)));
        }
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
	case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case ID_NEW:
			g_messages.clear();
			g_scrollPos = 0;
			UpdateScrollBar(hWnd);             
            break;
		case ID_HOME:
			g_scrollPos = 0;
			UpdateScrollBar(hWnd);
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		case ID_END:
			ScrollToBottom(hWnd);
            break;
        case ID_PRIOR:
            {
                SCROLLINFO si = { sizeof(SCROLLINFO), SIF_ALL };
                GetScrollInfo(hWnd, SB_VERT, &si);
                int pos = si.nPos - (int)si.nPage;
                if (pos < (int)si.nMin) pos = (int)si.nMin;
                if (pos > (int)si.nMax - (int)si.nPage + 1) pos = (int)si.nMax - (int)si.nPage + 1;
                si.fMask = SIF_POS;
                si.nPos = pos;
                SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
                g_scrollPos = pos;
                InvalidateRect(hWnd, NULL, TRUE);
                break;
            }
		case ID_NEXT:
		    {
			    SCROLLINFO si = { sizeof(SCROLLINFO), SIF_ALL };
			    GetScrollInfo(hWnd, SB_VERT, &si);
			    int pos = si.nPos + (int)si.nPage;
			    if (pos < (int)si.nMin) pos = (int)si.nMin;
			    if (pos > (int)si.nMax - (int)si.nPage + 1) pos = (int)si.nMax - (int)si.nPage + 1;
			    si.fMask = SIF_POS;
			    si.nPos = pos;
			    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
			    g_scrollPos = pos;
			    InvalidateRect(hWnd, NULL, TRUE);
			    break;
		    }
        }
        break;
    case WM_DESTROY:
        CleanupD2D();
        DeleteObject(hBrushOther);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hWnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int)
{
    INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icex);
    MSG msg;
    WNDCLASS wc = {
        CS_HREDRAW | CS_VREDRAW,
        WndProc,0,0,hInstance,
        LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)),
        LoadCursor(0,IDC_ARROW),
        0,0,L"chat"
    };
    RegisterClass(&wc);
    HWND hWnd = CreateWindow(
        L"chat", L"chat",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_VSCROLL,
        CW_USEDEFAULT, 0, CW_USEDEFAULT, 0,
        0, 0, hInstance, 0);
    ShowWindow(hWnd, SW_SHOWDEFAULT);
    UpdateWindow(hWnd);
    HACCEL hAccel = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));
    while (GetMessage(&msg, 0, 0, 0)) {
        if (!TranslateAccelerator(hWnd, hAccel, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    return (int)msg.wParam;
}
