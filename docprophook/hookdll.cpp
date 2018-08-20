#include "stdafx.h"
#include "olefunc.h"

#pragma data_seg(".sharedata")
//共有領域のデータは初期化必須
HHOOK hHookMouse = 0;
#pragma data_seg()

#define MARGIN 9   // プロパティ表示の左および上マージン(ピクセル)
#define ROUNDR 11  // ウィンドウの角の丸み(ピクセル)
#define FONTPOINT 12
#define FONTNAME L"ＭＳ ゴシック"

static HINSTANCE g_hDll = 0;
static HWND g_hWndDispInfo = 0;
static WCHAR gOutText[1024]; //全プロパティの表示文字列最大長

LRESULT CALLBACK CallMouseProc(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);
static HFONT getFont(void);

//フックを組み込む
__declspec(dllexport) void CALLBACK setHook(void)
{
	hHookMouse = SetWindowsHookEx(WH_MOUSE, CallMouseProc, g_hDll, 0);
}

//フックを解除する
__declspec(dllexport) void CALLBACK freeHook(void)
{
	UnhookWindowsHookEx(hHookMouse);
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	UNREFERENCED_PARAMETER(lpReserved);

	switch (ul_reason_for_call) {
	case DLL_PROCESS_ATTACH:
		g_hDll = hModule;     //DLLのハンドルを保存する
		break;

	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}

	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message) {
    case WM_CREATE:
		break;
	case WM_COMMAND:
		break;
	case WM_MOUSEMOVE:
		ShowWindow(g_hWndDispInfo, SW_HIDE);  // マウス瞬時移動で、このウィンドウに入ってしまった場合の措置
//		DestroyWindow(hWndDispInfo);
		break;
	case WM_PAINT:
		{
		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(g_hWndDispInfo, &ps);

		// 描画属性設定
		long textColorOld = SetTextColor(hdc, RGB(0, 0x40, 0x80));
		int bkModeOld = SetBkMode(hdc, TRANSPARENT);

		HFONT hFont = getFont();
		HFONT hFontOld = (HFONT)SelectObject(hdc, hFont);

		// ドキュメントプロパティ表示
		RECT rect = {MARGIN, MARGIN, 0, 0};  //NOCLIPなのでrectの大きさは関係ない
		DrawText(hdc, gOutText, wcslen(gOutText), &rect, DT_LEFT | DT_NOCLIP);

		//ウインドウの境界を表示
		RECT winrect;
		GetWindowRect(g_hWndDispInfo, &winrect);
		int width = winrect.right - winrect.left - 1;
		int height = winrect.bottom - winrect.top - 1;
		HRGN hRGN = CreateRoundRectRgn(0, 0, width, height, ROUNDR, ROUNDR);
		FrameRgn(hdc, hRGN, CreateSolidBrush(RGB(0xE0, 0xE0, 0xE0)), 3, 3);
		FrameRgn(hdc, hRGN, CreateSolidBrush(RGB(0x80, 0x80, 0x80)), 2, 2);
		DeleteObject(hRGN);

		// 描画属性復帰
		SelectObject(hdc, hFontOld);
		DeleteObject(hFont);

		SetBkMode(hdc, bkModeOld);
		SetTextColor(hdc, textColorOld);

		EndPaint(g_hWndDispInfo, &ps);
		}
		break;
	case WM_DESTROY:
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

void showDispInfo(int x, int y, int doctype)
{
	if (! IsWindow(g_hWndDispInfo)) {
		WNDCLASSEX wc;
		wc.cbSize = sizeof(wc);
		wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;   // スタイル
		wc.lpfnWndProc = WndProc;
		wc.cbClsExtra = 0;                    // 拡張情報１
		wc.cbWndExtra = 0;                    // 拡張情報２
		wc.hInstance = g_hDll;
		wc.hIcon = 0;
		wc.hIconSm = wc.hIcon;                // 子アイコン
		wc.hCursor = (HCURSOR)LoadImage(
			NULL, MAKEINTRESOURCE(IDC_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
		wc.hbrBackground = CreateSolidBrush(RGB(0xFF, 0xFF, 0xF0));
		wc.lpszMenuName = NULL;               // メニュー名
		wc.lpszClassName = L"DispInfo";       // ウィンドウクラス名
		
		// ウィンドウクラスを登録する
		if(RegisterClassEx(&wc) == 0) {
			return;
		}

		// ウィンドウを作成する
		g_hWndDispInfo = CreateWindowEx(WS_EX_TOOLWINDOW, wc.lpszClassName, L"", WS_POPUPWINDOW,
									0, 0, 0, 0, NULL, NULL, g_hDll, NULL);
		if (! g_hWndDispInfo) {
			return;
		}
	}

	// 属性値(文字列)取得
	if (! getFileProperty(gOutText, _countof(gOutText), doctype)) {
		wcsncpy_s(gOutText, L"ドキュメントの属性が取得できない", _TRUNCATE);
	}

	PAINTSTRUCT ps;
	HDC hdc = BeginPaint(g_hWndDispInfo, &ps);

	// 描画属性設定
	HFONT hFont = getFont();
	HFONT hFontOld = (HFONT)SelectObject(hdc, hFont);

	// 表示ウィンドウの大きさと位置再設定
	RECT rect = {0, 0, 0, 0};
	DrawText(hdc, gOutText, wcslen(gOutText), &rect, DT_CALCRECT);  //描画はしない 描画領域の大きさを調べる
	int width = rect.right - rect.left + MARGIN + MARGIN;
	int height = rect.bottom - rect.top + MARGIN + MARGIN;
	SetWindowPos(g_hWndDispInfo, HWND_TOP, x, y, width, height, 0);
	//ウインドウを丸くする
	HRGN hRGN = CreateRoundRectRgn(1, 1, width, height, ROUNDR, ROUNDR);
	SetWindowRgn(g_hWndDispInfo, hRGN, TRUE);
	DeleteObject(hRGN);

	// 描画属性復帰
	SelectObject(hdc, hFontOld);
	DeleteObject(hFont);

	EndPaint(g_hWndDispInfo, &ps);

	// 表示
	ShowWindow(g_hWndDispInfo, SW_SHOW);
	UpdateWindow(g_hWndDispInfo);
}

static HFONT getFont(void)
{
	return CreateFont(FONTPOINT, 0, 0, 0, FW_REGULAR, FALSE, FALSE, FALSE,
				SHIFTJIS_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY,
				FIXED_PITCH | FF_MODERN, FONTNAME);
}

static WCHAR *getFilename(HWND hwnd, WCHAR *filename, long filenamesize)
{
	WCHAR path[_MAX_PATH];
	WCHAR *s;

	GetWindowModuleFileName(hwnd, path, _countof(path));

	if ((s = wcsrchr(path, '\\')) == NULL) {
		s = path;
	} else {
		s++;
	}
	wcsncpy_s(filename, filenamesize, s, _TRUNCATE);

	return filename;
}

LRESULT CALLBACK CallMouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	HWND hwnd = ((MOUSEHOOKSTRUCT *)lParam)->hwnd;
	DWORD testcode = ((MOUSEHOOKSTRUCT *)lParam)->wHitTestCode;

	if (nCode == HC_NOREMOVE                                   // 処理すべきであり
	 && GetAncestor(hwnd, GA_PARENT) == GetDesktopWindow()     // かつトップレベルのウィンドウであり
	 && hwnd == GetForegroundWindow()                          // かつフォーカスのあるウィンドウであり
	 && ! (IsWindow(g_hWndDispInfo) && hwnd == g_hWndDispInfo) // かつプロパティ表示ウィンドウ自身でない
	 && (testcode != HTCLIENT && testcode != HTCAPTION)        // かつクライアント領域の外で、かつタイトル領域でない
	) {
		POINT pt = ((MOUSEHOOKSTRUCT *)lParam)->pt;
		RECT rect;
		GetWindowRect(hwnd, &rect);

		if (abs(pt.x - rect.left - 12) < 7
		 && abs(pt.y - rect.top - 12) < 7  // ウィンドウ左上の小さいアイコンの領域内である
		) {
			WCHAR progname[15]; // "powerpnt.exe"と比較できる十分な長さ
			getFilename(hwnd, progname, _countof(progname));
			int doctype;
			if (_wcsicmp(progname, L"excel.exe") == 0) {
				doctype = MS_EXCEL;
			} else if (_wcsicmp(progname, L"winword.exe") == 0) {
				doctype = MS_WORD;
			} else if (_wcsicmp(progname, L"wwlib.dll") == 0) {
				doctype = MS_WORD;
			} else if (_wcsicmp(progname, L"powerpnt.exe") == 0) {
				doctype = MS_PPT;
			} else if (_wcsicmp(progname, L"ppcore.dll") == 0) {
				doctype = MS_PPT;
			} else {
				doctype = OTHER;
			}

			if (doctype != OTHER) {
				showDispInfo(rect.left + 30, rect.top + 7, doctype);  //ウィンドウ左上から、右30下7ピクセルに表示
			}
		} else if (IsWindow(g_hWndDispInfo)) {
			ShowWindow(g_hWndDispInfo, SW_HIDE);
		}
	}

	return CallNextHookEx(hHookMouse, nCode, wParam, lParam); //次のフックを呼ぶ
}
