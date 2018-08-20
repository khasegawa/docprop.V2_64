// docprop.cpp : アプリケーションのエントリ ポイントを定義
//
#include "stdafx.h"
#include <stdlib.h>
#include <crtdbg.h>
#include "docprop.h"

extern void CALLBACK setHook(void);
extern void CALLBACK freeHook(void);

static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);
static BOOL CALLBACK AboutDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
static ATOM MyRegisterClass(HINSTANCE hInstance);

#define MAX_LOADSTRING 100
static WCHAR *gProgPath;                    // このプログラムファイルのフルパス
static HINSTANCE g_hInst;                   // 現在のインターフェイス
static TCHAR gTitle[MAX_LOADSTRING];        // タイトル バーのテキスト
static TCHAR gWindowClass[MAX_LOADSTRING];  // メイン ウィンドウ クラス名

int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
	HWND hWnd;

	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

    _CrtSetReportMode(_CRT_ASSERT, 0);

	gProgPath = __wargv[0];

	// グローバル文字列初期化
	LoadString(hInstance, IDS_APP_TITLE, gTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_DOCPROP, gWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// 動かすインスタンスは一つだけ
	hWnd = FindWindow(gWindowClass, NULL);
	if (hWnd) {
		MessageBox(NULL, L"このプログラムは既に動いています。\nタスクトレイを確認してください。", gTitle, MB_OK);
		SetForegroundWindow(hWnd);
		return 0;
	}

	g_hInst = hInstance; // グローバル変数にインスタンス処理を格納

	// アプリケーション初期化
	hWnd = CreateWindowW(gWindowClass, gTitle, WS_OVERLAPPEDWINDOW,
				CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
				NULL, NULL, hInstance, NULL);
	if (! hWnd) {
		return -1;
	}

#if 0
	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);
#endif

	// ショートカットキー読み込み
	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_DOCPROP));

	// フック開始
	setHook();
	// メイン メッセージ ループ:
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		if (! TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	// フック終了
	freeHook();

	return (int)msg.wParam;
}

//
//  メイン ウィンドウのメッセージ
//
static LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	static HMENU hMenu;
	static NOTIFYICONDATA nid;

	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message) {
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// 選択されたメニューの解析:
		switch (wmId) {
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;

		case IDD_ABOUT:{
			EnableMenuItem(hMenu, 0, MF_BYPOSITION | MF_GRAYED);  // Aboutダイアログメニューを重ねて呼べないようにする
			DialogBox(g_hInst, MAKEINTRESOURCE(IDD_ABOUT), hWnd, AboutDlgProc);
			EnableMenuItem(hMenu, 0, MF_BYPOSITION | MF_ENABLED); // Aboutダイアログメニューを元に戻す
			break;}

		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;

	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: 描画コードをここに追加してください...
		EndPaint(hWnd, &ps);
		break;

	case WM_TRAYICONMESSAGE:
		switch(lParam) {
		case WM_RBUTTONDOWN: // タスクトレイアイコン上で右クリック
			POINT pt;
			GetCursorPos(&pt);
			SetForegroundWindow(hWnd);  // これがないと、ポップアップメニューが消えなくなる
			TrackPopupMenuEx(hMenu, TPM_LEFTALIGN, pt.x, pt.y, hWnd, 0);
			break;
		}
		break;

	case WM_CREATE:
		hMenu = CreatePopupMenu();
		MENUITEMINFO menuiteminfo[2];
		menuiteminfo[0].cbSize = sizeof(*menuiteminfo);
		menuiteminfo[0].fMask = MIIM_STRING | MIIM_ID;
		menuiteminfo[0].wID = IDD_ABOUT;
		menuiteminfo[0].dwTypeData = L"このソフトウェアについて";
		menuiteminfo[0].cch = wcslen(menuiteminfo[0].dwTypeData);
		InsertMenuItem(hMenu, 0, true, &menuiteminfo[0]);
		menuiteminfo[1].cbSize = sizeof(*menuiteminfo);
		menuiteminfo[1].fMask = MIIM_STRING | MIIM_ID;
		menuiteminfo[1].wID = IDM_EXIT;
		menuiteminfo[1].dwTypeData = L"終了";
		menuiteminfo[1].cch = wcslen(menuiteminfo[1].dwTypeData);
		InsertMenuItem(hMenu, 1, true, &menuiteminfo[1]);

		nid.cbSize = sizeof(NOTIFYICONDATA);
		nid.hWnd = hWnd;
		nid.uID = 0;
		nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
		nid.uCallbackMessage = WM_TRAYICONMESSAGE;
		nid.hIcon = LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_SMALL));
		wcsncpy_s(nid.szTip, gTitle, _TRUNCATE);
		Shell_NotifyIcon(NIM_ADD, &nid);
		break;

	case WM_DESTROY:
		DestroyMenu(hMenu);
		Shell_NotifyIcon(NIM_DELETE, &nid);
		PostQuitMessage(0);
		break;

	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}

	return 0;
}

//
//  About Dialog のメッセージ処理
//
static BOOL CALLBACK AboutDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg) {

	case WM_INITDIALOG:
		{
			int bufSize = GetFileVersionInfoSize(gProgPath, 0);
			if (bufSize > 0) {
				void *buf = (void *)malloc(sizeof (BYTE) * bufSize);
				if (buf != NULL) {
					GetFileVersionInfo(gProgPath, 0, bufSize, buf);

					void *str;
					unsigned int strSize;

					VerQueryValue(buf, L"\\StringFileInfo\\041103a4\\ProductName", (LPVOID *)&str, &strSize);
					SetWindowText(hwndDlg, (WCHAR *)str);
					SetWindowText(GetDlgItem(hwndDlg, IDC_EDIT2), (WCHAR *)str);

					VerQueryValue(buf, L"\\StringFileInfo\\041103a4\\FileVersion", (LPVOID *)&str, &strSize);
					SetWindowText(GetDlgItem(hwndDlg, IDC_EDIT1), (WCHAR *)str);

					VerQueryValue(buf, L"\\StringFileInfo\\041103a4\\LegalCopyright", (LPVOID *)&str, &strSize);
					SetWindowText(GetDlgItem(hwndDlg, IDC_EDIT3), (WCHAR *)str);

					free(buf);
				}
			}
		}
		return TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK) {
			EndDialog(hwndDlg, TRUE);
			return TRUE;
		}
		break;

	}
	return FALSE;
}

//
//  ウィンドウ クラスを登録
//
//    この関数および使い方は、'RegisterClassEx' 関数が追加された
//    Windows 95 より前の Win32 システムと互換させる場合にのみ必要です。
//    アプリケーションが、関連付けられた
//    正しい形式の小さいアイコンを取得できるようにするには、
//    この関数を呼び出してください。
//
static ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_LARGE));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW + 1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDR_DOCPROP);
	wcex.lpszClassName	= gWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}
