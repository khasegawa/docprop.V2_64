// docprop.cpp : �A�v���P�[�V�����̃G���g�� �|�C���g���`
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
static WCHAR *gProgPath;                    // ���̃v���O�����t�@�C���̃t���p�X
static HINSTANCE g_hInst;                   // ���݂̃C���^�[�t�F�C�X
static TCHAR gTitle[MAX_LOADSTRING];        // �^�C�g�� �o�[�̃e�L�X�g
static TCHAR gWindowClass[MAX_LOADSTRING];  // ���C�� �E�B���h�E �N���X��

int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
	HWND hWnd;

	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

    _CrtSetReportMode(_CRT_ASSERT, 0);

	gProgPath = __wargv[0];

	// �O���[�o�������񏉊���
	LoadString(hInstance, IDS_APP_TITLE, gTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_DOCPROP, gWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// �������C���X�^���X�͈����
	hWnd = FindWindow(gWindowClass, NULL);
	if (hWnd) {
		MessageBox(NULL, L"���̃v���O�����͊��ɓ����Ă��܂��B\n�^�X�N�g���C���m�F���Ă��������B", gTitle, MB_OK);
		SetForegroundWindow(hWnd);
		return 0;
	}

	g_hInst = hInstance; // �O���[�o���ϐ��ɃC���X�^���X�������i�[

	// �A�v���P�[�V����������
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

	// �V���[�g�J�b�g�L�[�ǂݍ���
	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_DOCPROP));

	// �t�b�N�J�n
	setHook();
	// ���C�� ���b�Z�[�W ���[�v:
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		if (! TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	// �t�b�N�I��
	freeHook();

	return (int)msg.wParam;
}

//
//  ���C�� �E�B���h�E�̃��b�Z�[�W
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
		// �I�����ꂽ���j���[�̉��:
		switch (wmId) {
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;

		case IDD_ABOUT:{
			EnableMenuItem(hMenu, 0, MF_BYPOSITION | MF_GRAYED);  // About�_�C�A���O���j���[���d�˂ČĂׂȂ��悤�ɂ���
			DialogBox(g_hInst, MAKEINTRESOURCE(IDD_ABOUT), hWnd, AboutDlgProc);
			EnableMenuItem(hMenu, 0, MF_BYPOSITION | MF_ENABLED); // About�_�C�A���O���j���[�����ɖ߂�
			break;}

		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;

	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: �`��R�[�h�������ɒǉ����Ă�������...
		EndPaint(hWnd, &ps);
		break;

	case WM_TRAYICONMESSAGE:
		switch(lParam) {
		case WM_RBUTTONDOWN: // �^�X�N�g���C�A�C�R����ŉE�N���b�N
			POINT pt;
			GetCursorPos(&pt);
			SetForegroundWindow(hWnd);  // ���ꂪ�Ȃ��ƁA�|�b�v�A�b�v���j���[�������Ȃ��Ȃ�
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
		menuiteminfo[0].dwTypeData = L"���̃\�t�g�E�F�A�ɂ���";
		menuiteminfo[0].cch = wcslen(menuiteminfo[0].dwTypeData);
		InsertMenuItem(hMenu, 0, true, &menuiteminfo[0]);
		menuiteminfo[1].cbSize = sizeof(*menuiteminfo);
		menuiteminfo[1].fMask = MIIM_STRING | MIIM_ID;
		menuiteminfo[1].wID = IDM_EXIT;
		menuiteminfo[1].dwTypeData = L"�I��";
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
//  About Dialog �̃��b�Z�[�W����
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
//  �E�B���h�E �N���X��o�^
//
//    ���̊֐�����юg�����́A'RegisterClassEx' �֐����ǉ����ꂽ
//    Windows 95 ���O�� Win32 �V�X�e���ƌ݊�������ꍇ�ɂ̂ݕK�v�ł��B
//    �A�v���P�[�V�������A�֘A�t����ꂽ
//    �������`���̏������A�C�R�����擾�ł���悤�ɂ���ɂ́A
//    ���̊֐����Ăяo���Ă��������B
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
