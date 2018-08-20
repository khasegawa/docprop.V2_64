#include "stdafx.h"
#include "olefunc.h"

#pragma data_seg(".sharedata")
//���L�̈�̃f�[�^�͏������K�{
HHOOK hHookMouse = 0;
#pragma data_seg()

#define MARGIN 9   // �v���p�e�B�\���̍�����я�}�[�W��(�s�N�Z��)
#define ROUNDR 11  // �E�B���h�E�̊p�̊ۂ�(�s�N�Z��)
#define FONTPOINT 12
#define FONTNAME L"�l�r �S�V�b�N"

static HINSTANCE g_hDll = 0;
static HWND g_hWndDispInfo = 0;
static WCHAR gOutText[1024]; //�S�v���p�e�B�̕\��������ő咷

LRESULT CALLBACK CallMouseProc(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);
static HFONT getFont(void);

//�t�b�N��g�ݍ���
__declspec(dllexport) void CALLBACK setHook(void)
{
	hHookMouse = SetWindowsHookEx(WH_MOUSE, CallMouseProc, g_hDll, 0);
}

//�t�b�N����������
__declspec(dllexport) void CALLBACK freeHook(void)
{
	UnhookWindowsHookEx(hHookMouse);
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	UNREFERENCED_PARAMETER(lpReserved);

	switch (ul_reason_for_call) {
	case DLL_PROCESS_ATTACH:
		g_hDll = hModule;     //DLL�̃n���h����ۑ�����
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
		ShowWindow(g_hWndDispInfo, SW_HIDE);  // �}�E�X�u���ړ��ŁA���̃E�B���h�E�ɓ����Ă��܂����ꍇ�̑[�u
//		DestroyWindow(hWndDispInfo);
		break;
	case WM_PAINT:
		{
		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(g_hWndDispInfo, &ps);

		// �`�摮���ݒ�
		long textColorOld = SetTextColor(hdc, RGB(0, 0x40, 0x80));
		int bkModeOld = SetBkMode(hdc, TRANSPARENT);

		HFONT hFont = getFont();
		HFONT hFontOld = (HFONT)SelectObject(hdc, hFont);

		// �h�L�������g�v���p�e�B�\��
		RECT rect = {MARGIN, MARGIN, 0, 0};  //NOCLIP�Ȃ̂�rect�̑傫���͊֌W�Ȃ�
		DrawText(hdc, gOutText, wcslen(gOutText), &rect, DT_LEFT | DT_NOCLIP);

		//�E�C���h�E�̋��E��\��
		RECT winrect;
		GetWindowRect(g_hWndDispInfo, &winrect);
		int width = winrect.right - winrect.left - 1;
		int height = winrect.bottom - winrect.top - 1;
		HRGN hRGN = CreateRoundRectRgn(0, 0, width, height, ROUNDR, ROUNDR);
		FrameRgn(hdc, hRGN, CreateSolidBrush(RGB(0xE0, 0xE0, 0xE0)), 3, 3);
		FrameRgn(hdc, hRGN, CreateSolidBrush(RGB(0x80, 0x80, 0x80)), 2, 2);
		DeleteObject(hRGN);

		// �`�摮�����A
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
		wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;   // �X�^�C��
		wc.lpfnWndProc = WndProc;
		wc.cbClsExtra = 0;                    // �g�����P
		wc.cbWndExtra = 0;                    // �g�����Q
		wc.hInstance = g_hDll;
		wc.hIcon = 0;
		wc.hIconSm = wc.hIcon;                // �q�A�C�R��
		wc.hCursor = (HCURSOR)LoadImage(
			NULL, MAKEINTRESOURCE(IDC_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE | LR_SHARED);
		wc.hbrBackground = CreateSolidBrush(RGB(0xFF, 0xFF, 0xF0));
		wc.lpszMenuName = NULL;               // ���j���[��
		wc.lpszClassName = L"DispInfo";       // �E�B���h�E�N���X��
		
		// �E�B���h�E�N���X��o�^����
		if(RegisterClassEx(&wc) == 0) {
			return;
		}

		// �E�B���h�E���쐬����
		g_hWndDispInfo = CreateWindowEx(WS_EX_TOOLWINDOW, wc.lpszClassName, L"", WS_POPUPWINDOW,
									0, 0, 0, 0, NULL, NULL, g_hDll, NULL);
		if (! g_hWndDispInfo) {
			return;
		}
	}

	// �����l(������)�擾
	if (! getFileProperty(gOutText, _countof(gOutText), doctype)) {
		wcsncpy_s(gOutText, L"�h�L�������g�̑������擾�ł��Ȃ�", _TRUNCATE);
	}

	PAINTSTRUCT ps;
	HDC hdc = BeginPaint(g_hWndDispInfo, &ps);

	// �`�摮���ݒ�
	HFONT hFont = getFont();
	HFONT hFontOld = (HFONT)SelectObject(hdc, hFont);

	// �\���E�B���h�E�̑傫���ƈʒu�Đݒ�
	RECT rect = {0, 0, 0, 0};
	DrawText(hdc, gOutText, wcslen(gOutText), &rect, DT_CALCRECT);  //�`��͂��Ȃ� �`��̈�̑傫���𒲂ׂ�
	int width = rect.right - rect.left + MARGIN + MARGIN;
	int height = rect.bottom - rect.top + MARGIN + MARGIN;
	SetWindowPos(g_hWndDispInfo, HWND_TOP, x, y, width, height, 0);
	//�E�C���h�E���ۂ�����
	HRGN hRGN = CreateRoundRectRgn(1, 1, width, height, ROUNDR, ROUNDR);
	SetWindowRgn(g_hWndDispInfo, hRGN, TRUE);
	DeleteObject(hRGN);

	// �`�摮�����A
	SelectObject(hdc, hFontOld);
	DeleteObject(hFont);

	EndPaint(g_hWndDispInfo, &ps);

	// �\��
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

	if (nCode == HC_NOREMOVE                                   // �������ׂ��ł���
	 && GetAncestor(hwnd, GA_PARENT) == GetDesktopWindow()     // ���g�b�v���x���̃E�B���h�E�ł���
	 && hwnd == GetForegroundWindow()                          // ���t�H�[�J�X�̂���E�B���h�E�ł���
	 && ! (IsWindow(g_hWndDispInfo) && hwnd == g_hWndDispInfo) // ���v���p�e�B�\���E�B���h�E���g�łȂ�
	 && (testcode != HTCLIENT && testcode != HTCAPTION)        // ���N���C�A���g�̈�̊O�ŁA���^�C�g���̈�łȂ�
	) {
		POINT pt = ((MOUSEHOOKSTRUCT *)lParam)->pt;
		RECT rect;
		GetWindowRect(hwnd, &rect);

		if (abs(pt.x - rect.left - 12) < 7
		 && abs(pt.y - rect.top - 12) < 7  // �E�B���h�E����̏������A�C�R���̗̈���ł���
		) {
			WCHAR progname[15]; // "powerpnt.exe"�Ɣ�r�ł���\���Ȓ���
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
				showDispInfo(rect.left + 30, rect.top + 7, doctype);  //�E�B���h�E���ォ��A�E30��7�s�N�Z���ɕ\��
			}
		} else if (IsWindow(g_hWndDispInfo)) {
			ShowWindow(g_hWndDispInfo, SW_HIDE);
		}
	}

	return CallNextHookEx(hHookMouse, nCode, wParam, lParam); //���̃t�b�N���Ă�
}
