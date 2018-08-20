#include "stdafx.h"
#include <map>
#include <string>
#include "invoke.h"
#include "varstr.h"

static void myInvalidParameterHandler(
	const wchar_t *expression,
	const wchar_t *function, 
	const wchar_t *file, 
	unsigned int line, 
	uintptr_t pReserved)
{
	UNREFERENCED_PARAMETER(expression);
	UNREFERENCED_PARAMETER(function);
	UNREFERENCED_PARAMETER(file);
	UNREFERENCED_PARAMETER(line);
	UNREFERENCED_PARAMETER(pReserved);
	//�������Ȃ�
}

static BOOL _getFileProperty(WCHAR *out, int outsize, WCHAR *progID, WCHAR *fileID)
{
	HRESULT hr;
	CLSID clsid;

	hr = CLSIDFromProgID(progID, &clsid);
	if(FAILED(hr)) {
		return false;
	}

	IUnknown *pUnk = NULL;
	hr = GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);
	if (FAILED(hr) || pUnk == NULL) {
		return false;
	}

	IDispatch *pDispatch = NULL;
	hr = pUnk->QueryInterface(IID_IDispatch, (void **)&pDispatch);
	pUnk->Release();
	if (FAILED(hr) || pDispatch == NULL) {
		return false;
	}

	VARIANT varResult;
	VariantInit(&varResult);

	hr = getProperty(pDispatch, &varResult, fileID);
	IDispatch *pDispatch2 = varResult.pdispVal;
	if (FAILED(hr) || pDispatch2 == NULL) {
		pDispatch->Release();
		return false;
	}

	hr = getProperty(pDispatch2, &varResult, L"Name");
	if(FAILED(hr)) {
		pDispatch->Release();
		return false;
	}
	if(varResult.vt != VT_BSTR || varResult.bstrVal == NULL) {
		pDispatch2->Release();
		pDispatch->Release();
		return false;
	}
	wcsncpy_s(out, outsize, L"���O            : ", _TRUNCATE);
	varcat(out, outsize, varResult);

	hr = getProperty(pDispatch2, &varResult, L"Path");
	if(FAILED(hr)) {
		pDispatch->Release();
		return false;
	}
	if(varResult.vt != VT_BSTR || varResult.bstrVal == NULL) {
		pDispatch2->Release();
		pDispatch->Release();
		return false;
	}
	wcsncat_s(out, outsize, L"\n�ꏊ            : ", _TRUNCATE);
	varcat(out, outsize, varResult);

	hr = getProperty(pDispatch2, &varResult, L"BuiltInDocumentProperties");
	IDispatch *pDispatch3 = varResult.pdispVal;
	if (FAILED(hr) || pDispatch3 == NULL) {
		pDispatch2->Release();
		pDispatch->Release();
		return false;
	}

	hr = getProperty(pDispatch3, &varResult, L"Count");
	if (FAILED(hr) || varResult.vt != VT_I4) {  // �ǂ̊��ł� VT_I4 ���Ԃ邩�ǂ����͊m�M���Ȃ�
		pDispatch2->Release();
		pDispatch->Release();
		return false;
	}
	int num = varResult.lVal;

	pDispatch3->Release();

	std::map<std::string, WCHAR *> pairList;                              // 2000 | 2010
	pairList["Application name"]                   = L"���            : "; //  9 |  9
	pairList["Creation date"]                      = L"�쐬����        : "; // 11 | 11
	pairList["Language"]                           = L"����            : "; //    | 33
	pairList["Last save time"]                     = L"�X�V����        : "; // 12 | 12
	pairList["Last print date"]                    = L"�������        : "; // 10 | 10
	pairList["Last author"]                        = L"�X�V��          : "; //  7 |  7
	pairList["Revision number"]                    = L"�����ԍ�        : "; //  8 |  8
	pairList["Total editing time"]                 = L"�ҏW����        : "; // 13 | 13
	pairList["Security"]                           = L"�߽ܰ�ސݒ�(0/1): "; // 17 | 17
	pairList["Title"]                              = L"�^�C�g��        : "; //  1 |  1
	pairList["Subject"]                            = L"�T�u�^�C�g��    : "; //  2 |  2
	pairList["Author"]                             = L"�쐬��          : "; //  3 |  3
	pairList["Manager"]                            = L"�Ǘ���          : "; // 20 | 20
	pairList["Company"]                            = L"��Ж�          : "; // 21 | 21
	pairList["Category"]                           = L"����            : "; // 18 | 18
	pairList["Keywords"]                           = L"�L�[���[�h      : "; //  4 |  4
	pairList["Comments"]                           = L"�R�����g        : "; //  5 |  5
	pairList["Content type"]                       = L"���������       : "; //    | 31
	pairList["Content status"]                     = L"����̏��     : "; //    | 32
	pairList["Document version"]                   = L"�޷�����ް�ޮ�  : "; //    | 34
	pairList["Hyperlink base"]                     = L"ʲ�߰�ݸ�̊�_  : "; // 29 | 29
	pairList["Template"]                           = L"�e���v���[�g    : "; //  6 |  6
	pairList["Number of pages"]                    = L"�y�[�W��        : "; // 14 | 14
	pairList["Number of paragraphs"]               = L"�i����          : "; //    |
	pairList["Number of lines"]                    = L"�s��            : "; // 23 | 23
	pairList["Number of words"]                    = L"�P�ꐔ          : "; //    | 15
	pairList["Number of characters"]               = L"������          : "; // 16 | 16
	pairList["Number of characters (with spaces)"] = L"������(��߰���) : "; // 30 | 30
	pairList["Number of bytes"]                    = L"�o�C�g��        : "; // 22 | 22
	pairList["Number of slides"]                   = L"�X���C�h��      : "; // 25 | 25
	pairList["Number of paragraphs"]               = L"�i����          : "; // 24 | 24
	pairList["Number of words"]                    = L"�P�ꐔ          : "; // 15 |
	pairList["Number of bytes"]                    = L"�o�C�g��        : "; //    |
	pairList["Number of notes"]                    = L"�m�[�g��        : "; // 26 | 26
	pairList["Number of hidden Slides"]            = L"��\���ײ�ސ�   : "; // 27 | 27
	pairList["Number of multimedia clips"]         = L"�����ި��د�ߐ� : "; // 28 | 28
	pairList["Format"]                             = L"��ھ��ð��݌`�� : "; // 19 | 19

	VARIANT parm1;
	VariantInit(&parm1);
	parm1.vt = VT_INT;
	for (parm1.intVal = 0; parm1.intVal < num; parm1.intVal++) {
		hr = getProperty1(pDispatch2, &varResult, L"BuiltInDocumentProperties", parm1);
		pDispatch3 = varResult.pdispVal;
		if (SUCCEEDED(hr) && pDispatch3 != NULL) {
			hr = getProperty(pDispatch3, &varResult, L"Name");
			if (SUCCEEDED(hr) && varResult.vt == VT_BSTR && varResult.bstrVal != NULL) {
				wcsncat_s(out, outsize, L"\n", _TRUNCATE);
				size_t nc;
				char mbstr[50]; // �ł�����"Number of multimedia clips"������\���Ȓ���
				wcstombs_s(&nc, mbstr, _countof(mbstr), varResult.bstrVal, _TRUNCATE);
				WCHAR *p = pairList[mbstr];
				if (p != NULL) {
					wcsncat_s(out, outsize, p, _TRUNCATE);
				} else {
					varcat(out, outsize, varResult);
					wcsncat_s(out, outsize, L" : ", _TRUNCATE);
				}

				hr = getProperty(pDispatch3, &varResult, L"Value");
				if (SUCCEEDED(hr)) {
					varcat(out, outsize, varResult);
				}
			}
			pDispatch3->Release();
		}
	}

	hr = getProperty(pDispatch2, &varResult, L"MultiUserEditing");
	if(SUCCEEDED(hr)) {
		wcsncat_s(out, outsize, L"\n���L            : ", _TRUNCATE);
		varcat(out, outsize, varResult);
	}

	pDispatch2->Release();
	pDispatch->Release();

	return true;
}

BOOL getFileProperty(WCHAR *out, int outsize, int doctype)
{
	BOOL ret;

	CoInitialize(NULL);
	switch (doctype) {
	case MS_EXCEL:
		ret = _getFileProperty(out, outsize, L"Excel.Application", L"ActiveWorkbook");
		break;
	case MS_WORD:
		ret = _getFileProperty(out, outsize, L"Word.Application", L"ActiveDocument");
		break;
	case MS_PPT:
		ret = _getFileProperty(out, outsize, L"PowerPoint.Application", L"ActivePresentation");
		break;
	}
	CoUninitialize();

	return ret;
}
