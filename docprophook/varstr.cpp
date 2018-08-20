#include "stdafx.h"
#include <stdio.h>

void varcat(WCHAR *buf, int bufsize, VARIANT val)
{
	WCHAR ibuf[21];
	WCHAR ubuf[20];
	WCHAR fbuf[11];
	WCHAR dbuf[11];
	WCHAR cbuf[1];
	WCHAR dtbuf[20];

	switch (val.vt) {
	case VT_I1 :         //  [V][T][P][s]  signed char (char)
		_snwprintf_s(cbuf, _countof(cbuf), _TRUNCATE, L"%c", val.cVal);
		wcsncat_s(buf, bufsize, cbuf, _TRUNCATE);
		break;

	case VT_I2 :         //  [V][T][P][S]  2 byte signed int (short)
		_snwprintf_s(ibuf, _countof(ibuf), _TRUNCATE, L"%d", val.iVal);
		wcsncat_s(buf, bufsize, ibuf, _TRUNCATE);
		break;

	case VT_I4 :         //  [V][T][P][S]  4 byte signed int (long)
		_snwprintf_s(ibuf, _countof(ibuf), _TRUNCATE, L"%d", val.lVal);
		wcsncat_s(buf, bufsize, ibuf, _TRUNCATE);
		break;

	case VT_INT :        //  [V][T][P][S]  signed machine int (int)
		_snwprintf_s(ibuf, _countof(ibuf), _TRUNCATE, L"%d", val.intVal);
		wcsncat_s(buf, bufsize, ibuf, _TRUNCATE);
		break;

	case VT_UI1 :        //  [V][T][P][S]  unsigned char (Byte)
		_snwprintf_s(cbuf, _countof(cbuf), _TRUNCATE, L"%c", val.bVal);
		wcsncat_s(buf, bufsize, cbuf, _TRUNCATE);
		break;

	case VT_UI2 :        //  [V][T][P][S]  unsigned short (unsigned short)
		_snwprintf_s(ubuf, _countof(ubuf), _TRUNCATE, L"%u", val.uiVal);
		wcsncat_s(buf, bufsize, ubuf, _TRUNCATE);
		break;

	case VT_UI4 :        //  [V][T][P][S]  unsigned long (unsigned long)
		_snwprintf_s(ubuf, _countof(ubuf), _TRUNCATE, L"%u", val.ulVal);
		wcsncat_s(buf, bufsize, ubuf, _TRUNCATE);
		break;

	case VT_UINT :       //  [V][T]   [S]  unsigned machine int (unsigned int)
		_snwprintf_s(ubuf, _countof(ubuf), _TRUNCATE, L"%u", val.uintVal);
		wcsncat_s(buf, bufsize, ubuf, _TRUNCATE);
		break;

	case VT_R4 :         //  [V][T][P][S]  4 byte real (float)
		_snwprintf_s(fbuf, _countof(fbuf), _TRUNCATE, L"%5g", val.fltVal);
		wcsncat_s(buf, bufsize, fbuf, _TRUNCATE);
		break;

	case VT_R8 :         //  [V][T][P][S]  8 byte real (double)
		_snwprintf_s(dbuf, _countof(dbuf), _TRUNCATE, L"%5g", val.dblVal);
		wcsncat_s(buf, bufsize, dbuf, _TRUNCATE);
		break;

	case VT_BOOL :       //  [V][T][P][S]  True=-1, False=0 (VARIANT_BOOL)
		wcsncat_s(buf, bufsize, val.boolVal ? L"True" : L"False", _TRUNCATE);
		break;

	case VT_DATE :       //  [V][T][P][S]  date (DATE)
		if (val.date == 0.0) {
			wcsncat_s(buf, bufsize, L"(null)", _TRUNCATE);
		} else {
			SYSTEMTIME st;
			VariantTimeToSystemTime(val.date, &st);
			_snwprintf_s(dtbuf, _countof(dtbuf), _TRUNCATE, L"%04d-%02d-%02d %02d:%02d:%02d", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
			wcsncat_s(buf, bufsize, dtbuf, _TRUNCATE);
		}
		break;

	case VT_BSTR :       //  [V][T][P][S]  OLE Automation string (BSTR)
		wcsncat_s(buf, bufsize, val.bstrVal ? val.bstrVal : L"(null)", _TRUNCATE);
		break;

	case VT_EMPTY :      //  [V]   [P]     nothing
		wcsncat_s(buf, bufsize, L"", _TRUNCATE);
		break;

	default:
	/*
	 *   VT_NULL             [V]   [P]              SQL style Null
	 *   VT_CY               [V][T][P][S]  cyVal    currency (CY)
	 *   VT_DISPATCH         [V][T]   [S]  pdispVal IDispatch *
	 *   VT_ERROR            [V][T][P][S]           SCODE
	 *   VT_VARIANT          [V][T][P][S]           VARIANT *
	 *   VT_UNKNOWN          [V][T]   [S]  punkVal  IUnknown *
	 *   VT_DECIMAL          [V][T]   [S]           16 byte fixed point
	 *   VT_RECORD           [V]   [P][S]           user defined type
	 *   VT_VOID                [T]                 C style void
	 *   VT_HRESULT             [T]                 Standard return type
	 *   VT_PTR                 [T]                 pointer type
	 *   VT_INT_PTR             [T]                 signed machine register size width
	 *   VT_UINT_PTR            [T]                 unsigned machine register size width
	 *   VT_SAFEARRAY           [T]                 (use VT_ARRAY in VARIANT)
	 *   VT_CARRAY              [T]                 C style array
	 *   VT_USERDEFINED         [T]                 user defined type
	 *   VT_I8                  [T][P]              signed 64-bit int
	 *   VT_UI8                 [T][P]              unsigned 64-bit int
	 *   VT_LPSTR               [T][P]              null terminated string
	 *   VT_LPWSTR              [T][P]              wide null terminated string
	 *   VT_FILETIME               [P]              FILETIME
	 *   VT_BLOB                   [P]              Length prefixed bytes
	 *   VT_STREAM                 [P]              Name of the stream follows
	 *   VT_STORAGE                [P]              Name of the storage follows
	 *   VT_STREAMED_OBJECT        [P]              Stream contains an object
	 *   VT_STORED_OBJECT          [P]              Storage contains an object
	 *   VT_VERSIONED_STREAM       [P]              Stream with a GUID version
	 *   VT_BLOB_OBJECT            [P]              Blob contains an object 
	 *   VT_CF                     [P]              Clipboard format
	 *   VT_CLSID                  [P]              A Class ID
	 *   VT_VECTOR                 [P]              simple counted array
	 *   VT_ARRAY            [V]           parray   SAFEARRAY*
	 *   VT_BYREF            [V]                    void* for local use
	 *   VT_BSTR_BLOB                               Reserved for system use
	 */
		wcsncat_s(buf, bufsize, L"(unable to print)", _TRUNCATE);
		break;
	}
}
