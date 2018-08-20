#include "stdafx.h"

static HRESULT _invoke(IDispatch *pDispatch, VARIANT *pVarResult, LPOLESTR pName, WORD flags, int argc...) 
{
	HRESULT hr;
	DISPID dispID;
   
	if (! pDispatch) {
		return S_FALSE;
	}
	hr = pDispatch->GetIDsOfNames(IID_NULL, &pName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		return hr;
	}

	// InvokeÇ÷ìnÇ∑à¯êî(â¬ïœå¬)ÇÃèàóù(Ç±Ç±Ç©ÇÁ)
	va_list ap;
	va_start(ap, argc);

	VARIANT *argv = new VARIANT[argc + 1];

	for(int i = 0; i < argc; i++) {
		argv[i] = va_arg(ap, VARIANT);
	}
	va_end(ap);
	// InvokeÇ÷ìnÇ∑à¯êî(â¬ïœå¬)ÇÃèàóù(Ç±Ç±Ç‹Ç≈)


	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;

	dp.cArgs = argc;
	dp.rgvarg = argv;
   
	if (flags & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	} else {
		dp.cNamedArgs = 0;
		dp.rgdispidNamedArgs = NULL;
	}
	hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, flags, &dp, pVarResult, NULL, NULL);

	delete [] argv;
   
	return hr;
}

HRESULT getProperty(IDispatch *pDispatch, VARIANT *pVarResult, WCHAR *prop_name)
{
	return _invoke(pDispatch, pVarResult, prop_name, DISPATCH_PROPERTYGET, 0);
}

HRESULT getProperty1(IDispatch *pDispatch, VARIANT *pVarResult, WCHAR *prop_name, VARIANT parm)
{
	return _invoke(pDispatch, pVarResult, prop_name, DISPATCH_PROPERTYGET, 1, parm);
}
