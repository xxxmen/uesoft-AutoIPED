#include "stdafx.h"
#include "mobject.h"
#include <memory.h>

IUnknown* MAtlComQIPtrAssign(IUnknown** pp, IUnknown* lp, REFIID riid)
{
	IUnknown* pTemp = *pp;
	*pp = NULL;
	if (lp != NULL)
		lp->QueryInterface(riid, (void**)pp);
	try
	{
		if (pTemp)
			pTemp->Release();
	}
	catch(...)
	{
	}
	return *pp;
}

IDispatch* CMObject::operator=(IDispatch* lp)
{
	try
	{
		if(p!=NULL)
			p->Release();
	}
	catch(...)
	{
	}
	p=lp;
	if(p!=NULL)
		p->AddRef();
	return p;
}

IDispatch* CMObject::operator=(IUnknown* lp)
{
	return (IDispatch*)MAtlComQIPtrAssign((IUnknown**)&p, lp, IID_IDispatch);
}

_variant_t CMObject::GetPropertyByName(LPCTSTR lpsz)
{
	_variant_t vRet;
	vRet.ChangeType(VT_EMPTY);
	if(p==NULL)
		return vRet;
	DISPID dwDispID;
	HRESULT hr = GetIDOfName(lpsz, &dwDispID);
	if (SUCCEEDED(hr))
		GetProperty(p, dwDispID, &vRet);
	return vRet;
}

_variant_t CMObject::GetProperty(DISPID dwDispID)
{
	_variant_t vRet;
	if(p==NULL)
		return vRet;
	GetProperty(p, dwDispID, &vRet);
	return vRet;
}

HRESULT CMObject::PutPropertyByName(LPCTSTR lpsz, VARIANT* pVar)
{
	DISPID dwDispID;
	if(p==NULL)
		return -1;
	HRESULT hr = GetIDOfName(lpsz, &dwDispID);
	if (SUCCEEDED(hr))
		hr = PutProperty(p, dwDispID, pVar);
	return hr;
}

HRESULT CMObject::PutProperty(DISPID dwDispID, VARIANT* pVar)
{
	if(p==NULL)
		return -1;
	return PutProperty(p, dwDispID, pVar);
}

HRESULT CMObject::GetIDOfName(LPCTSTR lpsz, DISPID* pdispid)
{
	if(p==NULL)
		return -1;
	LPOLESTR lpOleSz=NULL;
	_bstr_t tmpbstr=_bstr_t(lpsz);
	lpOleSz=(LPOLESTR)tmpbstr.operator wchar_t *();
	HRESULT rt;
	rt= p->GetIDsOfNames(IID_NULL, (LPOLESTR*)&lpOleSz, 1, LOCALE_SYSTEM_DEFAULT, pdispid);
	return rt;
}

// Invoke a method by DISPID with no parameters
_variant_t CMObject::Invoke0(DISPID dispid)
{
	_variant_t varRet;
	if(p==NULL)
		return varRet;
	DISPPARAMS dispparams = { NULL, NULL, 0, 0};
	p->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &varRet, NULL, NULL);
	return varRet;
}

// Invoke a method by name with no parameters
_variant_t CMObject::Invoke0(LPCTSTR lpszName)
{
	HRESULT hr;
	DISPID dispid;
	_variant_t varRet;
	if(p==NULL)
		return varRet;
	hr = GetIDOfName(lpszName, &dispid);
	if (SUCCEEDED(hr))
		varRet = Invoke0(dispid);
	return varRet;
}

HRESULT CMObject::GetProperty(IDispatch* pDisp, DISPID dwDispID, VARIANT* pVar)
{
	DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
	return pDisp->Invoke(dwDispID, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
		&dispparamsNoArgs, pVar, NULL, NULL);
}

HRESULT CMObject::PutProperty(IDispatch* pDisp, DISPID dwDispID, VARIANT* pVar)
{
	DISPPARAMS dispparams = {NULL, NULL, 1, 1};
	dispparams.rgvarg = pVar;
	DISPID dispidPut = DISPID_PROPERTYPUT;
	dispparams.rgdispidNamedArgs = &dispidPut;

	if (pVar->vt == VT_UNKNOWN || pVar->vt == VT_DISPATCH || 
		(pVar->vt & VT_ARRAY) || (pVar->vt & VT_BYREF))
	{
		HRESULT hr = pDisp->Invoke(dwDispID, IID_NULL,
			LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF,
			&dispparams, NULL, NULL, NULL);
		if (SUCCEEDED(hr))
			return hr;
	}

	return pDisp->Invoke(dwDispID, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
		&dispparams, NULL, NULL, NULL);
}

HRESULT CMObject::CreateObject(LPCTSTR objName)
{
	HRESULT hr;
	IDispatchPtr tmpobj;
	Release();
#ifdef _UNICODE
	hr=tmpobj.CreateInstance( (LPOLESTR)objName);
#else
	hr=tmpobj.CreateInstance((LPCSTR)objName);
#endif
	if (SUCCEEDED(hr))
		p=tmpobj.Detach();
	return hr;
}

HRESULT CMObject::GetActiveObject(LPCTSTR objName)
{
	HRESULT hr;
	IDispatchPtr tmpobj;
	Release();
#ifdef _UNICODE
	hr=tmpobj.GetActiveObject((LPOLESTR)objName);
#else
	hr=tmpobj.GetActiveObject(objName);
#endif
	if(SUCCEEDED(hr))
		p=tmpobj.Detach();
	return hr;
}

HRESULT CMObject::CreateObject(const CLSID clsid)
{
	HRESULT hr;
	IDispatchPtr tmpobj;
	Release();
	hr=tmpobj.CreateInstance(clsid);
	if (SUCCEEDED(hr))
		p=tmpobj.Detach();
	return hr;

}

HRESULT CMObject::GetActiveObject(const CLSID clsid)
{
	HRESULT hr;
	IDispatchPtr tmpobj;
	Release();
	hr=tmpobj.GetActiveObject(clsid);
	if(SUCCEEDED(hr))
		p=tmpobj.Detach();
	return hr;

}


IDispatch* CMObject::operator =(_variant_t& newVar)
{
	try
	{
		if(p!=NULL)
			p->Release();
	}
	catch(...)
	{
	}
	p=NULL;
	if(newVar.vt==VT_DISPATCH)
	{
		p=newVar.pdispVal;
		p->AddRef();
		return p;
	}
	else if(newVar.vt==VT_UNKNOWN)
		return (IDispatch*)MAtlComQIPtrAssign((IUnknown**)&p, (IUnknown *)newVar, IID_IDispatch);
	return p;
}

CMObject::CMObject(_variant_t &tvar)
{
	p=NULL;
	if(tvar.vt==VT_DISPATCH)
	{
		((IDispatch*)tvar)->AddRef();
		p=(IDispatch*)tvar;
	}
	else if(tvar.vt==VT_UNKNOWN)
		((IUnknown *)tvar)->QueryInterface(IID_IDispatch,(void**)&p);

}

_variant_t CMObject::Invoke(LPCTSTR sName, long nParamCount, VARIANT * ...)
{
	_variant_t varRet;
	if(nParamCount<0 || p==NULL)
		return varRet;
	DISPID dispid;
	GetIDOfName(sName,&dispid);
	va_list argList;
	va_start(argList,nParamCount);
	varRet=InvokeV(dispid,nParamCount,argList);
	va_end(argList);
	return varRet;
}

_variant_t CMObject::Invoke(DISPID dispid, long nParamCount, VARIANT * ...)
{
	_variant_t varRet;
	va_list argList;
	va_start(argList,nParamCount);
	varRet=InvokeV(dispid,nParamCount,argList);
	va_end(argList);
	return varRet;
}

_variant_t CMObject::InvokeV(DISPID dispid, long nParamCount, va_list argList)
{
	_variant_t varRet;
	if(nParamCount<0 || p==NULL)
		return varRet;
	VARIANT* varArgs=NULL;
	if(nParamCount>0)
	{
		varArgs=new VARIANT[nParamCount];
		for(int i=nParamCount-1;i>=0;i--)
			memcpy(&varArgs[i],va_arg(argList,LPCTSTR),sizeof(VARIANT));
	}
	DISPPARAMS dispparams = { varArgs, NULL, nParamCount, 0};
	p->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &varRet, NULL, NULL);
	if(varArgs!=NULL)
		delete [] varArgs;
	return varRet;
}