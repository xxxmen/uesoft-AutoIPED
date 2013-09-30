#include "stdafx.h"
#include "VtoT.h"

inline int vtoi(_variant_t & v)
{
	CString tmps;
	int ret=0;
	
	switch(v.vt)
	{
	case VT_NULL:
	case VT_EMPTY:
		ret = 0;
		break;
	case VT_I2:
		ret= (int)V_I2(&v);
		break;
	case VT_I4:
		ret= V_I4(&v);
		break;
	case VT_R4:
		ret= (int)V_R4(&v);
		break;
	case VT_R8:
		ret =(int)V_R8(&v);
		break;
	case VT_BSTR:
		tmps=(char*)_bstr_t(&v);
		ret=atoi((LPCSTR)tmps);
		break;
	}
	return ret;
}

inline CString  vtos(_variant_t& v)
{
	CString ret;
	switch(v.vt)
	{
	case VT_NULL:
	case VT_EMPTY:
		ret="";
		break;
	case VT_BSTR:
		ret=V_BSTR(&v);
		break;
	case VT_R4:
		ret.Format("%g", (double)V_R4(&v));
		break;
	case VT_R8:
		ret.Format("%g",V_R8(&v));
		break;
	case VT_I2:
		ret.Format("%d",(int)V_I2(&v));
		break;
	case VT_I4:
		ret.Format("%d",(int)V_I4(&v));
		break;
	case VT_BOOL:
		ret=(V_BOOL(&v) ? "True" : "False");
		break;
	}
	ret.TrimLeft();ret.TrimRight();
	return ret;
}

inline double  vtof(_variant_t &v)
{
	double ret=0;
	CString tmps;
	switch(v.vt)
	{
	case VT_NULL:
	case VT_EMPTY:
		ret = 0;
		break;
	case VT_R4:
		ret = V_R4(&v);
		break;
	case VT_R8:
		ret=(double)V_R8(&v);
		break;
	case VT_I2:
		ret= (float)V_I2(&v);
		break;
	case VT_I4:
		ret= (float)V_I4(&v);
		break;
	case VT_BSTR:
		tmps=(char*)_bstr_t(&v);
		ret=static_cast<float>(atof((LPCSTR)tmps));
		break;
	case VT_DECIMAL:
		ret=(double)v;
	}
	return ret;
}

inline bool  Btob( BOOL Bn)
{
	return ( Bn == 0 ) ? false : true ;
}

inline bool  Dtob( double Dn)
{
	return Dn ? true : false;
}

inline bool vtob( _variant_t &v)
{
	switch(v.vt)
	{
	case VT_BOOL:
		return Btob( V_BOOL(&v) );
	case VT_I2:
		return Btob( V_I2(&v) );
	case VT_I4:
		return Btob( V_I4(&v) );
	case VT_R4:
		return Dtob( V_R4(&v) ); 
	case VT_R8:
		return Dtob( V_R8(&v) );
	}
	return false;
}

inline CString vtos(COleVariant& v)
{
	CString ret;
	switch(v.vt)
	{
	case VT_NULL:
	case VT_EMPTY:
		ret="";
		break;
	case VT_BSTR:
		ret=V_BSTRT(&v);
		break;
	case VT_R4:
		ret.Format("%g", (double)V_R4(&v));
		break;
	case VT_R8:
		ret.Format("%g",V_R8(&v));
		break;
	case VT_I2:
		ret.Format("%d",(int)V_I2(&v));
		break;
	case VT_I4:
		ret.Format("%d",(int)V_I4(&v));
		break;
	case VT_BOOL:
		ret=(V_BOOL(&v) ? "True" : "False");
		break;
	}
	ret.TrimLeft();ret.TrimRight();
	return ret;
}
