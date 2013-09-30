// CalcAlpha_Code.cpp: implementation of the CCalcAlpha_Code class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcAlpha_Code.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcAlpha_Code::CCalcAlpha_Code()
{

}

CCalcAlpha_Code::~CCalcAlpha_Code()
{

}


//------------------------------------------------------------------
// DATE         :[2005/08/11]
// Parameter(s) :
// Return       :
// Remark       :根据电力行业的规范计算平面的放热系数
//------------------------------------------------------------------
BOOL CCalcAlpha_Code::CalcPlain_Alpha(double &dresAlpha)
{
	double dAn=0;		//辐射传热系数
	double dAc=0;		//对流传系数

	dAn = GetEradiateCoefficient();
	
	dAc = (5.93-0.015*dA_ta) * pow(dA_W,0.8)/pow(dA_B,0.2);
	//返回平面的传热系数
	dresAlpha = dAn + dAc;

	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/08/11]
// Parameter(s) :D1,当为双层时应代入外层的外径, dresAlpha,返回计算所得的放热系数
// Return       :
// Remark       :根据电力行业的规范计算管道的放热系数
//------------------------------------------------------------------
BOOL CCalcAlpha_Code::CalcPipe_Alpha(double D1,double &dresAlpha)
{
	double dAn;		//辐射传热系数.
	double dAc;		//对流传热系数.

	dAn = GetEradiateCoefficient();
	dAc = 72.81*pow(dA_W,0.6)/pow(D1,0.4);

	dresAlpha = dAn + dAc;

	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/08/11]
// Parameter(s) :
// Return       :
// Remark       :获得辐射传热系数
//------------------------------------------------------------------
double CCalcAlpha_Code::GetEradiateCoefficient()
{
	double dAn;		//辐射传热系数
	dAn = 5.67*dA_hedu/(dA_ts-dA_ta)*(pow((273+dA_ts)/100,4.0)-pow((273+dA_ta)/100,4.0));
	
	return dAn;
}
