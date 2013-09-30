// CalcAlphaOriginalData.h: interface for the CCalcAlphaOriginalData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCALPHAORIGINALDATA_H__6B052138_5D6B_4DE6_BDA5_ACE00D6518AD__INCLUDED_)
#define AFX_CALCALPHAORIGINALDATA_H__6B052138_5D6B_4DE6_BDA5_ACE00D6518AD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

//------------------------------------------------------------------
// DATE         :[2005/08/16]
// Parameter(s) :
// Return       :
// Remark       :计算放热系数时到用到的数据.在调用时进行赋初值
//------------------------------------------------------------------
class CCalcAlphaOriginalData  
{
public:
	CCalcAlphaOriginalData();
	virtual ~CCalcAlphaOriginalData();

	double	dA_ta;		//环境温度
	double  dA_ts;		//外表面温度
	double	dA_W;		//年平均风速
	double	dA_hedu;	//黑度
	double	dA_B;		//沿风速方向的平壁宽度，m
	double	dA_AlphaIndex;//ALPHA的取值索引

};

#endif // !defined(AFX_CALCALPHAORIGINALDATA_H__6B052138_5D6B_4DE6_BDA5_ACE00D6518AD__INCLUDED_)
