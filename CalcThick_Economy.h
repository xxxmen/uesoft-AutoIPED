// CalcThick_Economy.h: interface for the CCalcThick_Economy class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICK_ECONOMY_H__A70D67DE_AD0F_4C44_84C2_7948207286A1__INCLUDED_)
#define AFX_CALCTHICK_ECONOMY_H__A70D67DE_AD0F_4C44_84C2_7948207286A1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "CalcThickBase.h"

class CCalcThick_Economy  : public CCalcThickBase   
{
public:
	CCalcThick_Economy();
	virtual ~CCalcThick_Economy();

	//¼ÆËã±£ÎÂºñ¶È
	virtual short CalcPipe_One(double &thick_resu, double &ts_resu);
	virtual short CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPlain_One(double &thick_resu, double &ts_resu);
	virtual short CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);

};

#endif // !defined(AFX_CALCTHICK_ECONOMY_H__A70D67DE_AD0F_4C44_84C2_7948207286A1__INCLUDED_)
