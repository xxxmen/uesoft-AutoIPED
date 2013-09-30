// CalcThick_HeatDensity.h: interface for the CCalcThick_HeatDensity class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICK_HEATDENSITY_H__FC5483CB_762C_40BB_AC36_9F9749B84C3C__INCLUDED_)
#define AFX_CALCTHICK_HEATDENSITY_H__FC5483CB_762C_40BB_AC36_9F9749B84C3C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcThickBase.h"

class CCalcThick_HeatDensity : public CCalcThickBase  
{
public:
	CCalcThick_HeatDensity();
	virtual ~CCalcThick_HeatDensity();

	//¼ÆËã±£ÎÂºñ¶È
	virtual short CalcPipe_One(double &thick_resu, double &ts_resu);
	virtual short CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPlain_One(double &thick_resu, double &ts_resu);
	virtual short CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);

};

#endif // !defined(AFX_CALCTHICK_HEATDENSITY_H__FC5483CB_762C_40BB_AC36_9F9749B84C3C__INCLUDED_)
