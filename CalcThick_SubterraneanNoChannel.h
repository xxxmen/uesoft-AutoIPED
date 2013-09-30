// CalcThick_SubterraneanNoChannel.h: interface for the CCalcThick_SubterraneanNoChannel class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICK_SUBTERRANEANNOCHANNEL_H__0C32A533_DF05_4CC6_88AB_F82A6DCD8CB8__INCLUDED_)
#define AFX_CALCTHICK_SUBTERRANEANNOCHANNEL_H__0C32A533_DF05_4CC6_88AB_F82A6DCD8CB8__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcThickSubterranean.h"

//////////////////////////////////////////////////////////////////////////
//
//计算地下直埋敷设(无管沟)，管道的保温厚度
//
//////////////////////////////////////////////////////////////////////////

class CCalcThick_SubterraneanNoChannel : public CCalcThickSubterranean 
{
public:
	short CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &ts1_resu, double &ts2_resu);
	// 管道之间的当量热阻 
	virtual void Calc_Ro(double &dRo);
	// 无管沟的单根保温层热阻
	void Calc_Rpipe_One(double D1, double dLamda, double& dR1);	
	// 无管沟的双根管道保温层热阻
	void Calc_Rpipe_Two(double dQ1, double dQ2, double& dR1, double& dR2);
	// 无管沟敷设土壤热阻
	void Calc_Rso(double &dRso);
	// 无管沟单根管道的热损失
	short Calc_Q_PipeOne(double &dOneQ);
	// 无管沟两根管道的热损失 
	void Calc_Q_PipeTwo(double &dReQ1, double &dReQ2); 
	// 无管沟管道 保温计算 
	short CalcPipe_One(double &thick_resu, double &ts_resu); 

public:
	CCalcThick_SubterraneanNoChannel();
	virtual ~CCalcThick_SubterraneanNoChannel();
	
};

#endif // !defined(AFX_CALCTHICK_SUBTERRANEANNOCHANNEL_H__0C32A533_DF05_4CC6_88AB_F82A6DCD8CB8__INCLUDED_)
