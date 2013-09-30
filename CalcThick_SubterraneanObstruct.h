// CalcThick_SubterraneanObstruct.h: interface for the CCalcThick_SubterraneanObstruct class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICK_SUBTERRANEANOBSTRUCT_H__B6C8E88A_C1BD_48BC_8700_005DEA0F22A2__INCLUDED_)
#define AFX_CALCTHICK_SUBTERRANEANOBSTRUCT_H__B6C8E88A_C1BD_48BC_8700_005DEA0F22A2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcThickSubterranean.h"
//////////////////////////////////////////////////////////////////////////
//
//计算地下敷设有管沟（不通行地沟中）管道的保温厚度
//
//////////////////////////////////////////////////////////////////////////
class CCalcThick_SubterraneanObstruct : public CCalcThickSubterranean 
{
public:
	inline	virtual short Calc_Q_PipeOne(double &dOneQ);
	inline	virtual	void Calc_Q_PipeTwo(double &dQ1, double & dQ2);

	virtual	short CalcPipe_One(double &thick_resu, double &ts_resu); // 埋地管道单层

	virtual short CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu); //计算埋地管道保温层厚

	CCalcThick_SubterraneanObstruct(); 
	virtual ~CCalcThick_SubterraneanObstruct();

private:
};

#endif // !defined(AFX_CALCTHICK_SUBTERRANEANOBSTRUCT_H__B6C8E88A_C1BD_48BC_8700_005DEA0F22A2__INCLUDED_)
