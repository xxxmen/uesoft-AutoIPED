// CalcThick_PreventCongeal.h: interface for the CCalcThick_PreventCongeal class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICK_PREVENTCONGEAL_H__2CF7030F_BD5C_4ABC_B069_A99ECCF5B200__INCLUDED_)
#define AFX_CALCTHICK_PREVENTCONGEAL_H__2CF7030F_BD5C_4ABC_B069_A99ECCF5B200__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcThickBase.h"

class CCalcThick_PreventCongeal : public CCalcThickBase   
{
public:
	CCalcThick_PreventCongeal();
	virtual ~CCalcThick_PreventCongeal();

	//计算保温厚度
	virtual short CalcPipe_One(double &thick_resu, double &ts_resu);
	virtual short CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPlain_One(double &thick_resu, double &ts_resu);
	virtual short CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);
 
	BOOL GetCongealOriginalData(long nID);

protected:
	short CalcThickFormula(double& thick_resu, double& ts_resu, BOOL bIsPipe=TRUE);
	
private:
	//防冻计算时新增加的变量
	double	Kr;					//管道通过支吊架处的热(或冷)损失的附加系数
	double	taofr;				//防冻结管道允许液体停留时间(小时)
	double	tfr;				//介质在管内冻结温度
	double	V;					//每米管长介质体积，m3/m;
	double	ro;					//介质密度
	double	C;					//介质比热
	double	Vp;					//每米管长管壁体积,m3/m;
	double	rop;				//管材密度
	double	Cp;					//管材比热
	double	Hfr;				//介质融解热
	double	dFlux;				//流量

};

#endif // !defined(AFX_CALCTHICK_PREVENTCONGEAL_H__2CF7030F_BD5C_4ABC_B069_A99ECCF5B200__INCLUDED_)
