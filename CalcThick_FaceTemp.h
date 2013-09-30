// CalcThick_FaceTemp.h: interface for the CCalcThick_FaceTemp class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICK_FACETEMP_H__28A479EC_25E2_4607_887A_811FDEE82188__INCLUDED_)
#define AFX_CALCTHICK_FACETEMP_H__28A479EC_25E2_4607_887A_811FDEE82188__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcThickBase.h"

class CCalcThick_FaceTemp : public CCalcThickBase  
{
public:
	virtual short CalcPlain_One_InputThick(double thick_resu, double &ts_resu);
	virtual short CalcPipe_One_InputThick(double thick_resu, double &ts_resu);
	virtual short CalcPipe_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPlain_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPipe_One(double &thick_resu, double &ts_resu);
	virtual short CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);

	virtual short CalcPlain_One(double &thick_resu, double &ts_resu);

	CCalcThick_FaceTemp();
	virtual ~CCalcThick_FaceTemp();

};

#endif // !defined(AFX_CALCTHICK_FACETEMP_H__28A479EC_25E2_4607_887A_811FDEE82188__INCLUDED_)
