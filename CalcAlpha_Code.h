// CalcAlpha_Code.h: interface for the CCalcAlpha_Code class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCALPHA_CODE_H__34A086DF_298A_4170_BB5C_FE07962E34B7__INCLUDED_)
#define AFX_CALCALPHA_CODE_H__34A086DF_298A_4170_BB5C_FE07962E34B7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcAlphaOriginalData.h"

class CCalcAlpha_Code : public CCalcAlphaOriginalData  
{
public:
	double GetEradiateCoefficient();
	BOOL CalcPipe_Alpha(double D1,double& dresAlpha);
	BOOL CalcPlain_Alpha(double &dresAlpha);
	CCalcAlpha_Code();
	virtual ~CCalcAlpha_Code();

};

#endif // !defined(AFX_CALCALPHA_CODE_H__34A086DF_298A_4170_BB5C_FE07962E34B7__INCLUDED_)
