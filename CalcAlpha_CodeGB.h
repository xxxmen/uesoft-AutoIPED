// CalcAlpha_CodeGB.h: interface for the CCalcAlpha_CodeGB class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCALPHA_CODEGB_H__A3E078A3_3952_4922_8006_6EBB6BBC7E09__INCLUDED_)
#define AFX_CALCALPHA_CODEGB_H__A3E078A3_3952_4922_8006_6EBB6BBC7E09__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcAlphaOriginalData.h"

class CCalcAlpha_CodeGB : public CCalcAlphaOriginalData  
{
public:
	BOOL CalcPreventScald_Alpha(double& dresAlpha);
	BOOL CalcEffectTest_Alpha(double D1,double& dresAlpha);
	BOOL CalcEconomy_Alpha(double& dresAlpha);

	CCalcAlpha_CodeGB();
	virtual ~CCalcAlpha_CodeGB();

};

#endif // !defined(AFX_CALCALPHA_CODEGB_H__A3E078A3_3952_4922_8006_6EBB6BBC7E09__INCLUDED_)
