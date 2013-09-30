// CalcAlpha_CodePC.h: interface for the CCalcAlpha_CodePC class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCALPHA_CODEPC_H__7093C31E_101E_454C_A781_8AACD9BFBE1F__INCLUDED_)
#define AFX_CALCALPHA_CODEPC_H__7093C31E_101E_454C_A781_8AACD9BFBE1F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcAlphaOriginalData.h"

class CCalcAlpha_CodePC : public CCalcAlphaOriginalData  
{
public:
	BOOL CalcEconomy_Alpha(double& dresAlpha);
	BOOL CalcFaceTemp_Alpha(double& dresAlpha);
	CCalcAlpha_CodePC();
	virtual ~CCalcAlpha_CodePC();

};

#endif // !defined(AFX_CALCALPHA_CODEPC_H__7093C31E_101E_454C_A781_8AACD9BFBE1F__INCLUDED_)
