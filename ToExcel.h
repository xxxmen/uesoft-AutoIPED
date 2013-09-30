// ToExcel.h: interface for the ToExcel class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_TOEXCEL_H__D731C3A6_B128_4D12_9E83_396F18B80E7B__INCLUDED_)
#define AFX_TOEXCEL_H__D731C3A6_B128_4D12_9E83_396F18B80E7B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "FoxBase.h"


class ToExcel :public CFoxBase
{
public:
	bool Menu511();
	bool Menu58();
	bool Menu54();
	bool Menu46();
	bool Menu43();
	void OutoExcel(_RecordsetPtr m_pRecordset,CString strPathDest,CString Sheetname,int startrow,int fieldnum);
	bool Menu34();
	ToExcel();
	virtual ~ToExcel();

};

#endif // !defined(AFX_TOEXCEL_H__D731C3A6_B128_4D12_9E83_396F18B80E7B__INCLUDED_)
