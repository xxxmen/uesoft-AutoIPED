#if !defined(AFX_REPORTCALERRORDLGBAR_H__EFDA8E35_E54E_4EC6_82DB_651597487C8F__INCLUDED_)
#define AFX_REPORTCALERRORDLGBAR_H__EFDA8E35_E54E_4EC6_82DB_651597487C8F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ReportCalErrorDlgBar.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// 
// 用于显示计算过程中的错误信息的扩张DialogBar
//
/////////////////////////////////////////////////////////////////////////////

#include "DialogBarEx.h"

#include "Resource.h"

class CReportCalErrorDlgBar : public CDialogBarEx
{
// Construction
public:
	CReportCalErrorDlgBar();
	virtual ~CReportCalErrorDlgBar();

	enum{IDD=IDD_REPORT_CAL_ERROR_DLG};

public:

public:
	CString GetReportContent();
	void SetReportContent(LPCTSTR szContent);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CReportCalErrorDlgBar)
	//}}AFX_VIRTUAL

private:
	CString	m_InformationToReport;

	// Generated message map functions
protected:
	//{{AFX_MSG(CReportCalErrorDlgBar)
	afx_msg void OnSize(UINT nType, int cx, int cy);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_REPORTCALERRORDLGBAR_H__EFDA8E35_E54E_4EC6_82DB_651597487C8F__INCLUDED_)
