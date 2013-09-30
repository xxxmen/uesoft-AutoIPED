#if !defined(AFX_OUTEXCELFILETHREAD_H__70F153EE_80C8_426D_81CF_1D2C4A65AF78__INCLUDED_)
#define AFX_OUTEXCELFILETHREAD_H__70F153EE_80C8_426D_81CF_1D2C4A65AF78__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// OutExcelFileThread.h : header file
//
#include "Autoipedview.h"


/////////////////////////////////////////////////////////////////////////////
// COutExcelFileThread thread

class COutExcelFileThread : public CWinThread
{
	DECLARE_DYNCREATE(COutExcelFileThread)
protected:

// Attributes
public:
	COutExcelFileThread();           // protected constructor used by dynamic creation
	virtual ~COutExcelFileThread();

// Operations
public:
	bool SetPow(long double _x, long _y);
	void OutInsulExplainTbl();
	void PutCurView(CAutoIPEDView* pCurView);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(COutExcelFileThread)
	public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(COutExcelFileThread)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
private:
	CAutoIPEDView* m_pCurView;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_OUTEXCELFILETHREAD_H__70F153EE_80C8_426D_81CF_1D2C4A65AF78__INCLUDED_)
