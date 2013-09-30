// MainFrm.h : interface of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_MAINFRM_H__8F862089_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
#define AFX_MAINFRM_H__8F862089_B060_11D7_9BCC_000AE616B8F1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ProgBar.h"

#include "ReportCalErrorDlgBar.h"	// 用于显示计算过程中的错误信息的扩张DialogBar

class CMainFrame : public CFrameWnd
{
	
protected: // create from serialization only
	DECLARE_DYNCREATE(CMainFrame)

// Attributes
public:
	CMainFrame();

// Operations
public:
	CProgStatusBar m_wndStatusBar;
	// 用于显示计算过程中的错误信息的扩张DialogBar
	CReportCalErrorDlgBar	m_ReportCalInfoBar;	

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMainFrame)
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	//}}AFX_VIRTUAL

// Implementation
public:
	void ShowCurrentProjectName();
	void ShowReportBar(bool bShow);
	CString GetReportBarContent();
	void SetReportBarContent(LPCTSTR szInfo);
	CBitmap bmp8;
	LRESULT RemoveProgress(WPARAM wParam,LPARAM lParam);

	LRESULT AddProgress(WPARAM wParam,LPARAM lParam);
	bool m_bIsExitMsg;
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:  // control bar embedded members
//	CStatusBar  m_wndStatusBar;
//	CProgStatusBar m_wndStatusBar;
//	CProgStatusBar m_wndStatusBar;
	CToolBar    m_wndToolBar;


// Generated message map functions
protected:
	//{{AFX_MSG(CMainFrame)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnMove(int x, int y);
	afx_msg void OnMenuHelp();
	afx_msg void OnMENUIndex();
	afx_msg void OnMENUSearch();
	afx_msg void OnViewInformationBar();
	afx_msg void OnUpdateViewInformationBar(CCmdUI* pCmdUI);
	afx_msg void OnClose();
	afx_msg void OnTimer(UINT nIDEvent);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__8F862089_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
