#if !defined(AFX_DLGPAINTRULEMODI_H__A9C62747_DAC5_4898_BD1D_30EE1EED8069__INCLUDED_)
#define AFX_DLGPAINTRULEMODI_H__A9C62747_DAC5_4898_BD1D_30EE1EED8069__INCLUDED_

#include "Page1.h"	// Added by ClassView
#include "Page2.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgPaintRuleModi.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDlgPaintRuleModi dialog

class CDlgPaintRuleModi : public CDialog
{
// Construction
public:
	CPage2 m_page2;
	CPropertySheet m_Sheet;
	CPage1 m_page1;
	CDlgPaintRuleModi(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgPaintRuleModi)
	enum { IDD = IDD_DLG_RULE };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgPaintRuleModi)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	BOOL OnInitDialog();
	virtual	void OnCancel();

	// Generated message map functions
	//{{AFX_MSG(CDlgPaintRuleModi)
	afx_msg void OnDestroy();
	afx_msg void OnClose();
	afx_msg void OnCancelMode();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	void DeleteObject();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPAINTRULEMODI_H__A9C62747_DAC5_4898_BD1D_30EE1EED8069__INCLUDED_)
