#if !defined(AFX_MYBUTTON_H__D4C5F357_25C6_4961_813B_3D23FB4EFB81__INCLUDED_)
#define AFX_MYBUTTON_H__D4C5F357_25C6_4961_813B_3D23FB4EFB81__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// MyButton.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CMyButton window

class CMyButton : public CButton
{
// Construction
public:
	CMyButton();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMyButton)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	//}}AFX_VIRTUAL

// Implementation
public:
	void SetTooltipText(LPCTSTR lpszText, BOOL bActivate);
	void InitToolTip();
	virtual ~CMyButton();

	// Generated message map functions
protected:
	//{{AFX_MSG(CMyButton)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG
private: 
	CToolTipCtrl m_ToolTip;

	DWORD		m_dwToolTipStyle;
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MYBUTTON_H__D4C5F357_25C6_4961_813B_3D23FB4EFB81__INCLUDED_)
