#if !defined(AFX_DIALOGBAREXCONTAINER_H__FA9796BD_78A2_4E7E_908D_66BCA6FD1FC2__INCLUDED_)
#define AFX_DIALOGBAREXCONTAINER_H__FA9796BD_78A2_4E7E_908D_66BCA6FD1FC2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DialogBarExContainer.h : header file
//

/////////////////////////////////////////////////////////////////////////////
//
// 当DialogBar在浮动时的包容窗口
//
////////////////////////////////////////////////////////////////////

class CDialogBarEx;

class CDialogBarExContainer : public CWnd
{
// Construction
public:
	CDialogBarExContainer();
	virtual ~CDialogBarExContainer();

public:

public:
	CDialogBarEx *m_pSubDialogBar;		//指向将包含在其中的DialogBar

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDialogBarExContainer)
	//}}AFX_VIRTUAL

public:
	CDialogBarEx* GetSubDialogBar();
	void SetSubDialogBar(CDialogBarEx *pDialogBar,CPoint pt);

	// Generated message map functions
protected:
	//{{AFX_MSG(CDialogBarExContainer)
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnMoving(UINT fwSide, LPRECT pRect);
	afx_msg void OnClose();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DIALOGBAREXCONTAINER_H__FA9796BD_78A2_4E7E_908D_66BCA6FD1FC2__INCLUDED_)
