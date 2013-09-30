#if !defined(AFX_RANGEDLG_H__4D924B6F_5869_407C_A1D1_E0EF9D6EA38E__INCLUDED_)
#define AFX_RANGEDLG_H__4D924B6F_5869_407C_A1D1_E0EF9D6EA38E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// RangeDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// RangeDlg dialog

class RangeDlg : public CDialog
{
// Construction
public:
	RangeDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(RangeDlg)
	enum { IDD = IDD_DIALOG_Range };
	CSpinButtonCtrl	c_start;
	CSpinButtonCtrl	c_end;
	long	m_start;
	long	m_end;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(RangeDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(RangeDlg)
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_RANGEDLG_H__4D924B6F_5869_407C_A1D1_E0EF9D6EA38E__INCLUDED_)
