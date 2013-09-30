#if !defined(AFX_CALC_H__04F0B96C_F454_47F3_96B5_DCEB9A880062__INCLUDED_)
#define AFX_CALC_H__04F0B96C_F454_47F3_96B5_DCEB9A880062__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// calc.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// calc dialog

class calc : public CDialog
{
// Construction
public:
	int i;
	calc(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(calc)
	enum { IDD = IDD_DIALOG_calc };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(calc)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(calc)
	afx_msg void OnBUTTONout();
	afx_msg void OnBUTTONcountall();
	afx_msg void OnBUTTONrange();
	afx_msg void OnBUTTONcustom();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CALC_H__04F0B96C_F454_47F3_96B5_DCEB9A880062__INCLUDED_)
