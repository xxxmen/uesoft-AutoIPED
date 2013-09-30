#if !defined(AFX_DLGAUTOTOTAL_H__D2086148_679E_4E80_AC65_C0B80D2E26AC__INCLUDED_)
#define AFX_DLGAUTOTOTAL_H__D2086148_679E_4E80_AC65_C0B80D2E26AC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgAutoTotal.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDlgAutoTotal dialog

class CDlgAutoTotal : public CDialog
{
// Construction
public:
	bool m_autoTotal[15];
	CDlgAutoTotal(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgAutoTotal)
	enum { IDD = IDD_DIALOG_AUTO_TOTAL };
	CButton	m_cTemperatureExplain;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgAutoTotal)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual BOOL OnInitDialog();

	// Generated message map functions
	//{{AFX_MSG(CDlgAutoTotal)
	afx_msg void OnTemperatureExplain();
	afx_msg void OnPaintExplain();
	afx_msg void OnTemperatureTotal1();
	afx_msg void OnPaintTotal();
	afx_msg void OnTempratureTotal2();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGAUTOTOTAL_H__D2086148_679E_4E80_AC65_C0B80D2E26AC__INCLUDED_)
