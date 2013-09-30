#if !defined(AFX_DLGOPTION_H__0ACABC94_E73B_4735_B61A_DEA5C86DC98E__INCLUDED_)
#define AFX_DLGOPTION_H__0ACABC94_E73B_4735_B61A_DEA5C86DC98E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgOption.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDlgOption dialog

class CDlgOption : public CDialog
{
// Construction
public:
	CDlgOption(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgOption)
	enum { IDD = IDD_DIALOG_OPTION };
	CProgressCtrl	m_ReplaceProgress;
	
private:
	BOOL	m_IsColseCompress;
	BOOL	m_IsReplOld;
	BOOL	m_IsMoveUpdate;
	BOOL	m_IsAutoSelectPre;
	BOOL	m_bIsHeatLoss;
	BOOL	m_IsAutoAddValve;
	double	m_dMaxRatio;
	double	m_dSurfaceMaxTemp;
	BOOL	m_bIsRunSelectEng;
	BOOL	m_bIsSeleTbl;
	BOOL	m_bIsReplaceUnit;
	BOOL	m_bAutoPaint120;
	BOOL	m_bWithoutProtectionCost;
	BOOL	m_bInnerByEconomic;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgOption)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CDlgOption)
	virtual void OnOK();
	afx_msg void OnReplaceOldnameNewname();
	afx_msg void OnReplaceCurrentOldnameNewname();
	afx_msg void OnRadioHandSet();
	afx_msg void OnRadioRenewSet();
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGOPTION_H__0ACABC94_E73B_4735_B61A_DEA5C86DC98E__INCLUDED_)
