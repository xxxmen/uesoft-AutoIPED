#if !defined(AFX_DLGSELCRITDB_H__674FA1A7_4088_4792_8F2B_E5E5B203F96E__INCLUDED_)
#define AFX_DLGSELCRITDB_H__674FA1A7_4088_4792_8F2B_E5E5B203F96E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgSelCritDB.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDlgSelCritDB dialog

class CDlgSelCritDB : public CDialog
{
// Construction
public:
	static short GetMaterialDBName(CString strMaterialDBName[] );
	static short GetCriterionDBName(CString strCritDBName[] );
	CDlgSelCritDB(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgSelCritDB)
	enum { IDD = IDD_DLG_SEL_CRITERION_DB };
	CListCtrl	m_ctrlShowCODE;
	CComboBox	m_ctlMaterCodeName;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgSelCritDB)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CDlgSelCritDB)
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	afx_msg void OnClickList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnColumnclickList1(NMHDR* pNMHDR, LRESULT* pResult);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	struct  KeyToCalling{
		long strKey;
		CString strCallingName;
		CString strCodeName;
	};

	int GetCallingFromDB(struct KeyToCalling sutKeyCalling[]);
	
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGSELCRITDB_H__674FA1A7_4088_4792_8F2B_E5E5B203F96E__INCLUDED_)
