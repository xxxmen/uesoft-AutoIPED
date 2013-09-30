#if !defined(AFX_DIALOG_PFORM_H__83773F90_2E08_4C25_B355_712F5CE235DD__INCLUDED_)
#define AFX_DIALOG_PFORM_H__83773F90_2E08_4C25_B355_712F5CE235DD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Dialog_pform.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// Dialog_pform dialog

class Dialog_pform : public CDialog
{
// Construction
public:
	~Dialog_pform();
	CString * engin_name;
	CString *engin;
	int i;
	int cancel;
	CString old1;
	CString m_pform1;
	int index;
	CString old;
	CString m_pform;
	CString m_strSql;
	CString m_strcolumn;
	
	Dialog_pform(CWnd* pParent = NULL);   // standard constructor


	_RecordsetPtr	m_pRecordset;
// Dialog Data
	//{{AFX_DATA(Dialog_pform)
	enum { IDD = IDD_DIALOG_pform };
	CListCtrl	c_list;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(Dialog_pform)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(Dialog_pform)
	virtual BOOL OnInitDialog();
	afx_msg void OnBUTTONselect();
	afx_msg void OnClickList(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnBUTTONcancel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DIALOG_PFORM_H__83773F90_2E08_4C25_B355_712F5CE235DD__INCLUDED_)
