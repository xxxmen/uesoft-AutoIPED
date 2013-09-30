#if !defined(AFX_NEWPROJECT_H__518B0954_0D7A_4694_B746_B4DDA0AA7362__INCLUDED_)
#define AFX_NEWPROJECT_H__518B0954_0D7A_4694_B746_B4DDA0AA7362__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// newproject.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// newproject dialog

class newproject : public CDialog
{
// Construction
public:
	bool SetAuto( bool bAuto );
	CString m_strCaption;
	void AddToValumeOnApply();
	newproject(CWnd* pParent = NULL);   // standard constructor
	BOOL insertitem();

// Dialog Data
	//{{AFX_DATA(newproject)
	enum { IDD = IDD_DIALOG_newproject };
	CProgressCtrl	c_progress;
	CListCtrl	c_list;
	CString	m_eng_code;
	CString	m_engin;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(newproject)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	bool GetImportTblNames(CString* &pTableName,int &nTblCount,CString pNotTbl[],int nNotCount=0);
	void AddCurrentVolume();
	void AddToValumeOnNew();

	// Generated message map functions
	//{{AFX_MSG(newproject)
	virtual BOOL OnInitDialog();
	afx_msg void OnBUTTONnew();
	afx_msg void OnClickList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnAddnewproject();
	afx_msg void OnDeleteSelectproject();
	afx_msg void OnSetfocusEDITengcode();
	afx_msg void OnSetfocusEDITengin();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
	BOOL addto_volume();
	BOOL addto_table(CString tablename);

private:
	// 从删除指定的工程
	BOOL DeleteProjectInTable(LPCTSTR TableName,LPCTSTR ProjectID);

	// 更新各个控件的状态
	void OnUpdateState();

	// 返回Volume表中的下一个可用的VolumeID号
	long GetNextVolumeID();
	BOOL tableExists(_ConnectionPtr pCon, CString tbn);
private:
	////自动新建工程.设置状态.
	bool m_bAuto;

	CString m_str4;
	int newpro;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NEWPROJECT_H__518B0954_0D7A_4694_B746_B4DDA0AA7362__INCLUDED_)
