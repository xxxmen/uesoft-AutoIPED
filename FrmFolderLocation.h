#if !defined(AFX_FRMFOLDERLOCATION_H__65B28A1D_C66B_4AF9_8593_9CD93A4D535A__INCLUDED_)
#define AFX_FRMFOLDERLOCATION_H__65B28A1D_C66B_4AF9_8593_9CD93A4D535A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FrmFolderLocation.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CFrmFolderLocation dialog

class CFrmFolderLocation : public CDialog
{
// Construction
public:
	bool IsDirExists(CString Directory);
	void OnSf(UINT uID);
	void SetRegValue(LPCTSTR pszKey, LPCTSTR pszName, const CString vValue);
	CString GetRegKey(LPCTSTR pszKey, LPCTSTR pszName,const CString Default);
	CString GetDir(CString key);
	void InitgstrEDIBdir();
	void SetWindowCenter(HWND hWnd);
	void LoadCaption();
//	CString GetRegKey(LPCTSTR pszKey, LPCTSTR pszName,const CString Default);

	CFrmFolderLocation(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CFrmFolderLocation)
	enum { IDD = IDD_FOLDER_LOCATION };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFrmFolderLocation)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CFrmFolderLocation)
	virtual BOOL OnInitDialog();
	afx_msg void OnSf1();
	afx_msg void OnSf2();
	afx_msg void OnSf3();
	afx_msg void OnSf4();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	CString m_strFormDir[4];
};


extern CString gstrEDIBdir[4];
//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.


#endif // !defined(AFX_FRMFOLDERLOCATION_H__65B28A1D_C66B_4AF9_8593_9CD93A4D535A__INCLUDED_)
