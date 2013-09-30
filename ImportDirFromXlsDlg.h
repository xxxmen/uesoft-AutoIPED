#if !defined(AFX_IMPORTDIRFROMXLSDLG_H__37CC9757_949D_4FBB_AC34_29D253A06473__INCLUDED_)
#define AFX_IMPORTDIRFROMXLSDLG_H__37CC9757_949D_4FBB_AC34_29D253A06473__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ImportDirFromXlsDlg.h : header file
//

#include "ImportFromXLSDlg.h"

/////////////////////////////////////////////////////////////
//
// 以CImportFromXLSDlg为父类从外部EXCEL文件导入数据到
// PRE_CALC表的对话框
//
//////////////////////////////////////////////////////////////

class CImportDirFromXlsDlg : public CImportFromXLSDlg
{
// Construction
public:
	CImportDirFromXlsDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CImportDirFromXlsDlg)
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CImportDirFromXlsDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual void BeginImport();
	virtual BOOL ValidateData();
	virtual BOOL InitPropertyWnd();

	// Generated message map functions
	//{{AFX_MSG(CImportDirFromXlsDlg)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_IMPORTDIRFROMXLSDLG_H__37CC9757_949D_4FBB_AC34_29D253A06473__INCLUDED_)
