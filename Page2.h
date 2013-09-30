//{{AFX_INCLUDES()
#include "datagrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_PAGE2_H__17F6AECC_AFF9_41D0_B0F6_514D207CB2DD__INCLUDED_)
#define AFX_PAGE2_H__17F6AECC_AFF9_41D0_B0F6_514D207CB2DD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Page2.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CPage2 dialog
#include "datagridex.h"

class CPage2 : public CPropertyPage
{
	DECLARE_DYNCREATE(CPage2)

// Construction
public:
	CPage2();
	~CPage2();

// Dialog Data
	//{{AFX_DATA(CPage2)
	enum { IDD = IDD_PROPPAG_2 };
	CDataGridEx	m_ctlArrow;
	CDataGridEx	m_ctlRing_Paint;
	//}}AFX_DATA


// Overrides
	// ClassWizard generate virtual function overrides
	//{{AFX_VIRTUAL(CPage2)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CPage2)
	virtual BOOL OnInitDialog();
	afx_msg void OnDestroy();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

private:
	_ConnectionPtr m_pConPrj;
	_ConnectionPtr m_pConCODE;
	_ConnectionPtr m_pConSort;
	BOOL UpdateA_RING_PAINT();
	BOOL UpdateA_ARROW();
	
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PAGE2_H__17F6AECC_AFF9_41D0_B0F6_514D207CB2DD__INCLUDED_)
