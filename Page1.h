//{{AFX_INCLUDES()
#include "datagrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_PAGE1_H__FB88B9F5_5400_4D33_9676_8EEA1CF4201C__INCLUDED_)
#define AFX_PAGE1_H__FB88B9F5_5400_4D33_9676_8EEA1CF4201C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Page1.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CPage1 dialog
#include "datagridex.h"
class CPage1 : public CPropertyPage
{
	DECLARE_DYNCREATE(CPage1)

// Construction
public:
	void SaveSetting();
	short UnMouseWheel();
	bool m_a_colorDeleteState;
	virtual BOOL OnInitDialog();
	CPage1();
	~CPage1();

// Dialog Data
	//{{AFX_DATA(CPage1)
	enum { IDD = IDD_PROPPAG_1 };
	CComboBox	m_conComboBox;
	CDataGridEx	m_a_colr_lGrid;
	CDataGridEx	m_a_colorGrid;
	CString	m_comboStr;
	CDataGridEx	m_a_paiGrid;
	//}}AFX_DATA


// Overrides
	// ClassWizard generate virtual function overrides
	//{{AFX_VIRTUAL(CPage1)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	bool m_a_colr_lDeleteState;
	CString m_curColr_name;
	CString m_lastColr_name;
	CString m_lastColr_arrow;
	CString m_lastColr_ring;
	CString m_lastColr_face;
	void changeA_colorField(CString str);
	short m_lastCol;
	BOOL InitComboBox();
	bool InitA_pai();
	_RecordsetPtr m_a_paiSet;
	_RecordsetPtr m_a_colorSet;
	_RecordsetPtr m_a_colr_lSet;
	virtual void  OnCancel();
	// Generated message map functions
	//{{AFX_MSG(CPage1)
	afx_msg void OnKillfocusCombo1();
	afx_msg void OnMouseUpDatagridAColor(short Button, short Shift, long X, long Y);
	afx_msg void OnRowColChangeDatagridAColor(VARIANT FAR* LastRow, short LastCol);
	afx_msg void OnSelchangeCombo1();
	afx_msg void OnOnAddNewDatagridAColor();
	afx_msg void OnAfterUpdateDatagridAColrL();
	afx_msg void OnBeforeUpdateDatagridAColrL(short FAR* Cancel);
	afx_msg void OnScrollDatagridAColor(short FAR* Cancel);
	afx_msg void OnBeforeColEditDatagridAColrL(short ColIndex, short KeyAscii, short FAR* Cancel);
	afx_msg void OnBeforeDeleteDatagridAColrL(short FAR* Cancel);
	afx_msg void OnAfterDeleteDatagridAColrL();
	afx_msg void OnOnAddNewDatagridAPai();
	afx_msg void OnBeforeUpdateDatagridAColor(short FAR* Cancel);
	afx_msg void OnAfterDeleteDatagridAColor();
	afx_msg void OnBeforeDeleteDatagridAColor(short FAR* Cancel);
	afx_msg void OnMouseUpDatagridAPai(short Button, short Shift, long X, long Y);
	afx_msg void OnMouseUpDatagridAColrL(short Button, short Shift, long X, long Y);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PAGE1_H__FB88B9F5_5400_4D33_9676_8EEA1CF4201C__INCLUDED_)
