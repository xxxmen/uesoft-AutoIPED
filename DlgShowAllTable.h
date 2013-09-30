#if !defined(AFX_DLGSHOWALLTABLE_H__7675EB0E_DE59_4B0F_A54F_FD32C258ECF5__INCLUDED_)
#define AFX_DLGSHOWALLTABLE_H__7675EB0E_DE59_4B0F_A54F_FD32C258ECF5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgShowAllTable.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDlgShowAllTable dialog
#include "DataShowDlg.h"
#include "mybutton.h"
class CDlgShowAllTable : public CDataShowDlg
{
// Construction
public:
	bool UpdateTblNameList();	//初始可显示的表。
	BOOL UpdateControlsPos();  // 更新各个控件的位置

	CDlgShowAllTable(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgShowAllTable)
//	enum { IDD = IDD_REPORT_CAL_ERROR_DLG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgShowAllTable)
	protected:
//	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CString m_strDescTblName;	//当前选择表的中文名称
	void UpdateDataGrid();		// 更新所选表的内容。
	CListBox m_TableNameList;
	CMyButton	 m_ctlDeleteAll;	//删除选择的表中的全部记录.
	// Generated message map functions
	//{{AFX_MSG(CDlgShowAllTable)
	virtual BOOL OnInitDialog();
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnPaint();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	//}}AFX_MSG
	afx_msg void OnTableSelChange(); // 响应不同的表的选择改变
	afx_msg void OnDeleteTblAll();	//删除表中的所有记录

	DECLARE_MESSAGE_MAP()
private:
	BOOL tableExists(_ConnectionPtr pCon, CString tbn);
	long m_nCurTblID;
	CSize			m_ListSize;	//材料名称列表的尺寸

	CRect	m_VDragArea;	// 竖向拖动的区域
	CRect	m_HDragArea;	// 横向拖动的区域
	bool	m_IsInVDragArea;	// 是否在竖向拖动的区域
	bool	m_IsInHDragArea;	// 是否在横向拖动的区域
	CPoint	m_LButtonPt;		// 用于记录左鼠标位置

	_RecordsetPtr m_pRsTableList;  //用于记录表的对应关系.
//	_RecordsetPtr m_pRsBrowseResult;	//对应浏览结果的记录集.
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGSHOWALLTABLE_H__7675EB0E_DE59_4B0F_A54F_FD32C258ECF5__INCLUDED_)
