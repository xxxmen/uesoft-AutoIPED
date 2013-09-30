#if !defined(AFX_XDIALOG_H__3BEC6B97_655B_4248_90D6_1E3379D1DF5B__INCLUDED_)
#define AFX_XDIALOG_H__3BEC6B97_655B_4248_90D6_1E3379D1DF5B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// XDialog.h : header file
//

//文件: XDialog.h
//作者: 许朝
//时间: 2002.8

//对话框扩展类,支持自动窗口滚动

/////////////////////////////////////////////////////////////////////////////
// CXDialog dialog

class CXDialog : public CDialog
{
// Construction
public:
	CXDialog(UINT nIDTemplate,CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CXDialog)
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

// Operations
public:
	static AFX_DATA const SIZE sizeDefault;
		// used to specify default calculated page and line sizes

	// in logical units - call one of the following Set routines
	void SetScaleToFitSize(SIZE sizeTotal);
	void SetScrollSizes(int nMapMode, SIZE sizeTotal,
				const SIZE& sizePage = sizeDefault,
				const SIZE& sizeLine = sizeDefault);
	void ScrollToPosition(POINT pt);    // set upper left position
	void ResizeParentToFit(BOOL bShrinkOnly = TRUE);
	CSize GetTotalSize() const {ASSERT(this != NULL); return m_totalLog; };
	CPoint GetScrollPosition() const;       // upper corner of scrolling

	// for device units
	CPoint GetDeviceScrollPosition() const;
	void GetDeviceScrollSizes(int& nMapMode, SIZE& sizeTotal,
			SIZE& sizePage, SIZE& sizeLine) const;
// Implementation
protected:
	int m_nMapMode;
	CSize m_totalLog;           // total size in logical units (no rounding)
	CSize m_totalDev;           // total size in device units
	CSize m_pageDev;            // per page scroll size in device units
	CSize m_lineDev;            // per line scroll size in device units

	BOOL m_bCenter;             // Center output if larger than total size
	BOOL m_bInsideUpdate;       // internal state for OnSize callback
	void CenterOnPoint(CPoint ptCenter);
	void ScrollToDevicePosition(POINT ptDev); // explicit scrolling no checking
	void UpdateBars();          // adjust scrollbars etc
	BOOL GetTrueClientSize(CSize& size, CSize& sizeSb);
		// size with no bars
	void GetScrollBarSizes(CSize& sizeSb);
	void GetScrollBarState(CSize sizeClient, CSize& needSb,
		CSize& sizeRange, CPoint& ptMove, BOOL bInsideClient);
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CXDialog)
	public:
	virtual BOOL OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	BOOL m_bEnableScroll;
	BOOL m_bAccelToParentFrm;
	virtual BOOL OnScrollBy(CSize sizeScroll, BOOL bDoScroll = TRUE);
	virtual BOOL OnScroll(UINT nScrollCode, UINT nPos, BOOL bDoScroll=TRUE);

	// Generated message map functions
	//{{AFX_MSG(CXDialog)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_XDIALOG_H__3BEC6B97_655B_4248_90D6_1E3379D1DF5B__INCLUDED_)
