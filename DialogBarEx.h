#if !defined(AFX_DIALOGBAREX_H__A885A54C_9195_48D4_AEA5_71CF5BFB7E24__INCLUDED_)
#define AFX_DIALOGBAREX_H__A885A54C_9195_48D4_AEA5_71CF5BFB7E24__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DialogBarEx.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// 
// 对CDialogBar的扩张
//
// 可以在停靠和浮动时改变大小
//
////////////////////////////////////////////////////////////////////////////
#include "DialogBarExContainer.h"

class CDialogBarEx : public CDialogBar
{
	friend class CDialogBarExContainer;
// Construction
public:
	CDialogBarEx();
	virtual ~CDialogBarEx();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDialogBarEx)
	//}}AFX_VIRTUAL

// Implementation
public:
	void GetClientRect(LPRECT lpRect);

	bool IsWindowVisible();	// 判断窗口是否可见
	void HideDlgBar();		// 隐藏窗口
	void ShowDlgBar();		// 显示窗口

	void GetFloatingSize(CSize &size);// 获得浮动时的窗口的尺寸

	CSize GetWndMinSize();			// 返回窗口的最小尺寸
	void SetWndMinSize(CSize size);	// 设置窗口的最小尺寸

	int GetDockingSide();			// 返回停靠的位置

protected:
	// 绘制标题栏
	virtual void PaintTitle(CDC *pDC);

	virtual CSize CalcFixedLayout(BOOL bStretch,BOOL bHorz);
	virtual CSize CalcDynamicLayout(int nLength,DWORD dwMode);


protected:
	CSize		m_InitSize;		//初始化的窗口的大小
	CSize		m_FloatingSize;	//浮动时窗口的大小
	CRect		m_TitleRect;	//标题栏的大小
	COLORREF	m_TitleBarColor;	// 标题栏的颜色
	CBrush		m_TitleBarBrush;	// 标题栏颜色的画刷


private:
	CPoint	m_CapturePt;	// 扑捉时的鼠标位置
	BOOL	m_IsCapture;	// 是否正在扑捉
	BOOL	m_bHorz;		// 是否是水平的

	// 是否在可改变大小的区域，只在停靠时有用,
	// 0,不可改变大小，1，上，2下，3左，4，右
	int		m_DragSizeArea;	
	CSize	m_MinSize;		// 窗口的最小尺寸（象素为单位）

	CDialogBarExContainer	*m_pContainWnd;		//用于包含浮动时的窗口

	CRect	m_CloseButtonRect;	//关闭窗口按钮的位置
	bool	m_IsInCloseButton;	// 是否在关闭窗口内
	bool	m_IsCursorInWnd;	// 鼠标是否在窗口内
	HCURSOR	m_hCursorH;
	HCURSOR	m_hCursorV;
protected:
	//{{AFX_MSG(CDialogBarEx)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnPaint();
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	//}}AFX_MSG
	afx_msg	LRESULT OnMouseLeave(WPARAM wParam,LPARAM lParam);
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DIALOGBAREX_H__A885A54C_9195_48D4_AEA5_71CF5BFB7E24__INCLUDED_)
