#if !defined(AFX_DATAGRIDEX_H__DCD2258E_FC6A_4C5E_B83C_CC66B285D5B8__INCLUDED_)
#define AFX_DATAGRIDEX_H__DCD2258E_FC6A_4C5E_B83C_CC66B285D5B8__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DataGridEx.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// 
// 对DataGrid控件进行扩张，加入水平与垂直滚动的立即显示，并加入提示的信息
// 实现鼠标的中间滚动球的滚动。
//
// 要鼠标的中间滚动球的滚动功能需要调用InitMouseWheel
//
///////////////////////////////////////////////////////////////////////////////

#include "datagrid.h"

#define WM_MOUSEWHEEL2 WM_USER+32

#define DATAGRIDEX_FIRSTID	30000

#define IDC_GRID_ASC	DATAGRIDEX_FIRSTID+1	// 升序菜当ID
#define IDC_GRID_DESC	DATAGRIDEX_FIRSTID+2	// 降序菜当ID
#define IDC_GRID_HIDE	DATAGRIDEX_FIRSTID+3	// 隐藏字段的ID 
#define IDC_GRID_CANCEL_HIDE	DATAGRIDEX_FIRSTID+4	// 取消隐藏字段的ID 


class CDataGridEx : public CDataGrid
{
// Construction
public:
	CDataGridEx();
	virtual ~CDataGridEx();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDataGridEx)
	//}}AFX_VIRTUAL

// Implementation
public:
	LRESULT OnSortDesc();
	LRESULT OnSortAsc();
	LRESULT OnHideField();
	LRESULT OnCancelHideField();

	// 钩子的会调函数
	static LRESULT CALLBACK GetMsgProc(int code,WPARAM wParam,LPARAM lParam);

	// 初始化中间滚动球的滚动响应
	void InitMouseWheel();
	// 释放钩子资源，应在窗口销毁前调用
	void UnInitMouseWheel();


	// Generated message map functions
protected:
	BOOL OnNotify(WPARAM wParam,LPARAM lParam,LRESULT *pResult);
	//{{AFX_MSG(CDataGridEx)
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	//}}AFX_MSG
	afx_msg LRESULT OnMouseWheel2(WPARAM wParam,LPARAM lParam);
	DECLARE_MESSAGE_MAP()

public:

	struct tagHookCountInfo
	{
		HHOOK hHook;
		DWORD Count;
	};

	//钩子的句柄Map,不同的线程的钩子不同
	static CMap<DWORD,DWORD,tagHookCountInfo*,tagHookCountInfo*> m_MouseThreadHookMap;

private:
	//放置ToolTip控件在合适显示的位置
	void PlaceToolTipWndToRightPosition();

	BOOL m_PreVScrollFlag;
	UINT m_PreVScrollPos;
	short m_SelectCol;

	int m_RecordCount;			//记录集的总数
	TCHAR m_ToolTipInfo[256];	//存放TOOLTIP控件的提示信息
	void InitToolTipControl();	//初始化TOOLTIP控件
	CToolTipCtrl m_ToolTip;		//TOOLTIP控件
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DATAGRIDEX_H__DCD2258E_FC6A_4C5E_B83C_CC66B285D5B8__INCLUDED_)
