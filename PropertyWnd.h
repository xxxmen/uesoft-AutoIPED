#if !defined(AFX_PROPERTYWND_H__40E4A8DD_5CB7_4CE9_995E_9EAA16BB733D__INCLUDED_)
#define AFX_PROPERTYWND_H__40E4A8DD_5CB7_4CE9_995E_9EAA16BB733D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// PropertyWnd.h : header file
//
// 设置右单元的数据
// 返回单元被选中的通知

/////////////////////////////////////////////////////////////////////////////
//
// Write By 黄伟卫
// 2004,10,12 
//
// 属性窗口，用于属性的输入
//
// 窗口左边是属性的名称，右边可输入属性
//
///////////////////////////////////////////////////////////////////////////////////

#define PWN_FIRST	2000U

#define PWN_CONTROL	PWN_FIRST+1		// 包含各个子控件的通知的通知,此通知转发
									// 右单元的子控的通知,并有PWNControlStruct结构

#define PWN_SELECTCHANGE	PWN_FIRST+2	// 当改变所选的单元行时发送,
										// 并有PWN_SELECTCHANGE结构随

class CPropertyWnd : public CWnd
{
// Construction
public:
	CPropertyWnd();
	virtual ~CPropertyWnd();

public:
	enum{
		TitleElement	=0x01,	// 左列标题单元
		ChildElement	=0x02,	// 左列标题下的子单元
		EditElement		=0x04,	// 右列是Edit控件
		ComBoxElement	=0x08,	// 右列是ComBox控件
		DefineBackColor	=0x10,	// 指定重新设置左列单元的背景色
		DefineTextColor	=0x20,	// 指定重新设置左列单元的字体颜色
		DefineControlWnd=0x40	// 用户指定在右单元所用的控件(如果用户指定，控件的ID应为单元的行号或大于所有的单元)
	};

	//
	// 用于用户对单元的定义结构
	//
	struct ElementDef
	{
		CString szElementName;		// 左列单元的内容字符串
		UINT Style;					// 整个单元的式样
		int Width;					// 左单元的宽（以象素为单位）当为负数时自动计算
		int Height;					// 左单元的高（以象素为单位）当为负数时自动计算
		COLORREF BackGroundColor;	// 左单元的背影色
		COLORREF TextColor;			// 左单元的字体颜色
		LPVOID pVoid;				// 用于附加的在单元内的数据
		CWnd *pControlWnd;			// 对应于右单元中的子控件
		CString RightElementText;	// 右单元现在需显示的内容
		CString strFieldName;		// 该项对应的原始数据表的字段名


		//
		// 设置默认的值
		//
		ElementDef()
		{
			Style=ChildElement|EditElement;
			Width=-1;
			Height=-1;
			TextColor=RGB(0,0,0);
			BackGroundColor=RGB(255,255,255);
			pVoid=NULL;
			pControlWnd=NULL;
		}
	};

	//
	// 随PWN_CONTROL通知的结构
	//
	struct PWNControlStruct
	{
		NMHDR nmh;

		LPCTSTR szElementName;		// 左列单元的内容字符串
		UINT Style;					// 整个单元的式样
		int Width;					// 左单元的宽（以象素为单位）
		int Height;					// 左单元的高（以象素为单位）
		COLORREF BackGroundColor;	// 左单元的背影色
		COLORREF TextColor;			// 左单元的字体颜色
		LPVOID pVoid;				// 用于附加的在单元内的数据
		LPCTSTR szRightElementText;	// 右单元现在需显示的内容
		int Row;					// 此单元的列号

		int		idCtrl;				// 右单元控件的ID
		WORD	wNotifyCode;		// 右单元控件的事件号
		HWND	hWndControl;		// 右单元控件的窗口句柄
	};

	//
	// 随PWN_SELECTCHANGE发送的结构
	//
	struct PWNSelectChangeStruct
	{
		NMHDR nmh;

		LPCTSTR szElementName;		// 左列单元的内容字符串
		UINT Style;					// 整个单元的式样
		int Width;					// 左单元的宽（以象素为单位）
		int Height;					// 左单元的高（以象素为单位）
		COLORREF BackGroundColor;	// 左单元的背影色
		COLORREF TextColor;			// 左单元的字体颜色
		LPVOID pVoid;				// 用于附加的在单元内的数据
		LPCTSTR szRightElementText;	// 右单元现在需显示的内容
		int Row;					// 此单元的列号
	};

protected:
	//
	// 此结构与用户对单元的定义结构相似，用于程序内部
	//
	struct ElementDef_Wnd
	{
		CString ElementName;
		UINT Style;
		int Width;
		int Height;
		COLORREF BackGroundColor;
		COLORREF TextColor;
		LPVOID pVoid;
		CWnd *pControlWnd;			// 对应于右单元中的子控件
		CString RightElementText;	// 右单元现在需显示的内容
		CString strFieldName;		// 该项对应的原始数据表的字段名

		int Row;					// 此单元的列号，0开始
		RECT LeftColumeRect;		// 左单元的位置（以象素为单位）
		RECT RightColumeRect;		// 右单元的位置（以象素为单位）

		RECT BitmapColumeRect;		// 为最左边预留的用于图标的位置（以象素为单位)
		BOOL IsHideThisElement;		// 是否不显示此单元的标志


		ElementDef_Wnd()
		{
			Row=-1;
			pControlWnd=NULL;
			IsHideThisElement=FALSE;	//默认显示此单元
		}
	};
// Attributes
public:
	// 返回总的单元数
	int GetElementCount();

	// 创建单元的子控件
	BOOL CreateElementControl(int ElementNum);

	// 返回指定位置的单元
	int GetElementAt(ElementDef *pElement,int ElementNum);

	// 返回位图列的宽度
	int GetBitmapColumeWidth();
	// 设置为位图列的宽度
	void SetBitmapColumeWidth(int Width);

	// 返回整个窗口的背影色
	COLORREF GetWndBackGroundColor();
	// 设置整个窗口的背影色
	void SetWndBackGroundColor(COLORREF Color);

// Operations
public:
	void RefreshData();
	int FindElement(int nStartAfter,LPCTSTR szString,BOOL IsLeftElement=TRUE);
	// 向右单元添加数据
	void SetRightElementContent(int Row,LPCTSTR szContent);

	// 重新定义GetClientRect
	void GetClientRect(LPRECT lpRect);

	// 插入单元
	int InsertElement(ElementDef *pElement,int InsertAfter=0x7FFFFFFF);

protected:
	// 重新计算各个单元的位置坐标
	void UpdateElementsListRect();

	// 重新更新滚动条的信息
	void UpdateScrollBarInfo();

	// 获得可见的列数
	int GetVisibleRow();

	// 返回指定列的上一个可见的列号
	int GetPrevVisibleRow(int row);

	// 返回指定列的下一个可见的列号
	int GetNextVisibleRow(int row);

	// 返回指定位置的单元
	ElementDef_Wnd* GetElementAt(int ElementNum);

private:
	// 绘造单元
	void PaintElement(CDC *pDC);

	// 对右单元的点击进行测试
	void RightElementsHitTest(CPoint point);

	// 对左单元的点击进行测试
	void LeftElementsHitTest(CPoint point);

	// 返回位图列的宽度
	void BitmapElementsHitTest(CPoint point);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPropertyWnd)
	protected:
	virtual BOOL OnCommand(WPARAM wParam, LPARAM lParam);
	//}}AFX_VIRTUAL

// Implementation

protected:
	int m_ElementCount;			// 总共包含单元格的数目
	int m_CurrentScrollPos;		// 当前滚动条的位置
	int m_BitmapColumeWidth;	// 设置位图列的宽度

private:
	CPtrList	m_ElementList;
	COLORREF	m_BitmapColumeColor;	// 位图列背景色
	CBrush		m_BitmapColumeBrush;	// 位图列背景色
	COLORREF	m_FrameColor;			// 绘制边框的颜色
	CPen		m_FramePen;				// 绘制边框的画笔
	COLORREF	m_WndBackGroundColor;	// 整个窗口的背景色
	CBrush		m_WndBackGroundBrush;	// 整个窗口的背影画刷
	BOOL		m_IsNeedReUpdateElementsListRect;	//是否需要重新计算各个单元的位置
	int			m_SelectRow;			// 已经选择的单元行
	COLORREF	m_SelectElementBackGroundColorBank;	// 用于保存已经被选中的单元原来的背景色
	HCURSOR		m_hDragCursor;			// 可用于拖动的鼠标图标句柄
	BOOL		m_IsInChangeElementRect;	//是否在可改变单元大小的区域
	CPoint		m_MouseMovePt;			// 用于记录鼠标滚动先前的位置
	CBitmap		m_BitmapSurface;		// 用于避免闪烁的位图

	CFont		m_Font;	//文字的字体

	// Generated message map functions
protected:
	void SendPWNSelectChangeNotify();
	void ChangeLeftElementWidth(int Width);
	//{{AFX_MSG(CPropertyWnd)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnPaint();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROPERTYWND_H__40E4A8DD_5CB7_4CE9_995E_9EAA16BB733D__INCLUDED_)
