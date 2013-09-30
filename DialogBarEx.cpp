// DialogBarEx.cpp : implementation file
//

#include "stdafx.h"
#include "DialogBarEx.h"
#include "winuser.h"
#include "resource.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CDialogBarEx::CDialogBarEx()
{
	m_TitleBarColor=RGB(0,0,255);	//设置标题栏颜色
	m_TitleBarBrush.CreateSolidBrush(m_TitleBarColor);	//创建标题栏颜色的画刷

	m_IsCapture=FALSE;	// 设置没有正在扑捉

	m_DragSizeArea=0;	// 设置默认不在可改变大小的区域

	m_MinSize.cx=30;	// 设置窗口最小的尺寸
	m_MinSize.cy=30;

	m_pContainWnd=new CDialogBarExContainer;	//创建窗口对象

	m_IsInCloseButton=false;	// 默认不在关闭窗口内

	m_IsCursorInWnd=false;		// 默认鼠标不在窗口内

	m_hCursorV = AfxGetApp()->LoadCursor(IDC_SPLITV);
	m_hCursorH = AfxGetApp()->LoadCursor(IDC_SPLITH);

}

CDialogBarEx::~CDialogBarEx()
{
	delete m_pContainWnd;	// 销毁窗口对象
}


BEGIN_MESSAGE_MAP(CDialogBarEx, CDialogBar)
	//{{AFX_MSG_MAP(CDialogBarEx)
	ON_WM_CREATE()
	ON_WM_LBUTTONDOWN()
	ON_WM_MOUSEMOVE()
	ON_WM_PAINT()
	ON_WM_LBUTTONUP()
	ON_WM_LBUTTONDBLCLK()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_MOUSELEAVE,OnMouseLeave)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CDialogBarEx message handlers

CSize CDialogBarEx::CalcDynamicLayout(int nLength, DWORD dwMode)
{
	return CDialogBar::CalcDynamicLayout(nLength,dwMode);
}

CSize CDialogBarEx::CalcFixedLayout(BOOL /*bStretch*/, BOOL bHorz)
{
	m_bHorz=bHorz;

	if(m_InitSize.cx<this->m_MinSize.cx)
		m_InitSize.cx=m_MinSize.cx;

	if(m_InitSize.cy<this->m_MinSize.cy)
		m_InitSize.cy=m_MinSize.cy;

	if(this->IsFloating())
	{
		return m_FloatingSize;
	}
	else
	{
		//停靠
		CRect FrameRect;
		this->GetDockingFrame()->GetClientRect(&FrameRect);

		if(bHorz)	//水平
		{
			this->m_TitleRect.SetRect(0,0,FrameRect.Width(),15);

			this->m_CloseButtonRect=m_TitleRect;
			this->m_CloseButtonRect.left=this->m_CloseButtonRect.right-m_TitleRect.Height();
			m_CloseButtonRect.OffsetRect(-10,0);

			return CSize(FrameRect.Width(),m_InitSize.cy);
		}
		else
		{
			this->m_TitleRect.SetRect(0,0,m_InitSize.cx,15);

			this->m_CloseButtonRect=m_TitleRect;
			this->m_CloseButtonRect.left=this->m_CloseButtonRect.right-m_TitleRect.Height();
			m_CloseButtonRect.OffsetRect(-10,0);

			return CSize(m_InitSize.cx,FrameRect.Height());
		}

		this->m_CloseButtonRect=m_TitleRect;
		m_CloseButtonRect.OffsetRect(-10,0);
	}

}

int CDialogBarEx::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CDialogBar::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	//
	// 初始化时浮动的窗口大小与初始窗口的大小相等
	//
	m_InitSize.cx=lpCreateStruct->cx;
	m_InitSize.cy=lpCreateStruct->cy;

	m_FloatingSize=m_InitSize;

	WNDCLASS WndClass;
	CString strText;

	WndClass.cbClsExtra=0;
	WndClass.cbWndExtra=0;
	WndClass.hbrBackground=NULL;
	WndClass.hCursor=NULL;
	WndClass.hIcon=NULL;
	WndClass.hInstance=NULL;
	WndClass.lpfnWndProc=::DefWindowProc;
	WndClass.lpszClassName=_T("DialogBarExContainer");
	WndClass.lpszMenuName=NULL;
	WndClass.style=0;

	::AfxRegisterClass(&WndClass);
	
	GetWindowText(strText);

	m_pContainWnd->CreateEx(WS_EX_TOOLWINDOW,_T("DialogBarExContainer"),strText,
							WS_OVERLAPPEDWINDOW,CRect(0,0,300,300),this->GetDockingFrame(),NULL);

	return 0;
}

void CDialogBarEx::OnLButtonDown(UINT nFlags, CPoint point) 
{
	//
	// 如果在关闭按钮内点击就隐藏窗口
	//
	if(this->m_IsInCloseButton)
	{
		this->HideDlgBar();
		return;
	}

	if(!IsFloating())
	{
		if(this->m_DragSizeArea!=0)
		{
			this->SetCapture();
			m_IsCapture=TRUE;
			m_CapturePt=point;
			return;
		}

		//是否在标题栏内
		if(m_TitleRect.PtInRect(point))
		{
			this->SetCapture();
			m_IsCapture=TRUE;
			m_CapturePt=point;
		}
	}
}

void CDialogBarEx::OnMouseMove(UINT nFlags, CPoint point) 
{
	//
	// 如果鼠标为在关闭按钮区则改变m_IsCursorInWnd并且启动
	// 监视鼠标是否离开窗口的事件
	//
	if(m_IsCursorInWnd==false)
	{
		TRACKMOUSEEVENT TraceStruct;

		TraceStruct.cbSize=sizeof(TRACKMOUSEEVENT);
		TraceStruct.dwFlags=TME_LEAVE;
		TraceStruct.dwHoverTime=HOVER_DEFAULT;
		TraceStruct.hwndTrack=this->GetSafeHwnd();

		TrackMouseEvent(&TraceStruct);

		m_IsCursorInWnd=true;
	}

	//停靠的并前没有被扑获
	if(!IsFloating() && !m_IsCapture)
	{
		CRect TopBorderRect,BottomBorderRect;
		CRect LeftBorderRect,RightBorderRect;
//		HCURSOR hCursor;

		CWnd::GetClientRect(&TopBorderRect);
		CWnd::GetClientRect(&BottomBorderRect);
		CWnd::GetClientRect(&LeftBorderRect);
		CWnd::GetClientRect(&RightBorderRect);

		TopBorderRect.top=TopBorderRect.bottom-8;
		BottomBorderRect.bottom=8;
		LeftBorderRect.left=LeftBorderRect.right-8;
		RightBorderRect.right=8;

		//如果停靠在下边并且在可改变大小的区域内
		if(BottomBorderRect.PtInRect(point) && this->GetDockingSide()==4)
		{
			this->m_DragSizeArea=1;
//			hCursor=AfxGetApp()->LoadCursor(IDC_SPLITV);
//			::SetCursor(hCursor);
//			::DeleteObject(hCursor);
			::SetCursor(m_hCursorV);
		}
		//如果停靠在上边并且在可改变大小的区域内
		else if(TopBorderRect.PtInRect(point) && this->GetDockingSide()==3)
		{
			this->m_DragSizeArea=2;
//			hCursor=AfxGetApp()->LoadCursor(IDC_SPLITV);
//			::SetCursor(hCursor);
//			::DeleteObject(hCursor);
			::SetCursor(m_hCursorV);

		}
		//如果停靠在左边并且在可改变大小的区域内 
		else if(LeftBorderRect.PtInRect(point) && this->GetDockingSide()==1)
		{
			this->m_DragSizeArea=4;
//			hCursor=AfxGetApp()->LoadCursor(IDC_SPLITH);
//			::SetCursor(hCursor);
//			::DeleteObject(hCursor);
			::SetCursor(m_hCursorH);

		}
		//如果停靠在右边并且在可改变大小的区域内 
		else if(RightBorderRect.PtInRect(point) && this->GetDockingSide()==2)
		{
			this->m_DragSizeArea=3;
//			hCursor=AfxGetApp()->LoadCursor(IDC_SPLITH);
//			::SetCursor(hCursor); 
//			::DeleteObject(hCursor);
			::SetCursor(m_hCursorH);
		}
		else
		{
			this->m_DragSizeArea=0;
		}
	}

	/////////////////////////////////////////////////
	//
	// 在标题拦类点击了左键
	//
	if(m_IsCapture && this->m_DragSizeArea==0)
	{
		if(abs(point.y-m_CapturePt.y)>35 || abs(point.x-m_CapturePt.x)>35)
		{
			CString strText;

			this->GetWindowText(strText);

			this->ClientToScreen(&point);
			this->m_pContainWnd->SetSubDialogBar(this,point);

			this->m_pContainWnd->SetWindowText(strText);
		}
	}

	if(m_IsCapture && this->m_DragSizeArea!=0)
	{
		CPoint pt1,pt2;

		//可以向上拖动从而改变大小
		if(this->m_DragSizeArea==1)
		{
			this->m_InitSize.cy+=m_CapturePt.y-point.y;

			::SetCursor(m_hCursorV);
		}
		//可以向下拖动从而改变大小
		else if(this->m_DragSizeArea==2)
		{
			pt1=m_CapturePt;
			pt2=point;

			this->ClientToScreen(&pt1);
			this->ClientToScreen(&pt2);

			this->m_InitSize.cy+=pt2.y-pt1.y;

			m_CapturePt=point;

			::SetCursor(m_hCursorV);			
		}
		//可以向右拖动从而改变大小
		else if(this->m_DragSizeArea==4)
		{
			pt1=m_CapturePt;
			pt2=point;

			this->ClientToScreen(&pt1);
			this->ClientToScreen(&pt2);

			this->m_InitSize.cx+=pt2.x-pt1.x;

			m_CapturePt=point;

			::SetCursor(m_hCursorH);			
		}
		//可以向左拖动从而改变大小
		else if(this->m_DragSizeArea==3)
		{
			this->m_InitSize.cx+=m_CapturePt.x-point.x;

			::SetCursor(m_hCursorH);
		}

		GetDockingFrame()->RecalcLayout();
		this->Invalidate();
	}

	//
	// 判断是否在关闭按钮内，如果状态改变重绘窗口
	//
	if(this->m_CloseButtonRect.PtInRect(point))
	{
		if(this->m_IsInCloseButton==false)
		{
			this->Invalidate();
		}
		HCURSOR hCursor2;

		this->m_IsInCloseButton=true;
		hCursor2 = LoadCursor(NULL, IDC_ARROW);
		::SetCursor(hCursor2);
		::DeleteObject(hCursor2);
	}
	else
	{
		if(this->m_IsInCloseButton==true)
		{
			this->Invalidate();
		}

		this->m_IsInCloseButton=false;
	}

}

void CDialogBarEx::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if(m_IsCapture)
	{
		ReleaseCapture();
		m_IsCapture=FALSE;
		this->m_DragSizeArea=0;
	}


}

void CDialogBarEx::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	
	//在停靠时才绘制
	if(!this->IsFloating())
		PaintTitle(&dc);

	// Do not call CDialogBar::OnPaint() for painting messages
}

void CDialogBarEx::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	
//	CDialogBar::OnLButtonDblClk(nFlags, point);
}

///////////////////////////////////////////////////////////


//////////////////////////////////////////
//
// 绘制标题栏
//
// pDC[in]	设备上下文
//
void CDialogBarEx::PaintTitle(CDC *pDC)
{
	CString strText;
	CRect	TextRect;
	COLORREF	OldColor;
	int			OldBkMode;
	CPen	Pen,*OldPen;
	CFont * pFont;

	CRect   gripper;
	gripper = m_TitleRect;
	gripper.left += 50;
	gripper.right -= 30;
	gripper.top += 6;
	gripper.bottom = gripper.top + 3;
	pDC->Draw3dRect(&gripper, ::GetSysColor(COLOR_BTNHIGHLIGHT),
        ::GetSysColor(COLOR_BTNSHADOW));
	gripper.OffsetRect(0, 3);
	pDC->Draw3dRect(&gripper, ::GetSysColor(COLOR_BTNHIGHLIGHT),
        ::GetSysColor(COLOR_BTNSHADOW));

	//刷标题栏 
//	pDC->FillRect(&m_TitleRect,&m_TitleBarBrush);

	//保存先前设置
	OldBkMode=pDC->SetBkMode(TRANSPARENT);
	OldColor=pDC->SetTextColor(RGB(0,0,0));
	pFont = pDC->SelectObject(GetFont());

	this->GetWindowText(strText);
	TextRect=m_TitleRect;

	TextRect.left+=3;
	TextRect.OffsetRect(0, 3);
	pDC->DrawText(strText,&TextRect,DT_LEFT);

	// 恢复先前设置
	pDC->SetTextColor(OldColor);
	pDC->SetBkMode(OldBkMode);
	pDC->SelectObject(pFont);

	//
	// 绘制关闭按钮
	// 如果鼠标在按钮内就用红色画笔
	//
	if(this->m_IsInCloseButton)
	{
		Pen.CreatePen(PS_SOLID,0,RGB(255,0,0));
		OldPen=pDC->SelectObject(&Pen);

	}

	CRect rClose;
	rClose = m_CloseButtonRect;
	rClose.DeflateRect(2, 2);
	
	pDC->MoveTo(rClose.left,rClose.top);
	pDC->LineTo(rClose.right,rClose.bottom);
	pDC->MoveTo(rClose.right,rClose.top);
	pDC->LineTo(rClose.left,rClose.bottom);

	if(this->m_IsInCloseButton)
	{
		pDC->SelectObject(OldPen);
	}
}

/////////////////////////////////////////
//
// 返回停靠的位置
//
// -1没有停靠
// 1，左
// 2，右
// 3，上
// 4，下
// 0,未知
//
int CDialogBarEx::GetDockingSide()
{
	if(IsFloating())
	{
		return -1;
	}

	LONG id;
	id=::GetWindowLong(this->GetParent()->GetSafeHwnd(),GWL_ID);

	if(!m_bHorz)
	{
		//是否在左边
		if(id==AFX_IDW_DOCKBAR_LEFT)
		{
			return 1;
		}

		//是否在右边
		if(id==AFX_IDW_DOCKBAR_RIGHT)
		{
			return 2;
		}
	}
	else
	{

		//是否在上边
		if(id==AFX_IDW_DOCKBAR_TOP)
		{
			return 3;
		}

		if(id==AFX_IDW_DOCKBAR_BOTTOM)
		{
			return 4;
		}
	}

	return 0;
}

///////////////////////////////////////////////
//
// 设置窗口的最小尺寸
//
// size[in]	窗口的最小尺寸
//
void CDialogBarEx::SetWndMinSize(CSize size)
{
	if(size.cx<0)
		m_MinSize.cx=0;
	else
		m_MinSize.cx=size.cx;

	if(size.cy<0)
		m_MinSize.cy=0;
	else
		m_MinSize.cy=size.cy;

}

///////////////////////////////////////////////
//
// 返回窗口的最小尺寸
//
CSize CDialogBarEx::GetWndMinSize()
{
	return m_MinSize;
}


////////////////////////////////////////////////
//
// 获得浮动时的窗口的尺寸
//
void CDialogBarEx::GetFloatingSize(CSize &size)
{
	size.cx=this->m_FloatingSize.cx;
	size.cy=this->m_FloatingSize.cy;
}

/////////////////////////////////////////////
//
// 显示窗口
//
void CDialogBarEx::ShowDlgBar()
{
	if(IsFloating())
	{
		this->m_pContainWnd->ShowWindow(SW_NORMAL);
	}
	else
	{
		this->GetDockingFrame()->ShowControlBar(this,TRUE,FALSE);
	}
}

////////////////////////////////////////////
//
// 隐藏窗口
//
void CDialogBarEx::HideDlgBar()
{

	if(IsFloating())
	{
		this->m_pContainWnd->ShowWindow(SW_HIDE);
	}
	else
	{
		this->GetDockingFrame()->ShowControlBar(this,FALSE,FALSE);
	}
}

////////////////////////////////////////
//
// 判断窗口是否可见
//
// 可见返回true,不可见返回false
//
bool CDialogBarEx::IsWindowVisible()
{
	if(CWnd::IsWindowVisible())
	{
		return true;
	}
	else
	{
		return false;
	}
}

void CDialogBarEx::GetClientRect(LPRECT lpRect)
{
	CRect ClientRect;

	CWnd::GetClientRect(&ClientRect);

	lpRect->left=ClientRect.left;
	lpRect->top=ClientRect.top;
	lpRect->right=ClientRect.right;
	lpRect->bottom=ClientRect.bottom;

	if(IsFloating())
	{
		return;
	}
	else
	{
		lpRect->top+=m_TitleRect.Height();

		// 留出位置给改变大小<->
		lpRect->left+=4;
		lpRect->right-=4;
		lpRect->bottom-=4;
	}
}

/////////////////////////////////////////////////////////////////
//
// 鼠标离开的消息响应函数
//
LRESULT CDialogBarEx::OnMouseLeave(WPARAM wParam, LPARAM lParam)
{
	m_IsCursorInWnd=false;

	if(m_IsInCloseButton)
	{
		this->m_IsInCloseButton=false;
		this->InvalidateRect(&this->m_CloseButtonRect,FALSE);
//		this->Invalidate();
	}

	return 0;
}
