// DialogBarExContainer.cpp : implementation file
//

#include "stdafx.h"
#include "DialogBarExContainer.h"

#include "DialogBarEx.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDialogBarExContainer

CDialogBarExContainer::CDialogBarExContainer()
{
	m_pSubDialogBar=NULL;	//初始指向将包含在其中的DialogBar
}

CDialogBarExContainer::~CDialogBarExContainer()
{
}


BEGIN_MESSAGE_MAP(CDialogBarExContainer, CWnd)
	//{{AFX_MSG_MAP(CDialogBarExContainer)
	ON_WM_SIZE()
	ON_WM_MOVING()
	ON_WM_CLOSE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CDialogBarExContainer message handlers

void CDialogBarExContainer::OnSize(UINT nType, int cx, int cy) 
{
	CWnd::OnSize(nType, cx, cy);

	if(m_pSubDialogBar && IsWindow(m_pSubDialogBar->GetSafeHwnd()))
	{
		m_pSubDialogBar->m_FloatingSize.cx=cx;
		m_pSubDialogBar->m_FloatingSize.cy=cy;

		m_pSubDialogBar->MoveWindow(0,0,cx,cy);
	}
	
}


//////////////////////////////////////////////////////////////
//
// 响应窗口的WM_MOVING
//
// 判断窗口是否应该停靠
//
// 
void CDialogBarExContainer::OnMoving(UINT fwSide, LPRECT pRect) 
{
	CWnd::OnMoving(fwSide, pRect);
	
	CFrameWnd *pFrame;
	CRect FrameRect;
	CPoint ScreenPt;

	this->MoveWindow(pRect);

	::GetCursorPos(&ScreenPt);

	if(m_pSubDialogBar && IsWindow(m_pSubDialogBar->GetSafeHwnd()))
	{
		pFrame=m_pSubDialogBar->GetDockingFrame();
		pFrame->GetWindowRect(&FrameRect);

		//停靠在左边 
		if(abs(ScreenPt.x-FrameRect.left)<10)
		{
			pFrame->DockControlBar(m_pSubDialogBar,AFX_IDW_DOCKBAR_LEFT);

			pFrame->RecalcLayout();
			this->ShowWindow(SW_HIDE);

			this->m_pSubDialogBar=NULL;
		}
		//停靠在右边
		else if(abs(ScreenPt.x-FrameRect.right)<10)
		{
			pFrame->DockControlBar(m_pSubDialogBar,AFX_IDW_DOCKBAR_RIGHT);
			pFrame->RecalcLayout();
			this->ShowWindow(SW_HIDE);

			this->m_pSubDialogBar=NULL;
		}
		// 停靠在上边
		else if(abs(ScreenPt.y-FrameRect.top)<10)
		{
			pFrame->DockControlBar(m_pSubDialogBar,AFX_IDW_DOCKBAR_TOP);

			pFrame->RecalcLayout();

			this->ShowWindow(SW_HIDE);

			this->m_pSubDialogBar=NULL;
		}
		// 停靠在下边
		else if(abs(ScreenPt.y-FrameRect.bottom)<10)
		{
			pFrame->DockControlBar(m_pSubDialogBar,AFX_IDW_DOCKBAR_BOTTOM);

			pFrame->RecalcLayout();

			this->ShowWindow(SW_HIDE);

			this->m_pSubDialogBar=NULL;
		}
	}	
}

void CDialogBarExContainer::OnClose() 
{
	//点击右上角的“X”不会销毁窗口，只会隐藏
	this->ShowWindow(SW_HIDE);	
//	CWnd::OnClose();
}

//////////////////////////////////////////////////////////////////////////////
//
// 设置包含在其中的DialogBar,当停靠后下一次需浮动前需重新设置
//
// pDialogBar[in]	DialogBar对象的指针
// pt[in]			初始的窗口浮动位置（屏幕坐标）
//
void CDialogBarExContainer::SetSubDialogBar(CDialogBarEx *pDialogBar,CPoint pt)
{
	CSize size;

	if(!IsWindow(this->GetSafeHwnd()))
		return;

	if(pDialogBar && IsWindow(pDialogBar->GetSafeHwnd()))
	{
		//
		// 如果窗口没有浮动使之浮动
		//
		if(!pDialogBar->IsFloating())
		{
			pDialogBar->GetDockingFrame()->FloatControlBar(pDialogBar,pt);
		}

		//
		// 使包容的窗口与DialogBar同样的大
		// 使DialogBar原来的父窗口隐藏
		// 改变父窗口的所属
		//
		if(pDialogBar->GetParentOwner()!=this)
		{
			pDialogBar->GetFloatingSize(size);
			this->MoveWindow(pt.x,pt.y,size.cx,size.cy);

			pDialogBar->GetParentOwner()->ShowWindow(SW_HIDE);

			pDialogBar->SetParent(this);

			this->ShowWindow(SW_NORMAL);
		}
		else
		{
			pDialogBar->GetFloatingSize(size);

			this->MoveWindow(pt.x,pt.y,size.cx,size.cy);
		}
	}

	m_pSubDialogBar=pDialogBar;

}

////////////////////////////////////////////////////////
//
// 返回包含在其中的DialogBar
//
CDialogBarEx* CDialogBarExContainer::GetSubDialogBar()
{
	return m_pSubDialogBar;
}

