// PropertyWnd.cpp : implementation file
//

#include "stdafx.h"
#include "PropertyWnd.h"

#include <stdlib.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define DEFAULT_BACKGROUND_COLOR RGB(192,192,192)	// 默认的背景色

#define DEFAULT_SELECT_COLOR_FOCUS RGB(170,170,255)			//默认的被选中并具有输入焦点的背景色
#define DEFAULT_SELECT_COLOR_NOFOCUS RGB(192,192,192)	//默认的被选中但失去输入焦点的背景色

/////////////////////////////////////////////////////////////////////////////
// CPropertyWnd

CPropertyWnd::CPropertyWnd()
{
	m_ElementCount=0;

	m_WndBackGroundColor=RGB(255,255,255);	// 背影默认为白色

	m_WndBackGroundBrush.CreateSolidBrush(m_WndBackGroundColor);

	m_IsNeedReUpdateElementsListRect=FALSE;	//默认不需要重新计算各个单元的位置

	m_CurrentScrollPos=0;	// 设置当前滚动条的位置

	m_BitmapColumeWidth=16;	// 设置位图列默认的宽度

	m_BitmapColumeColor=DEFAULT_BACKGROUND_COLOR;	// 设置位图列背景色

	m_BitmapColumeBrush.CreateSolidBrush(m_BitmapColumeColor);	// 创建位图列背景画刷

	m_FrameColor=DEFAULT_BACKGROUND_COLOR;			// 设置边框的颜色

	m_FramePen.CreatePen(PS_SOLID,0,m_FrameColor);	// 创建绘制边框用的画笔

	m_SelectRow=-1;			// 设置选择的单元行号

	m_hDragCursor=LoadCursor(NULL, IDC_SIZEWE); //加载图标

	m_IsInChangeElementRect=FALSE;	// 初始化不在可改变单元大小的区域

}

CPropertyWnd::~CPropertyWnd()
{
	POSITION pos;
	ElementDef_Wnd *pElementWnd;

	pos=this->m_ElementList.GetHeadPosition();

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		// 释放非自定义的子控件
		if(!(pElementWnd->Style & DefineControlWnd))
			delete pElementWnd->pControlWnd;

		delete pElementWnd;
	}

	this->m_ElementList.RemoveAll();

	::DeleteObject(m_hDragCursor);	// 释放图标资源
}


BEGIN_MESSAGE_MAP(CPropertyWnd, CWnd)
	//{{AFX_MSG_MAP(CPropertyWnd)
	ON_WM_CREATE()
	ON_WM_PAINT()
	ON_WM_ERASEBKGND()
	ON_WM_VSCROLL()
	ON_WM_LBUTTONDOWN()
	ON_WM_SETFOCUS()
	ON_WM_KILLFOCUS()
	ON_WM_MOUSEWHEEL()
	ON_WM_MOUSEMOVE()
	ON_WM_SETCURSOR()
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CPropertyWnd message handlers

int CPropertyWnd::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	CWnd *pParent;
	CFont *pFont;
	LOGFONT LogFont;

	//
	// 创建与与父窗口同样的字体
	//
	pParent=GetParent();

	if(pParent && IsWindow(pParent->GetSafeHwnd()))
	{
		pFont=pParent->GetFont();
		pFont->GetLogFont(&LogFont);

		this->m_Font.CreateFontIndirect(&LogFont);
	}
	
	// 计算各个单元的位置
	UpdateElementsListRect();
	UpdateScrollBarInfo();
	UpdateElementsListRect();

	return 0;
}

///////////////////////////////
//
// 响应WM_PAINT消息
//
void CPropertyWnd::OnPaint() 
{
	//
	// 未了减少闪烁,先在内存中绘制,最后一次性在屏幕上绘制
	//
	CPaintDC dc(this); // device context for painting
	CDC MemDC;
	CBitmap *pOldBitmap;
	CRect ClientRect;
	CFont *OldFont;

	MemDC.CreateCompatibleDC(&dc);

	// 选入绘图的字体
	if(m_Font.GetSafeHandle()!=NULL)
	{
		OldFont=MemDC.SelectObject(&m_Font);
	}

	pOldBitmap=MemDC.SelectObject(&this->m_BitmapSurface);

	//是否需要重新计算各个单元的位置
	if(m_IsNeedReUpdateElementsListRect)
	{
		UpdateElementsListRect();
		UpdateScrollBarInfo();
		UpdateElementsListRect();

	}

	GetClientRect(&ClientRect);

	//绘制各个单元
	PaintElement(&MemDC);

	dc.BitBlt(0,0,ClientRect.Width(),ClientRect.Height(),&MemDC,0,0,SRCCOPY);

	MemDC.SelectObject(pOldBitmap);

	//替换回原来的字体
	if(m_Font.GetSafeHandle()!=NULL)
	{
		MemDC.SelectObject(OldFont);
	}
}

BOOL CPropertyWnd::OnEraseBkgnd(CDC* pDC) 
{
	CDC Memdc;
	CBitmap *pOldBitmap;

	//
	// 用背景色重添背景位图
	//
	Memdc.CreateCompatibleDC(pDC);

	pOldBitmap=Memdc.SelectObject(&this->m_BitmapSurface);

	CRect rect;

	GetClientRect(&rect);

	Memdc.FillRect(&rect,&m_WndBackGroundBrush);

	Memdc.SelectObject(pOldBitmap);

	return 1;
}

///////////////////////////////////////////////////////////////////////////////
//
// 响应WM_VSCROLL
//
void CPropertyWnd::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	int iMax,iMin;
	int iPosTemp;

	this->GetScrollRange(SB_VERT,&iMin,&iMax);

	iPosTemp=m_CurrentScrollPos;

	if(nSBCode==SB_LINEUP)
	{
		m_CurrentScrollPos=GetPrevVisibleRow(m_CurrentScrollPos);

		m_CurrentScrollPos=__max(m_CurrentScrollPos,iMin);

	}
	else if(nSBCode==SB_LINEDOWN)
	{
		m_CurrentScrollPos=GetNextVisibleRow(m_CurrentScrollPos);

		m_CurrentScrollPos=__min(m_CurrentScrollPos,iMax);
	}
	else if(nSBCode==SB_PAGEUP)
	{
		m_CurrentScrollPos-=this->GetVisibleRow();

		m_CurrentScrollPos=__max(m_CurrentScrollPos,iMin);
	}
	else if(nSBCode==SB_PAGEDOWN)
	{
		m_CurrentScrollPos+=this->GetVisibleRow();

		m_CurrentScrollPos=__min(m_CurrentScrollPos,iMax);
	}
	else if(nSBCode==SB_THUMBPOSITION)
	{
		m_CurrentScrollPos=nPos;
	}
	else if(nSBCode==SB_THUMBTRACK)
	{
		m_CurrentScrollPos=nPos;
	}

	//
	// 如果未改变位置不需要重画
	//
	if(iPosTemp==m_CurrentScrollPos)
		return;

	SetScrollPos(SB_VERT,m_CurrentScrollPos);

	this->Invalidate();
	this->UpdateWindow();
}

////////////////////////////////////////////////////////////
//
// 响应鼠标左键消息
//
void CPropertyWnd::OnLButtonDown(UINT nFlags, CPoint point) 
{
	::SetFocus(this->GetSafeHwnd());

	if(this->m_IsInChangeElementRect)
	{
		if(m_SelectRow>=0 && GetElementAt(m_SelectRow)->pControlWnd!=NULL)
		{
			GetElementAt(this->m_SelectRow)->pControlWnd->MoveWindow(0,0,0,0);
		}
		return;
	}

	// 对位图列单元进行测试
	BitmapElementsHitTest(point);

	// 对坐单元列进行测试
	LeftElementsHitTest(point);

	// 对右单元列进行测试
	RightElementsHitTest(point);

	CWnd::OnLButtonDown(nFlags, point);
}

//////////////////////////////////////////
//
// 响应获得输入焦点
//
void CPropertyWnd::OnSetFocus(CWnd* pOldWnd) 
{
	CWnd::OnSetFocus(pOldWnd);
	
	//
	// 设置被选中单元当获得输入焦点时的背景色
	//
	if(this->m_SelectRow>=0)
	{
		GetElementAt(m_SelectRow)->BackGroundColor=DEFAULT_SELECT_COLOR_FOCUS;
	}
	
	this->Invalidate();
}

///////////////////////////////////////////////
//
// 响应失去输入焦点
//
void CPropertyWnd::OnKillFocus(CWnd* pNewWnd) 
{
	CWnd *pParentWnd;

	CWnd::OnKillFocus(pNewWnd);

	//
	// 当获得输入焦点的是THIS控件的子窗口时，不改变THIS控件的式样
	//
	if(pNewWnd!=NULL)
	{
		pParentWnd=pNewWnd->GetParent();


		while(pParentWnd)
		{
			if(pParentWnd==this)
				break;

			pParentWnd=pParentWnd->GetParent();
		}

		if(pParentWnd==this)
		{
			return;
		}
	}
	
	//
	// 设置被选中单元当失去输入焦点时的背景色
	// 并将其中的子控件不可见
	//
	if(this->m_SelectRow>=0)
	{
		GetElementAt(m_SelectRow)->BackGroundColor=DEFAULT_SELECT_COLOR_NOFOCUS;
		GetElementAt(m_SelectRow)->pControlWnd->MoveWindow(0,0,0,0);
	}

	this->Invalidate();
	
}

BOOL CPropertyWnd::OnCommand(WPARAM wParam, LPARAM lParam) 
{
	//
	// 当子控件失去输入焦点，改变选中单元为无焦点的背景色
	//
	if(HIWORD(wParam)==EN_KILLFOCUS || HIWORD(wParam)==CBN_KILLFOCUS)
	{
		if(this->m_SelectRow>=0)
		{
			GetElementAt(m_SelectRow)->BackGroundColor=DEFAULT_SELECT_COLOR_NOFOCUS;
			GetElementAt(m_SelectRow)->pControlWnd->MoveWindow(0,0,0,0);
		}

		this->Invalidate();
	}

	//
	// 通过PWN_CONTROL通报转发子控件的消息
	//
	if(lParam!=NULL && m_SelectRow>=0)
	{
		ElementDef_Wnd *pElementWnd;
		PWNControlStruct NofityStruct;
		CWnd *pParent;

		pElementWnd=GetElementAt(m_SelectRow);

		if(pElementWnd->pControlWnd==NULL || pElementWnd->pControlWnd->GetSafeHwnd()!=(HWND)lParam)
		{
			return CWnd::OnCommand(wParam, lParam);;
		}

		//
		// 填写NofityStruct结构
		//
		NofityStruct.nmh.code=PWN_CONTROL;
		NofityStruct.nmh.hwndFrom=this->GetSafeHwnd();
		NofityStruct.nmh.idFrom=GetWindowLong(this->GetSafeHwnd(),GWL_ID);

		NofityStruct.szElementName=pElementWnd->ElementName;
		NofityStruct.Style=pElementWnd->Style;
		NofityStruct.Width=pElementWnd->Width;
		NofityStruct.Height=pElementWnd->Height;
		NofityStruct.BackGroundColor=pElementWnd->BackGroundColor;
		NofityStruct.TextColor=pElementWnd->TextColor;
		NofityStruct.pVoid=pElementWnd->pVoid;
		NofityStruct.szRightElementText=pElementWnd->RightElementText;
		NofityStruct.Row=pElementWnd->Row;

		NofityStruct.idCtrl=(int)LOWORD(wParam);
		NofityStruct.wNotifyCode=HIWORD(wParam);
		NofityStruct.hWndControl=(HWND)lParam;

		pParent=this->GetParent();

		if(pParent==NULL || !IsWindow(pParent->GetSafeHwnd()))
		{
			return CWnd::OnCommand(wParam, lParam);
		}

		// 向父窗口发送消息
		pParent->SendMessage(WM_NOTIFY,NofityStruct.nmh.idFrom,(LPARAM)&NofityStruct);

	}

	return CWnd::OnCommand(wParam, lParam);
}

//////////////////////////////////////////////////////////////////////
//
// 响应鼠标中键的滚动
//
BOOL CPropertyWnd::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{

	if(zDelta>=0)
	{
		this->OnVScroll(SB_LINEUP,0,NULL);
	}
	else
	{
		this->OnVScroll(SB_LINEDOWN,0,NULL);
	}

	return CWnd::OnMouseWheel(nFlags, zDelta, pt);
}

//////////////////////////////////////////////////////////
//
// 响应鼠标移动
//
void CPropertyWnd::OnMouseMove(UINT nFlags, CPoint point) 
{
	CRect DragRect;
	CRect ClientRect;

	ElementDef_Wnd *pElementWnd;

	//
	// 如果按下鼠标左键，并且在此移动前不在可改变大小的范围内
	// 将不处理
	//
	if((nFlags & MK_LBUTTON) && m_IsInChangeElementRect==FALSE)
	{
		return;
	}

	if((nFlags & MK_LBUTTON) && m_IsInChangeElementRect==TRUE)
	{
		m_IsInChangeElementRect=TRUE;
	}
	else
	{
		m_IsInChangeElementRect=FALSE;
	}

	//
	// 获得第一个有ChildElement样式的单元
	//
	for(int i=0;;i++)
	{
		pElementWnd=(ElementDef_Wnd*)this->GetElementAt(i);

		if(!pElementWnd)
			return;

		if(pElementWnd->Style & ChildElement)
			break;
	}

	GetClientRect(&ClientRect);

	DragRect.SetRect(pElementWnd->LeftColumeRect.right-3,ClientRect.top,
					 pElementWnd->LeftColumeRect.right+3,ClientRect.bottom);


	if(DragRect.PtInRect(point))
	{

		// 设置可该变单元的大小
		m_IsInChangeElementRect=TRUE;
	}

	if(m_IsInChangeElementRect)
	{
		//
		//
		//
		if(nFlags & MK_LBUTTON)
		{
			ChangeLeftElementWidth(pElementWnd->LeftColumeRect.right-
							   pElementWnd->LeftColumeRect.left+
							   point.x-m_MouseMovePt.x);

			this->Invalidate();
			UpdateWindow();
		}
	}

	// 保存鼠标的位置
	m_MouseMovePt=point;

	CWnd::OnMouseMove(nFlags, point);
}

//////////////////////////////////////////////////////////////////////////
//
// 改变鼠标的样式
//
BOOL CPropertyWnd::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	//
	// 如果鼠标在指定的范围内将变为可该变单元大小的图标
	//
	if(m_IsInChangeElementRect)
	{
		SetCursor(m_hDragCursor);
	}
	else
	{
		return CWnd::OnSetCursor(pWnd, nHitTest, message);
	}

	return TRUE;
}

///////////////////////////////////////////////////////
//
// 响应WM_SIZE消息
//
void CPropertyWnd::OnSize(UINT nType, int cx, int cy) 
{
	CWnd *pWnd;

	CWnd::OnSize(nType, cx, cy);
	
	CClientDC dc(this);

	m_BitmapSurface.DeleteObject();

	//分配与窗口同大小的位图用于绘制
	m_BitmapSurface.CreateCompatibleBitmap(&dc,cx,cy);

	this->UpdateScrollBarInfo();
	this->UpdateElementsListRect();

	if(m_SelectRow>=0)
	{
		pWnd=this->GetElementAt(m_SelectRow)->pControlWnd;

		if(pWnd && IsWindow(pWnd->GetSafeHwnd()))
		{
			pWnd->MoveWindow(0,0,0,0);
		}
	}
}

///////////////////////////////////////////////////////////////////////

/////////////////////////////////////////
//
// 返回总的单元数
//
int CPropertyWnd::GetElementCount()
{
	return m_ElementCount;
}

/////////////////////////////////////////////////////////////////////////
//
// 插入单元
// pElement[in]		用于定义单元的ElementDef结构
// InsertAfter[in]	单元插入的位置,如果小于0则将数据加入在开头位置,
//					如果大于最大的单元将加入在最后位置
//
int CPropertyWnd::InsertElement(ElementDef *pElement, int InsertAfter)
{

	ElementDef_Wnd *pElementWnd;

	if(pElement==NULL)
		return -1;

	pElementWnd=new ElementDef_Wnd;

	pElementWnd->ElementName=pElement->szElementName;
	pElementWnd->RightElementText=pElement->RightElementText;
	pElementWnd->strFieldName = pElement->strFieldName;
	
	pElementWnd->Style=pElement->Style;
	pElementWnd->Width=pElement->Width;
	pElementWnd->Height=pElement->Height;

	if(pElementWnd->Style & DefineControlWnd)
	{
		pElementWnd->pControlWnd=pElement->pControlWnd;
	}

	//
	// 如果有DefineBackColor 样式将用用户自定义的背景色
	// 否则将会用默认的背景色
	//
	if(pElement->Style & DefineBackColor)
		pElementWnd->BackGroundColor=pElement->BackGroundColor;
	else if(pElement->Style & TitleElement)
		pElementWnd->BackGroundColor=DEFAULT_BACKGROUND_COLOR;
	else
		pElementWnd->BackGroundColor=RGB(255,255,255);

	//
	// 如果有DefineTextColor 样式将用用户自定义的字体色
	// 否则将会用默认的字体色
	//
	if(pElement->Style & DefineTextColor)
		pElementWnd->TextColor=pElement->TextColor;
	else
		pElementWnd->TextColor=RGB(0,0,0);

	pElementWnd->pVoid=pElement->pVoid;

	//
	// 如果InsertAfter小于0则加入在开始位置
	// 如果大于等于最后元素的位置
	// 新单元将插入在最后位置上
	//
	if(InsertAfter<0)
	{
		this->m_ElementList.AddHead(pElementWnd);
	}
	else if(InsertAfter>=m_ElementCount-1)
	{
		pElementWnd->Row=m_ElementCount;
		this->m_ElementList.AddTail(pElementWnd);
	}
	else
	{
		ElementDef_Wnd *pTest;
		POSITION pos;

		pos=this->m_ElementList.GetHeadPosition();

		while(pos!=NULL)
		{
			pTest=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

			if(pTest->Row>InsertAfter)
				break;
		}

		this->m_ElementList.InsertBefore(pos,pElementWnd);
	}

	m_ElementCount++;

	// 设置需要重新计算各个单元的位置的标志
	m_IsNeedReUpdateElementsListRect=TRUE;

	return pElementWnd->Row;
}

//////////////////////////////////////////////////////////////////////
//
// 返回指定位置的单元
//
// ElementNum[in]	需返回的单元行号
//
// 返回指定位置的单元的ElementDef_Wnd结构
//
CPropertyWnd::ElementDef_Wnd* CPropertyWnd::GetElementAt(int ElementNum)
{
	POSITION pos;
	ElementDef_Wnd *pElementWnd;

	if(ElementNum<0 || ElementNum>=m_ElementCount)
		return NULL;

	pos=this->m_ElementList.GetHeadPosition();

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		if(pElementWnd->Row==ElementNum)
			break;
	}

	if(pElementWnd->pControlWnd && IsWindow(pElementWnd->pControlWnd->GetSafeHwnd()))
	{
		pElementWnd->pControlWnd->GetWindowText(pElementWnd->RightElementText);
	}

	return pElementWnd;
}

/////////////////////////////////////////////////////////////////////
//
// 返回指定位置的单元
//
// pElement[out]	返回的单元结构
// ElementNum[in]	需返回的单元行号
//
// 如果成功返回单元的行号，否则返回-1
// 
// 注意：返回的pControlWnd如果不是用户自定义的不应该释放
//
int CPropertyWnd::GetElementAt(ElementDef *pElement, int ElementNum)
{
	ElementDef_Wnd *pElementWnd;

	//如过pElement为NULL或单元行好号大于单元总数返回-1
	if(pElement==NULL || ElementNum>=GetElementCount())
		return -1;

	pElementWnd=GetElementAt(ElementNum);
	if ( pElementWnd == NULL )
		return -1;

	pElement->BackGroundColor	=pElementWnd->BackGroundColor;
	pElement->Height			=pElementWnd->Height;
	pElement->pControlWnd		=pElementWnd->pControlWnd;
	pElement->pVoid				=pElementWnd->pVoid;
	pElement->RightElementText	=pElementWnd->RightElementText;
	pElement->strFieldName		=pElementWnd->strFieldName;
	pElement->Style				=pElementWnd->Style;
	pElement->szElementName		=pElementWnd->ElementName;
	pElement->TextColor			=pElementWnd->TextColor;
	pElement->Width				=pElementWnd->Width;

	return -1;
}

/////////////////////////////////////////////////
//
// 重新计算各个单元的位置坐标
//
void CPropertyWnd::UpdateElementsListRect()
{
	CRect ClientRect;
	POSITION pos;
	SIZE TextSize;
	int MaxWidth=0;
	int RowPos;
	ElementDef_Wnd *pElementWnd,*preElementWnd;

	if(!IsWindow(this->GetSafeHwnd()))
	{
		return;
	}

	CClientDC dc(this);

	GetClientRect(&ClientRect);

	pos=this->m_ElementList.GetHeadPosition();
	RowPos=0;

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		//重新设置行号
		pElementWnd->Row=RowPos++;

		//设置位图列
		pElementWnd->BitmapColumeRect.left=0;
		pElementWnd->BitmapColumeRect.right=pElementWnd->BitmapColumeRect.left+GetBitmapColumeWidth();

		//左单元应该紧连着位图列
		pElementWnd->LeftColumeRect.left=pElementWnd->BitmapColumeRect.right;

		if(pElementWnd->Row==0)
		{
			pElementWnd->LeftColumeRect.top=0;
		}
		else
		{
			pElementWnd->LeftColumeRect.top=preElementWnd->LeftColumeRect.bottom;
		}

		// 获得左单元内容的长度
		::GetTextExtentPoint32(dc.GetSafeHdc(),pElementWnd->ElementName,
							   lstrlen(pElementWnd->ElementName),&TextSize);

		TextSize.cy+=5;
		TextSize.cx+=5;

		//
		// 当输入的宽度，高度小于0时自动计算
		//
		if(pElementWnd->Width<0)
		{
			pElementWnd->LeftColumeRect.right=pElementWnd->LeftColumeRect.left+TextSize.cx;
			pElementWnd->Width=TextSize.cx;
		}
		else
		{
			pElementWnd->LeftColumeRect.right=pElementWnd->LeftColumeRect.left+pElementWnd->Width;
		}

		if(pElementWnd->Height<0)
		{
			pElementWnd->LeftColumeRect.bottom=pElementWnd->LeftColumeRect.top+TextSize.cy;
			pElementWnd->Height=TextSize.cy;
		}
		else
		{
			pElementWnd->LeftColumeRect.bottom=pElementWnd->LeftColumeRect.top+pElementWnd->Height;
		}

		// 左单元格的大小不能超过整个窗口的最右边
		if(pElementWnd->LeftColumeRect.right>ClientRect.right)
		{
			pElementWnd->LeftColumeRect.right=ClientRect.right;
			pElementWnd->Width=pElementWnd->LeftColumeRect.right-pElementWnd->LeftColumeRect.left;
		}

		if(MaxWidth<pElementWnd->LeftColumeRect.right-pElementWnd->LeftColumeRect.left)
			MaxWidth=pElementWnd->LeftColumeRect.right-pElementWnd->LeftColumeRect.left;

		//
		// 确定右边窗口的大小
		//
		pElementWnd->RightColumeRect.left=pElementWnd->LeftColumeRect.right;
		pElementWnd->RightColumeRect.top=pElementWnd->LeftColumeRect.top;
		
		// 右边单元窗口会充满窗口剩下的部分
		pElementWnd->RightColumeRect.right=ClientRect.right;
		pElementWnd->RightColumeRect.bottom=pElementWnd->LeftColumeRect.bottom;

		preElementWnd=pElementWnd;
	}

	//
	// 对齐左边列单元使宽度相等，即等于最大的宽度
	//
	pos=this->m_ElementList.GetHeadPosition();
	
	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		pElementWnd->LeftColumeRect.right=pElementWnd->LeftColumeRect.left+MaxWidth;

		pElementWnd->Width=pElementWnd->LeftColumeRect.right-pElementWnd->LeftColumeRect.left;

		pElementWnd->RightColumeRect.left=pElementWnd->LeftColumeRect.right;

		//如果是TitleElement样式将坐单元将充满整个窗口宽
		if(pElementWnd->Style & TitleElement)
		{
			pElementWnd->LeftColumeRect.right=ClientRect.right;
		}

		//设置位图列
		pElementWnd->BitmapColumeRect.top=pElementWnd->LeftColumeRect.top;
		pElementWnd->BitmapColumeRect.bottom=pElementWnd->LeftColumeRect.bottom;

	}

	m_IsNeedReUpdateElementsListRect=FALSE;
}

/////////////////////////////////////////////
//
// 重新更新滚动条的信息
//
// 对滚动条进行范围计算
// 此函数应该在UpdateElementsListRect后立刻调用
//
void CPropertyWnd::UpdateScrollBarInfo()
{
	POSITION pos;
	int VisibleElementCount=0;
	int VisibleElementHeight=0;
	CRect rect;
	ElementDef_Wnd *pElementWnd;

	if(!IsWindow(this->GetSafeHwnd()))
	{
		return;
	}

	this->GetClientRect(&rect);


//	this->SetScrollRange(SB_VERT,0,0,FALSE);

	pos=this->m_ElementList.GetHeadPosition();

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		VisibleElementHeight+=pElementWnd->Height;
		VisibleElementCount++;

		if(VisibleElementHeight>rect.Height())
		{
			this->SetScrollRange(SB_VERT,0,this->m_ElementCount-VisibleElementCount+1,FALSE);
			break;
		}
		else if(VisibleElementHeight==rect.Height())
		{
			this->SetScrollRange(SB_VERT,0,this->m_ElementCount-VisibleElementCount,FALSE);
			break;
		}
		
	}

	this->SetScrollPos(SB_VERT,0,FALSE);
	m_CurrentScrollPos=0;

	Invalidate();
	UpdateWindow();
}

//////////////////////////////////////////////////////////
//
// 设置整个窗口的背影色
//
// Color[in]	整个窗口的背景色
//
void CPropertyWnd::SetWndBackGroundColor(COLORREF Color)
{
	m_WndBackGroundColor=Color;
	m_WndBackGroundBrush.DeleteObject();

	m_WndBackGroundBrush.CreateSolidBrush(Color);
}

/////////////////////////////////////////////////////////
//
// 返回整个窗口的背影色
//
COLORREF CPropertyWnd::GetWndBackGroundColor()
{
	return m_WndBackGroundColor;
}

//////////////////////////////////////////////////////
//
// 绘造单元
//
void CPropertyWnd::PaintElement(CDC *pDC)
{
	POSITION pos;
	CRect rect;
	int VisibleHeight=0;
	int baseTop;
	ElementDef_Wnd *pElementWnd;
	CBrush *pOldBrush;
	CBrush WhiteBrush;
	CPen *pOldPen;
	BOOL bSelectElementIsDraw=FALSE;		// 被选中的单元是否被绘制

	GetClientRect(&rect);

	//
	// 绘制位图列的背景色
	//
	CRect BitmapColumeRect(0,0,
						   this->m_BitmapColumeWidth,rect.Height());

	pDC->FillRect(&BitmapColumeRect,&this->m_BitmapColumeBrush);


	//找到当前滚动位置上的单元
	pos=m_ElementList.FindIndex(m_CurrentScrollPos);

	if(pos==NULL)
	{
		return;
	}

	pOldPen=pDC->SelectObject(&this->m_FramePen);

	//
	// 绘制左单元
	//
	while(pos!=NULL)
	{
		CBrush Brush;

		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		//
		// 如果不显示此单元,并且不具有TitleElement样式,则取下一个单元
		//
		if(pElementWnd->IsHideThisElement && !(pElementWnd->Style & TitleElement))
		{
			continue;
		}

		baseTop=pElementWnd->LeftColumeRect.top-VisibleHeight;
		//
		// 绘制左单元边框
		//
		CRect LeftElementRect(pElementWnd->LeftColumeRect.left,pElementWnd->LeftColumeRect.top-baseTop-1,
							  pElementWnd->LeftColumeRect.right,pElementWnd->LeftColumeRect.bottom-baseTop);

		Brush.CreateSolidBrush(pElementWnd->BackGroundColor);
		pOldBrush=pDC->SelectObject(&Brush);

		pDC->Rectangle(&LeftElementRect);

		pDC->SelectObject(pOldBrush);

		//
		// 绘制右单元边框
		//
		CRect RightElementRect(pElementWnd->RightColumeRect.left-1,pElementWnd->RightColumeRect.top-baseTop-1,
							   pElementWnd->RightColumeRect.right,pElementWnd->RightColumeRect.bottom-baseTop);

		if(pElementWnd->Style & ChildElement)
			pDC->Rectangle(&RightElementRect);

		//
		// 绘制位图列的背景色
		//
		CRect BitmapColumeRect(pElementWnd->BitmapColumeRect.left,pElementWnd->BitmapColumeRect.top-baseTop,
							   pElementWnd->BitmapColumeRect.right,pElementWnd->BitmapColumeRect.bottom-baseTop);

		pDC->FillRect(&BitmapColumeRect,&this->m_BitmapColumeBrush);

		//
		// 绘制位图列中的内容
		//
		if(pElementWnd->Style & TitleElement)
		{
			pDC->SelectObject(pOldPen);

			POINT CenterPt;
			CRect FlagRect;
			int   nLineLength = 5 ;
			// 位图列单元的中心
			CenterPt.x=pElementWnd->BitmapColumeRect.left+
					   (pElementWnd->BitmapColumeRect.right-pElementWnd->BitmapColumeRect.left)/2;
			CenterPt.y=pElementWnd->BitmapColumeRect.top-baseTop+
					   (pElementWnd->BitmapColumeRect.bottom-pElementWnd->BitmapColumeRect.top)/2;

			//
			// 当IsHideThisElement为TRUE时在其中绘制一个正方形中加一个'+'
			// 当IsHideThisElement为FALSE时在其中绘制一个正方形加一个'-'
			//
			FlagRect.SetRect(CenterPt.x-nLineLength, CenterPt.y-nLineLength, CenterPt.x+nLineLength, CenterPt.y+nLineLength);

			pDC->Rectangle(&FlagRect);
			
			nLineLength -= 2;
			if(pElementWnd->IsHideThisElement)
			{
				pDC->MoveTo(CenterPt.x-nLineLength, CenterPt.y);
				pDC->LineTo(CenterPt.x+nLineLength, CenterPt.y);

				pDC->MoveTo(CenterPt.x, CenterPt.y-nLineLength);
				pDC->LineTo(CenterPt.x, CenterPt.y+nLineLength);
			}
			else
			{
				pDC->MoveTo(CenterPt.x-nLineLength, CenterPt.y);
				pDC->LineTo(CenterPt.x+nLineLength, CenterPt.y);
			}

			pDC->SelectObject(&this->m_FramePen);
		}

		//
		// 绘制左单元内容
		//
		pDC->SetTextColor(pElementWnd->TextColor);
		pDC->SetBkColor(pElementWnd->BackGroundColor);
		pDC->DrawText(pElementWnd->ElementName,&LeftElementRect,DT_SINGLELINE|DT_CENTER|DT_VCENTER);

		//
		// 显示右单元控件中的内容
		//
		if(pElementWnd->pControlWnd)
		{
			pElementWnd->pControlWnd->GetWindowText(pElementWnd->RightElementText);
			pDC->SetTextColor(RGB(0,0,0));
			pDC->SetBkColor(RGB(255,255,255));

			RightElementRect.left+=5;
			pDC->DrawText(pElementWnd->RightElementText,&RightElementRect,DT_SINGLELINE|DT_VCENTER);

			//
			// 如果被绘制的单元是选种的单元则调整对应的控件的位置
			//
			if(pElementWnd->Row==this->m_SelectRow)
			{
				pElementWnd->pControlWnd->SetWindowPos(NULL,pElementWnd->RightColumeRect.left,
													   pElementWnd->RightColumeRect.top-baseTop,
													   pElementWnd->RightColumeRect.right-pElementWnd->RightColumeRect.left,
													   pElementWnd->RightColumeRect.bottom-pElementWnd->RightColumeRect.top,
													   SWP_NOZORDER|SWP_NOSIZE|SWP_SHOWWINDOW);

				// 设置被选中的单元已经被绘制
				bSelectElementIsDraw=TRUE;
			}
		}


		//
		//当可见部分大于窗口时,停止绘
		//
		VisibleHeight+=pElementWnd->Height;

		if(VisibleHeight>rect.Height())
			break;
	}

	pDC->SelectObject(pOldPen);

	//
	// 如果被选中的单元未被绘制，则隐藏对应的子孔件
	//
	if(!bSelectElementIsDraw)
	{
		if(this->m_SelectRow>=0 && GetElementAt(this->m_SelectRow)->pControlWnd)
		{
			GetElementAt(this->m_SelectRow)->pControlWnd->ShowWindow(SW_HIDE);
		}
	}

}


/////////////////////////////////////////////////////
//
// 设置为位图列的宽度
//
void CPropertyWnd::SetBitmapColumeWidth(int Width)
{
	if(Width<0)
		Width=0;

	m_BitmapColumeWidth=Width;
}

////////////////////////////////////////////////////
//
// 返回位图列的宽度
//
int CPropertyWnd::GetBitmapColumeWidth()
{
	return m_BitmapColumeWidth;
}

////////////////////////////////////////////////////////
//
// 对位图列的点击进行测试
//
void CPropertyWnd::BitmapElementsHitTest(CPoint point)
{
	POSITION pos;
	int VisibleHeight=0;
	int baseTop;
	ElementDef_Wnd *pElementWnd;
	CRect rect;

	GetClientRect(&rect);

	//找到当前滚动位置上的单元
	pos=m_ElementList.FindIndex(m_CurrentScrollPos);

	if(pos==NULL)
	{
		return;
	}

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		//
		// 如果不显示此单元,并且不具有TitleElement样式,则取下一个单元
		//
		if(pElementWnd->IsHideThisElement && !(pElementWnd->Style & TitleElement))
		{
			continue;
		}

		baseTop=pElementWnd->LeftColumeRect.top-VisibleHeight;

		if(pElementWnd->Style & TitleElement)
		{
			BOOL HideFlag;
			CRect ElementRect(pElementWnd->BitmapColumeRect.left,pElementWnd->BitmapColumeRect.top-baseTop,
							  pElementWnd->BitmapColumeRect.right,pElementWnd->BitmapColumeRect.bottom-baseTop);

			if(ElementRect.PtInRect(point))
			{
				HideFlag=!(pElementWnd->IsHideThisElement);

				pElementWnd->IsHideThisElement=HideFlag;

				while(pos!=NULL)
				{
					pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

					if(pElementWnd->Style & TitleElement)
						break;

					pElementWnd->IsHideThisElement=HideFlag;
				}

				Invalidate();
				UpdateWindow();

				break;
			}
		}

		VisibleHeight+=pElementWnd->Height;

		if(VisibleHeight>rect.Height())
			break;
	}
}

////////////////////////////////////////////////////////
//
// 对左单元的点击进行测试
//
void CPropertyWnd::LeftElementsHitTest(CPoint point)
{
	POSITION pos;
	int VisibleHeight=0;
	int baseTop;
	ElementDef_Wnd *pElementWnd;
	CRect rect;

	GetClientRect(&rect);

	//找到当前滚动位置上的单元
	pos=m_ElementList.FindIndex(m_CurrentScrollPos);

	if(pos==NULL)
	{
		return;
	}

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		//
		// 如果不显示此单元,并且不具有TitleElement样式,则取下一个单元
		//
		if(pElementWnd->IsHideThisElement && !(pElementWnd->Style & TitleElement))
		{
			continue;
		}

		baseTop=pElementWnd->LeftColumeRect.top-VisibleHeight;

		//
		// 对具有ChildElement的样式进行点击处理
		//
		if(pElementWnd->Style & ChildElement)
		{
			CRect ElementRect(pElementWnd->LeftColumeRect.left,pElementWnd->LeftColumeRect.top-baseTop,
							  pElementWnd->LeftColumeRect.right,pElementWnd->LeftColumeRect.bottom-baseTop);

			if(ElementRect.PtInRect(point))
			{
				//
				// 如果控件子窗口还未创建则创建子窗口
				//
				if(pElementWnd->pControlWnd==NULL)
				{
					if(!CreateElementControl(pElementWnd->Row))
						return;
				}

				//
				// 恢复上次被选中单元的颜色
				// 并使上次被选中的子孔件不可见
				//
				if(m_SelectRow>=0)
				{
					GetElementAt(m_SelectRow)->BackGroundColor=m_SelectElementBackGroundColorBank;
					GetElementAt(m_SelectRow)->pControlWnd->MoveWindow(0,0,0,0);
				}

				// 保存被选中单元原来的颜色
				m_SelectElementBackGroundColorBank=pElementWnd->BackGroundColor;

				if(m_SelectRow!=pElementWnd->Row)
				{
					m_SelectRow=pElementWnd->Row;

					// 发送PWN_SELECTCHANGE通知
					SendPWNSelectChangeNotify();
				}

				// 保存被选中单元的行号
				m_SelectRow=pElementWnd->Row;

				pElementWnd->BackGroundColor=DEFAULT_SELECT_COLOR_FOCUS;

				//
				// 如果有子控件将放置子孔件在右单元内
				//
				if(pElementWnd->pControlWnd)
				{
					CRect PlaceRect(pElementWnd->RightColumeRect.left,pElementWnd->RightColumeRect.top-baseTop,
									pElementWnd->RightColumeRect.right-1,pElementWnd->RightColumeRect.bottom-baseTop-1);

					pElementWnd->pControlWnd->MoveWindow(&PlaceRect);
					pElementWnd->pControlWnd->ShowWindow(SW_SHOWNORMAL);
				}



				this->Invalidate();
			}

		}

		baseTop=pElementWnd->LeftColumeRect.top-VisibleHeight;


		VisibleHeight+=pElementWnd->Height;

		if(VisibleHeight>rect.Height())
			break;
	}
}

////////////////////////////////////////////////////////
//
// 对右单元的点击进行测试
//
void CPropertyWnd::RightElementsHitTest(CPoint point)
{
	POSITION pos;
	int VisibleHeight=0;
	int baseTop;
	ElementDef_Wnd *pElementWnd;
	CRect rect;

	GetClientRect(&rect);

	//找到当前滚动位置上的单元
	pos=m_ElementList.FindIndex(m_CurrentScrollPos);

	if(pos==NULL)
	{
		return;
	}

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		//
		// 如果不显示此单元,并且不具有TitleElement样式,则取下一个单元
		//
		if(pElementWnd->IsHideThisElement && !(pElementWnd->Style & TitleElement))
		{
			continue;
		}

		baseTop=pElementWnd->LeftColumeRect.top-VisibleHeight;

		//
		// 对具有ChildElement的样式进行点击处理
		//
		if(pElementWnd->Style & ChildElement)
		{
			CRect ElementRect(pElementWnd->RightColumeRect.left,pElementWnd->RightColumeRect.top-baseTop,
							  pElementWnd->RightColumeRect.right,pElementWnd->RightColumeRect.bottom-baseTop);

			if(ElementRect.PtInRect(point))
			{
				//
				// 如果控件子窗口还未创建则创建子窗口
				//
				if(pElementWnd->pControlWnd==NULL)
				{
					if(!CreateElementControl(pElementWnd->Row))
						return;
				}

				//
				// 恢复上次被选中单元的颜色
				// 并使上次被选中的子孔件不可见
				//
				if(m_SelectRow>=0)
				{
					GetElementAt(m_SelectRow)->BackGroundColor=m_SelectElementBackGroundColorBank;
					GetElementAt(m_SelectRow)->pControlWnd->MoveWindow(0,0,0,0);
				}

				// 保存被选中单元原来的颜色
				m_SelectElementBackGroundColorBank=pElementWnd->BackGroundColor;

				if(m_SelectRow!=pElementWnd->Row)
				{
					m_SelectRow=pElementWnd->Row;

					// 发送PWN_SELECTCHANGE通知
					SendPWNSelectChangeNotify();
				}

				// 保存被选中单元的行号
				m_SelectRow=pElementWnd->Row;

				pElementWnd->BackGroundColor=DEFAULT_SELECT_COLOR_FOCUS;

				//
				// 如果有子控件将放置子孔件在右单元内
				//
				if(pElementWnd->pControlWnd)
				{
					CRect PlaceRect(pElementWnd->RightColumeRect.left,pElementWnd->RightColumeRect.top-baseTop,
									pElementWnd->RightColumeRect.right-1,pElementWnd->RightColumeRect.bottom-baseTop-1);

					pElementWnd->pControlWnd->MoveWindow(&PlaceRect);
					pElementWnd->pControlWnd->ShowWindow(SW_SHOWNORMAL);

					pElementWnd->pControlWnd->SetFocus();
				}

				// 发送PWN_SELECTCHANGE通知
				SendPWNSelectChangeNotify();

				this->Invalidate();
			}

		}

		baseTop=pElementWnd->LeftColumeRect.top-VisibleHeight;

		VisibleHeight+=pElementWnd->Height;

		if(VisibleHeight>rect.Height())
			break;
	}
}

/////////////////////////////////////////
//
// 获得可见的列数
//
int CPropertyWnd::GetVisibleRow()
{
	POSITION pos;
	int VisibleHeight=0;
	int VisibleRowCount=0;
	int baseTop;
	ElementDef_Wnd *pElementWnd;
	CRect rect;

	GetClientRect(&rect);

	//找到当前滚动位置上的单元
	pos=m_ElementList.FindIndex(m_CurrentScrollPos);

	if(pos==NULL)
	{
		return 0;
	}

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		//
		// 如果不显示此单元,并且不具有TitleElement样式,则取下一个单元
		//
		if(pElementWnd->IsHideThisElement && !(pElementWnd->Style & TitleElement))
		{
			continue;
		}

		VisibleRowCount++;

		baseTop=pElementWnd->LeftColumeRect.top-VisibleHeight;

		VisibleHeight+=pElementWnd->Height;

		if(VisibleHeight>rect.Height())
			break;
	}

	return VisibleRowCount;
}

//////////////////////////////////////////////////
//
// 返回指定列的下一个可见的列号
//
// row[in]	指定的列
//
// 返回row下一个可见的列号
// 如果没有可见的单元，将返回负数
//
int CPropertyWnd::GetNextVisibleRow(int row)
{
	ElementDef_Wnd *pElementWnd;
	POSITION pos;

	//找row的下一个单元
	pos=m_ElementList.FindIndex(row+1);

	if(pos==NULL)
	{
		return 0;
	}

	//
	// 找下一个可见的单元，如果单元有TitleElement样式，始终可见
	//
	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		if(!pElementWnd->IsHideThisElement || (pElementWnd->Style & TitleElement))
		{
			return pElementWnd->Row;
		}

	}

	return -1;
}

//////////////////////////////////////////////////
//
// 返回指定列的上一个可见的列号
//
// row[in]	指定的列
//
// 返回row下一个可见的列号
// 如果没有可见的单元，将返回负数
//
int CPropertyWnd::GetPrevVisibleRow(int row)
{
	ElementDef_Wnd *pElementWnd;
	POSITION pos;

	//找row的上一个单元
	pos=m_ElementList.FindIndex(row-1);

	if(pos==NULL)
	{
		return 0;
	}

	//
	// 找上一个可见的单元，如果单元有TitleElement样式，始终可见
	//
	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetPrev(pos);

		if(!pElementWnd->IsHideThisElement || (pElementWnd->Style & TitleElement))
		{
			return pElementWnd->Row;
		}

	}

	return -1;
}

////////////////////////////////////////////////
//
// 重新定义GetClientRect
//
// 用来出去滚动条部分
//
void CPropertyWnd::GetClientRect(LPRECT lpRect)
{
	int iMax,iMin;

	if(lpRect==NULL)
		return;

	CWnd::GetClientRect(lpRect);

	GetScrollRange(SB_VERT,&iMin,&iMax);

	//如果滚动条可见就减去它的部分
	if(iMin<=iMax)
	{
//		lpRect->right-=18;
	}

}

///////////////////////////////////////////////////////
//
// 该边左单元的所有单元的宽度
// 但是不改变具有TitleElement样式的单元
//
// Width[in]	单元的宽度
//
void CPropertyWnd::ChangeLeftElementWidth(int Width)
{
	POSITION pos;
	ElementDef_Wnd *pElementWnd;

	if(Width<0)
		return;

	pos=this->m_ElementList.GetHeadPosition();

	while(pos)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		if(pElementWnd->Style & TitleElement)
		{
			continue;
		}

		pElementWnd->Width=Width;
		pElementWnd->LeftColumeRect.right=pElementWnd->LeftColumeRect.left+Width;

		pElementWnd->RightColumeRect.left=pElementWnd->LeftColumeRect.right;
	}
}

////////////////////////////////////////////////////////
//
// 创建单元的子控件
//
// 因为单元的子控件是在需要时自动创建，此函数可在用户
// 需要时创建,此函数需在窗口创建以后调用
//
// ElementNum[in]	需创建子控件的行号
//
// 如果函数成功返回TURE，如果失败，或已有子控件返回FALSE
//
BOOL CPropertyWnd::CreateElementControl(int ElementNum)
{
	ElementDef_Wnd *pElementWnd;
	BOOL bRet;

	if(ElementNum>=this->GetElementCount())
		return FALSE;

	pElementWnd=GetElementAt(ElementNum);

	//
	// 如果控件子窗口还未创建,并且不是用户自定义控件，就根据样式创建子窗口
	//
	if(pElementWnd->pControlWnd==NULL && !(pElementWnd->Style & DefineControlWnd))
	{
		CRect EmptyRect(0,0,0,0);

		if(pElementWnd->Style & EditElement)
		{
			pElementWnd->pControlWnd=new CEdit;
			bRet=((CEdit*)(pElementWnd->pControlWnd))->Create(ES_LEFT|ES_AUTOHSCROLL,
														 EmptyRect,this,pElementWnd->Row);

		}
		else if(pElementWnd->Style & ComBoxElement)
		{
			CRect ComboxRect(0,0,100,150);

			pElementWnd->pControlWnd=new CComboBox;
			bRet=((CComboBox*)(pElementWnd->pControlWnd))->Create(CBS_DROPDOWN|WS_CHILD|WS_VSCROLL|CBS_AUTOHSCROLL,
															 ComboxRect,this,pElementWnd->Row);

		}

		if(!bRet)
		{
			delete pElementWnd->pControlWnd;
			pElementWnd->pControlWnd=NULL;
			return FALSE;
		}

		// 初始子控件的内内容为右单元的内容
		if(!pElementWnd->RightElementText.IsEmpty())
			pElementWnd->pControlWnd->SetWindowText(pElementWnd->RightElementText);
	}
	else
	{
		return FALSE;
	}

	return TRUE;
}

/////////////////////////////////////////////////////////////////////
//
// 向右单元添加数据
//
// Row[in]			单元的行号
// szContent[in]	需要在右单元显示的内容
//
// 调用此函数应该在属性窗口创建之后
//
void CPropertyWnd::SetRightElementContent(int Row,LPCTSTR szContent)
{
	ElementDef_Wnd *pElementWnd;

	if(szContent==NULL)
		return;

	// 创建右单元内的子窗口
	this->CreateElementControl(Row);

	pElementWnd=GetElementAt(Row);

	pElementWnd->RightElementText=szContent;
	pElementWnd->pControlWnd->SetWindowText(szContent);

	this->Invalidate();
	UpdateData(FALSE);
}

////////////////////////////////////////////////////////////////////////////////
//
// 通过字符串找到指定单元的行号，通过行号可以找到指定的单元
//
// nStartAfter[in]		在指定的单元号之后开始查找
// szString[in]			需要查找单元的字符串
// IsLeftElement[in]	为TRUE时查找左单元的内容，否则查找右单元的内容
//
// 如果找到返回单元的行号，否则返回-1
// 
int CPropertyWnd::FindElement(int nStartAfter, LPCTSTR szString,BOOL IsLeftElement)
{
	POSITION pos;
	int iPos=0;
	ElementDef_Wnd *pElementWnd;

	if(szString==NULL)
		return -1;
	
	pos=this->m_ElementList.GetHeadPosition();

	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		if(nStartAfter<iPos)
		{
			iPos++;
			continue;
		}

		//
		// 如果是左单元内容匹配则返回左单元的行号
		// 如果是右单元的内容匹配则返回右单元的行号
		//
		if(IsLeftElement)
		{
			if(pElementWnd->ElementName==szString)
				return pElementWnd->Row;
		}
		else
		{
			if(pElementWnd->pControlWnd!=NULL && IsWindow(pElementWnd->pControlWnd->GetSafeHwnd()))
			{
				pElementWnd->pControlWnd->GetWindowText(pElementWnd->RightElementText);
			}

			if(pElementWnd->RightElementText==szString)
				return pElementWnd->Row;
		}
	}

	return -1;
}

/////////////////////////////////////////////////
//
// 用于发送PWN_SELECTCHANGE通知
//
void CPropertyWnd::SendPWNSelectChangeNotify()
{
	ElementDef_Wnd *pElementWnd;
	CWnd *pParent;

	PWNSelectChangeStruct NofityStruct;

	pParent=this->GetParent();

	if(pParent==NULL || !IsWindow(pParent->GetSafeHwnd()))
	{
		return;
	}

	if(this->m_SelectRow>0)
	{
		pElementWnd=(ElementDef_Wnd*)this->GetElementAt(m_SelectRow);
	}
	else
	{
		return;
	}

	NofityStruct.nmh.code=PWN_SELECTCHANGE;
	NofityStruct.nmh.hwndFrom=this->GetSafeHwnd();
	NofityStruct.nmh.idFrom=GetWindowLong(this->GetSafeHwnd(),GWL_ID);

	NofityStruct.szElementName=pElementWnd->ElementName;
	NofityStruct.Style=pElementWnd->Style;
	NofityStruct.Width=pElementWnd->Width;
	NofityStruct.Height=pElementWnd->Height;
	NofityStruct.BackGroundColor=pElementWnd->BackGroundColor;
	NofityStruct.TextColor=pElementWnd->TextColor;
	NofityStruct.pVoid=pElementWnd->pVoid;
	NofityStruct.szRightElementText=pElementWnd->RightElementText;
	NofityStruct.Row=pElementWnd->Row;

	// 向父窗口发送消息
	pParent->SendMessage(WM_NOTIFY,NofityStruct.nmh.idFrom,(LPARAM)&NofityStruct);
}

////////////////////////////////////////
//
// 当添加完数据后应该调用此函数
//
void CPropertyWnd::RefreshData()
{
	POSITION pos;
	ElementDef_Wnd *pElementWnd;
	int nRow;

	pos=this->m_ElementList.GetHeadPosition();

	//
	// 给各个单元编号
	//
	nRow=0;
	while(pos!=NULL)
	{
		pElementWnd=(ElementDef_Wnd*)this->m_ElementList.GetNext(pos);

		pElementWnd->Row=nRow++;
	}
}


#undef DEFAULT_BACKGROUND_COLOR
#undef DEFAULT_SELECT_COLOR_FOCUS
#undef DEFAULT_SELECT_COLOR_NOFOCUS



