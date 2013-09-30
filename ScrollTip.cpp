// ScrollTip.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "ScrollTip.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CScrollTip

CScrollTip::CScrollTip()
{
	m_crBkColor = RGB(255,255,255);
	m_crTextColor = RGB(0,0,0);
}

CScrollTip::~CScrollTip()
{
}


BEGIN_MESSAGE_MAP(CScrollTip, CStatic)
	//{{AFX_MSG_MAP(CScrollTip)
	ON_WM_PAINT()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CScrollTip message handlers

void CScrollTip::SetBkColor(COLORREF crBk)
{
	m_crBkColor = crBk;

	//重画
	Invalidate();
}

void CScrollTip::SetTextColor(COLORREF crText)
{
	m_crTextColor = crText;

	//重画
	Invalidate();
}

void CScrollTip::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	
	// TODO: Add your message handler code here
	// Do not call CStatic::OnPaint() for painting messages

	CRect rect;
    GetClientRect(&rect);
	
	//设置各项的值
	dc.SetBkColor(m_crBkColor);
	dc.SetTextColor(m_crTextColor);
	CString str;
	this->GetWindowText(str);
    dc.TextOut(rect.left,rect.top,str);
}
