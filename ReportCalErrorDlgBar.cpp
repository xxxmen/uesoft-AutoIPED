// ReportCalErrorDlgBar.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "ReportCalErrorDlgBar.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CReportCalErrorDlgBar

CReportCalErrorDlgBar::CReportCalErrorDlgBar()
{
}

CReportCalErrorDlgBar::~CReportCalErrorDlgBar()
{
}


BEGIN_MESSAGE_MAP(CReportCalErrorDlgBar, CDialogBarEx)
	//{{AFX_MSG_MAP(CReportCalErrorDlgBar)
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CReportCalErrorDlgBar message handlers

void CReportCalErrorDlgBar::OnSize(UINT nType, int cx, int cy) 
{
	CDialogBarEx::OnSize(nType, cx, cy);
	
	CWnd *pWnd;
	CRect Rect;

	pWnd=this->GetDlgItem(IDC_REPORT_EDIT);

	if(pWnd && IsWindow(pWnd->GetSafeHwnd()))
	{
		this->GetClientRect(&Rect);
		pWnd->MoveWindow(Rect.left,Rect.top,Rect.Width(),Rect.Height());
	}	
}

/////////////////////////////////////////////////////////////
//
// 设置需在对话框中的EDIT显示的内容
//
// szContent[in]	需显示的字符串
//
void CReportCalErrorDlgBar::SetReportContent(LPCTSTR szContent)
{
	CWnd *pWnd;

	if(szContent==NULL)
	{
		m_InformationToReport=_T("");
	}
	else
	{
		m_InformationToReport=szContent;
	}

	pWnd=this->GetDlgItem(IDC_REPORT_EDIT);

	if(pWnd && IsWindow(pWnd->GetSafeHwnd()))
	{
		pWnd->SetWindowText(m_InformationToReport);
	}

}

////////////////////////////////////////////////
//
// 返回在对话框中的EDIT显示的内容
//
CString CReportCalErrorDlgBar::GetReportContent()
{
	CWnd *pWnd;

	pWnd=this->GetDlgItem(IDC_REPORT_EDIT);

	if(pWnd && IsWindow(pWnd->GetSafeHwnd()))
	{
		pWnd->GetWindowText(m_InformationToReport);
	}

	return m_InformationToReport;
}
