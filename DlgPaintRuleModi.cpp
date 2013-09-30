// DlgPaintRuleModi.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DlgPaintRuleModi.h"
#include "Page1.h"
#include "page2.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDlgPaintRuleModi dialog


CDlgPaintRuleModi::CDlgPaintRuleModi(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgPaintRuleModi::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgPaintRuleModi)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CDlgPaintRuleModi::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgPaintRuleModi)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgPaintRuleModi, CDialog)
	//{{AFX_MSG_MAP(CDlgPaintRuleModi)
	ON_WM_DESTROY()
	ON_WM_CLOSE()
	ON_WM_CANCELMODE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgPaintRuleModi message handlers

BOOL CDlgPaintRuleModi::OnInitDialog()
{
	CDialog::OnInitDialog();
	m_Sheet.AddPage(&m_page1);
	m_Sheet.AddPage(&m_page2);
	m_Sheet.Create(this,WS_CHILD|WS_VISIBLE,0);
	m_Sheet.ModifyStyleEx(0,WS_EX_CONTROLPARENT);
	m_Sheet.ModifyStyle(0,WS_TABSTOP);
	m_Sheet.SetWindowPos(NULL,0,0,100,800,SWP_NOZORDER|SWP_NOSIZE|SWP_NOACTIVATE);
	
	ShowWindow(SW_SHOW);
	
	return true;
}

void CDlgPaintRuleModi::OnDestroy() 
{
	DeleteObject();
	CDialog::OnDestroy();		
}

void CDlgPaintRuleModi::OnClose() 
{	
	CDialog::OnClose();
	if ( IsWindow( this->GetSafeHwnd() ) )
	{
		DeleteObject();
		DestroyWindow();
	}
}

void CDlgPaintRuleModi::OnCancel() 
{
	CDialog::OnCancel();
// TODO: Add your message handler code here
	
}

void CDlgPaintRuleModi::OnCancelMode() 
{
	CDialog::OnCancelMode();
	
	// TODO: Add your message handler code here
	
}

void CDlgPaintRuleModi::DeleteObject()
{
	m_page1.UnMouseWheel();		//释放内存
	m_page1.SaveSetting();		//保存环境的设置
}
