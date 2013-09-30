// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include "AutoIPED.h"

#include "MainFrm.h"

#include "AutoIPEDView.h"
#include "htmlhelp.h"
#include "EDIBgbl.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWnd)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWnd)
	//{{AFX_MSG_MAP(CMainFrame)
	ON_WM_CREATE()
	ON_WM_MOVE()
	ON_COMMAND(ID_MENU_HELP, OnMenuHelp)
	ON_COMMAND(ID_MENU_Index, OnMENUIndex)
	ON_COMMAND(ID_MENU_Search, OnMENUSearch)
	ON_COMMAND(IDM_VIEW_INFORMATION_BAR, OnViewInformationBar)
	ON_UPDATE_COMMAND_UI(IDM_VIEW_INFORMATION_BAR, OnUpdateViewInformationBar)
	ON_WM_CLOSE()
	ON_WM_TIMER()
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_ADDPROGRESS,AddProgress)
	ON_MESSAGE(WM_REMOVEPROGRESS,RemoveProgress)
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // status line indicator
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
};
static UINT indicators1[] =
{
	ID_SEPARATOR,           // status line indicator
    ID_PROGRESS,
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
};

class CAutoIPEDView;

/////////////////////////////////////////////////////////////////////////////
// CMainFrame construction/destruction

CMainFrame::CMainFrame()
{
	bmp8.LoadBitmap(IDB_BITMAP9);
}

CMainFrame::~CMainFrame()
{
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP
		| CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC) ||
		!m_wndToolBar.LoadToolBar(IDR_MAINFRAME))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}

	if (!m_wndStatusBar.Create(this) ||
		!m_wndStatusBar.SetIndicators(indicators,
		  sizeof(indicators)/sizeof(UINT)))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	// 创建报告计算错误的BAR
	this->m_ReportCalInfoBar.Create(this,CReportCalErrorDlgBar::IDD,CBRS_BOTTOM,IDB_REPORT_ERROR_BAR);
	this->m_ReportCalInfoBar.SetWindowText(_T("信息栏"));

	m_wndToolBar.EnableDocking(CBRS_ALIGN_ANY);

	this->m_ReportCalInfoBar.EnableDocking(CBRS_ALIGN_ANY);

	EnableDocking(CBRS_ALIGN_ANY);
	DockControlBar(&m_wndToolBar);
	DockControlBar(&this->m_ReportCalInfoBar);

	//隐藏报告计算错误的BAR
	this->m_ReportCalInfoBar.HideDlgBar();

//	CWnd*parent=GetParent();
	CMenu*pmenubar=this->GetMenu();
//	CMenu*pmenu=pmenubar->GetSubMenu(6);
	pmenubar->SetMenuItemBitmaps(ID_APP_ABOUT,MF_BYCOMMAND,&bmp8,&bmp8);

	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWnd::PreCreateWindow(cs) )
		return FALSE;
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CMainFrame diagnostics

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWnd::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWnd::Dump(dc);
}

#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CMainFrame message handlers


void CMainFrame::OnMove(int x, int y) 
{
	CFrameWnd::OnMove(x, y);
	
    CAutoIPEDView *pView =(CAutoIPEDView *)GetActiveView(); 

	if(pView != NULL)
	{
		::SendMessage(pView->GetSafeHwnd(),WM_SIZE,0,0);
	}
}

LRESULT CMainFrame::AddProgress(WPARAM wParam,LPARAM lParam)
{
	m_wndStatusBar.SetIndicators(indicators1,sizeof(indicators1)/sizeof(UINT));

	m_wndStatusBar.SetProgressWndPos(1);
	//m_wndStatusBar.SetPaneText(0,_T("请稍后..."),true);
	return 1;
}


LRESULT CMainFrame::RemoveProgress(WPARAM wParam,LPARAM lParam)
{
	m_wndStatusBar.ShowProgressWnd(FALSE);
	m_wndStatusBar.SetIndicators(indicators,sizeof(indicators)/sizeof(UINT));
	return 1;
}

void CMainFrame::OnMenuHelp() 
{
	HtmlHelp(NULL,AfxGetApp()->m_pszHelpFilePath,HH_DISPLAY_TOC,NULL);

}

void CMainFrame::OnMENUIndex() 
{
	HtmlHelp(NULL,AfxGetApp()->m_pszHelpFilePath,HH_DISPLAY_INDEX,NULL);
}

void CMainFrame::OnMENUSearch() 
{
	HtmlHelp(NULL,AfxGetApp()->m_pszHelpFilePath,HH_HELP_FINDER,NULL);
}

////////////////////////////////////////////
//
// 响应"查看信息栏"
//
void CMainFrame::OnViewInformationBar() 
{
	//
	// 如果信息栏显示则隐藏
	// 隐藏则显示
	//
	if(this->m_ReportCalInfoBar.IsWindowVisible())
	{
		this->m_ReportCalInfoBar.HideDlgBar();
	}
	else
	{
		this->m_ReportCalInfoBar.ShowDlgBar();
	}
}

void CMainFrame::OnUpdateViewInformationBar(CCmdUI* pCmdUI) 
{
	if(this->m_ReportCalInfoBar.IsWindowVisible())
	{
		pCmdUI->SetCheck();
	}
	else
	{
		pCmdUI->SetCheck(0);
	}
}

//////////////////////////////////////////////////////
//
// 设置信息栏内的内容
//
// szInfo[in]	将显示在信息栏的内容
//
void CMainFrame::SetReportBarContent(LPCTSTR szInfo)
{
	this->m_ReportCalInfoBar.SetReportContent(szInfo);
}

//////////////////////////////////////////////////
//
// 返回信息栏内的内容
//
CString CMainFrame::GetReportBarContent()
{
	return this->m_ReportCalInfoBar.GetReportContent();
}

void CMainFrame::ShowReportBar(bool bShow)
{
	if(bShow)
	{
		this->m_ReportCalInfoBar.ShowDlgBar();
	}
	else
	{
		this->m_ReportCalInfoBar.HideDlgBar();
	}
}

/////////////////////////////////////////////
//
// 显示当前的工程名在标题栏上
//
void CMainFrame::ShowCurrentProjectName()
{
	_RecordsetPtr IRecordset;
	CString SQL,strTemp;
	HRESULT hr;
	_variant_t varTemp;

	//
	// 如果选择了工程将工程的名称显示在标题栏上
	//
	if(EDIBgbl::SelPrjID.IsEmpty())
	{
		this->SetWindowText(_T("优易保温油漆工程设计AutoIPED"));
	}
	else
	{
		hr=IRecordset.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{
			_com_error e(hr);
			ReportExceptionErrorLV1(e);
			return;
		}

		SQL.Format(_T("SELECT gcmc FROM Engin WHERE EnginID='%s'"),EDIBgbl::SelPrjID);

		try
		{
			IRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)(((CAutoIPEDApp*)::AfxGetApp())->m_pConAllPrj)),
							 adOpenDynamic,adLockOptimistic,adCmdText);

			if(IRecordset->adoEOF && IRecordset->BOF)
			{
				ReportExceptionErrorLV1(_T("没有此工程"));
				return;
			}

			varTemp=IRecordset->GetCollect(_variant_t("gcmc"));
		}
		catch(_com_error &e)
		{

			ReportExceptionErrorLV1(e);
			return;
		}
		//
		EDIBgbl::SelPrjName = vtos(varTemp);
		strTemp.Format(_T("优易保温油漆工程设计AutoIPED-[%s]    (%s %s)"),EDIBgbl::SelPrjName, EDIBgbl::\
			sCur_CodeName,EDIBgbl::sCur_CodeNO);

		this->SetWindowText(strTemp);
	}
}



void CMainFrame::OnClose() 
{
	//
	// 如果正在运算就关闭运算，并每隔0.3秒重新判断是否在运算
	// 如果停止运算了就关闭应用程序，否则继续判断
	//
	if(((CAutoIPEDDoc*)GetActiveDocument())->Operation())
	{
		((CAutoIPEDDoc*)GetActiveDocument())->OnButtonStopcal();

		this->SetTimer(ID_STOP_TIME,300,NULL);

		return;
	}

	CFrameWnd::OnClose();
}

void CMainFrame::OnTimer(UINT nIDEvent) 
{
	if(nIDEvent==ID_STOP_TIME)
	{
		//
		// 判断是否仍在计算，如果停止计算就关闭应用程序
		//
		if(((CAutoIPEDDoc*)GetActiveDocument())->Operation())
		{
			return;
		}

		KillTimer(nIDEvent);

		this->PostMessage(WM_CLOSE);
	}

	CFrameWnd::OnTimer(nIDEvent);
}
