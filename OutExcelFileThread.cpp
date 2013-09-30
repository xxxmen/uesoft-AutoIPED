// OutExcelFileThread.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "OutExcelFileThread.h"
#include "EDIBgbl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// COutExcelFileThread

IMPLEMENT_DYNCREATE(COutExcelFileThread, CWinThread)

COutExcelFileThread::COutExcelFileThread()
{
	m_pCurView = NULL;
}

COutExcelFileThread::~COutExcelFileThread()
{
}

BOOL COutExcelFileThread::InitInstance()
{
	::CoInitialize(NULL);
	if( m_pCurView == NULL )
	{
		return FALSE;
	}
//	::SendMessage(m_pCurView->m_hWnd, WM_COMMAND, ID_MENUITEM43, 0); // 输出保温说明表到EXCEL
	m_pCurView->DisplayRemarksInfo("");
	return TRUE;
}

int COutExcelFileThread::ExitInstance()
{
	// TODO:  perform any per-thread cleanup here
	::CoUninitialize();
	return CWinThread::ExitInstance();
}

BEGIN_MESSAGE_MAP(COutExcelFileThread, CWinThread)
	//{{AFX_MSG_MAP(COutExcelFileThread)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// COutExcelFileThread message handlers

void COutExcelFileThread::PutCurView(CAutoIPEDView *pCurView)
{
	m_pCurView = pCurView;
}

void COutExcelFileThread::OutInsulExplainTbl()
{
	if ( m_pCurView != NULL )
	{
//		::SendMessage(m_pCurView->m_hWnd, WM_COMMAND, ID_MENUITEM43, 0); // 输出保温说明表到EXCEL
	}
}

bool COutExcelFileThread::SetPow(long double _x, long _y)
{
	typedef struct E_Node
	{
		int j;
		double q;
		E_Node* pNext;
	}E_Node;
	typedef struct E_G
	{
		int count;
		E_Node* ListHead[20];
	}E_G;
	typedef struct Group
	{
		int c;
		E_G G;
	}Group;
	int ui=0;
	E_Node * tp;
	Group Group1;
	tp = Group1.G.ListHead[ui];
	while ( tp != NULL )
	{
		if ( tp->j == 0 )
		{
			// 
		}
		tp = tp->pNext;
	}
	
	

	return 1;
}
