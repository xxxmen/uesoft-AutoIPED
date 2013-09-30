// CalThread.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "CalThread.h"
#include "AutoIPEDDoc.h"

#include "EDIBgbl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CCalThread

IMPLEMENT_DYNCREATE(CCalThread, CWinThread)

CCalThread::CCalThread()
{
	m_IStream=NULL;
	m_IStreamCODE=NULL;
	m_pDoc=NULL;
}

CCalThread::~CCalThread()
{
}

BOOL CCalThread::InitInstance()
{
	::CoInitialize(NULL);
	{
		_ConnectionPtr IConnection;
		_ConnectionPtr IConnectionCODE;
		CString ProjectName;

		ProjectName=GetProjectName();

		CoGetInterfaceAndReleaseStream(m_IStream,__uuidof(_Connection),(LPVOID*)&IConnection);
		CoGetInterfaceAndReleaseStream(m_IStreamCODE,__uuidof(_Connection),(LPVOID*)&IConnectionCODE);

		m_pDoc->SetProjectName(ProjectName);
		m_pDoc->SetConnect(IConnection);
		m_pDoc->SetAssistantDbConnect(IConnectionCODE);
		m_pDoc->SetMaterialPath(EDIBgbl::sMaterialPath);

		try
		{
			m_pDoc->C_ANALYS();//±£ÎÂ·ÖÎö¼ÆËã
		}
		catch(_com_error)
		{
		}
		m_pDoc->IsStopCalc=TRUE;
		m_pDoc->OnEndCalculate();
	}
	return FALSE;
}

int CCalThread::ExitInstance()
{
	::CoUninitialize();
	return CWinThread::ExitInstance();
}

BEGIN_MESSAGE_MAP(CCalThread, CWinThread)
	//{{AFX_MSG_MAP(CCalThread)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCalThread message handlers

void CCalThread::SetProjectName(LPCTSTR ProjectName)
{
	m_ProjectName=ProjectName;
}

LPCTSTR CCalThread::GetProjectName()
{
	return m_ProjectName;
}
