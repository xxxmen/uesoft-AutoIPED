// RangeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "RangeDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// RangeDlg dialog


RangeDlg::RangeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(RangeDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(RangeDlg)
	m_start = 0;
	m_end = 0;
	//}}AFX_DATA_INIT
}


void RangeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(RangeDlg)
	DDX_Control(pDX, IDC_SPIN_start, c_start);
	DDX_Control(pDX, IDC_SPIN_end, c_end);
	DDX_Text(pDX, IDC_EDIT_start, m_start);
	DDX_Text(pDX, IDC_EDIT_end, m_end);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(RangeDlg, CDialog)
	//{{AFX_MSG_MAP(RangeDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// RangeDlg message handlers

BOOL RangeDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here

	c_start.SetRange(1,(int)m_end);
	c_end.SetRange(1,(int)m_end);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void RangeDlg::OnOK() 
{
	UpdateData();

	int nStart,nEnd,nTmp;
	c_start.GetRange(nStart,nEnd);

	if (m_end < m_start)
	{
		nTmp    = m_start;
		m_start = m_end;
		m_end   = nTmp;
		UpdateData(FALSE);
	}
	if (m_start < nStart || m_end  > nEnd)
	{
		AfxMessageBox("输入的数字不在范围之内!");
		return;
	}	
	CDialog::OnOK();
}
