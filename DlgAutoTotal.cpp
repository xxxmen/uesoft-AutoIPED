// DlgAutoTotal.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DlgAutoTotal.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDlgAutoTotal dialog


CDlgAutoTotal::CDlgAutoTotal(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgAutoTotal::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgAutoTotal)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CDlgAutoTotal::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgAutoTotal)
	DDX_Control(pDX, IDC_CHECK1, m_cTemperatureExplain);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgAutoTotal, CDialog)
	//{{AFX_MSG_MAP(CDlgAutoTotal)
	ON_BN_CLICKED(IDC_CHECK1, OnTemperatureExplain)
	ON_BN_CLICKED(IDC_CHECK2, OnPaintExplain)
	ON_BN_CLICKED(IDC_CHECK3, OnTemperatureTotal1)
	ON_BN_CLICKED(IDC_CHECK4, OnPaintTotal)
	ON_BN_CLICKED(IDC_CHECK5, OnTempratureTotal2)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgAutoTotal message handlers

BOOL CDlgAutoTotal::OnInitDialog()
{
	CDialog::OnInitDialog();
	((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK6))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK7))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK8))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK9))->SetCheck(true);
	((CButton*)GetDlgItem(IDC_CHECK10))->SetCheck(true);
	int i=0;
	for(;i<15;i++)
	{
		m_autoTotal[i]=false;
	}

	return true;
}

//生成说明表的消息响应函数
void CDlgAutoTotal::OnTemperatureExplain() 
{
	if(!((CButton*)GetDlgItem(IDC_CHECK1))->GetCheck())
	{
		((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK2))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK3))->EnableWindow(false);
		
		((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK4))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK5))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_CHECK6))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK6))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_CHECK7))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK7))->EnableWindow(false);
		
		((CButton*)GetDlgItem(IDC_CHECK8))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK8))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_CHECK9))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK9))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_CHECK10))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK10))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO1))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO2))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO2))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO3))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO3))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO4))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO4))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO5))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO5))->EnableWindow(false);
		((CButton*)GetDlgItem(IDOK))->EnableWindow(false);

	}
	else
	{
		((CButton*)GetDlgItem(IDC_CHECK2))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK3))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK4))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK5))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK6))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK6))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK7))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK7))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK8))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK8))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK9))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK9))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK10))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK10))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO1))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO2))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO3))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO4))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO5))->EnableWindow(true);
		((CButton*)GetDlgItem(IDOK))->EnableWindow(true);
	
	}
}

//生成油漆说明表消息响应函数
void CDlgAutoTotal::OnPaintExplain() 
{
	if(((CButton*)GetDlgItem(IDC_CHECK2))->GetCheck())
	{
		((CButton*)GetDlgItem(IDC_CHECK4))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK7))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK7))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_CHECK9))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK9))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO2))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO2))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO4))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO4))->SetCheck(false);

		((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO3))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO5))->SetCheck(false);
	}
	else
	{
		((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK4))->EnableWindow(false);
		
		((CButton*)GetDlgItem(IDC_CHECK7))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK7))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_CHECK9))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK9))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO2))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO2))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO4))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO4))->EnableWindow(false);
		
	}
}

//生成保温材料汇总表(标准+辅助)
void CDlgAutoTotal::OnTemperatureTotal1() 
{
	if(((CButton*)GetDlgItem(IDC_CHECK3))->GetCheck())
	{
		((CButton*)GetDlgItem(IDC_CHECK8))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK8))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO3))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO3))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO2))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO4))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO5))->SetCheck(false);

	}
	else
	{
		((CButton*)GetDlgItem(IDC_CHECK8))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK8))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO3))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO3))->EnableWindow(false);
	}
	
}

//生成油漆材料汇总表(标准+辅助)的消息响应函数
void CDlgAutoTotal::OnPaintTotal() 
{
	if(((CButton*)GetDlgItem(IDC_CHECK4))->GetCheck())
	{
		((CButton*)GetDlgItem(IDC_CHECK9))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK9))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO4))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO4))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO2))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO3))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO5))->SetCheck(false);

	}
	else
	{
		((CButton*)GetDlgItem(IDC_CHECK9))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK9))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO4))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO4))->EnableWindow(false);
	}
		
}

void CDlgAutoTotal::OnTempratureTotal2() 
{
	if(((CButton*)GetDlgItem(IDC_CHECK5))->GetCheck())
	{
		((CButton*)GetDlgItem(IDC_CHECK10))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_CHECK10))->SetCheck(true);

		((CButton*)GetDlgItem(IDC_RADIO5))->EnableWindow(true);
		((CButton*)GetDlgItem(IDC_RADIO5))->SetCheck(true);
		((CButton*)GetDlgItem(IDC_RADIO1))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO2))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO3))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO4))->SetCheck(false);

	}
	else
	{
		((CButton*)GetDlgItem(IDC_CHECK10))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_CHECK10))->EnableWindow(false);

		((CButton*)GetDlgItem(IDC_RADIO5))->SetCheck(false);
		((CButton*)GetDlgItem(IDC_RADIO5))->EnableWindow(false);
	}	
}

void CDlgAutoTotal::OnOK() 
{
	m_autoTotal[1]=(((CButton*)GetDlgItem(IDC_CHECK1))->GetCheck())?true:false;

	m_autoTotal[2]=(((CButton*)GetDlgItem(IDC_CHECK2))->GetCheck())?true:false;

	m_autoTotal[3]=(((CButton*)GetDlgItem(IDC_CHECK3))->GetCheck())?true:false;

	m_autoTotal[4]=(((CButton*)GetDlgItem(IDC_CHECK4))->GetCheck())?true:false;

	m_autoTotal[5]=(((CButton*)GetDlgItem(IDC_CHECK5))->GetCheck())?true:false;

	m_autoTotal[6]=(((CButton*)GetDlgItem(IDC_CHECK6))->GetCheck())?true:false;

	m_autoTotal[7]=(((CButton*)GetDlgItem(IDC_CHECK7))->GetCheck())?true:false;

	m_autoTotal[8]=(((CButton*)GetDlgItem(IDC_CHECK8))->GetCheck())?true:false;

	m_autoTotal[9]=(((CButton*)GetDlgItem(IDC_CHECK9))->GetCheck())?true:false;

	m_autoTotal[10]=(((CButton*)GetDlgItem(IDC_CHECK10))->GetCheck())?true:false;

	m_autoTotal[11]=(((CButton*)GetDlgItem(IDC_RADIO1))->GetCheck())?true:false;

	m_autoTotal[12]=(((CButton*)GetDlgItem(IDC_RADIO2))->GetCheck())?true:false;

	m_autoTotal[13]=(((CButton*)GetDlgItem(IDC_RADIO3))->GetCheck())?true:false;

	m_autoTotal[14]=(((CButton*)GetDlgItem(IDC_RADIO4))->GetCheck())?true:false;

	m_autoTotal[0]=(((CButton*)GetDlgItem(IDC_RADIO5))->GetCheck())?true:false;

	CDialog::OnOK();
}
