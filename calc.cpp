// calc.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "calc.h"
#include "HeatPreCal.h"
#include "Edibgbl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern CAutoIPEDApp theApp;
/////////////////////////////////////////////////////////////////////////////
// calc dialog


calc::calc(CWnd* pParent /*=NULL*/)
	: CDialog(calc::IDD, pParent)
{
	//{{AFX_DATA_INIT(calc)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void calc::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(calc)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(calc, CDialog)
	//{{AFX_MSG_MAP(calc)
	ON_BN_CLICKED(IDC_BUTTONout, OnBUTTONout)
	ON_BN_CLICKED(IDC_BUTTONcountall, OnBUTTONcountall)
	ON_BN_CLICKED(IDC_BUTTON_range, OnBUTTONrange)
	ON_BN_CLICKED(IDC_BUTTON_custom, OnBUTTONcustom)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// calc message handlers


/////////////////////////////////////////////////////
//   退出按纽
///////////////////////////////////////////
void calc::OnBUTTONout() 
{
	// TODO: Add your control notification handler code here
	i=4;
	OnCancel();
}

/////////////////////////////////////////
//    全部计算按纽
//////////////////////////////////////////
void calc::OnBUTTONcountall() 
{
	// TODO: Add your control notification handler code here
    i=1;
	OnOK();
}

////////////////////////////////////////
//    范围计算按纽
////////////////////////////////////////
void calc::OnBUTTONrange() 
{
	i=2;
	OnOK();
}

////////////////////////////////////
//   自定义计算按纽
/////////////////////////////////////////
void calc::OnBUTTONcustom() 
{
	// TODO: Add your control notification handler code here
	i=3;
	OnOK();
}
