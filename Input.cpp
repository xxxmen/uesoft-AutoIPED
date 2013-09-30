// Input.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "Input.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// Input dialog


Input::Input(CWnd* pParent /*=NULL*/)
	: CDialog(Input::IDD, pParent)
{
	//{{AFX_DATA_INIT(Input)
	m_static = _T("");
	m_input = 0.0;
	m_StringInput = _T("");
	//}}AFX_DATA_INIT

	m_IsStringInput=false;	// 默认的输入的不是字符串，即输入的是一个double值
}


void Input::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(Input)
	DDX_Text(pDX, IDC_EDIT_static, m_static);
	//}}AFX_DATA_MAP

	if(m_IsStringInput)
	{
		DDX_Text(pDX, IDC_EDIT_input, m_StringInput);
	}
	else
	{
		DDX_Text(pDX, IDC_EDIT_input, m_input);
	}
}


BEGIN_MESSAGE_MAP(Input, CDialog)
	//{{AFX_MSG_MAP(Input)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Input message handlers

///////////////////////////////////////////////
//
// 设置输入的是字符串还是数字
//
// IsStringInput[in]	如果为true输入的是字符串
//						如果为false输入的数字
//
void Input::SetStringInput(bool IsStringInput)
{
	m_IsStringInput=IsStringInput;
}

////////////////////////////////////////////
//
// 返回输入的是字符串还是数值
//
// 如果为true，获得的是字符串，访问m_StringInput获得
// 如果为false，获得的是数值，访问m_input获得
//
bool Input::GetStringInput()
{
	return m_IsStringInput;
}
