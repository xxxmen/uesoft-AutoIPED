// Page1.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "Page1.h"
#include "Columns.h"
#include "Column.h"
#include "EDIBgbl.h"

#include "Splits.h"
#include "Split.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPage1 property page

IMPLEMENT_DYNCREATE(CPage1, CPropertyPage)

CPage1::CPage1() : CPropertyPage(CPage1::IDD)
{
	//{{AFX_DATA_INIT(CPage1)
	m_comboStr = _T("");
	//}}AFX_DATA_INIT
}

CPage1::~CPage1()
{
}

void CPage1::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPage1)
	DDX_Control(pDX, IDC_COMBO1, m_conComboBox);
	DDX_Control(pDX, IDC_DATAGRID_A_COLR_L, m_a_colr_lGrid);
	DDX_Control(pDX, IDC_DATAGRID_A_COLOR, m_a_colorGrid);
	DDX_CBString(pDX, IDC_COMBO1, m_comboStr);
	DDX_Control(pDX, IDC_DATAGRID_A_PAI, m_a_paiGrid);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPage1, CPropertyPage)
	//{{AFX_MSG_MAP(CPage1)
	ON_CBN_KILLFOCUS(IDC_COMBO1, OnKillfocusCombo1)
	ON_CBN_SELCHANGE(IDC_COMBO1, OnSelchangeCombo1)
	ON_WM_CLOSE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPage1 message handlers
extern BOOL SetFieldCaptionAndWidth(CDataGridEx* pDataGrid, _ConnectionPtr& pConSort, CString strTblName);
extern BOOL SaveDataGridFieldWidth(CDataGridEx* pDataGrid, _ConnectionPtr& pConSort, CString strTblName);

extern CAutoIPEDApp theApp;
BOOL CPage1::OnInitDialog()
{
	CDialog::OnInitDialog();
	m_a_colorSet.CreateInstance(__uuidof(Recordset));
	m_a_colorSet->CursorLocation =adUseClient;
	m_a_colr_lSet.CreateInstance(__uuidof(Recordset));
	m_a_colr_lSet->CursorLocation = adUseClient;
	CString strSQL;
	strSQL="select * from a_colr_l";
	m_a_colr_lSet->Open(_bstr_t(strSQL),(IDispatch*)(theApp.m_pConnectionCODE),adOpenStatic,adLockOptimistic,adCmdText);
	m_a_colr_lGrid.SetRefDataSource(NULL);
	m_a_colr_lGrid.SetAllowAddNew(TRUE);
	m_a_colr_lGrid.SetAllowUpdate(TRUE);
	m_a_colr_lGrid.SetAllowDelete(TRUE);
	m_a_colr_lGrid.SetRefDataSource(m_a_colr_lSet->GetDataSource());
	m_a_colr_lGrid.InitMouseWheel();//滚动功能
	
	//设置字段名称和宽度。
	SetFieldCaptionAndWidth(&m_a_colr_lGrid, theApp.m_pConIPEDsort, "a_colr_l");


	strSQL="select * from a_color where	enginid='"+EDIBgbl::SelPrjID+"'";
	m_a_colorSet->Open(_bstr_t(strSQL),(IDispatch*)theApp.m_pConnection,adOpenStatic,adLockOptimistic,adCmdText);
	m_a_colorGrid.SetRefDataSource(m_a_colorSet->GetDataSource());
	m_a_colorGrid.GetColumns().GetItem(_variant_t("enginid")).SetVisible(false);
	m_a_colorGrid.InitMouseWheel();//滚动功能 
	m_a_colorGrid.SetAllowAddNew(true);
	m_a_colorGrid.SetAllowUpdate(true);
	m_a_colorGrid.SetAllowDelete(true);

	//设置字段名称和宽度。
	SetFieldCaptionAndWidth(&m_a_colorGrid, theApp.m_pConIPEDsort, "a_color");

/*	RECT rect;
	float fRect;
	m_a_colorGrid.GetWindowRect(&rect);
	fRect=rect.right-rect.left;
	int i;
	int ii;
	i=fRect/7;
	ii=fRect/6;
	m_a_colorGrid.GetColumns().GetItem(_variant_t("colr_media")).SetWidth(i*2);
	m_a_colorGrid.GetColumns().GetItem(_variant_t("colr_face")).SetWidth(ii);
	m_a_colorGrid.GetColumns().GetItem(_variant_t("colr_ring")).SetWidth(ii);
	m_a_colorGrid.GetColumns().GetItem(_variant_t("colr_arrow")).SetWidth(ii);
	m_a_colorGrid.GetColumns().GetItem(_variant_t("colr_code")).SetWidth(ii*0.8);
*/
	
	if(!InitA_pai())
	{
		return false;
	}
	InitComboBox();
	m_conComboBox.SetCurSel(0);
	m_conComboBox.ShowWindow(SW_HIDE);
	m_a_colr_lDeleteState=false;
	m_a_colorDeleteState=false;

	return true;

}

BEGIN_EVENTSINK_MAP(CPage1, CPropertyPage)
    //{{AFX_EVENTSINK_MAP(CPage1)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLOR, -607 /* MouseUp */, OnMouseUpDatagridAColor, VTS_I2 VTS_I2 VTS_I4 VTS_I4)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLOR, 218 /* RowColChange */, OnRowColChangeDatagridAColor, VTS_PVARIANT VTS_I2)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLOR, 217 /* OnAddNew */, OnOnAddNewDatagridAColor, VTS_NONE)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLR_L, 204 /* AfterUpdate */, OnAfterUpdateDatagridAColrL, VTS_NONE)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLR_L, 209 /* BeforeUpdate */, OnBeforeUpdateDatagridAColrL, VTS_PI2)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLOR, 112 /* Scroll */, OnScrollDatagridAColor, VTS_PI2)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLR_L, 205 /* BeforeColEdit */, OnBeforeColEditDatagridAColrL, VTS_I2 VTS_I2 VTS_PI2)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLR_L, 207 /* BeforeDelete */, OnBeforeDeleteDatagridAColrL, VTS_PI2)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLR_L, 202 /* AfterDelete */, OnAfterDeleteDatagridAColrL, VTS_NONE)
	ON_EVENT(CPage1, IDC_DATAGRID_A_PAI, 217 /* OnAddNew */, OnOnAddNewDatagridAPai, VTS_NONE)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLOR, 209 /* BeforeUpdate */, OnBeforeUpdateDatagridAColor, VTS_PI2)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLOR, 202 /* AfterDelete */, OnAfterDeleteDatagridAColor, VTS_NONE)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLOR, 207 /* BeforeDelete */, OnBeforeDeleteDatagridAColor, VTS_PI2)
	ON_EVENT(CPage1, IDC_DATAGRID_A_PAI, -607 /* MouseUp */, OnMouseUpDatagridAPai, VTS_I2 VTS_I2 VTS_I4 VTS_I4)
	ON_EVENT(CPage1, IDC_DATAGRID_A_COLR_L, -607 /* MouseUp */, OnMouseUpDatagridAColrL, VTS_I2 VTS_I2 VTS_I4 VTS_I4)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()



bool CPage1::InitA_pai()
{
	m_a_paiSet.CreateInstance(__uuidof(Recordset));
	CString strSQL;
	try
	{
		m_a_paiSet->CursorLocation =adUseClient;
		strSQL="select * from a_pai where enginid='"+EDIBgbl::SelPrjID+"'";
		m_a_paiSet->Open(_bstr_t(strSQL),(IDispatch*)theApp.m_pConnection,adOpenStatic,adLockOptimistic,adCmdText);
		m_a_paiGrid.SetRefDataSource(m_a_paiSet->GetDataSource());
		m_a_paiGrid.GetColumns().GetItem(_variant_t("enginid")).SetVisible(false);
		m_a_paiGrid.InitMouseWheel();
		m_a_paiGrid.SetAllowAddNew(true);
		m_a_paiGrid.SetAllowDelete(true);
		m_a_paiGrid.SetAllowUpdate(true);

		//设置字段的标题和宽度
		SetFieldCaptionAndWidth(&m_a_paiGrid, theApp.m_pConIPEDsort, "a_pai");
	}
	catch(_com_error e)
	{
		e.Description();
		return false;
	}
	return true;
}

BOOL CPage1::InitComboBox()
{
	_variant_t key;
	CString str;
	_RecordsetPtr a_colr_lSet;
	a_colr_lSet.CreateInstance(__uuidof(Recordset));
	a_colr_lSet->Open(_bstr_t("select * from a_colr_l"),(IDispatch*)(theApp.m_pConnectionCODE),adOpenStatic,adLockOptimistic,adCmdText);
	if(!(a_colr_lSet->GetRecordCount()>0))
	{
		return false;
	}
	m_conComboBox.ResetContent();
	try
	{
		for(a_colr_lSet->MoveFirst();!a_colr_lSet->adoEOF;a_colr_lSet->MoveNext())
		{
			key=a_colr_lSet->GetCollect(_T("colr_name"));
			str=(key.vt==VT_EMPTY||key.vt==VT_NULL)?"":(CString)key.bstrVal;
			m_conComboBox.AddString(str);
		}
		m_conComboBox.AddString("N");
		if(a_colr_lSet->State==adStateOpen)
		{
			a_colr_lSet->Close();
		}
	}
	catch(_com_error e)
	{
		e.Description();
		return false;
	}
	return true;
}

void CPage1::OnKillfocusCombo1() 
{
	m_conComboBox.BringWindowToTop();
	
}

void CPage1::OnMouseUpDatagridAColor(short Button, short Shift, long X, long Y) 
{
	m_conComboBox.ShowWindow(SW_HIDE);
	CRect dbRc,      //DataGrid的坐标。
    comRc,      //ComBox 在对话框里的坐标
    dlgRc;         //格式对话框在屏幕上的坐标

	m_lastCol=m_a_colorGrid.GetCol();
	POINT  m_point;                                                     
	HWND   m_hWnd;
	::GetCursorPos(&m_point);            //获得光标在屏幕上的坐标。         
	m_hWnd = ::WindowFromPoint(m_point);   //得到该坐标所在的窗口的名柄。
	::GetWindowRect(m_hWnd, &comRc);         
	this->ScreenToClient(&comRc);            //
	comRc.bottom += 100;  
	CString str;
	str=m_a_colorGrid.GetColumns().GetItem((_variant_t)m_lastCol).GetDataField();
	//10/26
	_variant_t key;
	CString str1;
	short indexCombo;
	//10/26
	if( (m_hWnd !=  m_a_colorGrid.m_hWnd)&&(str=="COLR_FACE"||str=="COLR_RING"||str=="COLR_ARROW") )
	{
		m_conComboBox.MoveWindow(&comRc);
		m_conComboBox.SetFocus();
    	m_conComboBox.ShowWindow(SW_SHOW);
		InitComboBox();
		key=m_a_colorSet->GetCollect(_variant_t(m_lastCol));
		str1=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"N":(CString)key.bstrVal;
		indexCombo=m_conComboBox.FindString(0,str1);
		m_conComboBox.SetCurSel(indexCombo);

	}
	
}

void CPage1::OnRowColChangeDatagridAColor(VARIANT FAR* LastRow, short LastCol) 
{
	m_lastCol=LastCol;		
}

void CPage1::OnSelchangeCombo1() 
{
	if(m_conComboBox.IsWindowVisible())
	{
		int i = m_conComboBox.GetCurSel();
		if(i!=LB_ERR)
		{
			CString str;
			m_conComboBox.GetLBText(i,str);
			m_a_colorGrid.GetColumns().GetItem(_variant_t((long)m_lastCol)).SetText(str);
		}
		
	}
}

void CPage1::OnOnAddNewDatagridAColor() 
{
	m_a_colorGrid.GetColumns().GetItem(_variant_t("enginid")).SetText(EDIBgbl::SelPrjID);	
}

void CPage1::OnAfterUpdateDatagridAColrL() 
{

	_variant_t key;
	key=m_a_colr_lSet->GetCollect(_T("colr_name"));
	m_curColr_name=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"":(CString)key.bstrVal;
	if(m_curColr_name.Compare(m_lastColr_name))
	{
		changeA_colorField(m_lastColr_name);
	}
	m_a_colr_lDeleteState=false;
}

void CPage1::OnBeforeUpdateDatagridAColrL(short FAR* Cancel) 
{
	if(m_a_colr_lDeleteState)
	{
		return;
	}
	_variant_t key;
	CString strTemp;
	key=m_a_colr_lSet->GetCollect(_T("colr_name"));
	strTemp=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"":(CString)key.bstrVal;
	if(strTemp.IsEmpty())
	{
		*Cancel=1;
		AfxMessageBox("提示：\n调和漆单位耗量库中的字段：COLR_NAME不能为空，操作已取消!");
		m_a_colr_lSet->CancelUpdate();
		return;
	}
	
}

void CPage1::OnScrollDatagridAColor(short FAR* Cancel) 
{
	m_conComboBox.ShowWindow(SW_HIDE);
	
}

void CPage1::changeA_colorField(CString str)
{
	_variant_t key;
	CString strColr_face,strColr_ring,strColr_arrow;
	try
	{
		if(m_a_colorSet->GetRecordCount()<=0)
		{
			return;
		}
		for(m_a_colorSet->MoveFirst();!m_a_colorSet->adoEOF;m_a_colorSet->MoveNext())
		{
			key=m_a_colorSet->GetCollect(_T("colr_face"));
			strColr_face=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"N":(CString)key.bstrVal;
			if(strColr_face.Compare(str)==0)
			{
				m_a_colorSet->PutCollect(_T("colr_face"),_variant_t(m_curColr_name));
			}

			key=m_a_colorSet->GetCollect(_T("colr_ring"));
			strColr_ring=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"N":(CString)key.bstrVal;
			if(strColr_ring.Compare(str)==0)
			{
				m_a_colorSet->PutCollect(_T("colr_ring"),_variant_t(m_curColr_name));
			}

			key=m_a_colorSet->GetCollect(_T("colr_arrow"));
			strColr_arrow=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"N":(CString)key.bstrVal;
			if(strColr_arrow.Compare(str)==0)
			{
				m_a_colorSet->PutCollect(_T("colr_arrow"),_variant_t(m_curColr_name));
			}
			m_a_colorSet->Update();

		}
	}
	catch(_com_error e)
	{
		e.Description();
		return;
	}
}

void CPage1::OnBeforeColEditDatagridAColrL(short ColIndex, short KeyAscii, short FAR* Cancel) 
{
	_variant_t key;
	key=m_a_colr_lSet->GetCollect(_T("colr_name"));
	m_lastColr_name=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"":(CString)key.bstrVal;
}

void CPage1::OnBeforeDeleteDatagridAColrL(short FAR* Cancel) 
{
	_variant_t key;
	key=m_a_colr_lSet->GetCollect(_T("colr_name"));
	m_lastColr_name=(key.vt==VT_NULL||key.vt==VT_EMPTY)?"":(CString)key.bstrVal;
	m_curColr_name="N";
	m_a_colr_lDeleteState=true;

}

void CPage1::OnAfterDeleteDatagridAColrL() 
{
	changeA_colorField(m_lastColr_name);	
}

void CPage1::OnOnAddNewDatagridAPai() 
{
	m_a_paiGrid.GetColumns().GetItem(_variant_t("enginid")).SetText(EDIBgbl::SelPrjID);	
}

void CPage1::OnBeforeUpdateDatagridAColor(short FAR* Cancel) 
{
	if(m_a_colorDeleteState)
	{
		return;
	}
	_variant_t key;
	double dbTemp;
	key=m_a_colorSet->GetCollect(_T("colr_code"));
	dbTemp=(key.vt==VT_NULL||key.vt==VT_EMPTY)?-1.0:key.dblVal;
	if(dbTemp==-1.0)
	{
		*Cancel=1;
		AfxMessageBox("提示：\n色环箭头油漆准则库中的字段：COLR_CODE不能为空!操作已取消!");
		m_a_colorSet->CancelUpdate();
		return;
	}
	
}

void CPage1::OnAfterDeleteDatagridAColor() 
{
	m_a_colorDeleteState=false;	
}

void CPage1::OnBeforeDeleteDatagridAColor(short FAR* Cancel) 
{
	m_a_colorDeleteState=true;	
}

void CPage1::OnMouseUpDatagridAPai(short Button, short Shift, long X, long Y) 
{
	m_conComboBox.ShowWindow(SW_HIDE);	
}

void CPage1::OnMouseUpDatagridAColrL(short Button, short Shift, long X, long Y) 
{
	m_conComboBox.ShowWindow(SW_HIDE);	
}

//------------------------------------------------------------------                
// DATE         :[2005/05/27]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :释放钩子的内存。
//------------------------------------------------------------------
short CPage1::UnMouseWheel()
{
	m_a_colr_lGrid.UnInitMouseWheel();		//释放钩子的内存
	m_a_colorGrid.UnInitMouseWheel();		//释放钩子的内存
	m_a_paiGrid.UnInitMouseWheel();
	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/17]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :保存当前环境的设置
//------------------------------------------------------------------
void CPage1::SaveSetting()
{
	SaveDataGridFieldWidth(&m_a_paiGrid, theApp.m_pConIPEDsort,"a_pai");
	SaveDataGridFieldWidth(&m_a_colr_lGrid, theApp.m_pConIPEDsort,"a_colr_l");
	SaveDataGridFieldWidth(&m_a_colorGrid, theApp.m_pConIPEDsort,"a_color");	
}

void CPage1::OnCancel() 
{
	// TODO: Add your message handler code here and/or call default
	
	CPropertyPage::OnCancel();
}
