// DlgShowAllTable.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DlgShowAllTable.h"
#include "DataShowDlg.h"
#include "ExceptionInfohandle.h"
#include "vtot.h"
#include "EDIBgbl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
extern CAutoIPEDApp theApp;
/////////////////////////////////////////////////////////////////////////////
// CDlgShowAllTable dialog


CDlgShowAllTable::CDlgShowAllTable(CWnd* pParent /*=NULL*/)
: CDataShowDlg(pParent)
{
	//{{AFX_DATA_INIT(CDlgShowAllTable)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	m_IsInHDragArea = false;
	m_IsInVDragArea = false;
	m_ListSize.cx=200;
	m_ListSize.cy =350;
}

/*
void CDlgShowAllTable::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgShowAllTable)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}
*/

BEGIN_MESSAGE_MAP(CDlgShowAllTable, CDataShowDlg)
	//{{AFX_MSG_MAP(CDlgShowAllTable)
	ON_WM_MOUSEMOVE()
	ON_WM_SETCURSOR()
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	//}}AFX_MSG_MAP
	ON_LBN_SELCHANGE(IDC_MATERIALNAME_LIST, OnTableSelChange)  //选择不同的表.
	ON_BN_CLICKED(IDC_DELETE_SEL_MATERIAL,  OnDeleteTblAll)		//删除表中的所有记录
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgShowAllTable message handlers

BOOL CDlgShowAllTable::tableExists(_ConnectionPtr pCon, CString tbn)
{
	if(pCon==NULL || tbn=="")
		return false;
	_RecordsetPtr rs;
	if(tbn.Left(1)!="[")
		tbn="["+tbn;
	if(tbn.Right(1)!="]")
		tbn+="]";
	try{
		rs=pCon->Execute(_bstr_t(tbn),NULL,adCmdTable);
		rs->Close();
		return true;
	}
	catch(_com_error e)
	{
		return false;
	}

}


BOOL CDlgShowAllTable::OnInitDialog() 
{
	CDataShowDlg::OnInitDialog();
	CRect ClientRect;
	CFont *pFont;

	m_HDragArea.SetRectEmpty();
	m_VDragArea.SetRectEmpty();

	GetWindowRect(&ClientRect);
	m_ListSize.cx = ClientRect.Width()/4;
	m_ListSize.cy = (long)(ClientRect.Height()/1.5);

/*
	try
	{
		if(m_pRsBrowseResult == NULL)
		{
			m_pRsBrowseResult.CreateInstance(__uuidof(Recordset));
			m_pRsBrowseResult->CursorLocation = adUseClient;
		}
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

*/	pFont = GetFont();
	//
	//创建各种表格名称List控件
	//
	m_TableNameList.Create(LBS_NOTIFY|WS_VISIBLE|WS_BORDER|WS_CHILD|WS_VSCROLL|LBS_NOINTEGRALHEIGHT,
									ClientRect,this,IDC_MATERIALNAME_LIST);
	//	m_MaterialNameList.SetFont(pFont);
	m_TableNameList.SetFont(pFont);
	m_TableNameList.ShowWindow(SW_SHOW);
	//创建删除全部记录的按钮.
	this->m_ctlDeleteAll.Create("全部删除", BS_PUSHBUTTON|WS_CHILD|WS_VISIBLE|WS_CLIPSIBLINGS,
								ClientRect,this,IDC_DELETE_SEL_MATERIAL);
	m_ctlDeleteAll.SetFont( pFont );
	m_ctlDeleteAll.SetTooltipText("删除选择的表中的所有记录！", TRUE);

	//初始须要显示的各表名称
	if( !UpdateTblNameList() )
	{
		return false;
	}
	// 更新控件的内容
	//初始为第一条记录。
	m_TableNameList.SetCurSel(0);
	m_nCurTblID = m_TableNameList.GetItemData(0);
	m_TableNameList.GetText(0, this->m_strDescTblName);
	UpdateDataGrid();

	// 更新各个控件的位置
	UpdateControlsPos();
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
///////////////////////////////////////////////
//
// 更新各个控件的位置
//
BOOL CDlgShowAllTable::UpdateControlsPos()
{
	CRect ClientRect;
	GetClientRect(&ClientRect);
//	int ctlWidth, left, right;
	GetClientRect(&ClientRect);

	//
	// 如果可横向拖动的区域为空就设置它
	//
	if(this->m_HDragArea.IsRectEmpty())
	{
		this->m_HDragArea=ClientRect;
		this->m_HDragArea.left=ClientRect.left+m_ListSize.cx;
		this->m_HDragArea.right=this->m_HDragArea.left+5;
	}

	//设置显示结果的DataGrid控件.
	ClientRect.left=m_HDragArea.right;
	this->m_ResultBrowseControl.MoveWindow(&ClientRect);

	//设置显示各表名称的位置
	GetClientRect(&ClientRect);
	ClientRect.right = ClientRect.left+m_ListSize.cx;
	ClientRect.bottom = ClientRect.top+m_ListSize.cy;
	this->m_TableNameList.MoveWindow(&ClientRect);

	//竖向区域
	this->m_VDragArea = ClientRect;
	this->m_VDragArea.top = m_VDragArea.bottom;
	this->m_VDragArea.bottom = this->m_VDragArea.top+5;

	//删除按钮.
	ClientRect.top = m_VDragArea.bottom + 10;
	ClientRect.bottom = ClientRect.top + 20;
	short width = ClientRect.Width();
	ClientRect.left += width/3;
	ClientRect.right -= width/3;
	this->m_ctlDeleteAll.MoveWindow(&ClientRect);

	//重画有用.
	this->Invalidate();
	return true;
}
//
//初始须要显示的各表名称.
//
bool CDlgShowAllTable::UpdateTblNameList()
{

/*	_RecordsetPtr pRs;
	pRs.CreateInstance(__uuidof(Recordset));
	pRs->CursorLocation = adUseClient;
*/	CString strSql, strTemp;

	if( m_pRsTableList == NULL )
	{
		m_pRsTableList.CreateInstance(__uuidof(Recordset));
	}
	if( m_pRsTableList->State == adStateOpen )
	{
		m_pRsTableList->Close();
	}
	strSql = "SELECT * FROM [TableINFO] WHERE SEQ IS NOT NULL ORDER BY SEQ";
	try
	{
		this->m_TableNameList.ResetContent();


		m_pRsTableList->Open(_bstr_t(strSql), theApp.m_pConIPEDsort.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		int count = m_pRsTableList->GetRecordCount();
		if( m_pRsTableList->GetRecordCount() <= 0 )
		{
			AfxMessageBox("没有数据!");
			return false;
		}
		//将表的描述加入到列表框中
		int Index;
		long dwItemData;
		for(; !m_pRsTableList->adoEOF; m_pRsTableList->MoveNext() )
		{
			strTemp = vtos(m_pRsTableList->GetCollect(_variant_t("ShortDescribe")));
			if( !strTemp.IsEmpty() )
			{
				Index = m_TableNameList.AddString(strTemp);
				if( Index != -1 )
				{
					//惟一标示一个表的序号.
					dwItemData = vtoi(m_pRsTableList->GetCollect(_variant_t("SEQ")));
					m_TableNameList.SetItemData(Index, dwItemData);
				}
			}
		}
//		this->SetRecordset(m_pRsTableList);
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return false;
	}

	return true;
}
//
// 如果在可拖动的区域就改变各个单元的大小与位置
//
void CDlgShowAllTable::OnMouseMove(UINT nFlags, CPoint point) 
{
	if( this->m_IsInHDragArea && this->m_LButtonPt != point )
	{
		//在横向拖动区域内.
		m_HDragArea.OffsetRect(point.x - m_LButtonPt.x, 0 );
		m_ListSize.cx += (point.x - m_LButtonPt.x);
		//更新各控件的位置
		UpdateControlsPos();
		m_LButtonPt = point;
	}
	else
	{
		if( m_IsInVDragArea && m_LButtonPt != point )
		{
			//在竖向拖动区域内.
			m_ListSize.cy += (point.y-m_LButtonPt.y);
			//更新各控件的位置
			UpdateControlsPos();
			m_LButtonPt = point;
		}
	}
	CDataShowDlg::OnMouseMove(nFlags, point);
}
// 在可拖动的区域改变合适个光标
BOOL CDlgShowAllTable::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	HCURSOR hCursor;
	CPoint   pt;
	::GetCursorPos(&pt);
	ScreenToClient(&pt);
	//
	// 在可拖动的区域改变合适个光标
	//
	if( this->m_HDragArea.PtInRect(pt) )
	{
		hCursor = LoadCursor(NULL,IDC_SIZEWE);
		SetCursor(hCursor);
		::DeleteObject(hCursor);
		return TRUE;
	}
	else
	{
		if( this->m_VDragArea.PtInRect(pt) )
		{
			hCursor = LoadCursor(NULL, IDC_SIZENS);
			SetCursor(hCursor);
			::DeleteObject(hCursor);
			return TRUE;
		}
	}

	return CDataShowDlg::OnSetCursor(pWnd, nHitTest, message);
}

void CDlgShowAllTable::OnSize(UINT nType, int cx, int cy) 
{
	CDataShowDlg::OnSize(nType, cx, cy);

	if( IsWindow(this->m_TableNameList.GetSafeHwnd()) && IsWindow(this->m_ResultBrowseControl.GetSafeHwnd()) )
	{
		//最大，最小时。改变窗口大小。
		this->m_HDragArea.SetRectEmpty();
		this->m_VDragArea.SetRectEmpty();

		CRect ClientRect;
		GetWindowRect(&ClientRect);
		m_ListSize.cx = ClientRect.Width()/4;
		m_ListSize.cy = (long)(ClientRect.Height()/1.5);
		//

		this->UpdateControlsPos();
	}	
	
}
//为可拖动的区域添加颜色
void CDlgShowAllTable::OnPaint() 
{
	CPaintDC dc(this); // device context for painting	
	CBrush hBrush;
	if( !this->m_HDragArea.IsRectEmpty() && !this->m_VDragArea.IsRectEmpty() )
	{
		hBrush.CreateSolidBrush(RGB(127,127,255));
		dc.FillRect(m_HDragArea,&hBrush);
		dc.FillRect(m_VDragArea,&hBrush);
	}
}
//
void CDlgShowAllTable::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CRect ClientRect;

	if( this->m_HDragArea.PtInRect(point) )
	{
		this->SetCapture();
		this->m_IsInHDragArea = true;
		GetClientRect(&ClientRect);
		ClientToScreen(ClientRect);

		//限定鼠标移动在客户区内
		ClipCursor(ClientRect);
	}
	else
	{
		if( this->m_VDragArea.PtInRect(point) )
		{
			this->SetCapture();
			this->m_IsInVDragArea = true;
			GetClientRect(&ClientRect);
			ClientToScreen(&ClientRect);

			//限定鼠标移动在客户区内
			ClipCursor(ClientRect);
		}
	}
	this->m_LButtonPt = point;

	CDataShowDlg::OnLButtonDown(nFlags, point);
}

void CDlgShowAllTable::OnLButtonUp(UINT nFlags, CPoint point) 
{
	if( m_IsInVDragArea || m_IsInHDragArea )
	{
		m_IsInVDragArea = false;
		m_IsInHDragArea = false;
		ReleaseCapture();
		ClipCursor(NULL);
	}
	CDataShowDlg::OnLButtonUp(nFlags, point);
}
//////////////////////////////////////////////////
//
// 响应不同的表的选择改变
//
void CDlgShowAllTable::OnTableSelChange()
{
//	CString strValue;
	int bT = this->m_TableNameList.GetCurSel();
	if( bT != CB_ERR )
	{
		m_nCurTblID = m_TableNameList.GetItemData(bT);
		m_TableNameList.GetText( bT, m_strDescTblName);
		UpdateDataGrid();
	}
}
//////////////////////////////////////////////////
//
// 更新所选表的内容。
//
void CDlgShowAllTable::UpdateDataGrid()
{
	//对应浏览结果的记录集.
	_RecordsetPtr m_pRsBrowseResult;
	try
	{
		if(m_pRsBrowseResult == NULL)
		{
			m_pRsBrowseResult.CreateInstance(__uuidof(Recordset));
			m_pRsBrowseResult->CursorLocation = adUseClient;
		}
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return ;
	}

	try
	{	//管理表没有记录。
		if( m_pRsTableList->GetRecordCount() <= 0 )
		{
			return ;
		}

		CString strSql,
				strTblName;
		int     FlgDB;
		//用表对应的ID来查找。
		strSql.Format("SEQ=%d", m_nCurTblID);
		try
		{
			m_pRsTableList->MoveFirst();
			m_pRsTableList->Find(_bstr_t(strSql), 0, adSearchForward);
		}
		catch(_com_error)
		{
			AfxMessageBox("没有该表");
			return;
		}
		if( m_pRsTableList->adoEOF || m_pRsTableList->BOF )
		{
			return ;
		}
		strTblName = vtos(m_pRsTableList->GetCollect(_variant_t("ManuTBName")) );
		FlgDB = vtoi(m_pRsTableList->GetCollect(_variant_t("FromDBMark")) );

		try
		{
			if( m_pRsBrowseResult->State == adStateOpen )
			{
				m_pRsBrowseResult->Close();
			}
		}
		catch(_com_error)
		{
			m_pRsBrowseResult->Cancel();
			m_pRsBrowseResult->Close();
		}
		
		_ConnectionPtr ptr = theApp.m_pConnection;
		strSql = "SELECT * FROM ["+strTblName+"] ";
		switch( FlgDB )
		{
		case 0:/*"AutoIPED.mdb":*/
			if ( !tableExists(ptr, strTblName) )
			{
				ptr = theApp.m_pConAllPrj;
			}
			strSql += " WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
			m_pRsBrowseResult->Open(_bstr_t(strSql), ptr.GetInterfacePtr(),
							adOpenStatic, adLockOptimistic, adCmdText);
			break;
		case 1://"IPEDCODE.mdb":*/
			m_pRsBrowseResult->Open(_bstr_t(strSql), theApp.m_pConnectionCODE.GetInterfacePtr(),
							adOpenStatic, adLockOptimistic, adCmdText);
			break;
		case 2:/*"Material.MDB":*/
			m_pRsBrowseResult->Open(_bstr_t(strSql), theApp.m_pConMaterial.GetInterfacePtr(),
							adOpenStatic, adLockOptimistic, adCmdText);
			break;
		}
		//设置显示的内容;
		this->SetRecordset(m_pRsBrowseResult);

		//隐藏ID，ENGINID。
		this->SetHideColumns(_T("ID|ENGINID"));
		//
		this->SetDataGridCaption(strTblName);
		//设置字段的标题和宽度
		this->SetFieldCaption(strTblName);
		
		//设置默认字段值.
		CDataShowDlg::tag_DefaultFieldValue DefaultFieldValue;

		DefaultFieldValue.strFieldName=_T("EnginID");
		DefaultFieldValue.varFieldValue=EDIBgbl::SelPrjID;
		this->SetDefaultFieldValue(&DefaultFieldValue,1);

	//	pdlg->SetDefaultFieldValue(&DefaultFieldValue,1);

	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return;
	}
}
//
//删除表中的所有记录
void CDlgShowAllTable::OnDeleteTblAll()
{
	try
	{	//管理表没有记录。
		if( m_pRsTableList->GetRecordCount() <= 0 )
		{
			return ;
		}

		CString strSql,
				strTblName;
		int     FlgDB;
	//	strSql = "ShortDescribe='"+m_strDescTblName+"'";
		//用表对应的ID来查找。
		strSql.Format("SEQ=%d", m_nCurTblID);

		try
		{
			m_pRsTableList->MoveFirst();
			m_pRsTableList->Find(_bstr_t(strSql), 0, adSearchForward);
		}
		catch(_com_error)
		{
			AfxMessageBox("没有该表");
			return;
		}
		if( m_pRsTableList->adoEOF || m_pRsTableList->BOF )
		{
			return ;
		}
		strTblName = vtos(m_pRsTableList->GetCollect(_variant_t("ManuTBName")) );
		FlgDB = vtoi(m_pRsTableList->GetCollect(_variant_t("FromDBMark")) );

		if( IDNO == AfxMessageBox("是否删除 '"+m_strDescTblName+"' 中的所有记录?", MB_YESNO+MB_DEFBUTTON2) )
		{
			return ;
		}

		_ConnectionPtr ptr = theApp.m_pConnection;
		strSql = "DELETE * FROM ["+strTblName+"] ";
		switch( FlgDB )
		{
		case 0://"AutoIPED.mdb"
			if ( !tableExists(ptr, strTblName) )
			{
				ptr = theApp.m_pConAllPrj;
			}

			strSql += " WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
			ptr->Execute(_bstr_t(strSql), NULL, adCmdText);
			break;
		case 1://"IPEDCODE.mdb"
			theApp.m_pConnectionCODE->Execute(_bstr_t(strSql), NULL, adCmdText);
			break;
		case 2://"Material.MDB"
			theApp.m_pConMaterial->Execute(_bstr_t(strSql), NULL, adCmdText);
			break;
		}
		UpdateDataGrid();
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return ;
	}
}