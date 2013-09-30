// Page2.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "Page2.h"
#include "edibgbl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPage2 property page

IMPLEMENT_DYNCREATE(CPage2, CPropertyPage)

BOOL SetFieldCaptionAndWidth(CDataGridEx* pDataGrid, _ConnectionPtr& pConSort, CString strTblName);

CPage2::CPage2() : CPropertyPage(CPage2::IDD)
{
	//{{AFX_DATA_INIT(CPage2)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}

CPage2::~CPage2()
{
}

void CPage2::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPage2)
	DDX_Control(pDX, IDC_DATAGRID_A_ARROW, m_ctlArrow);
	DDX_Control(pDX, IDC_DATAGRID_A_RING_PAINT, m_ctlRing_Paint);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPage2, CPropertyPage)
	//{{AFX_MSG_MAP(CPage2)
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPage2 message handlers

BOOL CPage2::OnInitDialog() 
{
	CPropertyPage::OnInitDialog();
	
	m_pConPrj = theApp.m_pConnection;
	m_pConCODE = theApp.m_pConnectionCODE;
	m_pConSort = theApp.m_pConIPEDsort;

	//更新色环箭头尺寸准则库
	UpdateA_ARROW();
	//更新色环箭头油漆材料名称库
	UpdateA_RING_PAINT();
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

// 销毁窗口  [11/23/2005]
void CPage2::OnDestroy() 
{
	m_ctlArrow.UnInitMouseWheel();
	m_ctlRing_Paint.UnInitMouseWheel();	
	
	CPropertyPage::OnDestroy();
}

//------------------------------------------------------------------
// DATE         :[2005/11/23]
// Parameter(s) :
// Return       :
// Remark       :更新色环箭头尺寸准则库
//------------------------------------------------------------------
BOOL CPage2::UpdateA_ARROW()
{
	if (!IsWindow(m_ctlArrow.GetSafeHwnd()))
	{
		return FALSE;
	}
	CString strSQL;
	CString strTblName="A_ARROW";
	_RecordsetPtr pRsArrow;
	pRsArrow.CreateInstance(__uuidof(Recordset));
	pRsArrow->CursorLocation = adUseClient;
	try
	{
		long nFieldCount;
		strSQL = "SELECT * FROM ["+strTblName+"] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
		pRsArrow->Open(_variant_t(strSQL), m_pConPrj.GetInterfacePtr(), adOpenStatic,adLockOptimistic, adCmdText);
		nFieldCount = pRsArrow->GetRecordCount();
		nFieldCount = pRsArrow->GetFields()->GetCount();
		//设置数据源
		m_ctlArrow.SetRefDataSource(pRsArrow->GetDataSource());
		//操作控制
		m_ctlArrow.SetAllowUpdate(TRUE);
		m_ctlArrow.SetAllowAddNew(TRUE);
		m_ctlArrow.SetAllowDelete(TRUE);

		SetFieldCaptionAndWidth(&m_ctlArrow,m_pConSort,strTblName);
		m_ctlArrow.InitMouseWheel();
	}
	catch (_com_error& e) {
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/11/23]
// Parameter(s) :
// Return       :
// Remark       :更新色环箭头油漆材料名称库(A_RING_PAINT):
//------------------------------------------------------------------
BOOL CPage2::UpdateA_RING_PAINT()
{
	CString strSQL;						//SQL语句
	CString strTblName = "A_RING_PAINT" ;//数据表的名称
	_RecordsetPtr pRsRing_Paint;		//色环箭头的油漆材料名称库 
	pRsRing_Paint.CreateInstance( __uuidof( Recordset ) );
	pRsRing_Paint->PutCursorLocation( adUseClient );
	try
	{
		strSQL.Format( "SELECT * FROM [%s]", strTblName );
		pRsRing_Paint->Open( _variant_t( strSQL ), m_pConCODE.GetInterfacePtr(), 
							adOpenStatic, adLockOptimistic, adCmdText );
		m_ctlRing_Paint.SetRefDataSource( pRsRing_Paint->GetDataSource() );
		//操制属性
		m_ctlRing_Paint.SetAllowUpdate( TRUE );
		m_ctlRing_Paint.SetAllowAddNew( TRUE );
		m_ctlRing_Paint.SetAllowDelete( TRUE );
		//设置字段的本地描述
		SetFieldCaptionAndWidth( &m_ctlRing_Paint, m_pConSort, strTblName );
	}
	catch( _com_error& e )
	{
		AfxMessageBox( e.Description() );
		return FALSE;
	}

	return TRUE;
}

