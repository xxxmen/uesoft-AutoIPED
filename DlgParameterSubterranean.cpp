// DlgParameterSubterranean.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DlgParameterSubterranean.h"
#include "vtot.h"
#include "foxbase.h"
#include "htmlhelp.h"
#include "editoriginaldata.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define nSubMaxPipeCount 2

/////////////////////////////////////////////////////////////////////////////
// CDlgParameterSubterranean dialog


CDlgParameterSubterranean::CDlgParameterSubterranean(CWnd* pParent /*=NULL*/)
	: CXDialog(CDlgParameterSubterranean::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgParameterSubterranean)
	m_dPipeSubDepth = 0.0;
	m_dEdaphicTemp = 0.0;
	m_dPipeSpan = 0.0;
	m_dEdaphicLamda = 0.0;
	m_nPipeLay = 0;
	m_dSupportK = 0.0;
	m_dPipeDag = 0.0;
	m_dCanalDepth = 0.0;
	//}}AFX_DATA_INIT
	m_pRsSubterranean = NULL;
	m_nID = -1;
}


void CDlgParameterSubterranean::DoDataExchange(CDataExchange* pDX)
{
	CXDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgParameterSubterranean)
	DDX_Control(pDX, IDC_COMBO_PIPE_COUNT, m_ctlPipeCount);
	DDX_Control(pDX, IDC_COMBO_LAY, m_comboPipeLay);
	DDX_Text(pDX, IDC_PIPE_SUB_DEPTH, m_dPipeSubDepth);
	DDX_Text(pDX, IDC_EDAPHIC_TEMPERATURE, m_dEdaphicTemp);
	DDX_Text(pDX, IDC_PIPE_SPAN, m_dPipeSpan);
	DDX_Text(pDX, IDC_EDAPHIC_LAMDA, m_dEdaphicLamda);
	DDX_CBIndex(pDX, IDC_COMBO_LAY, m_nPipeLay);
	DDX_Text(pDX, IDC_SUPPORT_K, m_dSupportK);
	DDX_Text(pDX, IDC_PIPE_DAG, m_dPipeDag);
	DDX_Text(pDX, IDC_CANAL_DEPTH, m_dCanalDepth);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgParameterSubterranean, CXDialog)
	//{{AFX_MSG_MAP(CDlgParameterSubterranean)
	ON_CBN_SELCHANGE(IDC_COMBO_LAY, OnSelchangeComboLay)
	ON_WM_HELPINFO()
	ON_CBN_SELCHANGE(IDC_COMBO_PIPE_COUNT, OnSelchangeComboPipeCount)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgParameterSubterranean message handlers

BOOL CDlgParameterSubterranean::OnInitDialog() 
{
	CXDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	InitPipeLayList();
	InitToolTip();
	
//	GetDataToSubterraneanControl(m_nID);

 	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
//------------------------------------------------------------------                
// DATE         : [2006/02/14]
// Author       : ZSY
// Parameter(s) : 
// Return       : 
// Remark       : 设置数据记录集
//------------------------------------------------------------------
void CDlgParameterSubterranean::SetCurrentRes(_RecordsetPtr pRsCur)
{
	ASSERT( pRsCur != NULL );

	m_pRsSubterranean = pRsCur;
}

//------------------------------------------------------------------                
// DATE         : [2006/02/14]
// Author       : ZSY
// Parameter(s) : 
// Return       : 
// Remark       : 设置当前防冻原始表的数据库连接
//------------------------------------------------------------------
void CDlgParameterSubterranean::SetCurrentProjectConnect(_ConnectionPtr pCon)
{
	ASSERT( pCon != NULL );
 
	m_pConSubter = pCon;
}

//------------------------------------------------------------------
// DATE         : [2006/02/17]
// Author       : ZSY 
// Parameter(s) : nID[in] 显示当前记录的序号
// Return       : 读取成功返回1，否则返回0
// Remark       : 读取埋地管道的保温数据到输入窗口中
//------------------------------------------------------------------
short CDlgParameterSubterranean::GetDataToSubterraneanControl(long nID)
{
	CString strSQL;
	int nSelect;
	try
	{
		if ( m_pRsSubterranean == NULL || ( m_pRsSubterranean->GetState() != adStateOpen ) )
		{
			return 1;
		}
		if ( !m_pRsSubterranean->adoEOF || !m_pRsSubterranean->BOF )
		{
			m_pRsSubterranean->MoveFirst();
			strSQL.Format("ID=%d", nID );
			m_pRsSubterranean->Find( _bstr_t( strSQL ), NULL, adSearchForward );
		}
		if ( m_pRsSubterranean->adoEOF || m_pRsSubterranean->BOF )// 没有找到当前工程指定序号的记录则新增一条
		{
			m_pRsSubterranean->AddNew();//
			m_pRsSubterranean->PutCollect( _variant_t("ID"), _variant_t(nID) );
			m_pRsSubterranean->PutCollect( _variant_t("EnginID"), _variant_t(m_strProjectID));
			m_pRsSubterranean->Update();

			m_ctlPipeCount.SetCurSel(0);
		}
		else
		{
			nSelect = vtoi( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Count") ) );
			nSelect = (nSelect > 0) ? nSelect -1 : 0;
			m_ctlPipeCount.SetCurSel(nSelect);
		}
		m_nID = nID;	//保存当前记录的序号		
		m_nPipeLay		= vtoi( m_pRsSubterranean->GetCollect( _variant_t("c_lay_mark") ) );	// 管道敷设状态
		m_dPipeSubDepth	= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Sub_Depth")));// 管道的埋设深度（地表面到管道中心的距离）
		m_dEdaphicTemp	= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Edaphic_Temp") ));	// 土壤的温度℃
		m_dPipeSpan		= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Span") ) );	// 两根管道之间的当量间距
		m_dEdaphicLamda = vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Edaphic_Lamda")));	// 土壤的导热系数
		m_dSupportK     = vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Support_K") ) );	// 支吊架的热损失校正系数
		m_dPipeDag		= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Dag") ) );	// 管沟当量直径
		m_dCanalDepth   = vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Canal_Depth") ) ); // 管沟深度

	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return 0;
	}
	UpdateData( FALSE );
	// 更新控件的状态
	UpdateControlState();

	
	return 1;
}

//------------------------------------------------------------------
// DATE         : [2006/02/17]
// Author       : ZSY
// Parameter(s) : lpstrProjectID【in】工程ID號
// Return       : void
// Remark       : 設置當前的工程ID
//------------------------------------------------------------------
void CDlgParameterSubterranean::SetProjectID(LPCTSTR lpstrProjectID)
{
	ASSERT(lpstrProjectID);
	
	m_strProjectID = lpstrProjectID;
}

//------------------------------------------------------------------
// DATE         : [2006/02/20]
// Author       : ZSY
// Parameter(s) : void
// Return       : 保存成功返回1，否则返回0
// Remark       : 将埋地管道保温计算时须要的数据保存到数据库中。
//------------------------------------------------------------------
short CDlgParameterSubterranean::PutDataToSubterraneanDB(long nID)
{
	UpdateData( TRUE );

	CString strSQL;
	try
	{
		if ( m_pRsSubterranean == NULL || m_pRsSubterranean->GetState() != adStateOpen )
		{
			return 0;
		}
		if (nID != -1 && m_nID != nID)
		{
			// 写入到指定的记录中
			if ( !m_pRsSubterranean->adoEOF || !m_pRsSubterranean->BOF )
			{
				m_pRsSubterranean->MoveFirst();
				strSQL.Format( "ID=%d", nID );
				m_pRsSubterranean->Find(_bstr_t(strSQL), NULL, adSearchForward);
			}
			if ( m_pRsSubterranean->adoEOF || m_pRsSubterranean->BOF )
			{
				//没有当前序号时，增加一条新记录
				m_pRsSubterranean->AddNew();
				//序号
				m_pRsSubterranean->PutCollect(_T("ID"), _variant_t(nID));
				//工程代号
				m_pRsSubterranean->PutCollect(_T("EnginID"), _variant_t( m_strProjectID ));			
				//
				m_pRsSubterranean->Update(vtMissing);
			}
		}
		else if (m_pRsSubterranean->adoEOF || m_pRsSubterranean->BOF)
		{
			return 0;
		}

		int nPipeCount = m_ctlPipeCount.GetCurSel();	// 管道根数
		if (nPipeCount < 0)
			nPipeCount = 1;
		else
			nPipeCount++;
		
		m_pRsSubterranean->PutCollect( _variant_t("c_Pipe_Sub_Depth"),_variant_t(m_dPipeSubDepth) );
		m_pRsSubterranean->PutCollect( _variant_t("c_Pipe_Span"),	  _variant_t(m_dPipeSpan) );
		m_pRsSubterranean->PutCollect( _variant_t("c_Edaphic_Temp"),  _variant_t(m_dEdaphicTemp) );
		m_pRsSubterranean->PutCollect( _variant_t("c_Edaphic_Lamda"), _variant_t(m_dEdaphicLamda) );
		m_pRsSubterranean->PutCollect( _variant_t("c_lay_mark"),	  _variant_t((short)m_nPipeLay) );
		m_pRsSubterranean->PutCollect( _variant_t("c_Support_K"),	  _variant_t(m_dSupportK) );
		m_pRsSubterranean->PutCollect( _variant_t("c_Pipe_Dag"),	  _variant_t(m_dPipeDag) );
		m_pRsSubterranean->PutCollect( _variant_t("c_Canal_Depth"),   _variant_t(m_dCanalDepth) );
		m_pRsSubterranean->PutCollect( _variant_t("c_Pipe_Count"),	  _variant_t((short)nPipeCount) );

		m_pRsSubterranean->Update();
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return 0;
	}
	return 1;
}

//------------------------------------------------------------------
// DATE         : [2006/02/25]
// Author       : ZSY
// Parameter(s) : void
// Return       : void
// Remark       : 初始输入框的提示信息
//------------------------------------------------------------------
void CDlgParameterSubterranean::InitToolTip()
{
	Txtup TipArr[20];
	int	pos = 0;

	TipArr[pos].dID   = IDC_PIPE_SUB_DEPTH; // 管道埋设深度
	TipArr[pos++].txt = _T("管道中心到地面的距离 m");

	TipArr[pos].dID   = IDC_PIPE_SPAN; // 两管道之间的间距
	TipArr[pos++].txt = _T("m");

	TipArr[pos].dID   = IDC_EDAPHIC_TEMPERATURE;	// 土壤的温度
	TipArr[pos++].txt = _T("℃");

	TipArr[pos].dID   = IDC_EDAPHIC_LAMDA;		//	土壤的导热系数
	TipArr[pos++].txt = _T("W/(m·K)");
	
	TipArr[pos].dID   = IDC_CANAL_DEPTH;	
	TipArr[pos++].txt = _T("m");

	TipArr[pos].dID   = IDC_CANAL_ALPHA;
	TipArr[pos++].txt = _T("W/(m·K)");	
	
	TipArr[pos].dID	  = IDC_CANAL_SIZE;
	TipArr[pos++].txt = _T("");
	
	if ( NULL == m_pctlTip.m_hWnd )
	{
		m_pctlTip.Create(this);
	}
	for ( int i = 0; i < pos; i++ )
	{
		if (!TipArr[i].txt.IsEmpty())
		{
			m_pctlTip.AddTool( GetDlgItem(TipArr[i].dID), TipArr[i].txt );
		}
	}
}

BOOL CDlgParameterSubterranean::PreTranslateMessage(MSG* pMsg) 
{
	m_pctlTip.RelayEvent(pMsg);
	return CXDialog::PreTranslateMessage(pMsg);
}

//------------------------------------------------------------------
// DATE         : [2006/03/01]
// Author       : ZSY
// Parameter(s) : void
// Return       : void
// Remark       : 加载管道敷设状态
//------------------------------------------------------------------
void CDlgParameterSubterranean::InitPipeLayList()
{
	m_comboPipeLay.ResetContent();
	CString strSQL;		// SQL语句
	CString strCount;
	int nIndex;		
	_RecordsetPtr pRsLay;
	pRsLay.CreateInstance( __uuidof( Recordset ) );

	try
	{
		m_comboPipeLay.ResetContent();

		strSQL.Format("SELECT * FROM [埋地管道敷设状态] ORDER BY Index");
		pRsLay->Open( _variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText );
		
		if ( pRsLay->GetadoEOF() && pRsLay->GetBOF() )
			return;		
		for (pRsLay->MoveFirst(); !pRsLay->adoEOF; pRsLay->MoveNext())
		{
			nIndex = m_comboPipeLay.AddString( vtos( pRsLay->GetCollect(_variant_t("PipeLay")) ) );
			m_comboPipeLay.SetItemData(nIndex, vtoi( pRsLay->GetCollect(_variant_t("Index")) ) );
		}

		// 同时敷设根数
		m_ctlPipeCount.ResetContent();
		for (int i = 1; i <= nSubMaxPipeCount; i++)
		{
			strCount.Format("%d", i); 
			m_ctlPipeCount.AddString(strCount);
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return ;
	}
}

void CDlgParameterSubterranean::OnSelchangeComboLay() 
{
	UpdateControlState();
}

//------------------------------------------------------------------
// DATE         : [2006/03/10]
// Author       : ZSY
// Parameter(s) : void
// Return       : void
// Remark       : 根据当前选择的数据更改各控件的状态
//------------------------------------------------------------------
void CDlgParameterSubterranean::UpdateControlState()
{
	UpdateData();
	BOOL bIsEnable = TRUE;

	switch(m_nPipeLay)
	{
	case nNoChannel:
		bIsEnable = FALSE;
		break;
	case nObstruct:
		bIsEnable = TRUE;
	}
	LONG pID[] = 
	{
		IDC_CANAL_DEPTH,
		IDC_CANAL_SIZE,
		IDC_SUPPORT_K,
		IDC_PIPE_DAG,
		IDC_CANAL_ALPHA,
	};
	for (int i = 0; i < sizeof(pID) / sizeof(LONG); i++)
	{
		(CButton*)GetDlgItem(pID[i])->EnableWindow(bIsEnable);
		//(CButton*)GetDlgItem(pID[i])->ShowWindow(bIsEnable ? SW_SHOW : SW_HIDE); 
	}

	int nPipeCount;
	CWnd *pParDlg;
	CEditOriginalData * pEditDlg;
	nPipeCount = m_ctlPipeCount.GetCurSel();
	nPipeCount = (nPipeCount > 0) ? nPipeCount + 1 : 1;
	try
	{
		pParDlg = this->GetParent();
		
		while (IsWindow(pParDlg->GetSafeHwnd()))
		{
			if (pParDlg->IsKindOf(RUNTIME_CLASS(CEditOriginalData)))
			{
				pEditDlg = (CEditOriginalData*)pParDlg;
				if ( IsWindow(pEditDlg->GetSafeHwnd()) )
				{
					pEditDlg->UpdateControlSubterranean(nPipeCount);
					break;
				}
			}
			pParDlg = pParDlg->GetParent();
		}
	}
	catch (COleDispatchException& e)
	{
		e.Delete();
	}
	if (m_nPipeLay == nObstruct && nPipeCount > 1)
	{
		(CButton*)GetDlgItem(IDC_PIPE_SPAN)->EnableWindow(TRUE);
	}
	else
	{
		(CButton*)GetDlgItem(IDC_PIPE_SPAN)->EnableWindow(FALSE);
	}
	
	UpdateData(FALSE);
}

// 当前工程的记录号
void CDlgParameterSubterranean::PutCurrentNO(long nCurNO)
{
	m_nID = nCurNO;
}

// 当前激活控制的帮助主题，F1键的消息
BOOL CDlgParameterSubterranean::OnHelpInfo(HELPINFO* pHelpInfo) 
{
	const DWORD hArr[] =
	{
		IDC_CANAL_SIZE, 3,
		IDC_PIPE_SPAN, 10,
		0,0
	};
	BOOL bIsExists = FALSE;

	if (pHelpInfo->iCtrlId > 0)
	{
		for (int pos=0, count = sizeof(hArr) / sizeof(DWORD); pos < count; pos += 2)
		{
			if (hArr[pos] == (DWORD)pHelpInfo->iCtrlId)
			{
				bIsExists = TRUE;
				break;
			}
		}	
	}

	if (pHelpInfo->iContextType == HELPINFO_WINDOW && bIsExists && AfxGetApp()->m_pszHelpFilePath!=NULL)
	{
/*
		CString strCmd = CString(AfxGetApp()->m_pszHelpFilePath) + _T("::/Subterranean.txt"); 

		return  (HtmlHelp((HWND)pHelpInfo->hItemHandle,	strCmd,HH_TP_HELP_WM_HELP, (DWORD)(LPVOID)hArr) != NULL);
*/ 
	}

	return CXDialog::OnHelpInfo(pHelpInfo);
}

void CDlgParameterSubterranean::OnSelchangeComboPipeCount() 
{
	UpdateControlState();
}
