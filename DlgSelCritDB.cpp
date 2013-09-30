// DlgSelCritDB.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DlgSelCritDB.h"
#include "MainFrm.h"

#include "EDIBgbl.h"
#include "vtot.h"
#include "GetPropertyofMaterial2.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDlgSelCritDB dialog
extern CAutoIPEDApp theApp;

CDlgSelCritDB::CDlgSelCritDB(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgSelCritDB::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgSelCritDB)
	//}}AFX_DATA_INIT
}


void CDlgSelCritDB::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgSelCritDB)
	DDX_Control(pDX, IDC_LIST1, m_ctrlShowCODE);
	DDX_Control(pDX, IDC_CBO_STUFF_DB, m_ctlMaterCodeName);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgSelCritDB, CDialog)
	//{{AFX_MSG_MAP(CDlgSelCritDB)
	ON_NOTIFY(NM_CLICK, IDC_LIST1, OnClickList1)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LIST1, OnColumnclickList1)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgSelCritDB message handlers
BOOL CDlgSelCritDB::OnInitDialog() 
{
	CDialog::OnInitDialog();

	int nFCount, pos;
	CString strCritDBName[20], strMaterialDBName[20];
	struct KeyToCalling sutKeyCalling[20];	// 行业名称和标示符

	int nKeyCount = GetCallingFromDB( sutKeyCalling );
	if ( nKeyCount <= 0 )
	{
		return FALSE;
	}
	//初始化标准数据库列表
//	nFCount = this->GetCriterionDBName( strCritDBName );
	//标准规范
	CMapStringToString mapCode_Mdb;
	CString strValue;
	short	nIndex = 0;	//初始的选择规范的索引号

	CRect rc;
	m_ctrlShowCODE.GetWindowRect(&rc);
	m_ctrlShowCODE.InsertColumn(0, _T("行业名称"), LVCFMT_LEFT, rc.Width()/3);
	m_ctrlShowCODE.InsertColumn(1, _T("数据库"), LVCFMT_LEFT, rc.Width()/3);
	m_ctrlShowCODE.InsertColumn(2, _T("标准名称"), LVCFMT_LEFT, rc.Width()/3);
	
	m_ctrlShowCODE.SetExtendedStyle( m_ctrlShowCODE.GetExtendedStyle()|LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES );

	try
	{
		_RecordsetPtr pRsDBRef;
		pRsDBRef.CreateInstance( __uuidof( Recordset ) );
		CString strSQL;
		CString strTmp;
		strSQL = "SELECT * FROM [DataBaseRef]";
		pRsDBRef->Open( _variant_t(strSQL), theApp.m_pConRefInfo.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText );
		for( int j=0, row=0; j < nKeyCount; j++ )
		{
			m_ctrlShowCODE.InsertItem( row, sutKeyCalling[j].strCallingName );
			if ( !pRsDBRef->adoEOF || !pRsDBRef->BOF)
			{
				strSQL.Format( "CallingKey='%d'",sutKeyCalling[j].strKey);
				pRsDBRef->MoveFirst();
				pRsDBRef->Find( _bstr_t(strSQL), 0, adSearchForward);
				if ( !pRsDBRef->adoEOF )
				{
					// 行业的标准数据库名
					strTmp = vtos( pRsDBRef->GetCollect( _variant_t("CodeDBName") ) );
					m_ctrlShowCODE.SetItemText( row, 1, strTmp );
					// 行业中材料的规范号
					m_ctrlShowCODE.SetItemText( row, 2, sutKeyCalling[j].strCodeName );
					// 
				}

				if( sutKeyCalling[j].strKey==EDIBgbl::iCur_CodeKey)
				{
					//记住上一次选择的规范的索引号
					nIndex = row;
				}
			}
			row++;
		}
		if ( m_ctrlShowCODE.GetItemCount() > 0 )
		{
			m_ctrlShowCODE.SetHotItem( nIndex );
			m_ctrlShowCODE.SetSelectionMark( nIndex );
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return FALSE;
	}
	
	//初始化规范列表
	CArray<CString,CString> mCodeNameArray;
	GetPropertyofMaterial mGetPropertyofMaterial;
	mGetPropertyofMaterial.GetAllCodeName( theApp.m_pConMaterial,mCodeNameArray );
	int iCount = mCodeNameArray.GetSize();
	for ( int i=0; i<iCount; i++ )
	{
		m_ctlMaterCodeName.AddString( mCodeNameArray.GetAt(i) );
	}
	
	int iCurSel = m_ctlMaterCodeName.FindString(-1,EDIBgbl::sCur_MaterialCodeName );
	if ( iCurSel == -1 )
		iCurSel = 0;
	m_ctlMaterCodeName.SetCurSel( iCurSel );

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
void CDlgSelCritDB::OnOK() 
{
	CString	strCurCodeName;
	CString	strCurDBName;
	CString strTmp;
	UpdateData();
	m_ctlMaterCodeName.GetWindowText( EDIBgbl::sCur_MaterialCodeName );

	int row = m_ctrlShowCODE.GetSelectionMark();
	if( row == -1 )
	{
		AfxMessageBox("请选择一个行业的规范数据库!");
		return;
	}
	char c[256];
	//规范数据库
	m_ctrlShowCODE.GetItemText(row, 1, c, 256);
	strCurDBName = c;

	//当前规范
	memset(c,'\0',256);
	m_ctrlShowCODE.GetItemText(row, 0, c, 256);
	strCurCodeName = c;

	//当前规范代号
	memset(c,'\0',256);
	m_ctrlShowCODE.GetItemText(row, 2, c, 256);
	EDIBgbl::sCur_CodeNO = c;

	try
	{
		if ( !FileExists( EDIBgbl::sCritPath + strCurDBName ) )
		{
			strTmp.Format(IDS_NOT_EXISTS_FILE, EDIBgbl::sCritPath + strCurDBName);
			AfxMessageBox( strTmp );
			return;
		}
		CString strCon;
		//重新连接标准库
		strCon = CONNECTSTRING + EDIBgbl::sCritPath + strCurDBName;
		if( theApp.m_pConnectionCODE->State == adStateOpen )
		{
			theApp.m_pConnectionCODE->Close();
		}
		theApp.m_pConnectionCODE->Open(_bstr_t(strCon), "", "", -1);
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.Description()+"\n\n选择的规范数据库（"+strCurCodeName+"）被破坏, 请重新安装 AutoIPED !");
		CDialog::OnCancel();
		return;
	}
	EDIBgbl::sCur_CritDbName = strCurDBName;
	EDIBgbl::sCur_CodeName = strCurCodeName;
	EDIBgbl::iCur_CodeKey=row+1;

	EDIBgbl::SetCurDBName();
	//显示当前的工程名(加上行业标准)
	((CMainFrame*)theApp.m_pMainWnd)->ShowCurrentProjectName();
	
	CDialog::OnOK();
}

//功能：获得标准数据库名称。
//函数返回数据库名称的个数，存于strCritDBName[]中
short CDlgSelCritDB::GetCriterionDBName(CString strCritDBName[])
{
	CFileFind  fileFind;
	CString    strFN;
	short      nFCount;
	short i=0;
	nFCount = fileFind.FindFile( EDIBgbl::sCritPath+"*.mdb" );
	if( nFCount == 0 )
	{
		return 0;
	}
	while( fileFind.FindNextFile() )
	{
		strFN = fileFind.GetFileName();
		if( strFN != "" )
		{
			strCritDBName[i++] = strFN;
		}
	}
	strFN = fileFind.GetFileName();
	if( strFN != "" )
	{
		strCritDBName[i++] = strFN;
	}

	return i;
}

//功能：获得所有材料数据库的名称。
//返回数据库名称的个数，并存于数组strMaterialDBName中
short CDlgSelCritDB::GetMaterialDBName(CString strMaterialDBName[])
{
	short nFCount, i=0;
	CString strFN;
	CFileFind fileFind;
	nFCount = fileFind.FindFile(EDIBgbl::sMaterialPath+"*.mdb");
	if( nFCount == 0 )
	{
		return 0;
	}
	while( fileFind.FindNextFile() )
	{
		strFN = fileFind.GetFileName();
		if( strFN != "" )
		{
			strMaterialDBName[i++] = strFN;
		}
	}
	strFN = fileFind.GetFileName();
	if( strFN != "" )
	{
		strMaterialDBName[i++] = strFN;
	}
	return i;
}

void CDlgSelCritDB::OnClickList1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	*pResult = 0;
}
//------------------------------------------------------------------
// DATE         :[2006/01/17]
// Parameter(s) :
// Return       :
// Remark       :从数据库中取出不同的行业名称
//------------------------------------------------------------------
int CDlgSelCritDB::GetCallingFromDB(struct KeyToCalling sutKeyCalling[])
{
	// 不同的行业名称用数据库统一进行管理 [2006/01/17]
	_RecordsetPtr pRsCalling;	// 行业管理表
	pRsCalling.CreateInstance( __uuidof( Recordset ) );
	CString strSQL;	// SQL语句
	int nCount = 0;	// 不同行业的个数
	try
	{
		strSQL = "SELECT * FROM [CallingManage]";
		pRsCalling->Open( _variant_t(strSQL), theApp.m_pConRefInfo.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText );
		if ( pRsCalling->adoEOF && pRsCalling->BOF )
		{
			return -1;
		}
		for (pRsCalling->MoveFirst(); !pRsCalling->adoEOF; pRsCalling->MoveNext())
		{
			sutKeyCalling[nCount].strKey = vtoi( pRsCalling->GetCollect( _variant_t("CallingKey") ) );
			sutKeyCalling[nCount].strCallingName = vtos( pRsCalling->GetCollect( _variant_t("CallingName") ) );
			sutKeyCalling[nCount].strCodeName = vtos( pRsCalling->GetCollect( _variant_t("CODE") ) );

			nCount++;
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return -1;
	}
	return nCount;
}

void CDlgSelCritDB::OnColumnclickList1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	// TODO: Add your control notification handler code here
	
	*pResult = 0;
}
