// EditOilPaintDataDlg.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "EditOilPaintDataDlg.h"
#include "EDIBgbl.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern CAutoIPEDApp theApp;
/////////////////////////////////////////////////////////////////////////////
// CEditOilPaintDataDlg dialog
//
// 油漆原始数据编辑对话框
//
////////////////////////////////////////////////////////////////////////////

CEditOilPaintDataDlg::CEditOilPaintDataDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CEditOilPaintDataDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CEditOilPaintDataDlg)
	m_ID = 0;
	m_VolNo = _T("");
	m_DeviceOutsizeArea = 0.0f;
	m_PipeSize = 0.0f;
	m_PipeLengthDevice = 0.0f;
	m_ReMark = _T("");
	m_IsUpdateByRoll = TRUE;
	//}}AFX_DATA_INIT
}


void CEditOilPaintDataDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CEditOilPaintDataDlg)
	DDX_Control(pDX, IDC_Del_All_Paint, m_DeleteAllPaint);
	DDX_Control(pDX, IDC_SAVECURRENT, m_SaveCurrent);
	DDX_Control(pDX, IDC_DELCURRENT, m_DelCurrent);
	DDX_Control(pDX, IDC_ADDNEW, m_AddNew);
	DDX_Control(pDX, IDC_LAST, m_MoveLast);
	DDX_Control(pDX, IDC_SUBSEQUENT, m_MoveSubsequent);
	DDX_Control(pDX, IDC_PREVIOUS, m_MovePrevious);
	DDX_Control(pDX, IDC_FOREFRONT, m_MoveForefront);
	DDX_Control(pDX, IDC_OILPAINT_TYPE, m_OilPaintType);
	DDX_Control(pDX, IDC_OILPAINT_COLOR, m_OilPaintColor);
	DDX_Control(pDX, IDC_PIPE_DEVICE_NAME, m_PipeDeviceName);
	DDX_Text(pDX, IDC_ID, m_ID);
	DDX_Text(pDX, IDC_VOLNO, m_VolNo);
	DDX_Text(pDX, IDC_DEVICE_OUTSIDE_AREA, m_DeviceOutsizeArea);
	DDX_Text(pDX, IDC_PIPE_SIZE, m_PipeSize);
	DDX_Text(pDX, IDC_PIPE_LENGTH_DEVICE, m_PipeLengthDevice);
	DDX_Text(pDX, IDC_REMARK, m_ReMark);
	DDX_Check(pDX, IDC_UPDATABYROLL, m_IsUpdateByRoll);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CEditOilPaintDataDlg, CDialog)
	//{{AFX_MSG_MAP(CEditOilPaintDataDlg)
	ON_BN_CLICKED(IDC_FOREFRONT, OnMoveForefront)
	ON_BN_CLICKED(IDC_PREVIOUS, OnMovePrevious)
	ON_BN_CLICKED(IDC_SUBSEQUENT, OnMoveSubsequent)
	ON_BN_CLICKED(IDC_LAST, OnMoveLast)
	ON_BN_CLICKED(IDC_ADDNEW, OnAddNew)
	ON_BN_CLICKED(IDC_DELCURRENT, OnDelCurrent)
	ON_BN_CLICKED(IDC_SAVECURRENT, OnSaveCurrent)
	ON_CBN_DROPDOWN(IDC_PIPE_DEVICE_NAME, OnDropdownPipeDeviceName)
	ON_BN_CLICKED(IDC_EXIT_PAINT, OnExit)
	ON_EN_CHANGE(IDC_ID, OnChangeId)
	ON_BN_CLICKED(IDC_Del_All_Paint, OnDelAllPaint)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

//////////////////////////////////////////////////////////////////////////
//
// 设置所选工程的数据库的连接
//
// IConnection[in]	已经连接的数据库的智能指针
//
void CEditOilPaintDataDlg::SetCurrentProjectConnect(_ConnectionPtr &IConnection)
{
	m_ICurrentProjectConnection=IConnection;
}

//////////////////////////////////////////////////////////////////////////
//
// 返回数据库连接的智能指针的引用
//
_ConnectionPtr& CEditOilPaintDataDlg::GetCurrentProjectConnect()
{
	return m_ICurrentProjectConnection;
}

//////////////////////////////////////////////////////////////////////////
//
// 设置公共的数据库的连接
//
// IConnection[in]	已经连接的数据库的智能指针
//
void CEditOilPaintDataDlg::SetPublicConnect(_ConnectionPtr &IConnection)
{
	m_IPublicConnection=IConnection;

}

//////////////////////////////////////////////////////////////////////////
//
// 返回公共数据库连接的智能指针的引用
//
_ConnectionPtr& CEditOilPaintDataDlg::GetPublicConnect()
{
	return m_IPublicConnection;
}

/////////////////////////////////////////////////////////////
//
// 设置所选工程的ID号
//
// ProjectID[in]	所选工程的ID号
//
void CEditOilPaintDataDlg::SetProjectID(LPCTSTR ProjectID)
{
	if(ProjectID==NULL)
		return;

	m_ProjectID=ProjectID;
}

//////////////////////////////////////////////////////////
//
// 返回所选工程的ID号
//
CString& CEditOilPaintDataDlg::GetProjectID()
{
	return m_ProjectID;
}

////////////////////////////////////////////////////////
//
// 打开PAINT表
//
BOOL CEditOilPaintDataDlg::InitCurrentProjectRecordset()
{
	CString SQL;

	if(m_ProjectID.IsEmpty())
	{
		AfxMessageBox(_T("未选择工程"));
		return FALSE;
	}

	if(m_ICurrentProjectConnection==NULL)
	{
		AfxMessageBox(_T("未连接数据库"));
		return FALSE;
	}

	try
	{
		if(m_ICurrentProjectRecordset==NULL)
		{
			m_ICurrentProjectRecordset.CreateInstance(__uuidof(Recordset));
		}
		else
		{
			if(m_ICurrentProjectRecordset->GetState() & adStateOpen)
			{
				m_ICurrentProjectRecordset->Close();
			}
		}

		SQL.Format(_T("SELECT * FROM paint WHERE EnginID='%s' ORDER BY ID,pai_vol,pai_name"),GetProjectID());

		m_ICurrentProjectRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
										 adOpenStatic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CEditOilPaintDataDlg message handlers
//

BOOL CEditOilPaintDataDlg::OnInitDialog() 
{
	BOOL bRet;

	CDialog::OnInitDialog();
	//移动时更新记录集
	m_IsUpdateByRoll = bIsMoveUpdate;
	// 提示信息 [2005/06/23]
	InitToolTip();

	//打开油漆原始数据表。
	bRet=InitCurrentProjectRecordset();
	if(bRet==FALSE)
	{
		OnCancel();
		return TRUE;
	}

	bRet=InitBitmapButton();
	if(bRet==FALSE)
	{
		OnCancel();
		return TRUE;
	}

	bRet=InitOilPaintColorCombox();
	if(bRet==FALSE)
	{
		OnCancel();
		return TRUE;
	}

	bRet=InitOilPaintTypeCombox();
	if(bRet==FALSE)
	{
		OnCancel();
		return TRUE;
	}

	bRet=PutDataToEveryControl();
	if(bRet==FALSE)
	{
		OnCancel();
		return TRUE;
	}

	UpdateControlsState();
	UpdatePipeDeviceNameCombox();
	//2005.2.24  移动窗口到正中央
	SetWindowCenter(this->m_hWnd);
	//设置“退出”为默认的按钮
	this->SetDefID( IDC_EXIT_PAINT );

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

/////////////////////////////////////////////////
//
// 初始化位图按扭
//
// 如果初始化成功返回TRUE，否则返回FALSE
//
BOOL CEditOilPaintDataDlg::InitBitmapButton()
{
	CBitmap Bitmap;

	struct
	{
		CMyButton *pButton;		
		UINT uBitmapID;			//位图的ID
		LPCTSTR szTooltipText;	//提示信息
	}ButtonInfo[]=
	{
		{&m_MoveForefront,	IDB_BITMAP1,	_T("第一条记录")},
		{&m_MovePrevious,	IDB_BITMAP2,	_T("前一条记录")},
		{&m_MoveSubsequent,	IDB_BITMAP3,	_T("后一条记录")},
		{&m_MoveLast,		IDB_BITMAP4,	_T("最后一条记录")},
		{&m_AddNew,			IDB_BITMAP5,	_T("新建")},
		{&m_DelCurrent,		IDB_BITMAP6,	_T("删除当前记录")},
		{&m_SaveCurrent,	IDB_BITMAP7,	_T("保存修改")},
		{&m_DeleteAllPaint, IDB_BITMAP10,   _T("删除所有的记录")}
	};

	for(int i=0;i<sizeof(ButtonInfo)/sizeof(ButtonInfo[0]);i++)
	{
		Bitmap.LoadBitmap(ButtonInfo[i].uBitmapID);

		ButtonInfo[i].pButton->SetBitmap((HBITMAP)Bitmap.Detach());
		ButtonInfo[i].pButton->SetTooltipText(ButtonInfo[i].szTooltipText,TRUE);
	}
	
	return TRUE;
}

///////////////////////////////////////////////////////////
//
// 初使化"油漆颜色"组合框
//
// 如果初始化成功返回TRUE，否则返回FALSE
//
BOOL CEditOilPaintDataDlg::InitOilPaintColorCombox()
{
	HRESULT hr;
	CString SQL,strTemp,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Color;

	m_OilPaintColor.GetWindowText(TextStr);

	do
	{
		iRet=m_OilPaintColor.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Color.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

//	SQL.Format(_T("SELECT COLR_MEDIA FROM A_COLOR"));
	SQL.Format(_T("SELECT COLR_MEDIA FROM A_COLOR WHERE EnginID='%s' ORDER BY COLR_CODE ASC"),GetProjectID());

	try
	{
//		IRecordset_Color->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							   adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Color->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							   adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Color->adoEOF)
		{
			GetTbValue(IRecordset_Color,_variant_t("COLR_MEDIA"),strTemp);

			if(!strTemp.IsEmpty())
				m_OilPaintColor.AddString(strTemp);

			IRecordset_Color->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_OilPaintColor.SetWindowText(TextStr);

	return TRUE;
}

///////////////////////////////////////////////////////////
//
// 初使化"油漆类型"组合框
//
// 如果初始化成功返回TRUE，否则返回FALSE
//
BOOL CEditOilPaintDataDlg::InitOilPaintTypeCombox()
{
	HRESULT hr;
	CString SQL,strTemp,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Pai;

	m_OilPaintType.GetWindowText(TextStr);

	do
	{
		iRet=m_OilPaintType.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Pai.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
//		AfxMessageBox(_T("Get interface error"));
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

//	SQL.Format(_T("SELECT PAI_TYPE FROM A_PAI"));
	SQL.Format(_T("SELECT PAI_TYPE FROM A_PAI WHERE EnginID='%s'"),GetProjectID());

	try
	{
//		IRecordset_Pai->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							 adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Pai->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							 adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Pai->adoEOF)
		{
			GetTbValue(IRecordset_Pai,_variant_t("PAI_TYPE"),strTemp);

			if(!strTemp.IsEmpty())
				m_OilPaintType.AddString(strTemp);

			IRecordset_Pai->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_OilPaintType.SetWindowText(TextStr);

	return TRUE;
}

/////////////////////////////////////////////////////////
//
// 从以打开的记录集的当前游标所在的位置取数据到各个控件
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOilPaintDataDlg::PutDataToEveryControl()
{
	CString strTemp;
	int iSelPos;
	try
	{
		if(m_ICurrentProjectRecordset==NULL || m_ICurrentProjectRecordset->GetState()==adStateClosed)
			return FALSE;
		
		if(m_ICurrentProjectRecordset->adoEOF)
			return TRUE;
		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("ID"),m_ID);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("pai_vol"),m_VolNo);
		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("Pai_name"),strTemp);
		m_PipeDeviceName.SetWindowText(strTemp);
		
		//	油漆颜色
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("Pai_code"),iSelPos);
		m_OilPaintColor.SetCurSel(iSelPos-1);
		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("pai_type"),strTemp);
		m_OilPaintType.SetWindowText(strTemp);
		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("pai_area"),m_DeviceOutsizeArea);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("pai_size"),m_PipeSize);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("pai_amou"),m_PipeLengthDevice);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("pai_mark"),m_ReMark);
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return FALSE;
	}

	UpdateData(FALSE);

	return TRUE;
}

/////////////////////////////////////////////////////////
//
// 从各个控件取数据到数据库
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOilPaintDataDlg::PutDataToDatabaseFromControl()
{
	CString strTemp;
	long iSelPos;

	UpdateData(TRUE);

	try
	{
		if(m_ICurrentProjectRecordset==NULL || m_ICurrentProjectRecordset->GetState()==adStateClosed)
			return FALSE;

		if(m_ICurrentProjectRecordset->adoEOF || m_ICurrentProjectRecordset->BOF)
			return FALSE;
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	try
	{
//		m_ICurrentProjectRecordset->PutCollect(_variant_t("ID"),_variant_t((long)m_ID));

//		PutTbValue(m_ICurrentProjectRecordset,_variant_t("ID"),m_ID);

		PutTbValue(m_ICurrentProjectRecordset,_variant_t("pai_vol"),m_VolNo);

		m_PipeDeviceName.GetWindowText(strTemp);
		PutTbValue(m_ICurrentProjectRecordset,_variant_t("Pai_name"),strTemp);

		m_OilPaintColor.GetWindowText(strTemp); //zsy
		strTemp.TrimLeft();
		strTemp.TrimRight();
		iSelPos = m_OilPaintColor.FindString(0, strTemp);
		iSelPos++;

	//	iSelPos=m_OilPaintColor.GetCurSel()+1;
		PutTbValue(m_ICurrentProjectRecordset,_variant_t("Pai_code"),iSelPos);

		m_OilPaintType.GetWindowText(strTemp);
		PutTbValue(m_ICurrentProjectRecordset,_variant_t("pai_type"),strTemp);

		PutTbValue(m_ICurrentProjectRecordset,_variant_t("pai_area"),m_DeviceOutsizeArea);
		PutTbValue(m_ICurrentProjectRecordset,_variant_t("pai_size"),m_PipeSize);
		PutTbValue(m_ICurrentProjectRecordset,_variant_t("pai_amou"),m_PipeLengthDevice);
		PutTbValue(m_ICurrentProjectRecordset,_variant_t("pai_mark"),m_ReMark);

		m_ICurrentProjectRecordset->Update();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	UpdateVolTable();

	return TRUE;
}

//////////////////////////////////////////////
//
// 响应“第一条记录”按扭
//
void CEditOilPaintDataDlg::OnMoveForefront() 
{
	try
	{
		if(m_ICurrentProjectRecordset==NULL || m_ICurrentProjectRecordset->GetState()==adStateClosed)
			return;

		UpdateData();

		if(m_ICurrentProjectRecordset->BOF)
			return;

		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
		}

		m_ICurrentProjectRecordset->MoveFirst();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	PutDataToEveryControl();

	UpdateData(FALSE);
	UpdateControlsState();
}

/////////////////////////////////////////
//
// 响应“前一条记录”按扭
//
void CEditOilPaintDataDlg::OnMovePrevious() 
{
	try
	{
		if(m_ICurrentProjectRecordset==NULL || m_ICurrentProjectRecordset->GetState()==adStateClosed)
			return;

		UpdateData();

		if(m_ICurrentProjectRecordset->BOF)
			return;
		
		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
		}

		m_ICurrentProjectRecordset->MovePrevious();

		if(m_ICurrentProjectRecordset->BOF)
			m_ICurrentProjectRecordset->MoveFirst();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	PutDataToEveryControl();

	UpdateData(FALSE);
	UpdateControlsState();
}

//////////////////////////////////////////////
//
// 响应“后一条记录”按扭
//
void CEditOilPaintDataDlg::OnMoveSubsequent() 
{
	try
	{
		if(m_ICurrentProjectRecordset==NULL || m_ICurrentProjectRecordset->GetState()==adStateClosed)
			return;

		UpdateData();

		if(m_ICurrentProjectRecordset->adoEOF)
			return;

		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
		}

		m_ICurrentProjectRecordset->MoveNext();

		if(m_ICurrentProjectRecordset->adoEOF)
			m_ICurrentProjectRecordset->MoveLast();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	PutDataToEveryControl();

	UpdateData(FALSE);
	UpdateControlsState();
}

////////////////////////////////////////////////
//
// 响应“最后一条记录”按扭
//
void CEditOilPaintDataDlg::OnMoveLast() 
{
	try
	{
		if(m_ICurrentProjectRecordset==NULL || m_ICurrentProjectRecordset->GetState()==adStateClosed)
			return;

		UpdateData();

		if(m_ICurrentProjectRecordset->adoEOF)
			return;
		
		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
		}

		m_ICurrentProjectRecordset->MoveLast();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	PutDataToEveryControl();

	UpdateData(FALSE);
	UpdateControlsState();
}

//////////////////////////////////////////
//
// 响应“增加记录”按扭
//
void CEditOilPaintDataDlg::OnAddNew() 
{
	UpdateData(TRUE);

	int Num=0;
	if(GetProjectID().IsEmpty())
		return;

	try
	{
		if(!m_ICurrentProjectRecordset->adoEOF || !m_ICurrentProjectRecordset->BOF)
		{
			//
			// 判断是否应该将数据存入数据库
			//
			if(m_IsUpdateByRoll)
			{
				PutDataToDatabaseFromControl();
			}

			m_ICurrentProjectRecordset->MoveFirst();

			while(!m_ICurrentProjectRecordset->adoEOF)
			{
				Num++;
				m_ICurrentProjectRecordset->MoveNext();
			}
		}

		Num++;

		m_ICurrentProjectRecordset->AddNew();

		PutTbValue(m_ICurrentProjectRecordset,_variant_t("ID"),Num);
		PutTbValue(m_ICurrentProjectRecordset,_variant_t("EnginID"),GetProjectID());

	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}
	//新建一条记录时，只将记录号自动加1。其他的继承上一条记录的值。 3.5
	//只继承卷册号.  2005.3.6
	CString strVolumeID;
	strVolumeID = m_VolNo;

	PutDataToEveryControl();        //该语句不继承。全部赋空值。

	m_VolNo = strVolumeID; //只继承卷册号
	UpdateData(FALSE);

	UpdateControlsState();
}

////////////////////////////////////////////////
//
// 响应“删除当前记录”按扭
//
void CEditOilPaintDataDlg::OnDelCurrent() 
{
	long Num,CurPos;

	if(AfxMessageBox(_T("是否删除当前记录"),MB_YESNO) == IDNO)
	{
		return;
	}
	
	try
	{
		if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
		{
			return;
		}

		CurPos=RecNo(m_ICurrentProjectRecordset);

		m_ICurrentProjectRecordset->Delete(adAffectCurrent);
		m_ICurrentProjectRecordset->Update();

		m_ICurrentProjectRecordset->MoveFirst();

		if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
		{
			UpdateControlsState();
			return;
		}

		Num=1;
		while(!m_ICurrentProjectRecordset->adoEOF)
		{
			m_ICurrentProjectRecordset->PutCollect(_variant_t("ID"),_variant_t(Num));
			m_ICurrentProjectRecordset->Update();

			Num++;
			m_ICurrentProjectRecordset->MoveNext();
		}

		m_ICurrentProjectRecordset->MoveFirst();
		while(TRUE)
		{
			if(m_ICurrentProjectRecordset->adoEOF)
				break;

			CurPos--;
			if(CurPos>0)
				m_ICurrentProjectRecordset->MoveNext();
			else
				break;
		}

		if(m_ICurrentProjectRecordset->adoEOF)
			m_ICurrentProjectRecordset->MovePrevious();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	PutDataToEveryControl();
	UpdateControlsState();
}

//////////////////////////////////////////////////
//
// 响应“保存当前”按扭
//
void CEditOilPaintDataDlg::OnSaveCurrent() 
{
	PutDataToDatabaseFromControl();	
}

//////////////////////////////////////////////
//
// 更新各个控件的状态
//
void CEditOilPaintDataDlg::UpdateControlsState()
{
	CString TempStr;

	UINT uControlIDs[]=
	{
		IDC_ID,
		IDC_VOLNO,
		IDC_PIPE_DEVICE_NAME,
		IDC_OILPAINT_COLOR,
		IDC_OILPAINT_TYPE,
		IDC_DEVICE_OUTSIDE_AREA,
		IDC_PIPE_SIZE,
		IDC_PIPE_LENGTH_DEVICE,
		IDC_REMARK,
		IDC_UPDATABYROLL,
		IDC_FOREFRONT,
		IDC_PREVIOUS,
		IDC_SUBSEQUENT,
		IDC_LAST,
		IDC_DELCURRENT,
		IDC_SAVECURRENT,
		IDC_Del_All_Paint
	};

	if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
	{
		for(int i=0;i<sizeof(uControlIDs)/sizeof(UINT);i++)
		{
			GetDlgItem(uControlIDs[i])->EnableWindow(FALSE);
		}
	}
	else
	{
		for(int i=0;i<sizeof(uControlIDs)/sizeof(UINT);i++)
		{
			GetDlgItem(uControlIDs[i])->EnableWindow(TRUE);
		}
	}

	UpdateData();
	try
	{
		if(m_ICurrentProjectRecordset->BOF)
		{
			m_MoveForefront.EnableWindow(FALSE);
			m_MovePrevious.EnableWindow(FALSE);
		}
		else
		{
			m_ICurrentProjectRecordset->MovePrevious();
			
			if(m_ICurrentProjectRecordset->BOF)
			{
				m_MoveForefront.EnableWindow(FALSE);
				m_MovePrevious.EnableWindow(FALSE);
				m_ICurrentProjectRecordset->MoveFirst();
				
			}
			else
			{
				m_MoveForefront.EnableWindow(TRUE);
				m_MovePrevious.EnableWindow(TRUE);
				m_ICurrentProjectRecordset->MoveNext();
			}		
		}
		
		if(m_ICurrentProjectRecordset->adoEOF)
		{
			m_MoveLast.EnableWindow(FALSE);
			m_MoveSubsequent.EnableWindow(FALSE);
		}
		else
		{
			m_ICurrentProjectRecordset->MoveNext();
			
			if(m_ICurrentProjectRecordset->adoEOF)
			{
				m_MoveLast.EnableWindow(FALSE);
				m_MoveSubsequent.EnableWindow(FALSE);
				m_ICurrentProjectRecordset->MoveLast();
			}
			else
			{
				m_MoveLast.EnableWindow(TRUE);
				m_MoveSubsequent.EnableWindow(TRUE);
				m_ICurrentProjectRecordset->MovePrevious();
			}
		}
	}
	catch (_com_error& e) 
	{
		AfxMessageBox( e.Description() );
		return;
	}
}

///////////////////////////////////////////////////////
//
// 更新“管道设备名称”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOilPaintDataDlg::UpdatePipeDeviceNameCombox()
{
	CString SQL,strTemp,strText;
	CString strVolNo,strPipeDeviceName;
	_RecordsetPtr IRecordset_Vol;
	int iRet;
	HRESULT hr;

	m_PipeDeviceName.GetWindowText(strText);

	do
	{
		iRet=m_PipeDeviceName.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Vol.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	UpdateData();

	strVolNo=m_VolNo;

	strVolNo.TrimLeft();
	strVolNo.TrimRight();

	if(strVolNo.IsEmpty())
		return FALSE;

	SQL.Format(_T("SELECT DISTINCT [VOL_NAME] FROM [a_vol] WHERE LEFT('%s',5)=LEFT(vol,5) OR VOL_BAK='UESoft'  "),strVolNo);

	try
	{
		IRecordset_Vol->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
							adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Vol->adoEOF)
		{
			GetTbValue(IRecordset_Vol,_variant_t("vol_name"),strTemp);

			strTemp.TrimLeft();
			strTemp.TrimRight();

			if(!strTemp.IsEmpty())
				m_PipeDeviceName.AddString(strTemp);

			IRecordset_Vol->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_PipeDeviceName.SetWindowText(strText);

	return TRUE;
}

////////////////////////////////////////////////////////////
//
// 当在a_vol中找不到对话框中的卷册号时,更新a_vol表
//
// 函数成功返回TRUE,否则返回FALSE
//
BOOL CEditOilPaintDataDlg::UpdateVolTable()
{
	CString SQL,strPipeDeviceName,strVol,strTemp;
	HRESULT hr;
	_RecordsetPtr IRecordset_Vol;

	hr=IRecordset_Vol.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	UpdateData();

	strVol=m_VolNo;
	m_PipeDeviceName.GetWindowText(strPipeDeviceName);

	strVol.TrimLeft();
	strVol.TrimRight();
	strPipeDeviceName.TrimLeft();
	strPipeDeviceName.TrimRight();

	strVol=strVol.Left(5);

	if(strVol.IsEmpty() || strPipeDeviceName.IsEmpty())
		return TRUE;

	SQL.Format(_T("SELECT * FROM a_vol WHERE LEFT(vol,5)='%s' AND vol_name='%s'"),strVol,strPipeDeviceName);

	try
	{
		IRecordset_Vol->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
							adOpenDynamic,adLockOptimistic,adCmdText);

		if(IRecordset_Vol->adoEOF && IRecordset_Vol->BOF)
		{
			IRecordset_Vol->AddNew();

//			IRecordset_Vol->PutCollect(_variant_t("vol"),_variant_t(strVol));
			PutTbValue(IRecordset_Vol,_variant_t("vol"),strVol);
//			IRecordset_Vol->PutCollect(_variant_t("vol_name"),_variant_t(strPipeDeviceName));
			PutTbValue(IRecordset_Vol,_variant_t("vol_name"),strPipeDeviceName);

//			IRecordset_Vol->Update();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}


void CEditOilPaintDataDlg::OnDropdownPipeDeviceName() 
{
	UpdatePipeDeviceNameCombox();
	
}

/////////////////////////////////////////////
//
// 响应"退出"按钮
//
void CEditOilPaintDataDlg::OnExit() 
{
	UpdateData();

	if(m_IsUpdateByRoll)
	{
		PutDataToDatabaseFromControl();
	}
	
	ProcessInExit();
	//移动时是否更新记录集
	bIsMoveUpdate = m_IsUpdateByRoll;
	OnOK();	
	//2005.2.24
	EndDialog(0);
	if( IsWindow(this->GetSafeHwnd()) )
	{
		DestroyWindow();
	}

	WritePassTbl();
}

void CEditOilPaintDataDlg::ProcessInExit()
{
	CString SQL;
	CString e_type,strTemp;
	_RecordsetPtr IRecordset_Pai;
	HRESULT hr;

	if(GetPublicConnect()==NULL || GetPublicConnect()->GetState()==adStateClosed)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return;
	}

	if(m_ICurrentProjectRecordset==NULL || m_ICurrentProjectRecordset->GetState()==adStateClosed)
	{
		AfxMessageBox(_T("记录集未打开"));
		return;
	}

	if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
	{
//		AfxMessageBox(_T("记录集为空"));
		return;
	}

	hr=IRecordset_Pai.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return;
	}

//	SQL.Format(_T("SELECT * FROM A_PAI"));
	SQL.Format(_T("SELECT * FROM A_PAI WHERE EnginID='%s'"),GetProjectID());

	try
	{
//		IRecordset_Pai->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Pai->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							adOpenStatic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	try
	{
		if (IRecordset_Pai->GetRecordCount() > 0 )
		{
			IRecordset_Pai->MoveFirst();
			while(!IRecordset_Pai->adoEOF)
			{
				GetTbValue(IRecordset_Pai,_variant_t("PAI_TYPE"),e_type);
				
				m_ICurrentProjectRecordset->MoveFirst();
				
				//		REPL ALL pai_name1  WITH p->pai_n1,pai_times1 WITH p->pai_t1,;
				//				pai_cost1  WITH p->pai_c1,pai_name2  WITH p->pai_n2,;
				//				pai_times2 WITH p->pai_t2,pai_cost2  WITH p->pai_c2,;
				//				pai_a1     WITH p->pai_a1,pai_a_t1   WITH p->pai_a_t1,;
				//				pai_a2     WITH p->pai_a2,pai_a_t2   WITH p->pai_a_t2;
				//				FOR pai_type=e_type
				
				while(!m_ICurrentProjectRecordset->adoEOF)
				{
					GetTbValue(m_ICurrentProjectRecordset,_variant_t("pai_type"),strTemp);
					
					if(strTemp==e_type)
					{
						struct
						{
							LPCTSTR szFrom;
							LPCTSTR szTo;
						}FromTo[]=
						{
							{_T("pai_n1"),		_T("pai_name1")},
							{_T("pai_t1"),		_T("pai_times1")},
							{_T("pai_c1"),		_T("pai_cost1")},
							{_T("pai_n2"),		_T("pai_name2")},
							{_T("pai_t2"),		_T("pai_times2")},
							{_T("pai_c2"),		_T("pai_cost2")},
							{_T("pai_a1"),		_T("pai_a1")},
							{_T("pai_a_t1"),	_T("pai_a_t1")},
							{_T("pai_a2"),		_T("pai_a2")},
							{_T("pai_a_t2"),	_T("pai_a_t2")}
						};
						
						for(int i=0;i<sizeof(FromTo)/sizeof(FromTo[0]);i++)
						{
							GetTbValue(IRecordset_Pai,_variant_t(FromTo[i].szFrom),strTemp);
							PutTbValue(m_ICurrentProjectRecordset,_variant_t(FromTo[i].szTo),strTemp);
						}
						// pai_name1  WITH p->pai_n1
						//				GetTbValue(IRecordset_Pai,_variant_t("pai_n1"),strTemp);
						//				PutTbValue(m_ICurrentProjectRecordset,_variant_t("pai_name1"),strTemp);
						
					}// if(strTemp==e_type)
					
					m_ICurrentProjectRecordset->MoveNext();
					
				}// while(!m_ICurrentProjectRecordset->adoEOF)
				
				IRecordset_Pai->MoveNext();
				
			}// while(!IRecordset_Pai->adoEOF)
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}
	
//	SELE ed_pai
//	***************
//	pi=3.1415926
//	REPL ALL pai_tcost1 WITH pai_cost1*pai_times1*pai_area*pai_amou,;
//		pai_tcost2 WITH pai_cost2*pai_times2*pai_area*pai_amou,
//	pai_p_area WITH pai_area*pai_amou FOR pai_area<>0 .AND. pai_size=0

	try
	{
		SQL.Format(_T("UPDATE PAINT SET pai_tcost1=pai_cost1*pai_times1*pai_area*pai_amou WHERE EnginId='%s' AND		\
					pai_area<>0 AND pai_size=0"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);

		SQL.Format(_T("UPDATE PAINT SET pai_tcost2=pai_cost2*pai_times2*pai_area*pai_amou WHERE EnginId='%s' AND		\
					pai_area<>0 AND pai_size=0"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);

		SQL.Format(_T("UPDATE PAINT SET pai_p_area=pai_area*pai_amou WHERE EnginId='%s' AND							\
					pai_area<>0 AND pai_size=0"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);

	//	REPL ALL pai_tcost1 WITH pai_cost1*pai_times1*pai_size*pai_amou*pi/1000,
	//	pai_tcost2 WITH pai_cost2*pai_times2*pai_size*pai_amou*pi/1000,
	//	pai_p_area WITH	pai_amou*pai_size*pi/1000 
	//	FOR pai_area=0 .AND. pai_size<>0
		SQL.Format(_T("UPDATE PAINT SET pai_tcost1=pai_cost1*pai_times1*pai_size*pai_amou*3.1415926/1000 WHERE		\
			EnginId='%s' AND pai_area=0 AND pai_size<>0	"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);

		SQL.Format(_T("UPDATE PAINT SET pai_tcost2=pai_cost2*pai_times2*pai_size*pai_amou*3.1415926/1000 WHERE		\
			EnginId='%s' AND pai_area=0 AND pai_size<>0"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);

		SQL.Format(_T("UPDATE PAINT SET pai_p_area=pai_amou*pai_size*3.1415926/1000	WHERE							\
			EnginId='%s' AND pai_area=0 AND pai_size<>0"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);

	//	REPL ALL pai_a_c1 WITH pai_a_t1*pai_p_area,;
	//			 pai_a_c2 WITH pai_a_t2*pai_p_area
		SQL.Format(_T("UPDATE PAINT SET pai_a_c1=pai_a_t1*pai_p_area WHERE EnginId='%s'"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);

		SQL.Format(_T("UPDATE PAINT SET pai_a_c2=pai_a_t2*pai_p_area WHERE EnginId='%s'"),GetProjectID());
		GetCurrentProjectConnect()->Execute(_bstr_t(SQL),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
	}

}

void CEditOilPaintDataDlg::OnChangeId() 
{
	// TODO: If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.
	
	int Num;

	UpdateData();

	try
	{
		if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
		{
			m_ID=0;
			UpdateData(FALSE);
			return;
		}

		if(m_ID<=0)
			Num=1;
		else
			Num=m_ID;

		m_ICurrentProjectRecordset->MoveFirst();
		
		while(!m_ICurrentProjectRecordset->adoEOF)
		{
			Num--;
			if(Num>0)
				m_ICurrentProjectRecordset->MoveNext();
			else
				break;
		}

		if(m_ICurrentProjectRecordset->adoEOF)
			m_ICurrentProjectRecordset->MovePrevious();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	PutDataToEveryControl();
	UpdateControlsState();
}

//设置窗口为正中
BOOL CEditOilPaintDataDlg::SetWindowCenter(HWND hWnd)
{
	if( hWnd == NULL )
		return false;
	long SW, SH, xPos, yPos;

	SW = ::GetSystemMetrics(SM_CXSCREEN);
	SH = ::GetSystemMetrics(SM_CYSCREEN);
	CRect rc;
	::GetWindowRect(hWnd, &rc);

	xPos = (SW - rc.Width() )/2;
	yPos = (SH - rc.Height() )/2;

	MoveWindow( xPos, yPos, rc.Width(), rc.Height() );

	return true;
}
void CEditOilPaintDataDlg::OnCancel(void)
{
 	EndDialog(0);
 	if( IsWindow(this->GetSafeHwnd()) )
	{
 		DestroyWindow();
 	}

 	CDialog::OnClose();
}
//删除所有的油漆原始数据。
void CEditOilPaintDataDlg::OnDelAllPaint() 
{
	if(	AfxMessageBox(_T("是否删除所有的原始数据!"),MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON2) == IDNO )
	{
		return ;
	}
	try
	{
		CString strSQL;
		strSQL = "DELETE * FROM [PAINT] WHERE EnginID='"+GetProjectID()+"' ";
		m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, -1);

		if( this->InitCurrentProjectRecordset() != TRUE )
		{
			OnCancel();
			return;
		}
		this->UpdateControlsState();
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return;
	}
}
//油漆说明表之后的操作。
BOOL CEditOilPaintDataDlg::WritePassTbl()
{
	//
	// 更改PASS表
	//
	CString SQL,strTemp;
	_RecordsetPtr IRecordset_Paint;
	_RecordsetPtr IRecordset_Pass;
	_variant_t varTemp;
	HRESULT hr;
	int Pos;

	hr=IRecordset_Paint.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		ReportExceptionErrorLV1(_com_error((HRESULT)hr));
		return false;
	}
		
	IRecordset_Paint->CursorLocation=adUseClient;

	hr=IRecordset_Pass.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		ReportExceptionErrorLV1(_com_error((HRESULT)hr));
		return false;
	}

	SQL.Format(_T("SELECT * FROM PAINT WHERE EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		IRecordset_Paint->Open(_variant_t(SQL),_variant_t((IDispatch*)theApp.m_pConnection),
							   adOpenStatic,adLockOptimistic,adCmdText);

		SQL.Format(_T("SELECT * FROM  PASS WHERE EnginID='%s'"),EDIBgbl::SelPrjID);

		IRecordset_Pass->Open(_variant_t(SQL),_variant_t((IDispatch*)theApp.m_pConnection),
							   adOpenStatic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return false;
	}

	Pos=0;

	try
	{
		//在控制操作流程的表中没有当前工程的记录.
		if ( IRecordset_Pass->GetRecordCount() <= 0)
		{
			return FALSE;
		}
		IRecordset_Pass->MoveFirst();
		while(!IRecordset_Pass->adoEOF)
		{
			Pos++;

			varTemp=IRecordset_Pass->GetCollect(_variant_t("pass_mod2"));

			if(varTemp.vt==VT_NULL)
			{
				strTemp=_T("");
			}
			else
			{
				strTemp=(LPCTSTR)(_bstr_t)varTemp;
				strTemp.TrimLeft();
				strTemp.TrimRight();

				strTemp.MakeUpper();
			}

			if(strTemp==_T("E_PAIDAT"))
				break;

			IRecordset_Pass->MoveNext();
		}

		if(IRecordset_Pass->adoEOF)
		{
			ReportExceptionErrorLV1(_T(" cann't find pass_mod1==E_PREDAT"));
			return false;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return false;
	}

	try
	{
		IRecordset_Paint->Filter=_variant_t("pai_p_area=0");

		if(IRecordset_Paint->adoEOF && IRecordset_Paint->BOF)
		{

			IRecordset_Pass->MoveFirst();

			while(Pos>0)
			{
				IRecordset_Pass->PutCollect(_variant_t("pass_mark2"),_variant_t("T"));
				IRecordset_Pass->MoveNext();
				
				Pos--;
			}

			while(!IRecordset_Pass->adoEOF)
			{
				IRecordset_Pass->PutCollect(_variant_t("pass_mark2"),_variant_t("F"));
				IRecordset_Pass->MoveNext();
			}

		}
		else
		{
			while(!IRecordset_Pass->adoEOF)
			{
				IRecordset_Pass->PutCollect(_variant_t("pass_mark2"),_variant_t("F"));
				IRecordset_Pass->MoveNext();
			}
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return false;
	}
	return true;
}

BOOL CEditOilPaintDataDlg::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	m_pctlTip.RelayEvent(pMsg);		//提示信息

	return CDialog::PreTranslateMessage(pMsg);
}

//------------------------------------------------------------------                
// DATE         :[2005/06/23]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :加载需要提示信息的编辑框与之对应的提示符。
//------------------------------------------------------------------
BOOL CEditOilPaintDataDlg::InitToolTip()
{

	int	pos = 0;		//数组下标
	Txtup	TipArr[20];
	
	TipArr[pos].dID = IDC_DEVICE_OUTSIDE_AREA;	//设备外表面积
	TipArr[pos++].txt = _T("㎡");				//单位

	TipArr[pos].dID = IDC_PIPE_SIZE;			//管道外径
	TipArr[pos++].txt = _T("mm");				//单位

	TipArr[pos].dID = IDC_PIPE_LENGTH_DEVICE;	//管道长度/设备数
	TipArr[pos++].txt = _T("m 或 件");				//单位

	m_pctlTip.Create(this);
	m_pctlTip.SetTipBkColor(RGB(255,255,255));
	for (int i=0; i<pos; i++)
	{
		m_pctlTip.AddTool(GetDlgItem(TipArr[i].dID), (TipArr[i].txt));
	}

	return TRUE;

}
