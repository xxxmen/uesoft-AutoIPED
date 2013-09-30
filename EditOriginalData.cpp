// EditOriginalData.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "EditOriginalData.h"
#include "vtot.h"
#include "MainFrm.h"
#include "AutoIPEDDoc.h"

#include "uewasp.h"
#include "EDIBgbl.h"

extern CAutoIPEDApp theApp;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define atmospheric_pressure 0.101325
/////////////////////////////////////////////////////////////////////////////
// CEditOriginalData dialog

CEditOriginalData::CEditOriginalData(CWnd* pParent /*=NULL*/)
	: CDialog(CEditOriginalData::IDD, pParent)
{
	//{{AFX_DATA_INIT(CEditOriginalData)
	m_VolNo = _T("");
	m_ID = 0;
	m_ReMark = _T("");
	m_IsVolListEnable = FALSE;
	m_IsHeatPreservationTypeEnable = FALSE;
	m_WideSpeed = 0.0f;
	m_PriceRatio = 0.0f;
	m_RunPerHour = 0.0f;
	m_MediumTemperature = 0.0f;
	m_IsUpdateByRoll = TRUE;
	m_Amount = 0;
	m_bIsAutoSelectMatOnRoll = FALSE;
	m_dHeatLoss = "";
	m_dHeatDensity = "";
	m_dThick1 = 0.0;
	m_dThick2 = 0.0;
	m_dTa = 0.0;
	m_dTs = 0.0;
	m_dPressure = 0.0;
	m_bIsCalInThick = FALSE;
	m_bIsCalPreThick = FALSE;
	m_dWindPlaimWidth = 0.0;
	//}}AFX_DATA_INIT

//	m_bIsAutoSelectMatOnRoll=FALSE;
	m_ProjectID=_T("");
	m_dMaxD0 = 2000;

	m_pDlgCongeal		= NULL;
	m_pDlgSubterranean	= NULL;
	m_pDlgCurChild		= NULL;
	m_pMatProp			= NULL;
} 


CEditOriginalData::~CEditOriginalData()
{
	// 释放所有动态创建的类成员变量
	if ( m_pDlgCongeal != NULL )
	{
		delete m_pDlgCongeal;
	}
	if ( m_pDlgSubterranean != NULL )
	{
		delete m_pDlgSubterranean;
	}
}


void CEditOriginalData::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CEditOriginalData)
	DDX_Control(pDX, IDC_TITLE_TAB, m_ctlTitleTab);
	DDX_Control(pDX, IDC_PIPE_MEDIUM, m_ctlPipeMedium);
	DDX_Control(pDX, IDC_EDIT_METHOD, m_ctlCalcMethod);
	DDX_Control(pDX, IDC_HEAT_TRANSFER_COEF, m_ctlHeatTran);
	DDX_Control(pDX, IDC_ENVIRON_TEMP, m_ctlEnvTemp);
	DDX_Control(pDX, IDC_DEL_ALL_INSU, m_DeleteAll);
	DDX_Control(pDX, IDC_COLOR, m_ColorRing);
	DDX_Control(pDX, IDC_SAVECURRENT, m_SaveCurrent);
	DDX_Control(pDX, IDC_DELCURRENT, m_DelCurrent);
	DDX_Control(pDX, IDC_ADDNEW, m_AddNew);
	DDX_Control(pDX, IDC_LAST, m_MoveLast);
	DDX_Control(pDX, IDC_SUBSEQUENT, m_MoveSubsequent);
	DDX_Control(pDX, IDC_PREVIOUS, m_MovePrevious);
	DDX_Control(pDX, IDC_FOREFRONT, m_MoveForefront);
	DDX_Control(pDX, IDC_SAFEGUARD_MATNAME, m_SafeguardMatName);
	DDX_Control(pDX, IDC_OUTSIZE_MATNAME, m_OutSizeMatName);
	DDX_Control(pDX, IDC_INSIDE_MATNAME, m_InsideMatName);
	DDX_Control(pDX, IDC_HEAT_PRESERVATION_TYPE, m_HeatPreservationTypeList);
	DDX_Control(pDX, IDC_VOLLIST, m_VolList);
	DDX_Control(pDX, IDC_PIPE_MAT, m_PipeMat);
	DDX_Control(pDX, IDC_BUILDIN_PLACE, m_BuidInPlace);
	DDX_Control(pDX, IDC_PIPE_THICK, m_PipeThick);
	DDX_Control(pDX, IDC_PIPE_SIZE, m_PipeSize);
	DDX_Control(pDX, IDC_PIPE_DEVICE_NAME, m_PipeDeviceName);
	DDX_Text(pDX, IDC_VOLNO, m_VolNo);
	DDX_Text(pDX, IDC_ID, m_ID);
	DDX_Text(pDX, IDC_REMARK, m_ReMark);
	DDX_Check(pDX, IDC_VOLLIST_ISENABLE, m_IsVolListEnable);
	DDX_Check(pDX, IDC_HEAT_PRESERVATION_TYPE_ENABLE, m_IsHeatPreservationTypeEnable);
	DDX_Text(pDX, IDC_WIDE_SPEED, m_WideSpeed);
	DDX_Text(pDX, IDC_PRICE, m_PriceRatio);
	DDX_Text(pDX, IDC_RUN_PERHOUR, m_RunPerHour);
	DDX_Text(pDX, IDC_MEDIUM_TEMPERATURE, m_MediumTemperature);
	DDX_Check(pDX, IDC_UPDATABYROLL, m_IsUpdateByRoll);
	DDX_Text(pDX, IDC_AMOUNT, m_Amount);
	DDX_Check(pDX, IDC_CHECK_AUTOSELECT_PRE_EDIT, m_bIsAutoSelectMatOnRoll);
	DDX_Text(pDX, IDC_EDIT_HEAT_LOSS, m_dHeatLoss);
	DDX_Text(pDX, IDC_EDIT_HEAT_DENSITY, m_dHeatDensity);
	DDX_Text(pDX, IDC_EDIT_THICK1, m_dThick1);
	DDX_Text(pDX, IDC_EDIT_THICK2, m_dThick2);
	DDX_Text(pDX, IDC_EDIT_TA, m_dTa);
	DDX_Text(pDX, IDC_EDIT_TS, m_dTs);
	DDX_Text(pDX, IDC_PIPE_PRESSURE, m_dPressure);
	DDX_Check(pDX, IDC_CHECK_THICK1, m_bIsCalInThick);
	DDX_Check(pDX, IDC_CHECK_THICK2, m_bIsCalPreThick);  
	DDX_Text(pDX, IDC_EDIT_PLAIM_THICK, m_dWindPlaimWidth);
	//}}AFX_DATA_MAP
}

IMPLEMENT_DYNCREATE(CEditOriginalData, CDialog)

BEGIN_MESSAGE_MAP(CEditOriginalData, CDialog)
	//{{AFX_MSG_MAP(CEditOriginalData)
	ON_BN_CLICKED(IDC_FOREFRONT, OnForefront)
	ON_BN_CLICKED(IDC_PREVIOUS, OnPrevious)
	ON_BN_CLICKED(IDC_SUBSEQUENT, OnSubsequent)
	ON_BN_CLICKED(IDC_LAST, OnLast)
	ON_BN_CLICKED(IDC_ADDNEW, OnAddNew)
	ON_BN_CLICKED(IDC_DELCURRENT, OnDelCurrent)
	ON_BN_CLICKED(IDC_SAVECURRENT, OnSaveCurrent)
	ON_CBN_DROPDOWN(IDC_COLOR, OnDropdownColor)
	ON_CBN_DROPDOWN(IDC_PIPE_DEVICE_NAME, OnDropdownPipeDeviceName)
	ON_CBN_DROPDOWN(IDC_PIPE_SIZE, OnDropdownPipeSize)
	ON_CBN_DROPDOWN(IDC_PIPE_THICK, OnDropdownPipeThick)
	ON_CBN_DROPDOWN(IDC_INSIDE_MATNAME, OnDropdownInsideMatname)
	ON_CBN_DROPDOWN(IDC_OUTSIZE_MATNAME, OnDropdownOutsizeMatname)
	ON_CBN_DROPDOWN(IDC_SAFEGUARD_MATNAME, OnDropdownSafeguardMatname)
	ON_NOTIFY(NM_CLICK, IDC_HEAT_PRESERVATION_TYPE, OnItemchangedHeatPreservationType)
	ON_NOTIFY(NM_CLICK, IDC_VOLLIST, OnItemchangedVollist)
	ON_BN_CLICKED(IDC_VOLLIST_ISENABLE, OnUpdateAllCheckBox)
	ON_EN_CHANGE(IDC_VOLNO, OnChangeVolno)
	ON_BN_CLICKED(IDC_AUTO_SELECTMAT, OnAutoSelectmat)
	ON_EN_CHANGE(IDC_ID, OnChangeId)
	ON_BN_CLICKED(IDC_SORT_BY_VOL, OnSortByVol)
	ON_BN_CLICKED(IDC_EXIT, OnExit)
	ON_BN_CLICKED(IDC_DEL_ALL_INSU, OnDelAllInsu)
	ON_BN_CLICKED(IDC_CHECK_METHOD, OnCheckMethod)
	ON_BN_CLICKED(IDC_CHECK_TA, OnCheckTa)
	ON_BN_CLICKED(IDC_CHECK_THICK1, OnCheckThick1)
	ON_BN_CLICKED(IDC_CHECK_THICK2, OnCheckThick2)
	ON_BN_CLICKED(IDC_CHECK_TS, OnCheckTs)
	ON_CBN_SELCHANGE(IDC_ENVIRON_TEMP, OnSelchangeEnvironTemp)
	ON_CBN_SELCHANGE(IDC_HEAT_TRANSFER_COEF, OnSelchangeHeatTransferCoef)
	ON_BN_CLICKED(IDC_CHECK_ENV_TEMP, OnCheckEnvTemp)
	ON_BN_CLICKED(IDC_CHECK_ALPHA, OnCheckAlpha)
	ON_CBN_SELCHANGE(IDC_EDIT_METHOD, OnSelchangeEditMethod)
	ON_CBN_SELCHANGE(IDC_BUILDIN_PLACE, OnSelchangeBuildinPlace)
	ON_CBN_SELCHANGE(IDC_PIPE_MEDIUM, OnSelchangePipeMedium)
	ON_BN_CLICKED(IDC_BUTTON_ADD_VALVE, OnButtonAddValve)
	ON_BN_CLICKED(IDC_BUTTON_CALC_RATIO, OnButtonCalcRatio)
	ON_BN_CLICKED(IDC_HEAT_PRESERVATION_TYPE_ENABLE, OnUpdateAllCheckBox)
	ON_CBN_SELCHANGE(IDC_PIPE_MAT, OnSelchangePipeMat)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEditOriginalData message handlers

//////////////////////////////////////////////////////////////////////////
//
// 设置所选工程的数据库的连接
//
// IConnection[in]	已经连接的数据库的智能指针
//
void CEditOriginalData::SetCurrentProjectConnect(_ConnectionPtr &IConnection)
{
	m_ICurrentProjectConnection=IConnection;
}

//////////////////////////////////////////////////////////////////////////
//
// 返回数据库连接的智能指针的引用
//
_ConnectionPtr& CEditOriginalData::GetCurrentProjectConnect()
{
	return m_ICurrentProjectConnection;
}

//////////////////////////////////////////////////////////////////////////
//
// 设置公共的数据库的连接
//
// IConnection[in]	已经连接的数据库的智能指针
//
void CEditOriginalData::SetPublicConnect(_ConnectionPtr &IConnection)
{
	m_IPublicConnection=IConnection;

}

//////////////////////////////////////////////////////////////////////////
//
// 返回公共数据库连接的智能指针的引用
//
_ConnectionPtr& CEditOriginalData::GetPublicConnect()
{
	return m_IPublicConnection;
}

/////////////////////////////////////////////////////////////
//
// 设置所选工程的ID号
//
// ProjectID[in]	所选工程的ID号
//
void CEditOriginalData::SetProjectID(LPCTSTR ProjectID)
{
	if(ProjectID==NULL)
		return;

	m_ProjectID=ProjectID;
}

//////////////////////////////////////////////////////////
//
// 返回所选工程的ID号
//
CString& CEditOriginalData::GetProjectID()
{
	return m_ProjectID;
}

////////////////////////////////////////////////////////
//
// 打开PRE_CALC表
// 打开防冻结计算时须要的参数表		ZSY		[2005/06/01] 
// 打开埋地管道的保温参数表
BOOL CEditOriginalData::InitCurrentProjectRecordset()
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
		if( m_ICurrentProjectRecordset==NULL )
		{
			m_ICurrentProjectRecordset.CreateInstance(__uuidof(Recordset));
		}
		else if( m_ICurrentProjectRecordset->GetState() & adStateOpen )
		{
			m_ICurrentProjectRecordset->Close();
		}
		//打开PRE_CALC表
		SQL.Format(_T("SELECT * FROM pre_calc WHERE EnginID='%s' ORDER BY id"),GetProjectID());
		m_ICurrentProjectRecordset->Open(_variant_t(SQL),GetCurrentProjectConnect().GetInterfacePtr(),adOpenStatic,adLockOptimistic,adCmdText);

		//打开防冻结计算时须要的参数表 
		if ( m_pRsCongealData == NULL )
		{
			m_pRsCongealData.CreateInstance(__uuidof(Recordset));
		}
		else if ( m_pRsCongealData->GetState() == adStateOpen )
		{
			m_pRsCongealData->Close();
		}
		SQL.Format("SELECT * FROM [PRE_CALC_CONGEAL] WHERE EnginID = '"+GetProjectID()+"' ORDER BY ID ");
		m_pRsCongealData->Open(_variant_t(SQL), GetCurrentProjectConnect().GetInterfacePtr(),
						  adOpenStatic, adLockOptimistic, adCmdText);

		// 埋地管道的参数表
		if ( NULL == m_pRsSubterranean )
		{
			m_pRsSubterranean.CreateInstance( __uuidof( Recordset ) );
		}
		else if ( m_pRsSubterranean->GetState() == adStateOpen )
		{
			m_pRsSubterranean->Close();
		}
		SQL.Format("SELECT * FROM [PRE_CALC_SUBTERRANEAN] WHERE EnginID='"+GetProjectID()+"' ORDER BY ID");
		m_pRsSubterranean->Open( _variant_t(SQL), GetCurrentProjectConnect().GetInterfacePtr(),
							adOpenStatic, adLockOptimistic, adCmdText);

	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}

BOOL CEditOriginalData::OnInitDialog() 
{
	BOOL bRet;
	m_pRsEnvironment.CreateInstance(__uuidof(Recordset));		//环境温度取值索引记录集
	m_pRsHeat.CreateInstance(__uuidof(Recordset));				//放热系数取值索引记录集

	CDialog::OnInitDialog();
	//移动时是否更新记录
	m_IsUpdateByRoll = bIsMoveUpdate;
	//TRUE--平壁标志,FALSE--管道
	m_bIsPlane = FALSE;
	//提示信息(显示单位)			 [2005/06/23]
	InitToolTip();

	//打开原始数据。
	bRet=InitCurrentProjectRecordset();
	if(bRet==FALSE)
	{
		CDialog::OnCancel();
		return FALSE;
	}

	bRet=InitVolListControl();
	if(bRet==FALSE)
	{
		CDialog::OnCancel();
		return TRUE;
	}
	//初始环境温度的取值.
	UpdateEnvironmentComBox();
	//
	//放热系数.
	UpdateHeatTransferCoef();
	//选择计算方法
	InitCalcMethod();
	//
	//管内介质
	InitPipeMedium();
	//
	//安装地点
	UpdateBuildPlaceComBox();


	bRet=InitHeatPreservationTypeList();
	if(bRet==FALSE)
	{
		CDialog::OnCancel();
		return TRUE;
	}
	
	bRet=InitBitmapButton();
	if(bRet==FALSE)
	{
		CDialog::OnCancel();
		return TRUE;
	}

	bRet=PutDataToEveryControl();
	if(bRet==FALSE)
	{
		CDialog::OnCancel();
		return TRUE;
	}
	//初始常量
	InitConstant();
	//更新各控件的状态。
	UpdateControlsState();

	//初始化材料控件
	GetPropertyofMaterial mGetPropertyofMaterial;
	int iCodeID = mGetPropertyofMaterial.GetCodeID( EDIBgbl::sCur_MaterialCodeName,theApp.m_pConMaterial );
	if ( iCodeID == -1 )
		iCodeID = 0;
	CArray<CString,CString> mAllMatName;
	mGetPropertyofMaterial.GetAllMatLanguageName( iCodeID,GetPropertyofMaterial::eSolid,theApp.m_pConMaterialName,theApp.m_pConMaterial,mAllMatName );
	int iCount = mAllMatName.GetSize();
	for ( int i=0; i<iCount; i++ )
	{
		m_PipeMat.AddString( mAllMatName.GetAt(i) );
	}

	UpdateData();

	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{	//没有记录时，不考虑.
		if( m_ICurrentProjectRecordset->GetRecordCount() >  0 )
		{
			OnAutoSelectmat();
		}
	}

	m_ctlTitleTab.ShowWindow(SW_SHOW);

	//设置“退出”为默认的按钮
	this->SetDefID(IDC_EXIT);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

///////////////////////////////////////////////////////
//
// 初始化卷册LIST控件 
//
// 如果初始化成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::InitVolListControl()
{
	CString SQL;
	CString TempStr1,TempStr2;
	_RecordsetPtr IRecordset;
	_variant_t varTemp;
	HRESULT hr;
	int nItem;
	RECT rect;
	CString strPriceRatio;	// 热价比主汽价

	m_VolList.GetWindowRect(&rect);
	m_VolList.InsertColumn(1,_T("卷册号"),LVCFMT_LEFT,(int)((rect.right-rect.left)/10.0*3.0));
	m_VolList.InsertColumn(2,_T("卷册名称"),LVCFMT_LEFT,(int)((rect.right-rect.left)/10.0*7.3));
	m_VolList.InsertColumn(3,_T("热价比主汽价"),LVCFMT_LEFT,(int)0);

	m_VolList.SetExtendedStyle(LVS_EX_FULLROWSELECT);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未建立公有数据库的连接"));
		return FALSE;
	}

	hr=IRecordset.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	SQL.Format(_T("SELECT * FROM a_vol WHERE [vol] IS NOT NULL AND [vol_name] IS NOT NULL AND  [VOL_BAK] <> 'UESoft' ORDER BY VOL,vol_name"));

	try
	{
		IRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
						adOpenDynamic,adLockOptimistic,adCmdText);

		nItem=0;
		while(!IRecordset->adoEOF)
		{

			TempStr1 =	vtos( IRecordset->GetCollect(_variant_t("vol")));		//卷册号
			TempStr2 =	vtos( IRecordset->GetCollect(_variant_t("vol_name")));	//卷册名
			strPriceRatio = vtos( IRecordset->GetCollect( _variant_t("vol_price")));	// 热价比主汽价
			if( TempStr1.IsEmpty() || TempStr2.IsEmpty() )
			{
				IRecordset->MoveNext();
				continue;
			}

			m_VolList.InsertItem(nItem,TempStr1);
			m_VolList.SetItemText(nItem,1,TempStr2);
			m_VolList.SetItemText(nItem,2,strPriceRatio);

			nItem++;
			IRecordset->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}

///////////////////////////////////////////////////////
//
// 初始化保温类型LIST控件 
//
// 如果初始化成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::InitHeatPreservationTypeList()
{
	CString SQL;
	CString TempStr1,TempStr2;
	_RecordsetPtr IRecordset;
	_variant_t varTemp;
	int nItem;
	RECT rect;

	m_HeatPreservationTypeList.GetWindowRect(&rect);
	m_HeatPreservationTypeList.InsertColumn(1,_T("类型代号"),LVCFMT_LEFT,(int)((rect.right-rect.left)/10.0*4.0));
	m_HeatPreservationTypeList.InsertColumn(2,_T("保温类型名称"),LVCFMT_LEFT,(int)((rect.right-rect.left)/10.0*7.3));

	m_HeatPreservationTypeList.SetExtendedStyle(LVS_EX_FULLROWSELECT);

	if(GetPublicConnect()==NULL)
	{
#ifdef _DEBUG
		AfxMessageBox(_T("为建立公有数据库的连接"));
#endif
		return FALSE;
	}

	IRecordset.CreateInstance(__uuidof(Recordset));

//	SQL.Format(_T("SELECT * FROM a_pre"));
	SQL.Format(_T("SELECT * FROM a_pre WHERE EnginID='%s'"), GetProjectID());

	try
	{
//		IRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//						adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset->Open(_variant_t(SQL),GetCurrentProjectConnect().GetInterfacePtr(), adOpenDynamic,adLockOptimistic,adCmdText);

		nItem=0;
		while(!IRecordset->adoEOF)
		{
			varTemp=IRecordset->GetCollect(_variant_t("AUTO_CODE"));
			if(varTemp.vt!=VT_NULL)
			{
				TempStr1 = vtos(varTemp);
			}
			else
			{
				IRecordset->MoveNext();
				continue;
			}

			varTemp=IRecordset->GetCollect(_variant_t("AUTO_MARK"));
			if(varTemp.vt!=VT_NULL)
			{
				TempStr2 = vtos(varTemp);
			}
			else
			{
				IRecordset->MoveNext();
				continue;
			}

			m_HeatPreservationTypeList.InsertItem(nItem,TempStr1);
			m_HeatPreservationTypeList.SetItemText(nItem,1,TempStr2);

			nItem++;
			IRecordset->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}

/////////////////////////////////////////////////
//
// 初始化位图按扭
//
// 如果初始化成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::InitBitmapButton()
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
		{&m_DeleteAll,      IDB_BITMAP10,   _T("删除所有的记录")}
	};

	for(int i=0;i<sizeof(ButtonInfo)/sizeof(ButtonInfo[0]);i++)
	{
		Bitmap.LoadBitmap(ButtonInfo[i].uBitmapID);

		ButtonInfo[i].pButton->SetBitmap((HBITMAP)Bitmap.Detach());
		ButtonInfo[i].pButton->SetTooltipText(ButtonInfo[i].szTooltipText,TRUE);
	}
	
	return TRUE;
}

/////////////////////////////////////////////////////////
//
// 从以打开的记录集的当前游标所在的位置取数据到各个控件
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::PutDataToEveryControl()
{
	CString TempStr;
//	int		nTmp;
	short   flg;
	short	nIndex;			//
		
	if(m_ICurrentProjectRecordset->adoEOF)
	{
		return TRUE;
	}
	try
	{
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("id"),m_ID);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_vol"),m_VolNo);		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_color"),TempStr);
		m_ColorRing.SetWindowText(TempStr);
		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_name1"),TempStr);
		m_PipeDeviceName.SetWindowText(TempStr);
		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_size"),TempStr);
		m_PipeSize.SetWindowText(TempStr);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_pi_thi"),TempStr);
		m_PipeThick.SetWindowText(TempStr);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_temp"),m_MediumTemperature);
		
		//安装地点
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_place"),TempStr);

		if (TempStr.IsEmpty())
		{
			m_BuidInPlace.SetCurSel(0);
		}
		else
		{
			TempStr = TempStr.Left(4);
			m_BuidInPlace.SelectString(-1,TempStr);
		}

		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_pi_mat"),TempStr);
		m_PipeMat.SetWindowText(TempStr);

		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_mark"),m_ReMark);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_name_in"),TempStr);
		m_InsideMatName.SetWindowText(TempStr);

		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_name2"),TempStr);
		m_OutSizeMatName.SetWindowText(TempStr);

		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_name3"),TempStr);
		m_SafeguardMatName.SetWindowText(TempStr);

		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_wind"),m_WideSpeed);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_price"),m_PriceRatio);
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_hour"),m_RunPerHour);
		
		//油管道的厚度,指定的厚度不由此输入。
		//	GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_pre_thi"),m_HeatPreservationThick);
		
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_amount"),m_Amount);		
		GetTbValue(m_ICurrentProjectRecordset, _variant_t("c_Env_Temp_Index"), TempStr);//环境温度对照 
		nIndex = (short)_tcstol(TempStr, NULL, 10);
		if(	nIndex >=0 && nIndex < m_ctlEnvTemp.GetCount() )
		{	
			m_ctlEnvTemp.GetLBText(nIndex, TempStr);
			m_ctlEnvTemp.SelectString( 0, TempStr );
			//从气象表中获得环境温度取值索引对应的环境温度值
			//	GetConditionTemp(m_dTa,nTmp,nIndex);		
		}
		else
		{
			m_ctlEnvTemp.SetWindowText("");
		}
		//放热系数取值.
		TempStr = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_Alfa_Index")));
		nIndex = (short)_tcstol(TempStr, NULL, 10);
		if( nIndex >=0 && nIndex < this->m_ctlHeatTran.GetCount() )
		{
			m_ctlHeatTran.GetLBText(nIndex, TempStr);
			m_ctlHeatTran.SelectString( 0, TempStr ); // 如果选择的字符串不在列表中，则保持当前的选择
		}
		else
		{
			m_ctlHeatTran.SetWindowText("");
		}
		//计算方法
		TempStr = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_CalcMethod_Index")));
		nIndex = (short)_tcstol(TempStr, NULL, 10);
		if( nIndex >=0 && nIndex < this->m_ctlCalcMethod.GetCount() )
		{
			//
			m_ctlCalcMethod.GetLBText(nIndex, TempStr);
			
			m_ctlCalcMethod.SelectString(-1,TempStr);
			
			//为外表面温度计算方法时
			if (nSurfaceTemperatureMethod == nIndex)
			{
				flg = TRUE;
			}
			else
			{
				flg = FALSE;
			}
			((CButton*)GetDlgItem(IDC_CHECK_TS))->SetCheck(flg);
			((CButton*)GetDlgItem(IDC_EDIT_TS))->EnableWindow(flg);
		}
		else
		{
			m_ctlCalcMethod.SetWindowText("");
		}

		//散热量只留二位小数
		m_dHeatLoss.Format("%.2f",vtof(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_lost"))) );
		//热流密度
		m_dHeatDensity.Format("%.2f", vtof(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_HeatFlowrate"))) );
		//内保温厚
		TempStr.Format("%.2f",vtof(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_in_thi"))));
		m_dThick1 =  strtod(TempStr,NULL);
		//外保温厚
		TempStr.Format("%.2f",vtof( m_ICurrentProjectRecordset->GetCollect(_variant_t("c_pre_thi")) ));
		m_dThick2 = strtod(TempStr,NULL);
		//环境温度
		TempStr.Format("%.2f",vtof( m_ICurrentProjectRecordset->GetCollect(_variant_t("c_con_temp"))) );
		m_dTa	  = strtod(TempStr,NULL);
		//外表面温度,只留二位小数
		TempStr.Format("%.2f",vtof( m_ICurrentProjectRecordset->GetCollect(_variant_t("c_tb3"))));
		m_dTs	 = strtod(TempStr,NULL);
		//管内压力
		m_dPressure = vtof( m_ICurrentProjectRecordset->GetCollect(_variant_t("c_Pressure")) );
		//管内介质
		TempStr	= vtos( m_ICurrentProjectRecordset->GetCollect(_variant_t("c_Medium")) );
		if (TempStr.IsEmpty())
		{
			TempStr = _T("主蒸汽");
		}
		m_ctlPipeMedium.SelectString(0, TempStr);
		
		//是否手动输入内保温层的厚度
		m_bIsCalInThick =  vtob( m_ICurrentProjectRecordset->GetCollect(_variant_t("C_CalInThi")));
		//是否手动输入外保温层的厚度
		m_bIsCalPreThick = vtob( m_ICurrentProjectRecordset->GetCollect(_variant_t("c_CalPreThi")));
		//沿风速方向的平壁宽度
		m_dWindPlaimWidth = vtof(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_B")));

		ShowWindowRect();
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
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
//调用之前先UpdateData(TRUE)对话框.
BOOL CEditOriginalData::PutDataToDatabaseFromControl(long nID)
{
	CString TempStr;	//
	short	nIndex;		//临时的索引值
	short	nMethodIndex;// 保温计算方法的索引
//	UpdateData(TRUE);

	try
	{
		if(m_ICurrentProjectRecordset->adoEOF || m_ICurrentProjectRecordset->BOF)
			return FALSE;

		//卷册号
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_vol"),_variant_t(m_VolNo));

		//色环
		m_ColorRing.GetWindowText(TempStr);
		TempStr=TempStr.Left(4);
		TempStr.TrimLeft();
		TempStr.Format("%d",_tcstol(TempStr,NULL,10));
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_color"),_variant_t(TempStr));
		
		//环境温度对照
		m_ctlEnvTemp.GetWindowText(TempStr);
		nIndex = m_ctlEnvTemp.FindString(-1, TempStr);
		if( nIndex != -1 )
		{
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_Env_Temp_Index"), _variant_t((long)nIndex));
		}
		//将数据表中的相应数据写入.
	//	m_ICurrentProjectRecordset->PutCollect(_variant_t("c_con_temp"), _variant_t());
		
		//放热系数的取值索引
		m_ctlHeatTran.GetWindowText(TempStr);
		nIndex = m_ctlHeatTran.FindString(-1,TempStr);
		if(nIndex != -1)
		{
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_Alfa_Index"),_variant_t((long)nIndex));
		}
		//计算方法的索引
		m_ctlCalcMethod.GetWindowText(TempStr);
		nMethodIndex = m_ctlCalcMethod.FindString(-1,TempStr);
		if( nMethodIndex == -1 )
		{
			//当没有选择计算方法时,默认为经济厚度法.
			nMethodIndex = 0 ;
		}
		//保存计算方法的索引
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_CalcMethod_Index"),_variant_t((long)nMethodIndex));

		//外表面温度
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_tb3"),_variant_t(m_dTs));
		//环境温度
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_con_temp"),_variant_t(m_dTa));

		//内保温厚
		if (((CButton*)GetDlgItem(IDC_CHECK_THICK1))->GetCheck())
		{
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_in_thi"),_variant_t(m_dThick1));
		}
		//外保温厚
		if (((CButton*)GetDlgItem(IDC_CHECK_THICK2))->GetCheck())
		{
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_pre_thi"),_variant_t(m_dThick2));
		}
		//管内压力
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_Pressure"),_variant_t(m_dPressure));

		//管内介质
		m_ctlPipeMedium.GetWindowText(TempStr);
		if ( -1 != m_ctlPipeMedium.FindString(-1, TempStr) )
		{
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_Medium"),_variant_t(TempStr));
		}
		//设备或管道名称
		m_PipeDeviceName.GetWindowText(TempStr);
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_name1"),_variant_t(TempStr));
		//外径
		m_PipeSize.GetWindowText(TempStr);
		TempStr.TrimLeft();
		if(!TempStr.IsEmpty())
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_size"),_variant_t(TempStr));
		//厚度
		m_PipeThick.GetWindowText(TempStr);
		TempStr.TrimLeft();
		if(!TempStr.IsEmpty())
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_pi_thi"),_variant_t(TempStr));
		//介质温度
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_temp"),_variant_t(m_MediumTemperature));
		//安装地点
		m_BuidInPlace.GetWindowText(TempStr);
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_place"),_variant_t(TempStr));
		//管道材质
		m_PipeMat.GetWindowText(TempStr);
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_pi_mat"),_variant_t(TempStr));
		//备注
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_mark"),_variant_t(m_ReMark));

		//内保温材料
		m_InsideMatName.GetWindowText(TempStr);
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_name_in"),_variant_t(TempStr));

		//是否手动输入内保温层的厚度.
		if ( TempStr.IsEmpty() )
		{
			//材料名为空，没有指定的厚度.
			m_bIsCalInThick = 0;
		}
		m_ICurrentProjectRecordset->PutCollect(_variant_t("C_CalInThi"),_variant_t(Btob(m_bIsCalInThick)));
		
		//外保温材料
		m_OutSizeMatName.GetWindowText(TempStr);
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_name2"),_variant_t(TempStr));

		//是否手动输入外保温层的厚度
		if ( TempStr.IsEmpty() )
		{
			m_bIsCalPreThick = 0;
		}//
		m_ICurrentProjectRecordset->PutCollect(_variant_t("C_CalPreThi"),_variant_t(Btob(m_bIsCalPreThick)));

		//保护材料
		m_SafeguardMatName.GetWindowText(TempStr);
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_name3"),_variant_t(TempStr)); 
		//风速 
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_wind"),_variant_t(m_WideSpeed)); 
		//热价比主汽价 
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_price"),_variant_t(m_PriceRatio)); 
		//年运行小时数
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_hour"),_variant_t(m_RunPerHour)); 
		//数量 
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_amount"),_variant_t(m_Amount));
		//沿风速方向的平壁宽度
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_B"),_variant_t(m_dWindPlaimWidth));

		m_ICurrentProjectRecordset->Update();
	}
	catch(_com_error &e)
	{
		m_ICurrentProjectRecordset->CancelUpdate();
		ReportExceptionError(e);
		return FALSE;
	}
	switch( nMethodIndex )	// 根据方法索引，有些方法需要另外增加一些参
	{	
	case nPreventCongealMethod:	
		PutDataToPreventCongealDB(nID);// 将防冻对话框中的数据写入到数据库中 [2005/06/01] 
		break;
	case nSubterraneanMethod:	// 将埋地管道的保温数据写入到数据库中	 [2006/02/20]
		PutDataToSubterraneanDB(nID);
		break;
	}

	//增加新的卷册名
	UpdateVolTable();

	return TRUE;
}

//////////////////////////////////////////////
//
// 响应“第一条记录”按扭
//
void CEditOriginalData::OnForefront() 
{
	UpdateData();
//先滚动再选材
/*	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}

*/

	try
	{
		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
			if( !Refresh() )
				return;
		}

		if(m_ICurrentProjectRecordset->BOF)
			return;

		m_ICurrentProjectRecordset->MoveFirst();

		PutDataToEveryControl();
		UpdateControlsState();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
	}
	//2005.2.24
	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}

	UpdateWindow();

}

/////////////////////////////////////////
//
// 响应“前一条记录”按扭
//
void CEditOriginalData::OnPrevious() 
{
	UpdateData();
//2005.2.24
/*
	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}

*/

	try
	{
		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
			if( !Refresh() )
				return;
		}
		if(m_ICurrentProjectRecordset->BOF)
			return;

		m_ICurrentProjectRecordset->MovePrevious();

		if(m_ICurrentProjectRecordset->BOF)
			return;

		PutDataToEveryControl();
		UpdateControlsState();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
	}
//2005.2.24
	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}

	UpdateWindow();
}

//////////////////////////////////////////////
//
// 响应“后一条记录”按扭
//
void CEditOriginalData::OnSubsequent() 
{
	UpdateData();
//2005.2.24
/*	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}

*/

	try
	{
		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
			if( !Refresh() )
				return;
		}
		if(m_ICurrentProjectRecordset->adoEOF)
			return;
		//选项中的移动记录时,自动增加阀门
		if (bIsAutoAddValve)
		{
			CString	strSize;	//外径
			CString strType;	//管道的类型
			
			strType = m_VolNo.Right(1);
			
			if (!strType.CompareNoCase("O") || !strType.CompareNoCase("M") || !strType.CompareNoCase("S"))
			{
				m_PipeSize.GetWindowText(strSize);
				if (m_Amount > 10 && strtod(strSize,NULL) < m_dMaxD0)
				{
					if (AfxMessageBox("是否自动增加阀门!",MB_YESNO) == IDYES)
					{
						OnButtonAddValve();
						
						return;
					}
				}
			}
		}
		//
		m_ICurrentProjectRecordset->MoveNext(); 
		if(m_ICurrentProjectRecordset->adoEOF)
			return;

		PutDataToEveryControl();
		UpdateControlsState();		
	} 
	catch(_com_error &e)
	{
		ReportExceptionError(e);
	}	
//2005.2.24 
	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}
 
	UpdateWindow();

}

////////////////////////////////////////////////
// 
// 响应“最后一条记录”按扭
//
void CEditOriginalData::OnLast() 
{
	UpdateData();
//2005.2.24
/*
	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}
*/

	try
	{
		if(m_IsUpdateByRoll)
		{
			PutDataToDatabaseFromControl();
			if( !Refresh() )
				return;
		}
		if(m_ICurrentProjectRecordset->adoEOF)
			return;

		m_ICurrentProjectRecordset->MoveLast();

		if(m_ICurrentProjectRecordset->adoEOF)
			return;
		
/*		//// // TEST ADD VALVE  [2005/07/07]
		_RecordsetPtr pRs;
		pRs.CreateInstance(__uuidof(Recordset));
		CString strSQL;
		strSQL.Format("SELECT * FROM [PIPE_VALVE] WHERE EnginID='%s' AND ValveID=%d",
					m_ProjectID, vtoi(m_ICurrentProjectRecordset->GetCollect("ID") ));
		pRs->Open(_bstr_t(strSQL), m_ICurrentProjectConnection.GetInterfacePtr(),
					adOpenStatic,adLockOptimistic,adCmdText);
		if (pRs->GetRecordCount() <= 0)
		{
			if(bIsAutoAddValve)
			{
				CString strSize;
				m_PipeSize.GetWindowText(strSize);
				GetDlgItem(IDC_PIPE_SIZE)->GetWindowText(strSize.GetBuffer(50),3);
				
				if (strtod(strSize,NULL) < m_dMaxD0 && (!m_VolNo.Right(1).CompareNoCase("O") || !m_VolNo.Right(1).CompareNoCase("S") || !m_VolNo.Right(1).CompareNoCase("M") )) 
				{
					if( IDYES == AfxMessageBox("是否自动增加阀门!",MB_YESNO) )
					{
						OnButtonAddValve();
						return;
					}
				}
			}
		}

//*/
		PutDataToEveryControl();
		UpdateControlsState();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
	}	
//2005/.2.24
	//自动选材
	if(GetIsAutoSelectMatOnRoll())
	{
		OnAutoSelectmat();
	}

	UpdateWindow();

}

//////////////////////////////////////////
//
// 响应“增加记录”按扭
//
void CEditOriginalData::OnAddNew() 
{
	UpdateData(TRUE);
	long	Num=0;
	int		nTmp;

	if(GetProjectID().IsEmpty())
		return;

	try
	{
		if(!m_ICurrentProjectRecordset->adoEOF || !m_ICurrentProjectRecordset->BOF)
		{
			//自动选材
			if(GetIsAutoSelectMatOnRoll())
			{
				OnAutoSelectmat();
			}

			//
			// 判断是否应该将数据存入数据库
			//
			if(m_IsUpdateByRoll)
			{
				PutDataToDatabaseFromControl();
				if( !Refresh() )
					return;
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
		m_ICurrentProjectRecordset->PutCollect(_variant_t("id"),_variant_t(Num));
		m_ICurrentProjectRecordset->PutCollect(_variant_t("EnginID"),_variant_t(GetProjectID()));
		m_ICurrentProjectRecordset->Update();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	//新建一条记录时，只将记录号自动加1。其他的继承上一条记录的值。
	//只继承卷册号.  2005.3.6
	CString strVolumeID;
	strVolumeID = m_VolNo;

	PutDataToEveryControl();        //该语句不继承。全部赋空值。

	m_VolNo = strVolumeID; //只继承卷册号
	//获得环境温度的取值索引对应的值.
	GetConditionTemp(m_dTa,nTmp);

	UpdateData(FALSE);

	UpdateControlsState();
}
////////////////////////////////////////////////
//
// 响应“删除当前记录”按扭
//
void CEditOriginalData::OnDelCurrent() 
{
	long	CurPos;
	CString	strSQL;

	if(AfxMessageBox(_T("是否删除当前记录"),MB_YESNO) == IDNO)
	{
		return;
	}

	try
	{
		if(!m_ICurrentProjectRecordset->adoEOF || !m_ICurrentProjectRecordset->BOF)
		{
			CurPos=RecNo(m_ICurrentProjectRecordset);

			m_ICurrentProjectRecordset->Delete(adAffectCurrent);
			m_ICurrentProjectRecordset->Update();

			// 删除防冻的记录
			if (!m_pRsCongealData->adoEOF || !m_pRsCongealData->BOF)
			{
				m_pRsCongealData->MoveFirst();
				strSQL.Format("ID=%d", CurPos);
				m_pRsCongealData->Find( _bstr_t(strSQL), NULL, adSearchForward );
				if (!m_pRsCongealData->GetadoEOF())
				{
					m_pRsCongealData->Delete(adAffectCurrent);
					m_pRsCongealData->Update();
				}
				NumberSubtractOne(m_pRsCongealData, CurPos);
			}
			// 删除埋地保温的记录
			if (!m_pRsSubterranean->adoEOF || !m_pRsSubterranean->BOF)
			{
				m_pRsSubterranean->MoveFirst();
				strSQL.Format("ID=%d", CurPos);
				m_pRsSubterranean->Find( _bstr_t(strSQL), NULL, adSearchForward);
				if (!m_pRsSubterranean->GetadoEOF())
				{
					m_pRsSubterranean->Delete(adAffectCurrent);
					m_pRsSubterranean->Update();
				}
				NumberSubtractOne( m_pRsSubterranean, CurPos );				
			}

			m_ICurrentProjectRecordset->MoveFirst();
			if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
			{
				UpdateControlsState();
				return;
			}
			//记录集重新编号(从1开始),并定位到相应的CurPos位置
			RenewNumberFind(m_ICurrentProjectRecordset,1,CurPos);
			//  [2005/06/27]
			//删除某条管道时,在管道-阀门映射表中也要删除对应的记录.
			strSQL.Format("DELETE * FROM [PIPE_VALVE] WHERE (PipeID=%d OR ValveID=%d) AND EnginID='%s' ",
						CurPos,CurPos,m_ProjectID);
			m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
			//改变序号
			strSQL.Format("UPDATE [PIPE_VALVE] SET PipeID=PipeID-1 WHERE PipeID>%d AND EnginID='%s' ",
							CurPos, m_ProjectID);
			m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
			//
			strSQL.Format("UPDATE [PIPE_VALVE] SET ValveID=ValveID-1 WHERE ValveID>%d AND EnginID='%s' ",
							CurPos, m_ProjectID);
			m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

			PutDataToEveryControl();
			UpdateControlsState();
		}
	}
	catch(_com_error &e)
	{ 
		ReportExceptionError(e);
		return;
	}

}

//////////////////////////////////////////////////
//
// 响应“保存当前”按扭
//
void CEditOriginalData::OnSaveCurrent() 
{
	BOOL bRet;
	try
	{
		UpdateData();
		bRet=PutDataToDatabaseFromControl();
		if(bRet)
		{
			if( !Refresh() )
				return;
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
	}
}

/////////////////////////////////////////////////
//
// 更新“色环”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::UpdateColorComBox()
{
	HRESULT hr;
	CString SQL,TempStr1,TempStr2,TempStr,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Color;

	m_ColorRing.GetWindowText(TextStr);

	do
	{
		iRet=m_ColorRing.DeleteString(0);
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

//	SQL.Format(_T("SELECT * FROM a_color"));
	SQL.Format(_T("SELECT * FROM a_color WHERE EnginID='%s'"),GetProjectID());

	try
	{
//		IRecordset_Color->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							   adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Color->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							   adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Color->adoEOF)
		{
			GetTbValue(IRecordset_Color,_variant_t("colr_code"),TempStr1);
			GetTbValue(IRecordset_Color,_variant_t("colr_media"),TempStr2);

			if(!TempStr1.IsEmpty() && !TempStr2.IsEmpty())
			{
				TempStr.Format("%-4s %s",TempStr1,TempStr2);
				m_ColorRing.AddString(TempStr);
			}

			IRecordset_Color->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_ColorRing.SetWindowText(TextStr);
	return TRUE;
}

///////////////////////////////////////////////////////
//
// 更新“管道设备名称”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::UpdatePipeDeviceNameComBox()
{
	HRESULT hr;
	CString SQL,TempStr,TextStr;
	_RecordsetPtr IRecordset_vol;

	m_PipeDeviceName.GetWindowText(TextStr);
	UpdateData();

/*
	int iRet;
	do
	{
		iRet=m_PipeDeviceName.DeleteString(0);
	}while(iRet!=CB_ERR);
*/
	m_PipeDeviceName.ResetContent();
	
	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}
	hr=IRecordset_vol.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	SQL.Format(_T("SELECT DISTINCT [VOL_NAME] FROM [a_vol] WHERE LEFT('%s',5)=LEFT(vol,5) OR VOL_BAK='UESoft'"),m_VolNo);

	try
	{
		IRecordset_vol->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
							 adOpenDynamic,adLockOptimistic,adCmdText);

		if(IRecordset_vol->adoEOF && IRecordset_vol->BOF)
		{
//			AfxMessageBox(_T("库中无此卷册号!"));
			return TRUE;
		}

		while(!IRecordset_vol->adoEOF)
		{
			GetTbValue(IRecordset_vol,_variant_t("vol_name"),TempStr); 

			if(!TempStr.IsEmpty())
				m_PipeDeviceName.AddString(TempStr);

			IRecordset_vol->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}
	
	m_PipeDeviceName.SetWindowText(TextStr);
	return TRUE;
}

///////////////////////////////////////////////////
//
// 更新“管道外径规格”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::UpdatePipeSizeComBox()
{
	HRESULT hr;
	CString SQL,TempStr,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Pipe;

	m_PipeSize.GetWindowText(TextStr);

	do
	{
		iRet=m_PipeSize.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Pipe.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	SQL.Format(_T("SELECT DISTINCT(PIPE_DW) FROM a_pipe ORDER BY PIPE_DW"));

	try
	{
		IRecordset_Pipe->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
							  adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Pipe->adoEOF)
		{
			GetTbValue(IRecordset_Pipe,_variant_t("PIPE_DW"),TempStr);

			if(!TempStr.IsEmpty())
				m_PipeSize.AddString(TempStr);

			IRecordset_Pipe->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_PipeSize.SetWindowText(TextStr);
	return TRUE;
}

////////////////////////////////////////////////////
//
// 更新“管道壁厚”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::UpdatePipeThickComBox()
{
	HRESULT hr;
	CString SQL,TempStr,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Pipe;

	m_PipeThick.GetWindowText(TextStr);
	UpdateData();

	do
	{
		iRet=m_PipeThick.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Pipe.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	m_PipeSize.GetWindowText(TempStr);
	
	TempStr.TrimLeft();
	if(TempStr.IsEmpty())
		return TRUE;

	SQL.Format(_T("SELECT DISTINCT(PIPE_S) FROM a_pipe WHERE abs(PIPE_DW-%s)<0.001 ORDER BY PIPE_S"),TempStr);

	try
	{
		IRecordset_Pipe->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
							  adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Pipe->adoEOF)
		{
			GetTbValue(IRecordset_Pipe,_variant_t("PIPE_S"),TempStr);

			if(!TempStr.IsEmpty())
				m_PipeThick.AddString(TempStr);

			IRecordset_Pipe->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_PipeThick.SetWindowText(TextStr);
	return TRUE;

}

////////////////////////////////////////////////////
//
// 更新“安装地点”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::UpdateBuildPlaceComBox()
{
	HRESULT hr;
	CString SQL;		//SQL语句
	CString TempStr;	//临时字符串
	CString TextStr;	//前一次选择的安装地点
	short	nIndex;		//安装地点对应的索引号
	short	pos;		//
	_RecordsetPtr IRecordset_Loc;

	m_BuidInPlace.GetWindowText(TextStr);

	//重新排列
	m_BuidInPlace.ResetContent();		
	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Loc.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	SQL.Format(_T("SELECT * FROM [安装地点表] "));

	try
	{
		IRecordset_Loc->Open(_variant_t(SQL),GetPublicConnect().GetInterfacePtr(),
							 adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Loc->adoEOF)
		{
			TempStr = vtos(IRecordset_Loc->GetCollect(_variant_t("LOC_NAME")));
			if(!TempStr.IsEmpty())
			{
				pos = m_BuidInPlace.AddString(TempStr);
				
				//地点对应的索引号
				nIndex = vtoi(IRecordset_Loc->GetCollect(_variant_t("SiteIndex")));
				m_BuidInPlace.SetItemData(pos, nIndex);				
			}
			IRecordset_Loc->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}
	
	if( -1 == m_BuidInPlace.SelectString(-1, TextStr) )
	{
		m_BuidInPlace.SetCurSel(0);
	}
	return TRUE;
}


////////////////////////////////////////////////////////
//
// 更新“内保温层材料名称”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::UpdateInsideMatNameCombox()
{
	HRESULT hr;
	CString SQL,TempStr,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Mat;

	m_InsideMatName.GetWindowText(TextStr);

	do
	{
		iRet=m_InsideMatName.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Mat.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;		
	}

//	SQL.Format(_T("SELECT * FROM a_mat WHERE mat_code='PRE'"));
	SQL.Format(_T("SELECT * FROM a_mat WHERE mat_code='PRE' AND EnginID='%s' ORDER BY MAT_NAME"),GetProjectID());

	try
	{
//		IRecordset_Mat->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							 adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Mat->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							 adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Mat->adoEOF)
		{
			GetTbValue(IRecordset_Mat,_variant_t("MAT_NAME"),TempStr);

			if(!TempStr.IsEmpty())
				m_InsideMatName.AddString(TempStr);

			IRecordset_Mat->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_InsideMatName.SetWindowText(TextStr);
	return TRUE;
}

//////////////////////////////////////////////////////
//
// 更新“外保温层材料名称”组合框
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CEditOriginalData::UpdateOutSizeMatNameComBox()
{
	HRESULT hr;
	CString SQL,TempStr,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Pipe;

	m_OutSizeMatName.GetWindowText(TextStr);

	do
	{
		iRet=m_OutSizeMatName.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Pipe.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

//	SQL.Format(_T("SELECT * FROM a_mat WHERE mat_code='PRE'"));
	SQL.Format(_T("SELECT * FROM a_mat WHERE mat_code='PRE' AND EnginID='%s' ORDER BY MAT_NAME"),GetProjectID());

	try
	{
//		IRecordset_Pipe->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							  adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Pipe->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							  adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Pipe->adoEOF)
		{
			GetTbValue(IRecordset_Pipe,_variant_t("mat_name"),TempStr);

			if(!TempStr.IsEmpty())
				m_OutSizeMatName.AddString(TempStr);

			IRecordset_Pipe->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_OutSizeMatName.SetWindowText(TextStr);
	return TRUE;
}

//////////////////////////////////////////////////////////
//
// 更新"保护层材料名称"组合框
//
// 如果函数成功返回TRUE,否则返回FALSE
//
BOOL CEditOriginalData::UpdateSafeguardMatNameComBox()
{
	HRESULT hr;
	CString SQL,TempStr,TextStr;
	int iRet;
	_RecordsetPtr IRecordset_Mat;

	m_SafeguardMatName.GetWindowText(TextStr);

	do
	{
		iRet=m_SafeguardMatName.DeleteString(0);
	}while(iRet!=CB_ERR);

	if(GetPublicConnect()==NULL)
	{
		AfxMessageBox(_T("未连接公共数据库"));
		return FALSE;
	}

	hr=IRecordset_Mat.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

//	SQL.Format(_T("SELECT * FROM a_mat WHERE mat_code='PRO'"));
	SQL.Format(_T("SELECT * FROM a_mat WHERE mat_code='PRO' AND EnginID='%s' ORDER BY MAT_NAME"),GetProjectID());

	try
	{
//		IRecordset_Mat->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							 adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Mat->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							 adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Mat->adoEOF)
		{
			GetTbValue(IRecordset_Mat,_variant_t("mat_name"),TempStr);

			if(!TempStr.IsEmpty())
				m_SafeguardMatName.AddString(TempStr);

			IRecordset_Mat->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	m_SafeguardMatName.SetWindowText(TextStr);
	return TRUE;
}

////////////////////////////////////////////////////////////
//
// 当在a_vol中找不到对话框中的卷册号时,更新a_vol表
//
// 函数成功返回TRUE,否则返回FALSE
//
BOOL CEditOriginalData::UpdateVolTable()
{
	CString SQL,csVol,csPipeDeviceName,csTemp;
	HRESULT hr;
	_RecordsetPtr IRecordset_Vol;

	hr=IRecordset_Vol.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
		return FALSE;

	csVol=m_VolNo;
	m_PipeDeviceName.GetWindowText(csPipeDeviceName);

	csVol.TrimLeft();
	csVol.TrimRight();
	csPipeDeviceName.TrimLeft();
	csPipeDeviceName.TrimRight();

	csVol=csVol.Left(5);

	if(csVol.IsEmpty() || csPipeDeviceName.IsEmpty())
		return TRUE;

	SQL.Format(_T("SELECT * FROM a_vol WHERE LEFT(vol,5)='%s' AND vol_name='%s'"),csVol,csPipeDeviceName);


	try
	{
		IRecordset_Vol->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
							adOpenDynamic,adLockOptimistic,adCmdText);

		if(IRecordset_Vol->adoEOF && IRecordset_Vol->BOF)
		{
			//此处，将新的卷册增加到表A_VOL中。输出到EXCEL时又从Volume中取。
			IRecordset_Vol->AddNew();
			
			IRecordset_Vol->PutCollect(_variant_t("vol"),_variant_t(csVol));
			IRecordset_Vol->PutCollect(_variant_t("vol_name"),_variant_t(csPipeDeviceName));
			IRecordset_Vol->PutCollect(_variant_t("vol_price"),_variant_t(m_PriceRatio));

			this->m_ColorRing.GetWindowText(csTemp);
			csTemp.Format("%d",_tcstol(csTemp, NULL, 10 ));
			IRecordset_Vol->PutCollect(_variant_t("vol_colr"),_variant_t(csTemp));

			IRecordset_Vol->Update();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}

void CEditOriginalData::OnDropdownColor() 
{
	UpdateColorComBox();	
}

void CEditOriginalData::OnDropdownPipeDeviceName() 
{
	UpdatePipeDeviceNameComBox();	
}

void CEditOriginalData::OnDropdownPipeSize() 
{
	UpdatePipeSizeComBox();	
}

void CEditOriginalData::OnDropdownPipeThick() 
{
	UpdatePipeThickComBox();	
}

void CEditOriginalData::OnDropdownInsideMatname() 
{
	UpdateInsideMatNameCombox();	
}

void CEditOriginalData::OnDropdownOutsizeMatname() 
{
	UpdateOutSizeMatNameComBox();	
}

void CEditOriginalData::OnDropdownSafeguardMatname() 
{
	UpdateSafeguardMatNameComBox();	
}

void CEditOriginalData::OnItemchangedHeatPreservationType(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString TempStr;

	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	
	if(pNMListView->iItem==-1)
	{
		*pResult=0;
		return;
	}

	UpdateData();

	m_VolNo.TrimLeft();
	if(m_VolNo.IsEmpty())
	{
		*pResult=0;
		return;
	}

	while(m_VolNo.GetLength()/sizeof(TCHAR)<5)
	{
		m_VolNo+=_T(" ");
	}

	m_VolNo=m_VolNo.Left(5);
	TempStr=m_HeatPreservationTypeList.GetItemText(pNMListView->iItem,0);
	m_VolNo+=TempStr;

	UpdateData(FALSE);
	UpdateControlsState();
	
	*pResult = 0;
}

void CEditOriginalData::OnItemchangedVollist(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

	if(pNMListView->iItem==-1)
	{
		*pResult=0;
		return;
	}
	CString	strDeviceName;		//管道/设备名称
	CString	strPriceRatio;		// 热价比主汽价

	m_VolNo=m_VolList.GetItemText(pNMListView->iItem,0);
	
	strDeviceName = m_VolList.GetItemText(pNMListView->iItem,1);
	if ( !strDeviceName.IsEmpty() )
	{
		m_PipeDeviceName.SetWindowText(strDeviceName);
	}
	if ( 1 == gbIsSelTblPrice )
	{
		strPriceRatio = m_VolList.GetItemText(pNMListView->iItem, 2);
		m_PriceRatio = strtod( strPriceRatio, NULL );
	}
	UpdateData(FALSE);
	UpdateControlsState();

	*pResult = 0;
}

void CEditOriginalData::OnUpdateAllCheckBox() 
{
	UpdateControlsState();
}

void CEditOriginalData::OnChangeVolno() 
{
	// TODO: If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.
	
	UpdateControlsState();	
}
//响应记录号的直接输出入
void CEditOriginalData::OnChangeId() 
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
	//2005.2.24
	//自动选择材料。
	if( this->GetIsAutoSelectMatOnRoll() )
	{
		this->OnAutoSelectmat();
	}
}

//////////////////////////////////////////////
//
// 更新各个控件的状态
//
void CEditOriginalData::UpdateControlsState()
{
	CString TempStr;
	CString strSize;
	CString	strSQL;

	UINT uControlIDs[]=
	{
		IDC_ID,
		IDC_FOREFRONT,
		IDC_PREVIOUS,
		IDC_SUBSEQUENT,
		IDC_LAST,
		IDC_DELCURRENT,
		IDC_DEL_ALL_INSU,
		IDC_SAVECURRENT,
		IDC_UPDATABYROLL,
		IDC_VOLNO,
		IDC_COLOR,
		IDC_PIPE_DEVICE_NAME,
		IDC_PIPE_SIZE,
		IDC_PIPE_THICK,
		IDC_MEDIUM_TEMPERATURE,
		IDC_SAFEGUARD_MATNAME,
		IDC_BUILDIN_PLACE,
		IDC_PIPE_MAT,
		IDC_REMARK,
		IDC_VOLLIST_ISENABLE,
		IDC_HEAT_PRESERVATION_TYPE_ENABLE,
		IDC_VOLLIST,
		IDC_HEAT_PRESERVATION_TYPE,
		IDC_WIDE_SPEED,
		IDC_PRICE,
		IDC_RUN_PERHOUR,
		IDC_AUTO_SELECTMAT,
		IDC_CHECK_AUTOSELECT_PRE_EDIT,
		IDC_SORT_BY_VOL,
		IDC_AMOUNT,						// 数量
		IDC_ENVIRON_TEMP,
		IDC_CHECK_TS,
		IDC_EDIT_TS,
		IDC_CHECK_TA,
		IDC_EDIT_TA,
		IDC_CHECK_METHOD,
		IDC_EDIT_METHOD,
		IDC_HEAT_TRANSFER_COEF,
		IDC_CHECK_ENV_TEMP,
		IDC_CHECK_ALPHA,
		IDC_PIPE_PRESSURE,
		IDC_PIPE_MEDIUM,
		IDC_BUTTON_ADD_VALVE,
		IDC_EDIT_PLAIM_THICK,
		IDC_BUTTON_CALC_RATIO
	};

	if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
	{
		//没有原始数据
		int i=0;
		for(i=0;i<sizeof(uControlIDs)/sizeof(UINT);i++)
		{
			GetDlgItem(uControlIDs[i])->EnableWindow(FALSE);
		}
		m_ctlCalcMethod.SetCurSel(-1);
		ShowWindowRect();
		UINT controlID[] = // 需要特殊处理的一些控件(埋地计算时）
		{
			IDC_CHECK_THICK1,
			IDC_EDIT_THICK1,
			IDC_CHECK_THICK2,
			IDC_EDIT_THICK2,
			IDC_INSIDE_MATNAME,
			IDC_OUTSIZE_MATNAME,
		};
		for ( i = 0 ; i < sizeof(controlID) / sizeof(UINT) ; i ++)
		{
			GetDlgItem(controlID[i])->EnableWindow(FALSE);
		}

		return ;
	}
	else
	{
		for(int i=0;i<sizeof(uControlIDs)/sizeof(UINT);i++)
		{
			GetDlgItem(uControlIDs[i])->EnableWindow(TRUE);
		}
	}

	UpdateData();

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

	//请选择卷册：
	if(m_IsVolListEnable)
	{
		if( GetDlgItem(IDC_VOLLIST_ISENABLE)->IsWindowEnabled() )
			m_VolList.EnableWindow(TRUE);
		else
			m_VolList.EnableWindow(FALSE);
	}
	else
	{
		m_VolList.EnableWindow(FALSE);
	}
	//请选择保温数据类型：
	if(m_IsHeatPreservationTypeEnable)
	{
		if( GetDlgItem(IDC_HEAT_PRESERVATION_TYPE_ENABLE)->IsWindowEnabled() )
			m_HeatPreservationTypeList.EnableWindow(TRUE);
		else
			m_HeatPreservationTypeList.EnableWindow(FALSE);
	}
	else
	{
		m_HeatPreservationTypeList.EnableWindow(FALSE);
	}
	//选择计算方法.
	GetDlgItem(IDC_EDIT_METHOD)->EnableWindow(((CButton*)GetDlgItem(IDC_CHECK_METHOD))->GetCheck()?TRUE:FALSE);
	
	//外表面温度.
	GetDlgItem(IDC_EDIT_TS)->EnableWindow(((CButton*)GetDlgItem(IDC_CHECK_TS))->GetCheck()?TRUE:FALSE);
	//环境温度.
	GetDlgItem(IDC_EDIT_TA)->EnableWindow(((CButton*)GetDlgItem(IDC_CHECK_TA))->GetCheck()?TRUE:FALSE);

	//只有当管道类型为油管道，汽水管道，主蒸汽管道时才能在其上增加阀门
	TempStr = m_VolNo;
	TempStr.TrimRight();
	TempStr = TempStr.Right(1);

	m_PipeSize.GetWindowText(strSize);
	
	if (strtod(strSize,NULL) < m_dMaxD0 ) // 管道记录
	{
		if (!TempStr.CompareNoCase("O") || !TempStr.CompareNoCase("M") || !TempStr.CompareNoCase("S"))
		{
			GetDlgItem(IDC_BUTTON_ADD_VALVE)->EnableWindow(TRUE);
		}else
		{
			GetDlgItem(IDC_BUTTON_ADD_VALVE)->EnableWindow(FALSE);
		}

		GetDlgItem(IDC_EDIT_PLAIM_THICK)->EnableWindow(FALSE);
	}else	// 平壁
	{
		GetDlgItem(IDC_EDIT_PLAIM_THICK)->EnableWindow(TRUE);
		GetDlgItem(IDC_BUTTON_ADD_VALVE)->EnableWindow(FALSE);
	}
	try
	{
		_RecordsetPtr pRs;
		pRs.CreateInstance(__uuidof(Recordset));
		strSQL.Format("SELECT * FROM [Pipe_Valve] WHERE ValveID=%d AND EnginID='%s' ",
				m_ID, m_ProjectID);	
		pRs->Open(_variant_t(strSQL), m_ICurrentProjectConnection.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		if (pRs->GetRecordCount() > 0)
		{
			m_PipeSize.EnableWindow(FALSE);
		}
		//选项中选中了“选择表中的热价比主汽价”时可用，当介质为,水和蒸汽 时按公式计算热价比主汽价
		if (  1 == gbIsSelTblPrice )
		{
			GetDlgItem( IDC_BUTTON_CALC_RATIO )->EnableWindow( TRUE );
		}
		else
		{
			CString strMediumName;
			CString strTmp;
			GetDlgItem( IDC_BUTTON_CALC_RATIO )->EnableWindow( FALSE );
			int pos = m_ctlPipeMedium.GetCurSel();
			if ( -1 != pos )
			{
				m_ctlPipeMedium.GetLBText( pos, strMediumName);
			}else
			{
				m_ctlPipeMedium.GetWindowText( strMediumName );
			}
			if ( !strMediumName.IsEmpty() )
			{
				CHeatPreCal::GetMediumDensity( strMediumName, &strTmp );
				if ( !strTmp.CompareNoCase( "水" ) || !strTmp.CompareNoCase("蒸汽") )
				{
					GetDlgItem( IDC_BUTTON_CALC_RATIO )->EnableWindow( TRUE );
				}
			}		
		}
	}
	catch (_com_error)
	{
	}

	m_ctlCalcMethod.GetWindowText(TempStr);
	if (m_ctlCalcMethod.FindString(-1, TempStr) != nSubterraneanMethod)
	{
		UpdateControlSubterranean( -1 );
//		//内保温厚度
//		GetDlgItem(IDC_EDIT_THICK1)->EnableWindow(((CButton*)GetDlgItem(IDC_CHECK_THICK1))->GetCheck()?TRUE:FALSE);
//		//外保温厚度
//		GetDlgItem(IDC_EDIT_THICK2)->EnableWindow(((CButton*)GetDlgItem(IDC_CHECK_THICK2))->GetCheck()?TRUE:FALSE);
	}
	
}

//功能；
//以原始数据中的一些字段的值做为参考，更新其它的字段值
BOOL CEditOriginalData::Refresh()
{
	double	e_rate = 0;
	double	con_temp1 = 0;
	double	con_temp2 = 0;
	double	e_hours = 0;
	double	e_wind = 0;

	double	Pro_thi;				//保护层厚度
//	double	dConTemp;				//环境温度。

	_RecordsetPtr IRecord_Config;
	_variant_t TempVar;
	CString SQL,TempStr;
	HRESULT hr;

	hr=IRecord_Config.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return FALSE;
	}

	SQL.Format(_T("SELECT * FROM a_config WHERE EnginID='%s'"),GetProjectID());

	if(m_ICurrentProjectRecordset->adoEOF || m_ICurrentProjectRecordset->BOF)
	{
		return TRUE;
	}

	try
	{
		IRecord_Config->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							 adOpenStatic,adLockOptimistic,adCmdText);

		if ( IRecord_Config->GetRecordCount() > 0 )
		{
			GetTbValue(IRecord_Config,_variant_t("年费用率"),e_rate);
			GetTbValue(IRecord_Config,_variant_t("室内温度"),con_temp1);
			GetTbValue(IRecord_Config,_variant_t("室外温度"),con_temp2);
			GetTbValue(IRecord_Config,_variant_t("年运行小时"),e_hours);
			GetTbValue(IRecord_Config,_variant_t("室外风速"),e_wind);
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}
 
	try
	{
		//////////应该在新建工程中////////////////////////
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_rate"),_variant_t(e_rate));
		m_ICurrentProjectRecordset->Update();
		////////////////////////////////////////////
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_vol"),TempStr);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}


	/////////////////////////有可能不需要///////////////////
	TempStr.TrimLeft();
	if(TempStr.IsEmpty())
	{
		GetDlgItem(IDC_VOLNO)->SetFocus();
	}
	////////////////////////////////////////////////////
/*		取消选择油管道时，可以输入厚度 //  [2005/06/02] 
	if(TempStr.Right(1)==_T("O"))
	{
		GetDlgItem(IDC_HEAT_PRESERVATION_THICK)->EnableWindow(TRUE);
	}
	else
	{
		GetDlgItem(IDC_HEAT_PRESERVATION_THICK)->EnableWindow(FALSE);
		GetDlgItem(IDC_HEAT_PRESERVATION_THICK)->SetWindowText(_T("0"));
		try
		{
			TempVar.Clear();
			TempVar.vt=VT_NULL;
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_pre_thi"),TempVar);
			m_ICurrentProjectRecordset->Update();
		}
		catch(_com_error &e)
		{
			AfxMessageBox(e.Description());
			return FALSE;
		}
	}
*/
	try
	{
		GetTbValue(m_ICurrentProjectRecordset,_variant_t("c_name3"),TempStr);
		if(!TempStr.IsEmpty())
		{
			if(TempStr.Find(_T("("))==-1)
			{
				Pro_thi=0.5;//默认保护层厚度0.5mm
			}
			else
			{
				TCHAR *StopString;

				TempStr=TempStr.Mid(TempStr.Find(_T("("))+1);
				Pro_thi=static_cast<double>(_tcstod(TempStr,&StopString));
			}

			try
			{
				m_ICurrentProjectRecordset->PutCollect(_variant_t("c_pro_thi"),_variant_t(Pro_thi));
				m_ICurrentProjectRecordset->Update();
			}
			catch(_com_error &e)
			{
				ReportExceptionError(e);
				return FALSE;
			}
		}
/*		//安装地点。
		TempStr = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_place")));
		if(!TempStr.IsEmpty())
		{
			CString place,steam;
			place=TempStr.Left(4);
			steam=TempStr.Right(4);
			try
			{
				m_ICurrentProjectRecordset->PutCollect(_variant_t("c_place"),_variant_t(place));
				m_ICurrentProjectRecordset->PutCollect(_variant_t("c_steam"),_variant_t(steam));
				if(place==_T("室内"))
				{
					m_ICurrentProjectRecordset->PutCollect(_variant_t("c_con_temp"),_variant_t(con_temp1));
				}
				else
				{
					m_ICurrentProjectRecordset->PutCollect(_variant_t("c_con_temp"),_variant_t(con_temp2));
				}

			}
			catch(_com_error &e)
			{
				ReportExceptionError(e);
				return FALSE;
			}
		}
*/		//环境温度取值由用户选择的索引确定.
		//如果没有,则按上面的室内.室外取值.
	//	if( GetConditionTemp(dConTemp, nMethodIndex) )
	//	{
	//		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_con_temp"), _variant_t(dConTemp));
	//	}
		//环境温度值
	//	m_dTa = dConTemp;

		//更新介质的类型
		
		//检查数据的正确性
		if ( !CheckDataValidity() )
		{
			return false;
		}
		

		m_ICurrentProjectRecordset->Update();
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return TRUE;
}

//////////////////////////////////////////////
//
// 响应“自动选择保温材料”
//
void CEditOriginalData::OnAutoSelectmat() 
{
	double e_hours,e_wind;

	_RecordsetPtr IRecord_Config;
	_variant_t TempVar;
	CString SQL;
	CString txtC_vol,txtC_size,txtC_temp,txtC_p_s;
	CString name_pre,name_pro,name_prein;
	HRESULT hr;

	UpdateData();

	hr=IRecord_Config.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return;
	}
	
	SQL.Format(_T("SELECT * FROM a_config WHERE EnginID='%s' "),GetProjectID());

	try
	{
		IRecord_Config->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							 adOpenDynamic,adLockOptimistic,adCmdText);

		if (IRecord_Config->adoEOF && IRecord_Config->BOF)
		{
			AfxMessageBox("保温设计常数库为空!");
		}
		else
		{
			GetTbValue(IRecord_Config,_variant_t("年运行小时"),e_hours);
			GetTbValue(IRecord_Config,_variant_t("室外风速"),e_wind);
			
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_hour"),_variant_t(e_hours));
			// 改变风速时,先判断安装地点.		 [2005/06/23]
/*
			m_BuidInPlace.GetWindowText(txtC_p_s);
			if (-1 != txtC_p_s.Find("室内"))
			{
				//室内的风速为0;
				e_wind = 0;
			}
*/			
			m_ICurrentProjectRecordset->PutCollect(_variant_t("c_wind"),_variant_t(e_wind));
			
			//将对话框中的控件数据进行更新
			m_RunPerHour=e_hours;
			m_WideSpeed=e_wind;
			
			this->UpdateData(FALSE);
			
			m_ICurrentProjectRecordset->Update();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	GetDlgItem(IDC_VOLNO)->GetWindowText(txtC_vol);
	GetDlgItem(IDC_PIPE_SIZE)->GetWindowText(txtC_size);
	GetDlgItem(IDC_MEDIUM_TEMPERATURE)->GetWindowText(txtC_temp);
	GetDlgItem(IDC_BUILDIN_PLACE)->GetWindowText(txtC_p_s);

	txtC_vol.TrimLeft();
	txtC_size.TrimLeft();
	txtC_temp.TrimLeft();
	txtC_p_s.TrimLeft();

	if(txtC_vol.IsEmpty() || txtC_size.IsEmpty() || txtC_temp.IsEmpty() || txtC_p_s.IsEmpty())
	{
		return;
	}	

	try
	{
		AUTO(txtC_vol,txtC_size,txtC_temp,name_pre,name_pro,txtC_p_s,name_prein);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}
/*	//test  计算热价比主汽价
	double P,T,H,S;
	int R;
	P = 1;
	T = 350;
	int stdid=97;
    SETSTD_WASP(stdid);  //设置计算标准 IF97/IFC67  可以在项选对话框中加一个条件
	
	PT2H(P,T,&H,&R);
	PT2S(P,T,&S,&R);
//*/
	GetDlgItem(IDC_INSIDE_MATNAME)->SetWindowText(name_prein);
	GetDlgItem(IDC_OUTSIZE_MATNAME)->SetWindowText(name_pre);
	GetDlgItem(IDC_SAFEGUARD_MATNAME)->SetWindowText(name_pro);

}

void CEditOriginalData::compound(CString &mat_in, CString &mat)
{
	CString SQL;
	_RecordsetPtr IRecordset_Compnd;

	IRecordset_Compnd.CreateInstance(__uuidof(Recordset));

//	SQL.Format(_T("SELECT * FROM a_compnd"));
	SQL.Format(_T("SELECT * FROM a_compnd WHERE EnginID='%s'"),GetProjectID());

	try
	{
//		IRecordset_Compnd->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//								adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_Compnd->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
								adOpenDynamic,adLockOptimistic,adCmdText);

		if (LocateFor(IRecordset_Compnd,_variant_t("complex"),CFoxBase::EQUAL,_variant_t(mat)))
		{
			if (!IRecordset_Compnd->adoEOF)
			{
				mat_in	= vtos( IRecordset_Compnd->GetCollect(_variant_t("comp_in")) );
				mat		= vtos( IRecordset_Compnd->GetCollect(_variant_t("comp_out")) );
			}
		}		
	}
	catch(_com_error)
	{
		throw;
	}
}

void CEditOriginalData::cond(_RecordsetPtr &IRecordset, CString &mat,CString &v,
							 CString &p,CString t,CString d,CString h)
{
//	CString SQL;
	TCHAR *StopString;
	CString a,k,TempStr,con;
	BOOL Bool,bRet;
	CParseExpression Parse;
	int c1,rec,max_k;

	CParseExpression::tagExpressionVar Var[5];

	try
	{
		if (IRecordset->adoEOF && IRecordset->BOF)
		{
			return;
		}
		max_k=IRecordset->GetFields()->GetCount()-2;

		IRecordset->MoveFirst();
		while(!IRecordset->adoEOF)
		{
			this->GetTbValue(IRecordset,_variant_t("auto_obj"),TempStr);

			TempStr.TrimLeft();
			if(!TempStr.IsEmpty())
				break;

			IRecordset->MoveNext();
		}

		if(IRecordset->adoEOF)
			return;

		c1=RecNo(IRecordset);
	}
	catch(_com_error)
	{
		throw;
	}

	Bool=FALSE;

	Var[0].VarName=_T("v");
	Var[0].VarValue=v;
	Var[1].VarName=_T("p");
	Var[1].VarValue=p;
	Var[2].VarName=_T("t");
	Var[2].VarValue=_tcstod(t,&StopString);
	Var[3].VarName=_T("d");
	Var[3].VarValue=_tcstod(d,&StopString);
	Var[4].VarName=_T("h");
	Var[4].VarValue=h;

	Parse.SetConnectionForParse(GetPublicConnect());

	try
	{
		while(!IRecordset->adoEOF)
		{
			GetTbValue(IRecordset,_variant_t("auto_obj"),a);

			bRet=Parse.ParseUseConn_Bool(a,Var,5,NULL);
			if(bRet)
			{
				rec=RecNo(IRecordset);
				Bool=TRUE;
				break;
			}
			IRecordset->MoveNext();
		}

		IRecordset->MoveFirst();
	}
	catch(_com_error)
	{
		throw;
	}
	
	k=_T("1");

//	DO WHILE RECNO()<C1 .AND. VAL(k)<=max_k
	try
	{
		while(RecNo(IRecordset)<c1 && atof(k)<=max_k)
		{
	//		con=auto_con&k
			TempStr=_T("auto_con");
			TempStr+=k;

			GetTbValue(IRecordset,_variant_t(TempStr),con);
			con.TrimLeft();
	//		IF ""=TRIM(con)
			if(con.IsEmpty())
			{
	//			SKIP
				IRecordset->MoveNext();
				continue;
	//			LOOP
			}
	//		ENDIF

			bRet=Parse.ParseUseConn_Bool(con,Var,5,NULL);
	//		IF &con
	//			SKIP
	//		ELSE
	//			k=STR(VAL(k)+1,1)
	//			GO TOP
	//			LOOP
	//		ENDIF

			if(bRet)
			{
				IRecordset->MoveNext();
			}
			else
			{
				TempStr.Format(_T("%d"),_tcstol(k,&StopString,10)+1);
				k=TempStr;
				IRecordset->MoveFirst();
				continue;
			}
		}//	ENDDO

		IRecordset->MoveFirst();
		if(Bool)
		{
			while(TRUE)
			{
				rec--;
				if(rec>0)
					IRecordset->MoveNext();
				else
					break;
			}

			TempStr=_T("auto_con");
			TempStr+=k;
			GetTbValue(IRecordset,_variant_t(TempStr),mat);
		}
	}
	catch(_com_error)
	{
		throw;
	}
}

void CEditOriginalData::AUTO(CString &v, CString &s, CString &t, CString &pre, CString &pro, CString &p, CString &pre_in)
{
	CString d;
	CString mat,mat_in,h;
	CString SQL;
	CString csTemp;
	int max_k;

	_RecordsetPtr IRecordset_pre;

	d=s;
	p=p.Left(4);
	mat=_T("找不到!");

	IRecordset_pre.CreateInstance(__uuidof(Recordset));
//	SQL.Format(_T("SELECT * FROM a_pre"));
	SQL.Format(_T("SELECT * FROM a_pre WHERE EnginID='%s'"),GetProjectID());

	try
	{
//		IRecordset_pre->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							 adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_pre->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							 adOpenDynamic,adLockOptimistic,adCmdText);

		max_k=IRecordset_pre->GetFields()->GetCount()-2;

		cond(IRecordset_pre,mat,v,p,t,d);

		mat_in=pre_in=_T("");

		if(mat.Find(_T("复合型式"))!=-1)
		{
			compound(mat_in,mat);
			pre_in=mat_in;
		}
	}
	catch(_com_error)
	{
		throw;
	}

	pre=mat;

	_RecordsetPtr IRecordset_mat;
	IRecordset_mat.CreateInstance(__uuidof(Recordset));

//	SQL.Format(_T("SELECT * FROM a_mat"));
	SQL.Format(_T("SELECT * FROM a_mat WHERE EnginID='%s'"),GetProjectID());

	try
	{
//		IRecordset_mat->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							adOpenDynamic,adLockOptimistic,adCmdText);
		IRecordset_mat->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							adOpenDynamic,adLockOptimistic,adCmdText);

//		LocateFor(IRecordset_mat,_variant_t("mat_name"),CFoxBase::EQUAL,_variant_t(pre));

		while(!IRecordset_mat->adoEOF)
		{
			GetTbValue(IRecordset_mat,_variant_t("mat_name"),csTemp);

			if(csTemp==pre)
				break;

			IRecordset_mat->MoveNext();
		}

		if(IRecordset_mat->adoEOF)
		{
			h=_T("");
		}
		else
		{
			GetTbValue(IRecordset_mat,_variant_t("mat_h"),h);
		}

		mat=_T("找不到!");

	//	SELE 5
	//	USE a_pro
		_RecordsetPtr IRecordset_pro;
		IRecordset_pro.CreateInstance(__uuidof(Recordset));

//		SQL.Format(_T("SELECT * FROM a_pro"));
//		IRecordset_pro->Open(_variant_t(SQL),_variant_t((IDispatch*)GetPublicConnect()),
//							adOpenDynamic,adLockOptimistic,adCmdText);

		SQL.Format(_T("SELECT * FROM a_pro WHERE EnginID='%s'"),GetProjectID());
		IRecordset_pro->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							adOpenDynamic,adLockOptimistic,adCmdText);


	//	max_k=FCOUNT()-2
		max_k=IRecordset_pro->GetFields()->GetCount()-2;
	//	DO cond WITH mat
		cond(IRecordset_pro,mat,v,p,t,d,h);
	//	pro=mat
		pro=mat;
	}
	catch(_com_error)
	{
		throw;
	}
}

///////////////////////////////////////////
//
// 响应“按卷册号重排”按扭
//
void CEditOriginalData::OnSortByVol() 
{
	CString csSQL;
	long iNumber;

	try
	{
		UpdateData();
		
		if(m_ProjectID.IsEmpty())
		{
			AfxMessageBox(_T("未选择工程"));
			return;
		}

		if(m_ICurrentProjectConnection==NULL)
		{
			AfxMessageBox(_T("未连接数据库"));
			return;
		}

		if(!PutDataToDatabaseFromControl())
			return;

		if(!Refresh())
			return;
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return;
	}

	try
	{
		if(!(m_ICurrentProjectRecordset->GetState()==adStateClosed))
		{
			m_ICurrentProjectRecordset->Close();
		}

		csSQL.Format(_T("SELECT * FROM pre_calc WHERE EnginID='%s' ORDER BY c_vol ASC,c_size DESC"),GetProjectID());

		m_ICurrentProjectRecordset->Open(_variant_t(csSQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
										 adOpenStatic,adLockOptimistic,adCmdText);

		iNumber=1;
		while(!m_ICurrentProjectRecordset->adoEOF)
		{
			m_ICurrentProjectRecordset->PutCollect(_variant_t("id"),_variant_t(iNumber));
			m_ICurrentProjectRecordset->Update();
			iNumber++;
			
			m_ICurrentProjectRecordset->MoveNext();
		}

		if(m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
			return;

		m_ICurrentProjectRecordset->MoveFirst();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return;
	}

	PutDataToEveryControl();
	UpdateControlsState();

}

////////////////////////////////////////////////////
//
// 响应”退出“按扭
//
void CEditOriginalData::OnExit() 
{
	UpdateData();
	try
	{
		if(m_IsUpdateByRoll)
		{
			if ( PutDataToDatabaseFromControl() )
				Refresh();
		}
	}
	catch(_com_error)
	{	
	}
	bIsMoveUpdate = m_IsUpdateByRoll;			//移动更新.
	bIsAutoSelectPre = m_bIsAutoSelectMatOnRoll;//移动选材.

	OnOK();
	//2005.2.24
	EndDialog(0);
	if( IsWindow(this->GetSafeHwnd()) )
		DestroyWindow();
}

/////////////////////////////////////////////////////////////////
//
// 设置在滚动时是否自动选材
//
// bSelect[in]	当为TRUE时将在向前，或向后移动记录时自动选材
//				当为FALSE则反之
//
void CEditOriginalData::SetIsAutoSelectMatOnRoll(BOOL bSelect)
{
	m_bIsAutoSelectMatOnRoll=bSelect;
}

///////////////////////////////////////////////////////////////
//
// 返回是否在滚动时自动选材
// 
// 当返回为TRUE时自动选材，否则不自动选材
//
BOOL CEditOriginalData::GetIsAutoSelectMatOnRoll()
{
	return m_bIsAutoSelectMatOnRoll;
}


void CEditOriginalData::OnCancel()
{
	EndDialog( 0 );
	if( IsWindow( this->GetSafeHwnd()) )
	{
		DestroyWindow();
	}
	CDialog::OnClose();
}
//删除所有的保温原始数据。
void CEditOriginalData::OnDelAllInsu() 
{
	if(	AfxMessageBox(_T("是否删除所有的原始数据!"),MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON2) == IDNO )
	{
		return ;
	}
	try
	{
		CString strSQL;
		//清空原始数据时。
		strSQL = "DELETE * FROM [PRE_CALC] WHERE EnginID='"+GetProjectID()+"' ";
		m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, -1);

		//清空原始数据时,在管道-阀门映射表中也要删除对应的记录		 [2005/06/27]
		strSQL.Format("DELETE * FROM [PIPE_VALVE] WHERE EnginID='%s'",m_ProjectID);
		m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

		// 防冻记录
		strSQL.Format("DELETE * FROM [PRE_CALC_CONGEAL] WHERE EnginID='%s'", m_ProjectID);
		m_ICurrentProjectConnection->Execute( _bstr_t(strSQL), NULL, adCmdText);

		// 埋地数据
		strSQL.Format("DELETE * FROM [PRE_CALC_SUBTERRANEAN] WHERE EnginID='%s'", m_ProjectID);
		m_ICurrentProjectConnection->Execute( _bstr_t(strSQL), NULL, adCmdText);

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
/////////////////////////////////
//
//将所有的保温原始数据自动选择材料。
bool CEditOriginalData::AutoSelAllMat()
{
	//打开原始数据表
	if( !this->InitCurrentProjectRecordset() )
		return false;

	int Count = m_ICurrentProjectRecordset->GetRecordCount();	//记录条数.
	if( Count <= 0 )
	{
		return false;
	}

	double e_hours=0,e_wind=0;
	_RecordsetPtr IRecord_Config;
//	_variant_t TempVar;
	CString SQL, strPass;
	CString txtC_vol,txtC_size,txtC_temp,txtC_p_s;
	CString name_pre,name_pro,name_prein;
	HRESULT hr;
	hr=IRecord_Config.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		ReportExceptionError(_com_error((HRESULT)hr));
		return false;
	}	
	SQL.Format(_T("SELECT * FROM a_config WHERE EnginID='%s'"),GetProjectID());
	try
	{
		IRecord_Config->Open(_variant_t(SQL),_variant_t((IDispatch*)GetCurrentProjectConnect()),
							 adOpenDynamic,adLockOptimistic,adCmdText);
		if( !IRecord_Config->adoEOF && !IRecord_Config->BOF )
		{
			e_hours = vtof(IRecord_Config->GetCollect(_variant_t("年运行小时")));
			e_wind =  vtof(IRecord_Config->GetCollect(_variant_t("室外风速")));
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
	int pos = 1;
	//将所有记录重新选择保温结构
	for(m_ICurrentProjectRecordset->MoveFirst(); !m_ICurrentProjectRecordset->adoEOF; m_ICurrentProjectRecordset->MoveNext(),pos++ )
	{

		if(1)    //显示进度条.  可以设置满足条件时才显示.
		{
			((CMainFrame*)AfxGetApp()->GetMainWnd())->m_wndStatusBar.GetProgressCtrl().SetRange(0, Count);
			((CMainFrame*)AfxGetApp()->GetMainWnd())->m_wndStatusBar.OnProgress(pos);
			
			if( gbIsStopCalc )	//判断是否停止计算
			{
				((CMainFrame*)AfxGetApp()->GetMainWnd())->m_wndStatusBar.OnProgress(0);
				return false;
			}
		}

		//年运行小时
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_hour"),_variant_t(e_hours));
		//室外风速
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_wind"),_variant_t(e_wind));
		

		txtC_vol  = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("C_vol")));
		txtC_size = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("C_size")));
		txtC_temp = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("C_temp")));
		txtC_p_s  = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_place")));

		strPass   = vtos(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_pass")));
		if( strPass.CompareNoCase("Y") || txtC_vol.IsEmpty() || txtC_size.IsEmpty() || txtC_temp.IsEmpty() || txtC_p_s.IsEmpty() )
		{
			continue;
		}	
		try
		{	//自动选材
			AUTO(txtC_vol,txtC_size,txtC_temp,name_pre,name_pro,txtC_p_s,name_prein);
		}
		catch(_com_error &e)
		{
			ReportExceptionError(e);
			return false;
		}
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_name_in"),_variant_t(name_prein)); //内保温材料
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_name2"),_variant_t(name_pre));		//外保温材料
		m_ICurrentProjectRecordset->PutCollect(_variant_t("c_name3"),_variant_t(name_pro));		//保护材料

		m_ICurrentProjectRecordset->Update();
	}
	((CMainFrame*)AfxGetApp()->GetMainWnd())->m_wndStatusBar.OnProgress(0);
	return true;
}
///////////////////////////////////////////
//更新环境温度的取值表.
//
bool CEditOriginalData::UpdateEnvironmentComBox(CString strFilter)
{
	CString strSQL;			//SQL语句
	CString strTemp;
	CString strEnv;			//环境温度的取值
	int     nTmp;
	bool	bPlace=false;			//是否和安装地点有关.
	try
	{
		//首先获得当前显示的环境温度取值
		m_ctlEnvTemp.GetWindowText(strEnv);

		m_ctlEnvTemp.ResetContent();	//清空组合框
		if(m_pRsEnvironment->State == adStateOpen )
		{
			m_pRsEnvironment->Close();
		}
		//
		if( strFilter.IsEmpty() )
		{
			strSQL = "SELECT * FROM [环境温度取值表] ORDER BY Index ";	//从CODE(规范)库中取.
		}
		else
		{	
			//有过滤条件,与放热系数相关联.
			strSQL = "SELECT * FROM [环境温度取值表] WHERE "+strFilter+" ORDER BY Index ";	//从CODE(规范)库中取.
		}
		m_pRsEnvironment->Open(_variant_t(strSQL), this->GetPublicConnect().GetInterfacePtr(),
			adOpenStatic, adLockOptimistic, adCmdText);
		if( m_pRsEnvironment->GetRecordCount() <= 0 )
		{
			return true;
		}
		int nIndex;
		for(m_pRsEnvironment->MoveFirst(); !m_pRsEnvironment->adoEOF; m_pRsEnvironment->MoveNext())
		{
			strTemp = vtos( m_pRsEnvironment->GetCollect(_variant_t("Ta_Mode_Desc_Local")) );
			nIndex = m_ctlEnvTemp.AddString(strTemp);

			nTmp = vtoi( m_pRsEnvironment->GetCollect(_variant_t("Index")));			//索引号
			nIndex = 	m_ctlEnvTemp.SetItemData(nIndex, nTmp );
		}
		//重新恢复原来的选择.
		m_ctlEnvTemp.SelectString(-1,strEnv);
		
		//以下处理选择不同的放热系数,更新环境温度的取值.
		//考虑安装位置时,室内室外要相同.
/*		if ( -1 == ( pos = m_ctlHeatTran.GetCurSel() ) )
		{
			return false;
		}
		

		//原来选择的取值.
		strTemp = strEnv.Mid(0,4);
		if ( -1 != m_BuidInPlace.FindString(-1, strTemp) )
		{
			//和安装地点相关
			bPlace = true;
		}

		//选择的放热系数条件
		m_ctlHeatTran.GetLBText(pos, strTemp);
		//取前两个字符,是否为安装地点
		strTemp = strTemp.Mid(0,4);				
		//在所有可选用的环境取值中查找安装地点.
		pos = m_ctlEnvTemp.FindString(-1,strTemp);
		if ( -1 != pos && bPlace )  //
		{
			//原来已经选择的取值中查找
			nTmp = strEnv.Find(strTemp, 0);
			if ( -1 != nTmp )
			{
				//找到,选择原来的设置
				m_ctlEnvTemp.SelectString(-1,strEnv);
			}
			else
			{
				m_ctlEnvTemp.SetCurSel(pos);
			}
		}
		else
		{
			//用原来的选择在当前可选择的列表中查找.
			pos = m_ctlEnvTemp.SelectString(-1,strEnv);
			if ( -1 == pos)
			{
				m_ctlEnvTemp.SetCurSel(0);
			}
		}
*/
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return false;
	}
	return true;
}

//
//更新放热系数下拉列表.
//
bool CEditOriginalData::UpdateHeatTransferCoef(CString strFilter)
{
	CString strSQL;			//SQL语句
	CString strALPHA;		//放热系数的取值
	CString	strTmp;
	CString strValName;
	short	nTmp;
	short	nIndex;			//与环境温度对应的索引

	bool	bPlace=false;

	try
	{
		//获得当前的放热系数的取值
		m_ctlHeatTran.GetWindowText(strALPHA);
		
		this->m_ctlHeatTran.ResetContent();		//空清列表中的内容/
//		m_ctlHeatTran.SetDroppedWidth(200);
		if(m_pRsHeat->State == adStateOpen)
		{
			m_pRsHeat->Close();
		}
		if( strFilter.IsEmpty() )
		{
			strSQL = "SELECT * FROM [放热系数取值表] ORDER BY Index";
		}
		else
		{
			//有过滤条件,与环境温度相关联. 
			strSQL = "SELECT * FROM [放热系数取值表] WHERE "+strFilter+" ORDER BY Index";
		}
		
//		strSQL.Format("SELECT * FROM [放热系数取值表] %s ORDER BY Index",strFilter.IsEmpty()?"":"WHERE "+strFilter+" ");

		m_pRsHeat->Open(_variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		if( m_pRsHeat->GetRecordCount() <= 0 )
		{
			return true;
		}
		//列出符合参数条件的记录.
		for(; !m_pRsHeat->adoEOF; m_pRsHeat->MoveNext())
		{
			strTmp = vtos(m_pRsHeat->GetCollect(_T("Ta_Mode_Desc_Local")));
			strValName = vtos(m_pRsHeat->GetCollect(_T("Variable_Name_Desc")));

			if( !strTmp.IsEmpty() && !strValName.CompareNoCase("ALPHA") )
			{
				nIndex = m_ctlHeatTran.AddString(strTmp);
				//放热系数对应的索引号
				nTmp = vtoi(m_pRsHeat->GetCollect(_T("Index")));
				m_ctlHeatTran.SetItemData(nIndex, nTmp);
			}
		}

		//重新恢复原来的选择
		m_ctlHeatTran.SelectString(-1,strALPHA);

		//处理与环境温度取值的记录
/*		//考虑安装位置时,室内室外要相同.
		if ( -1 == ( pos = m_ctlEnvTemp.GetCurSel() ) )
		{
			return false;
		}

		//原来选择的取值.
		strTmp = strALPHA.Mid(0,4);
		if ( -1 != m_BuidInPlace.FindString(-1, strTmp) )
		{
			//和安装地点相关
			bPlace = true;
		}
		//选择的放热系数条件
		m_ctlEnvTemp.GetLBText(pos, strTmp);
		//取前两个字符,是否为安装地点
		strTmp = strTmp.Mid(0,4);				
		//在所有可选用的环境取值中查找安装地点.
		pos = m_ctlHeatTran.FindString(-1,strTmp);
		if ( -1 != pos && bPlace )  //
		{
			//原来已经选择的取值中查找
			nTmp = strALPHA.Find(strTmp, 0);
			if ( -1 != nTmp )
			{
				//找到,选择原来的设置
				m_ctlHeatTran.SelectString(-1,strALPHA);
			}
			else
			{
				m_ctlHeatTran.SetCurSel(pos);
			}
		}
		else
		{
			//用原来的选择在当前可选择的列表中查找.
			pos = m_ctlHeatTran.SelectString(-1,strALPHA);
			if ( -1 == pos)
			{
				m_ctlHeatTran.SetCurSel(0);
			}
		}
*/

	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;

}

/////////////////////////////////////////////////////
//功能:
//根据环境温度取值的索引,在[Ta_Variable]表中找出对应的值.
int CEditOriginalData::GetConditionTemp(double &dConTemp, int & nMethodIndex, int nIndex)
{
	CString strTmp;			//临时变量
	CString	strSQL;			//SQL语名.

	_RecordsetPtr	pRs;		//数据表的记录集.
	_variant_t		var;
	pRs.CreateInstance(__uuidof(Recordset));

	try
	{
		//不是由参数传过来的.重新从数据库中取出索引号
		if (nIndex == -1)
		{
			//获得对应的安装地点,保温条件的取值索引号
			nIndex = vtoi(m_ICurrentProjectRecordset->GetCollect(_variant_t("c_Env_Temp_Index") ));
		}
		if(m_pRsEnvironment==NULL || m_pRsEnvironment->State == adStateClosed || m_pRsEnvironment->GetRecordCount() <= 0  )
		{	
			return 0;
		}
		m_pRsEnvironment->MoveFirst();
		strSQL.Format("Index=%d",nIndex);
		m_pRsEnvironment->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( m_pRsEnvironment->adoEOF )
		{
			//没有找到索引值所对应的记录.
			return 0;
		}
		//计算方法的索引值.
		var	= m_pRsEnvironment->GetCollect(_T("InsulationCalcMethodIndex"));
		if( var.vt != VT_EMPTY && var.vt != VT_NULL )
		{
			nMethodIndex = vtoi(var);
		}
		//对应的外表面温度值.
		double dTmp;
		dTmp = vtof(m_pRsEnvironment->GetCollect(_T("SurfaceTemperature")));
		if ( DZero  < dTmp )
		{
			//只有当外表面温度大于0时才返回.
			strTmp.Format("%.2f",dTmp);
			m_dTs = strtod(strTmp,NULL);
		}
		else
		{
			//没有特定的外表面温度。
		}
		//环境温度的索引 
		nIndex = vtoi(m_pRsEnvironment->GetCollect(_T("TaModeIndex")));		//如果没有设置索引值勤，取出的值还是0
		strSQL.Format("SELECT * FROM [Ta_Variable] WHERE Index=%d",nIndex);
		pRs->Open(_variant_t(strSQL), GetPublicConnect().GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		if( !pRs->adoEOF && !pRs->BOF )
		{
			pRs->MoveFirst();
			//获得给定索引所对应的环境温度值.
 			dConTemp = vtof( pRs->GetCollect(_variant_t("Ta_Variable_VALUE")) );
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	return 1;
}

//功能：
//控制计算方法的选择.
void CEditOriginalData::OnCheckMethod() 
{
	this->UpdateControlsState();	
}
//控制环境温度的手动输入
void CEditOriginalData::OnCheckTa() 
{
	this->UpdateControlsState();		
}
//控制内层保温厚
void CEditOriginalData::OnCheckThick1() 
{
	CString	strMatName;
	short	bEnable;
	m_OutSizeMatName.GetWindowText(strMatName);
	if ( !strMatName.IsEmpty() )
	{
		//设置外保温层输入厚度的状态和内层相同.(手工输入或自动计算)
		bEnable = ((CButton*)GetDlgItem(IDC_CHECK_THICK1))->GetCheck(); 
		((CButton*)GetDlgItem(IDC_CHECK_THICK2))->SetCheck( bEnable );
	}
	else
	{
		//外保温层不可手动输入.
		((CButton*)GetDlgItem(IDC_CHECK_THICK2))->SetCheck( FALSE );
	}
	this->UpdateControlsState();

}
//控制外层保温厚
void CEditOriginalData::OnCheckThick2() 
{
	CString	strMatName;
	short	bEnable;
	m_InsideMatName.GetWindowText(strMatName);
	if ( !strMatName.IsEmpty() )
	{
		//设置内保温层输入厚度的状态和外层相同.(手工输入或自动计算)
		bEnable = ((CButton*)GetDlgItem(IDC_CHECK_THICK2))->GetCheck();
		((CButton*)GetDlgItem(IDC_CHECK_THICK1))->SetCheck( bEnable );
	}
	else
	{
		//内保温层不可手动输入.
		((CButton*)GetDlgItem(IDC_CHECK_THICK1))->SetCheck( FALSE );
	}
	this->UpdateControlsState();	

}
//控制外表面温度
void CEditOriginalData::OnCheckTs() 
{
	//在关联的选项中,一些被关联的值被自动改变之后, 则要进行改变之后的处理.
	AutoChangeValue();
	//更新控件的显示状态，是否可用。
	this->UpdateControlsState();	
}

//
//初始计算方法的选项
short CEditOriginalData::InitCalcMethod()
{
	CString strSQL;
	_RecordsetPtr pRsMethod;
	pRsMethod.CreateInstance(__uuidof(Recordset));

	try
	{
		this->m_ctlCalcMethod.ResetContent();		//空清列表中的内容/
		strSQL = "SELECT * FROM [保温计算方法表] ORDER BY Index";
		pRsMethod->Open(_variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		if( pRsMethod->GetRecordCount() <= 0 )
		{
			return true;
		}
		CString strTmp, strValName;
		int nIndex,nTmp;

		for(; !pRsMethod->adoEOF; pRsMethod->MoveNext())
		{
			strTmp = vtos(pRsMethod->GetCollect(_T("Ta_Mode_Desc_Local")));
			if( !strTmp.IsEmpty() )
			{
				nIndex = m_ctlCalcMethod.AddString(strTmp);
				//对应的索引号
				nTmp = vtoi(pRsMethod->GetCollect(_T("Index")));
				m_ctlCalcMethod.SetItemData(nIndex, nTmp);
			}
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}

	return 1;
}

//
//选择环境温度的取值.
void CEditOriginalData::OnSelchangeEnvironTemp() 
{
	CString strTmp,			//临时变量
			strSQL;			//SQL语名.

	int				nIndex;			//当前记录,环境温度取值的索引
	int				nPos;
	int				nMethodIndex;	//计算方法的取值索引
	_RecordsetPtr	pRs;			//数据表的记录集.
	_variant_t		var;
	short			SiteIndex;		//安装地点的索引
	pRs.CreateInstance(__uuidof(Recordset));

	try
	{
		UpdateData(TRUE);
		nPos = m_ctlEnvTemp.GetCurSel();
		if(nPos == -1)
		{
			return ;
		}
		
		//获得对应环境温度的取值索引号
		nIndex = m_ctlEnvTemp.GetItemData( nPos );

		//环境温度取值由用户选择的索引确定.
		//给定温度的索引(nIndex) , 取出温度值
		if( GetConditionTemp(m_dTa, nMethodIndex,nIndex) )
		{
		}			
		
		if( m_pRsEnvironment->GetRecordCount() <= 0  )
		{	
			return ;
		}
		m_pRsEnvironment->MoveFirst();
		strSQL.Format("Index=%d",nIndex);

		m_pRsEnvironment->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( m_pRsEnvironment->adoEOF )
		{
			//没有找环境温度取值所对应的记录.
			return ;
		}
		//计算方法的索引值.
		var	= m_pRsEnvironment->GetCollect(_T("InsulationCalcMethodIndex"));
		if( var.vt != VT_EMPTY && var.vt != VT_NULL )
		{
			nMethodIndex = vtoi(var);
			m_ctlCalcMethod.SetCurSel(nMethodIndex);
			((CButton*)GetDlgItem(IDC_CHECK_METHOD))->SetCheck(FALSE);
			((CButton*)GetDlgItem(IDC_EDIT_METHOD))->EnableWindow(FALSE);
			//环境温度与放热系数的关联
			strSQL.Format("InsulationCalcMethodIndex=%d OR InsulationCalcMethodIndex IS NULL",nMethodIndex);
		}
		else
		{
			//m_ctlCalcMethod.SetCurSel(-1);

			((CButton*)GetDlgItem(IDC_CHECK_METHOD))->SetCheck(TRUE);
			((CButton*)GetDlgItem(IDC_EDIT_METHOD))->EnableWindow(TRUE);
			//环境温度与放热系数的关联
			strSQL.Format("InsulationCalcMethodIndex IS NULL");
		}
		//环境温度设为不可改写的状态
		((CButton*)GetDlgItem(IDC_CHECK_TA))->SetCheck(FALSE);
		((CButton*)GetDlgItem(IDC_EDIT_TA))->EnableWindow(FALSE);

		//改变放热系数
		if( !((CButton*)GetDlgItem(IDC_CHECK_ALPHA))->GetCheck() )
		{
			//以计算方法为条件,确定相应的放热系数.
			this->UpdateHeatTransferCoef(strSQL);
		}

		//校验安装地点{{
		nPos = m_ctlEnvTemp.GetCurSel();
		//获得对应环境温度的取值索引号
		nIndex = m_ctlEnvTemp.GetItemData( nPos );
		if ( -1 != GetSiteIndex(m_pRsEnvironment,nIndex,SiteIndex) )
		{
			//
			UpdateSite(SiteIndex);
		}
		//校验安装地点}}

		ShowWindowRect();
		
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return ;
	}
	UpdateData(FALSE);
}

//
//选择放热系数确定计算方法
void CEditOriginalData::OnSelchangeHeatTransferCoef() 
{
	CString strTmp;			//临时变量
	CString	strSQL;			//SQL语名.

	int		nIndex;			//当前记录,放热系数取值的索引
	int		nPos;
	int		nMethodIndex;	//计算方法的取值索引
	short	SiteIndex;		//安装地点索引		

	_variant_t		var;
	
	try
	{
		UpdateData();
		
		nPos = m_ctlHeatTran.GetCurSel();
		if(nPos == -1)
		{
			return ;
		}

		//获得对应的安装地点,保温条件的取值索引号
		nIndex = m_ctlHeatTran.GetItemData( nPos );
			
		if( m_pRsHeat->GetRecordCount() <= 0  )
		{	
			return ;
		}

		m_pRsHeat->MoveFirst();
		strSQL.Format("Index=%d",nIndex);
		//
		m_pRsHeat->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( m_pRsHeat->adoEOF )
		{
			//没有找到放热系数,所对应的记录.
			return ;
		}
		//计算方法的索引值.
		var	= m_pRsHeat->GetCollect(_T("InsulationCalcMethodIndex"));
		if( var.vt != VT_EMPTY && var.vt != VT_NULL )
		{
			nMethodIndex = vtoi(var);
			m_ctlCalcMethod.SetCurSel(nMethodIndex);	
			((CButton*)GetDlgItem(IDC_CHECK_METHOD))->SetCheck(FALSE);
			((CButton*)GetDlgItem(IDC_EDIT_METHOD))->EnableWindow(FALSE);
			//设置放热系数与环境温度的关联.
			strSQL.Format("InsulationCalcMethodIndex=%d OR InsulationCalcMethodIndex IS NULL",nMethodIndex);
		}
		else
		{
			//m_ctlCalcMethod.SetCurSel(-1);

			((CButton*)GetDlgItem(IDC_CHECK_METHOD))->SetCheck(TRUE);
			((CButton*)GetDlgItem(IDC_EDIT_METHOD))->EnableWindow(TRUE);

			//设置放热系数与环境温度的关联.
			strSQL.Format("InsulationCalcMethodIndex IS NULL");			
		}

		//设置环境温度取值的关联
		//
		if( !((CButton*)GetDlgItem(IDC_CHECK_ENV_TEMP))->GetCheck() )
		{
			this->UpdateEnvironmentComBox( strSQL );
		}

		//校验安装地点{{		
		nPos = m_ctlHeatTran.GetCurSel();			
		//获得对应环境温度的取值索引号
		nIndex = m_ctlHeatTran.GetItemData( nPos );
		if ( -1 != GetSiteIndex(m_pRsHeat,nIndex,SiteIndex) )
		{
			//
			UpdateSite(SiteIndex);
		}
		//校验安装地点	}}

		ShowWindowRect();

	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());		
		return ;
	}
	UpdateData(FALSE);
}

//列出所有的环境温度取值
void CEditOriginalData::OnCheckEnvTemp() 
{
	if( ((CButton*)GetDlgItem(IDC_CHECK_ENV_TEMP))->GetCheck() )
	{
		this->UpdateEnvironmentComBox();
	}
}

//列出所有的放热系数的取值
void CEditOriginalData::OnCheckAlpha() 
{
	if( ((CButton*)GetDlgItem(IDC_CHECK_ALPHA))->GetCheck() )
	{
		this->UpdateHeatTransferCoef();
	}
}

//选择不同的计算方法
void CEditOriginalData::OnSelchangeEditMethod() 
{
	UpdateData(TRUE);
	
	int nPos = m_ctlCalcMethod.GetCurSel();		//选择的位置
	CString STR;								//方法名称
	long	nMethodIndex;						//方法的索引
	STR.Format("%d",m_ctlCalcMethod.GetItemData(nPos));
	nMethodIndex = m_ctlCalcMethod.GetItemData(nPos);
	if ( nSurfaceTemperatureMethod == nMethodIndex )
	{
		//外表面温度设为可以改写.
		if( !((CButton*)GetDlgItem(IDC_CHECK_TS))->GetCheck() )
		{
			//
			((CButton*)GetDlgItem(IDC_CHECK_TS))->SetCheck(TRUE);
			((CButton*)GetDlgItem(IDC_EDIT_TS))->EnableWindow(TRUE);
			((CButton*)GetDlgItem(IDC_EDIT_TS))->SetFocus();
		}
	}
	else
	{
		((CButton*)GetDlgItem(IDC_CHECK_TS))->SetCheck(FALSE);
		((CButton*)GetDlgItem(IDC_EDIT_TS))->EnableWindow(FALSE);
	}
	//如果防冻保温法需要加外增加一些参数.
	ShowWindowRect();
	//
	UpdateData(FALSE);
	
}

//初始管内介质 
short CEditOriginalData::InitPipeMedium()
{
	_RecordsetPtr pRsMedium;	//管内介质
	CString		  strSQL;		//SQL语句
	CString		  strMedium;	//介质名称
	short		  nIndex;		//介质对应的序号
	short		  pos;			//控件的位置

	try
	{
		
		pRsMedium.CreateInstance(__uuidof(Recordset));

		//
		strSQL.Format("SELECT * FROM [MediumAlias] ORDER BY MediumID");
		pRsMedium->Open(_variant_t(strSQL), theApp.m_pConMedium.GetInterfacePtr(),
						adOpenStatic, adLockOptimistic, adCmdText);
		

		//清空介质列表
		m_ctlPipeMedium.ResetContent();
		//
		for (; !pRsMedium->adoEOF; pRsMedium->MoveNext())
		{
			strMedium = vtos(pRsMedium->GetCollect(_variant_t("MediumAlias")));
			if ( !strMedium.IsEmpty() )
			{
				//介质别名对应介质的索引
				nIndex = vtoi(pRsMedium->GetCollect(_variant_t("MediumID")));
				//
				pos = m_ctlPipeMedium.AddString(strMedium);
				//设置管内介质别名的索引
				m_ctlPipeMedium.SetItemData(pos, nIndex);
			}
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	
	return 1;
}

//功能:
//将环境温度、放热系数和安装地点联成一起,各受控制
short CEditOriginalData::UpdateSite(short SiteIndex)
{
	CString		strTmp;		//临时变量
	CString		strSQL;		//SQL语句
	short		pos;		//
	short		nIndex;		//当前选项的索引值
	short		nCurrSiteIndex;		//当前的安装地点的索引
	bool		bEmpty=false;		//
	short		bReturnFlg;			//调用函时的返回状态
	//检查安装地点选项.
	nIndex = m_BuidInPlace.GetCurSel();
	if ( -1 == SiteIndex )
	{	//没有安装地点的控制
		if ( -1 == nIndex )
		{
			m_BuidInPlace.SetCurSel(0);
		}
		SiteIndex = m_BuidInPlace.GetCurSel();
	}
	else
	{		
		if ( -1 != nIndex && SiteIndex != nIndex )
		{
			//设置为参数传过来的索引.
			if( -1 == m_BuidInPlace.SetCurSel(SiteIndex) )
				m_BuidInPlace.SetCurSel(nIndex);			//还原
		}
	}

	//环境温度的选项
	if ( m_ctlEnvTemp.GetCount() > 0 )
	{
		pos = m_ctlEnvTemp.GetCurSel();
		if ( -1 == pos )
		{
			//没有当前项,默认为第一条
			pos = 0;
			bEmpty = true;
		}
		for (; pos < m_ctlEnvTemp.GetCount(); pos++ )
		{
			nIndex = (short)m_ctlEnvTemp.GetItemData(pos);
			//获得当前环境温度取值对应的安装地点的索引
			bReturnFlg = GetSiteIndex(m_pRsEnvironment,nIndex,nCurrSiteIndex);
			if ( !bEmpty && 0 == bReturnFlg )
			{	
				//当前的选择没有安装地点的控制时
				//
				m_ctlEnvTemp.SetCurSel(pos);
				break;
			}
			else
			{
				if ( 1 == bReturnFlg && SiteIndex == nCurrSiteIndex )
				{
					//设置当前判断的选项
					m_ctlEnvTemp.SetCurSel(pos);
					break;
				}
			}
			//有当前值时,再从第一条记录比较
			if( !bEmpty )
			{
				pos = -1;
				bEmpty = true;
			}
		}
		if (pos == m_ctlEnvTemp.GetCount())
		{
			//都没有限制安装地点时,设为第一条记录
			m_ctlEnvTemp.SetCurSel(0);
		}
		
	}


	bEmpty = false;
	//放热系数的选项
	if ( m_ctlHeatTran.GetCount() > 0 )
	{
		pos = m_ctlHeatTran.GetCurSel();
		if ( -1 == pos )
		{
			//没有当前项,默认为第一条.
			pos = 0;
			bEmpty = true;
		}
		for (; pos < m_ctlHeatTran.GetCount(); pos++ )
		{
			nIndex = (short)m_ctlHeatTran.GetItemData(pos);
			//获得当前环境温度取值对应的安装地点的索引.
			bReturnFlg = GetSiteIndex(m_pRsHeat,nIndex,nCurrSiteIndex);
			if ( (!bEmpty && 0 == bReturnFlg) || (1 == bReturnFlg && SiteIndex == nCurrSiteIndex) )
			{
				//设置为当前判断的选项
				m_ctlHeatTran.SetCurSel(pos);		
				break;
			}

			//有当前值时,再从第一条记录比较
			if( !bEmpty )
			{
				pos = -1;
				bEmpty = true;
			}
		}
		if ( pos == m_ctlHeatTran.GetCount() )
		{
			//都没有限制安装地点时,设为第一条记录		
			m_ctlHeatTran.SetCurSel(0);
		}
	}
	//在关联的选项中,一些被关联的值被自动改变之后, 则要进行改变之后的处理.
	AutoChangeValue();
    
	return 1;
}

//功能:
//获得安装地点的索引.
//返回值,   -1			出错或没有找到对应的索引
//			0			没有特别指定安装地点
//			1			找到,成功返回
short CEditOriginalData::GetSiteIndex(_RecordsetPtr &pRs, short nIndex, short& SiteIndex)
{
	try
	{
		SiteIndex = -1;
		if (pRs == NULL || pRs->GetRecordCount()<=0)
		{
			return -1;
		}
		
		CString			strSQL;
		_variant_t		var;
		pRs->Filter = (long)adFilterNone;
		strSQL.Format("Index=%d",nIndex);
		//定位到给定索引的记录.
		pRs->Filter = _variant_t(strSQL);
		if ( !pRs->adoEOF )
		{
			//获得安装地点的索引.
			var = pRs->GetCollect(_variant_t("SiteIndex"));
			SiteIndex = vtoi( var );
			pRs->Filter = (long)adFilterNone;
			if(var.vt == VT_NULL || var.vt == VT_EMPTY)
			{
				SiteIndex = -1;
				return 0;
			}
			else
			{
				return 1;
			}
		}
		pRs->Filter = (long)adFilterNone;
	}
	catch (_com_error)
	{
		return -1;
	}

	return -1;
}

//
//选择安装地点
void CEditOriginalData::OnSelchangeBuildinPlace() 
{
	UpdateData();
	short pos;
	pos = m_BuidInPlace.GetCurSel();
	if (-1 != pos)
	{
		UpdateSite(pos);
	}
	UpdateData(FALSE);
}


//--------------------------------------------------------------------------
//DATE:			[2005/5/20]
//
//功能:			选择管内介质				
//--------------------------------------------------------------------------
void CEditOriginalData::OnSelchangePipeMedium() 
{
	UpdateMediumMaterialDensity();
}

//--------------------------------------------------------------------------
// DATE         :[2005/5/23]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :检测输入数据的正确性,合法性.
//--------------------------------------------------------------------------
bool CEditOriginalData::CheckDataValidity()
{
	int		nMethodIndex=-1;		//计算方法的索引
	CString strTmp;
	double  t = m_MediumTemperature;	// 介质温度
	try
	{
		//如果计算方法为外表面温度法时，外表面温度一定要有值
		nMethodIndex = vtoi( m_ICurrentProjectRecordset->GetCollect(_variant_t("c_CalcMethod_Index")) );
		if( nSurfaceTemperatureMethod == nMethodIndex )
		{
			if( DZero >= fabs(m_dTs) )
			{
				AfxMessageBox("提示：\n\n计算方法为外表面温度法时，外表面温度不能为0。");
				((CButton*)GetDlgItem(IDC_CHECK_TS))->SetCheck(TRUE);
				((CButton*)GetDlgItem(IDC_EDIT_TS))->EnableWindow(TRUE);
				((CButton*)GetDlgItem(IDC_EDIT_TS))->SetFocus();
				return false;
			}
			else if ( !(m_dTa > t && m_dTs <= m_dTa && m_dTs >= t) ) // 保冷
					if ( !(m_dTa < t && m_dTs >= m_dTa && m_dTs <= t) )	// 保温
						if ( !(m_dTa == t && m_dTs == t) )		// 三个温度都相等
			{				
				AfxMessageBox("介质温度，外表温度或环境温度输入有错！");
				return false;
			}
		}
		
		//如果选择手动输入保温厚,则保温厚要大于0
		if ( 1 == m_bIsCalInThick || 1 == m_bIsCalPreThick )
		{ 
			m_InsideMatName.GetWindowText(strTmp);		//内层保温
			if ( !strTmp.IsEmpty() )
			{
				if( DZero >= fabs(m_dThick1))
				{
					AfxMessageBox("提示：\n\n手工输入保温层厚度时，内保温层厚度的值必须大于０ ");
					return false;
				}
			}
			m_OutSizeMatName.GetWindowText(strTmp);		//外层保温
			if ( !strTmp.IsEmpty() )				
			{
				if( DZero >= fabs(m_dThick2))
				{
					AfxMessageBox("提示：\n\n手工输入保温层厚度时，外保温层厚度的值必须大于０ ");
					return false;
				}
			}
		}
		CString	strInMat;		//内保温层材料
		CString	strOutMat;		//外保温层材料
		CString strProMat;		//保护层材料
		m_InsideMatName.GetWindowText(strInMat);
		m_OutSizeMatName.GetWindowText(strOutMat);
		if(!strInMat.IsEmpty() && strOutMat.IsEmpty())
		{
			AfxMessageBox("提示：\n\n请选择外保温层材料!");
			return false;
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}


//--------------------------------------------------------------------------
// DATE         :[2005/5/24]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :在关联的选项中,一些被关联的值被自动改变之后, 则要进行改变之后的处理.
//--------------------------------------------------------------------------
short CEditOriginalData::AutoChangeValue()
{
	CString	strTemp;			//临时字符串
	int		nMethodIndex;		//计算方法的索引值
	int		nIndex;
	short	pos;				//位置
	bool	bEnable = FALSE;			//是否激活。

	try
	{	//对应的环境温度的选项.
		m_ctlEnvTemp.GetWindowText(strTemp);
		if ( !strTemp.IsEmpty() )
		{
			pos = m_ctlEnvTemp.FindString(-1, strTemp);
			if ( -1 != pos )
			{
				//设置当前环境温度的取值。和对应的外表面温度值
				nIndex = m_ctlEnvTemp.GetItemData(pos);
				GetConditionTemp(m_dTa,nMethodIndex,nIndex);
			}
		}
		//计算方法的选项,与外表面温度关联。
		m_ctlCalcMethod.GetWindowText(strTemp);
		if ( !strTemp.IsEmpty() )
		{
			pos = m_ctlCalcMethod.FindString(-1, strTemp);
			if ( -1 != pos )
			{
				nIndex = m_ctlCalcMethod.GetItemData(pos);
				if ( nIndex == nSurfaceTemperatureMethod )
				{
					//外表面温度计算方法时.需要输入外表面温度的值。
					bEnable = TRUE;
				}
				else
				{
					bEnable = FALSE;
				}
			}
		}
		((CButton*)GetDlgItem(IDC_CHECK_TS))->SetCheck(bEnable);
		((CButton*)GetDlgItem(IDC_EDIT_TS))->EnableWindow(bEnable);
		//
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	return	1;
}

//------------------------------------------------------------------                
// DATE         :[2005/05/31]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :设置当前对话框的大小
//------------------------------------------------------------------
short CEditOriginalData::ShowWindowRect()
{
	//如果防冻保温法需要加外增加一些参数.
	CRect	rCongeal;		//显示输入防冻计算的各参数的窗口大小
	CRect	rWindow;		//当前对话框窗口的大小
	int		nWidth = 10;	//右边多加的空白处的宽度
	bool	bMinMove = false;	//是否改变现在的窗口的大小
	long	nMethodIndex = -1;	//计算方法的索引
	short	pos;

	pos = m_ctlCalcMethod.GetCurSel();
	if ( -1 != pos )
	{
		nMethodIndex = m_ctlCalcMethod.GetItemData(pos);
	}
	try
	{
		if ( nMethodIndex != nSubterraneanMethod )
		{
			UpdateControlSubterranean( -1 );
		}
		// 需要增加新的参数的方法
		switch( nMethodIndex )
		{
		case nPreventCongealMethod:	// 防冻保温计算
			if ( NULL == m_pDlgCongeal )
			{
				m_pDlgCongeal = new CDlgParameterCongeal();
			}
			AddDynamicDlg( m_pDlgCongeal, IDD_PARAMETER_CONGEAL );
			m_pDlgCongeal->SetProjectID( GetProjectID() );
			m_pDlgCongeal->SetCurrentRes( m_pRsCongealData );
			m_pDlgCongeal->GetDataToPreventCongealControl( m_ID ); // 获得防冻保温的参数
			break;
			
		case nSubterraneanMethod: // 埋地管道的计算方法
			if ( NULL == m_pDlgSubterranean )
			{
				m_pDlgSubterranean = new CDlgParameterSubterranean();
			}
			AddDynamicDlg( m_pDlgSubterranean, IDD_PARAMETER_SUBTERRANEAN );
			m_pDlgSubterranean->SetProjectID( GetProjectID() );
			m_pDlgSubterranean->SetCurrentRes( m_pRsSubterranean );
			m_pDlgSubterranean->GetDataToSubterraneanControl( m_ID ); // 获得埋地管道保温计算的参数
			break;
			
		default:
			bMinMove = true;	// 不须要显示增加的窗口
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		// 
		return 0 ;
	}
	// 移动窗口
	this->GetWindowRect(&rWindow);
	m_ctlTitleTab.GetClientRect(rCongeal);
	
	if ( bMinMove )	// 只显示标准的输入窗口 
	{
		if ( m_ctlTitleTab.IsWindowEnabled() )	// 但是当前显示了增加的窗口
		{
			rWindow.right -= rCongeal.right + nWidth;
			//须要改变现在的窗口的大小
			MoveWindow( &rWindow );
			m_ctlTitleTab.EnableWindow( FALSE );
		}
	}
	else if ( !m_ctlTitleTab.IsWindowEnabled() ) // 显示增加的窗口
	{
		rWindow.right += rCongeal.right + nWidth;
		MoveWindow( &rWindow );
		m_ctlTitleTab.EnableWindow( TRUE );
	}

/*	if ( nPreventCongealMethod == nMethodIndex || nSubterraneanMethod == nMethodIndex )	//防冻保温法
	{
		rWindow.right += rCongeal.right+nWidth;
		if ( !m_ctlTitleTab.IsWindowEnabled())
		{
			bMove = true;
			//显示了输入防冻保温法的和参数
			m_ctlTitleTab.EnableWindow(TRUE);
		}
	}
	else
	{
		rWindow.right -= rCongeal.right+nWidth;
		if ( m_ctlTitleTab.IsWindowEnabled())
		{
			bMove = true;
			//隐藏
			m_ctlTitleTab.EnableWindow(FALSE);
		}
	}
	if ( bMove )
	{
		//须要改变现在的窗口的大小
		this->MoveWindow(&rWindow);
	}

	// 将数据库中的数据写入到防冻结的参数中
	GetDataToPreventCongealControl();
*/
	return	1;
}


//------------------------------------------------------------------                
// DATE         :[2005/05/31]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :从数据库中取出数据到防冻结计算参数中.
//------------------------------------------------------------------
short CEditOriginalData::GetDataToPreventCongealControl()
{
//	double		dTmp;
	CString		strTmp;		//临时字符
	CString		strSQL;		//SQL语句
	
	try
	{
		//如果不是防冻保温法计算时，将不取出相应的数据。
		if ( !m_ctlTitleTab.IsWindowEnabled() )
		{
			return 0;
		}
		//
		if ( m_pRsCongealData == NULL || m_pRsCongealData->State != adStateOpen )
		{
			return 0;
		}

/*
		if (m_pRsCongealData->GetRecordCount() > 0 )
		{
			m_pRsCongealData->MoveFirst();
			strSQL.Format("ID=%d",m_ID);
			m_pRsCongealData->Find(_bstr_t(strSQL), NULL, adSearchForward);
		}
		if ( m_pRsCongealData->adoEOF || m_pRsCongealData->BOF )
		{
			//没有当前序号时，增加一条新记录.
			m_pRsCongealData->AddNew();

			//序号
			m_pRsCongealData->PutCollect(_T("ID"), _variant_t((long)m_ID));

			//工程代号
			m_pRsCongealData->PutCollect(_T("EnginID"), _variant_t(GetProjectID()));
			
			m_pRsCongealData->Update(vtMissing);
		}
		
		//介质密度
		m_dMediumDensity = vtof(m_pRsCongealData->GetCollect(_variant_t("c_Medium_Density")));

		//管材密度
		m_dPipeDensity	= vtof( m_pRsCongealData->GetCollect(_T("c_Pipe_Density")) );

		//介质比热
		m_dMediumHeat	= vtof( m_pRsCongealData->GetCollect(_T("c_Medium_Heat")) );

		//管材比热
		m_dPipeHeat		= vtof( m_pRsCongealData->GetCollect(_T("c_Pipe_Heat")) );

		//流量
		m_dFlux			= vtof( m_pRsCongealData->GetCollect(_T("c_Flux")) );

		//介质融解热
		m_dMediumMeltHeat	= vtof( m_pRsCongealData->GetCollect(_T("c_Medium_Melt_Heat")) );

		//介质在管内冻结温度
		m_dMediumCongealTemp = vtof( m_pRsCongealData->GetCollect(_T("c_Medium_Congeal_Temp")) );

		//防冻结管道允许液体停留时间(小时)
		m_dStopTime		= vtof( m_pRsCongealData->GetCollect(_T("c_Stop_Time")) );

		//管道通过支吊架处的热(或冷)损失的附加系数
		dTmp			= vtof( m_pRsCongealData->GetCollect(_T("c_Heat_Loss_Data")) );	
		strTmp.Format("%.2f",dTmp);
		m_ctlHeatLossData.SetWindowText(strTmp);
		
		//	= vtof( m_pRsCongealData->GetCollect(_T("EnginID")) );
*/
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	
	return	1;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/01]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :将防冻保温法须要的数据输入到数据库中。
//------------------------------------------------------------------
short CEditOriginalData::PutDataToPreventCongealDB(long nID)
{
	try
	{
		//如果不是防冻保温法计算时，将不取出相应的数据。
		if ( !m_ctlTitleTab.IsWindowEnabled() )
		{
			return 0 ;
		}
		//test zsy 2006
		if ( m_pDlgCongeal != NULL )
		{
			m_pDlgCongeal->PutDataToPreventCongealDB(nID);
			return 1;
		}
		//test zsy 2006

		if ( m_pRsCongealData == NULL || m_pRsCongealData->State != adStateOpen )
		{
			return 0;
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
/*
		CString strTmp;
		double	dTmp;

		if (m_pRsCongealData->adoEOF || m_pRsCongealData->BOF)
		{
			return 0;
		}

		//介质密度
		m_pRsCongealData->PutCollect(_T("c_Medium_Density"), _variant_t(m_dMediumDensity));

		//管材密度
		m_pRsCongealData->PutCollect( _T("c_Pipe_Density"), _variant_t(m_dPipeDensity) );

		//介质比热
		m_pRsCongealData->PutCollect( _T("c_Medium_Heat"), _variant_t(m_dMediumHeat) );

		//管材比热
		m_pRsCongealData->PutCollect( _T("c_Pipe_Heat"), _variant_t(m_dPipeHeat) );

		//流量
		m_pRsCongealData->PutCollect( _T("c_Flux"), _variant_t(m_dFlux) );

		//介质融解热
		m_pRsCongealData->PutCollect( _T("c_Medium_Melt_Heat"), _variant_t(m_dMediumMeltHeat) );

		//介质在管内冻结温度
		m_pRsCongealData->PutCollect( _T("c_Medium_Congeal_Temp"), _variant_t(m_dMediumCongealTemp) );

		//防冻结管道允许液体停留时间(小时)
		m_pRsCongealData->PutCollect( _T("c_Stop_Time"), _variant_t(m_dStopTime) );

		//管道通过支吊架处的热(或冷)损失的附加系数
		m_ctlHeatLossData.GetWindowText(strTmp);
		dTmp = strtod(strTmp, NULL);
		m_pRsCongealData->PutCollect( _T("c_Heat_Loss_Data"), _variant_t(dTmp) );

		//序号
	//	m_pRsCongealData->PutCollect(_T("ID"), _variant_t((long)m_ID));

		//工程代号
	//	m_pRsCongealData->PutCollect(_T("EnginID"), _variant_t(GetProjectID()));
		 
		m_pRsCongealData->Update(vtMissing);

	}
	catch (_com_error& e) 
	{
		AfxMessageBox(e.Description());
		return 0;
	}
*/
	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :在一根直管道上增加阀门.
//------------------------------------------------------------------
BOOL CEditOriginalData::AddValveData()
{
	short ID = 1;			//序号
	CString strSQL;
 	try
	{
		if (m_ICurrentProjectRecordset->adoEOF && m_ICurrentProjectRecordset->BOF)
		{
			ID = 0;		//原始记录为空
		}
		else
		{
			ID = vtoi(m_ICurrentProjectRecordset->GetCollect(_variant_t("ID")));
		}

		//在当前的位置增加一条新的记录
		this->InsertNew(m_ICurrentProjectRecordset, 1);
		m_ICurrentProjectRecordset->PutCollect(_variant_t("ID"),_variant_t((long)++ID));
		m_ICurrentProjectRecordset->PutCollect(_variant_t("EnginID"),_variant_t(m_ProjectID));

		if (NULL != m_pRsCongealData && (!m_pRsCongealData->GetadoEOF() || !m_pRsCongealData->GetBOF()))
		{
			//改变序号
			strSQL.Format("UPDATE [PRE_CALC_CONGEAL] SET ID=ID+1 WHERE ID>=%d AND EnginID='%s' ", ID, m_ProjectID);
			m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		}
		if (NULL != m_pRsSubterranean && (!m_pRsSubterranean->GetadoEOF() || !m_pRsSubterranean->GetBOF()))
		{
			strSQL.Format("UPDATE [PRE_CALC_SUBTERRANEAN] SET ID=ID+1 WHERE ID>=%d AND EnginID='%s' ", ID, m_ProjectID);
			m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		}
		
		//将当前的数据保存到新增加的记录中
		this->PutDataToDatabaseFromControl(ID);
		this->PutDataToEveryControl();
		
		//重新编号并定位到指定的记录
		RenewNumberFind(m_ICurrentProjectRecordset, ID, ID);

		//更新控件的状态
		UpdateControlsState();
	}
	catch (_com_error)
	{
		return FALSE;
	}

	return TRUE;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :响应增加一个阀门按钮的消息.
//------------------------------------------------------------------
void CEditOriginalData::OnButtonAddValve() 
{
	UpdateData();
	
	CString	strTmp;
	long	pipeID;			//管道的序号
	long	valveID;		//阀门的序号
	pipeID = m_ID;

	//根据外径和壁厚来判断是否为管道或平壁.
	m_PipeSize.GetWindowText(strTmp);
	if (strtod(strTmp, NULL) >= m_dMaxD0)
	{
		m_bIsPlane = TRUE;
	}
	else
	{
		m_bIsPlane = FALSE;
	}
	//判断是否先保存数据.
	if(m_IsUpdateByRoll)
	{
		PutDataToDatabaseFromControl();
		if( !Refresh() )
			return;
	}

/*
	//test  [2005/06/28]
	UpdateData();
	if (!m_bIsPlane && m_Amount > m_dValveSpace)
	{
		//为管道,同时管道长度大天10m时.自动增加一个阀门
		AddValveData();
	}
*/
	//将管道与阀门的关系映射到一个表中
	try
	{
		//管道时
		if (!m_bIsPlane)
		{
			//为管道,增加一个阀门
			strTmp.Format("%f",m_dMaxD0);
			m_PipeSize.SetWindowText(strTmp);		//设置为平壁。
			m_PipeThick.SetWindowText("0");
			//类型
			if ( !m_VolNo.IsEmpty() )
			{
				m_VolNo.TrimRight();
				m_VolNo.TrimLeft();
				while (m_VolNo.GetLength() <= 5)
				{
					m_VolNo+=" ";					
				}
				m_VolNo.SetAt(5,'V');
	 			m_VolNo = m_VolNo.Mid(0,6);
			}
			//设备名
			m_PipeDeviceName.GetWindowText(strTmp);
			if (strTmp.IsEmpty() || -1 == strTmp.Find("阀门"))
			{
				strTmp += "阀门";
				m_PipeDeviceName.SetWindowText(strTmp); 				 
			}
			UpdateData(FALSE);
			
			AddValveData();
		}
		else
		{
			//平壁上不能增加阀门
			AfxMessageBox("平壁上不能增加阀门");
			return;
		}

		_RecordsetPtr pRs;			//管道与阀门的关系表
		CString	strSQL;				//SQL语句
										   												
		pRs.CreateInstance(__uuidof(Recordset));
		//

 		strSQL.Format("SELECT * FROM [pipe_valve] WHERE PipeID=%d AND EnginID='%s' ORDER BY ValveID",
						pipeID, m_ProjectID);
		

		pRs->Open(_variant_t(strSQL), theApp.m_pConnection.GetInterfacePtr(),
				adOpenStatic, adLockOptimistic, adCmdText);
		if (pRs->adoEOF && pRs->BOF && pRs->GetRecordCount() <= 0 )
		{
			//在管道后面加一条阀门记录。
			valveID = vtoi(  m_ICurrentProjectRecordset->GetCollect(_variant_t("ID")) );
		}
		else
		{
			//在“管道－阀门”映射表中取当前管道对应的阀门的最大序号。
			pRs->MoveLast();
			valveID = vtoi(pRs->GetCollect(_variant_t("valveID")));
			valveID++;
		}

		//增加管道-阀门的映射
		pRs->AddNew();
		pRs->PutCollect(_variant_t("PipeID"),_variant_t(pipeID));
		pRs->PutCollect(_variant_t("ValveID"),_variant_t(valveID));
		pRs->PutCollect(_variant_t("EnginID"),_variant_t(m_ProjectID));
		
		pRs->Update();
   
		//改变序号
		strSQL.Format("UPDATE [PIPE_VALVE] SET PipeID=PipeID+1 WHERE PipeID>%d AND EnginID='%s' ",
						pipeID, m_ProjectID);
		m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		//
		strSQL.Format("UPDATE [PIPE_VALVE] SET VALVEID=VALVEID+1 WHERE VALVEID>%d AND EnginID='%s' ",
						valveID, m_ProjectID);
		m_ICurrentProjectConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		
		//更新控件的状态
		UpdateControlsState();

		UpdateData(FALSE);
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());
		return;
	}
}


//------------------------------------------------------------------                
// DATE         :[2005/06/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :对指定的记录集重新编号,并定位到相应的位置.
//------------------------------------------------------------------
BOOL CEditOriginalData::RenewNumberFind(_RecordsetPtr &pRs, long ID, long FindNO)
{
	try
	{
		if (pRs == NULL || pRs->adoEOF && pRs->BOF)
		{
			return FALSE;
		}
		//重新编号
		for (; !pRs->adoEOF; pRs->MoveNext(), ID++)
		{
			pRs->PutCollect(_variant_t("ID"),ID);
		}
		//定位到指定的记录
		for(pRs->MoveFirst(); FindNO > 1 && !pRs->adoEOF; pRs->MoveNext(),FindNO--)
		{
			;
		}
		if (pRs->adoEOF)
		{
			pRs->MovePrevious();
		}
		return TRUE;
	}
	catch (_com_error)
	{
		return FALSE;
	}
	return TRUE;
}

BOOL CEditOriginalData::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	m_pctlTip.RelayEvent(pMsg);
	return CDialog::PreTranslateMessage(pMsg);
}

//------------------------------------------------------------------                
// DATE         :[2005/06/23]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :加载需要提示信息的编辑框与之对应的提示符。
//------------------------------------------------------------------
BOOL CEditOriginalData::InitToolTip()
{
	int	pos = 0;		//数组下标
	Txtup	TipArr[20];
	
	TipArr[pos].dID = IDC_PIPE_SIZE;	//管道外径/规格
	TipArr[pos++].txt = _T("mm");		//单位

	TipArr[pos].dID = IDC_PIPE_THICK;	//管道壁厚
	TipArr[pos++].txt = _T("mm");		//单位

	TipArr[pos].dID = IDC_PIPE_PRESSURE;	//管内压力
	TipArr[pos++].txt = _T("MPa");			//单位
	
	TipArr[pos].dID = IDC_MEDIUM_TEMPERATURE;	//介质温度℃
	TipArr[pos++].txt = _T("℃");				//单位

	TipArr[pos].dID = IDC_EDIT_TS;				//表面温度℃
	TipArr[pos++].txt = _T("℃");				//单位
	
	TipArr[pos].dID = IDC_EDIT_TA;				//环境温度℃
	TipArr[pos++].txt = _T("℃");				//单位

	TipArr[pos].dID = IDC_AMOUNT;				//管道长度或平壁面积
	TipArr[pos++].txt = _T("m 或 ㎡");			//单位

	TipArr[pos].dID = IDC_EDIT_THICK2;			//外保温厚.
	TipArr[pos++].txt = _T("mm");				//单位

	TipArr[pos].dID = IDC_EDIT_THICK1;			//内保温厚.
	TipArr[pos++].txt = _T("mm");				//单位

	TipArr[pos].dID = IDC_WIDE_SPEED;			//风速（m/s）
	TipArr[pos++].txt = _T("m/s");				//单位

	TipArr[pos].dID = IDC_EDIT_HEAT_DENSITY;	//热流密度
	TipArr[pos++].txt = _T("W/㎡");				//单位
	
	TipArr[pos].dID = IDC_EDIT_HEAT_LOSS;		//散热量
	TipArr[pos++].txt = _T("W");				//单位

	TipArr[pos].dID = IDC_EDIT_PLAIM_THICK;		//沿风速方向的平壁宽度
	TipArr[pos++].txt = _T("m");

	//////////////////////////////////////////////////////////////////////////
/*
	// 防冻结计算新增加的参数
	TipArr[pos].dID = IDC_EDIT_MEDIUM_DENSITY1;	//介质密度
	TipArr[pos++].txt = _T("Kg/m3");
	
	TipArr[pos].dID = IDC_EDIT_PIPE_DENSITY;	//管材密度
	TipArr[pos++].txt = _T("Kg/m3");
	
	TipArr[pos].dID = IDC_EDIT_MEDIUM_HEAT;		//介质比热
	TipArr[pos++].txt = _T("J/(Kg/℃)");

	TipArr[pos].dID = IDC_EDIT_PIPE_HEAT;		//管道材料的比热
	TipArr[pos++].txt = _T("J/(Kg/℃)");
	
	TipArr[pos].dID = IDC_EDIT_MEDIUM_MELT_HEAT;//介质溶解热
	TipArr[pos++].txt = _T("J/Kg");

	TipArr[pos].dID = IDC_EDIT_MEDIUM_CONGEAL_TEMP;//介质在管内冻结温度
	TipArr[pos++].txt = _T("℃");

	TipArr[pos].dID = IDC_EDIT_STOP_TIME;		//防冻结管道允许液体停留时间
	TipArr[pos++].txt = _T("h");

*/
	m_pctlTip.Create(this);
//	m_pctlTip.SetTipBkColor(RGB(255,255,255));
	for (int i=0; i<pos; i++)
	{
		m_pctlTip.AddTool(GetDlgItem(TipArr[i].dID), (TipArr[i].txt));
	}

	return TRUE;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/24]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :打开常量数据库,初始一些常数.
//------------------------------------------------------------------
BOOL CEditOriginalData::InitConstant()
{
	_RecordsetPtr pRs;			// 常数库的记录集
	pRs.CreateInstance(__uuidof(Recordset));
	CString strSQL;				// SQL语句
	try
	{
		strSQL = "SELECT * FROM [A_CONFIG] WHERE EnginID='"+m_ProjectID+"' ";
		pRs->Open(_variant_t(strSQL), GetCurrentProjectConnect().GetInterfacePtr(),
				adOpenStatic, adLockOptimistic, adCmdText);
		if (pRs->adoEOF && pRs->BOF)
		{
			// 常量库中没有记录。
			m_dMaxD0 = 2000;		//分界外径赋一个默认值  
			m_dValveSpace = 10;		//阀门的间距 
			return	FALSE; 
		}
		m_dMaxD0 = vtof(pRs->GetCollect(_variant_t("平壁与圆管的分界外径")) );		
		if( m_dMaxD0 <= 0 )
		{
			// 当分界外径不合理时,赋一个默认值 
			m_dMaxD0 = 2000;
		}
		m_dValveSpace = vtof(pRs->GetCollect(_variant_t("阀门间距"))); 
		if (m_dValveSpace <= 0) 
		{
			m_dValveSpace = 10;
		}
		
		//获得锅炉出口过热蒸汽的压力
		m_dBoilerPressure = vtof(pRs->GetCollect(_variant_t("锅炉出口过热蒸汽压力")));
		//锅炉出口过热蒸汽温度
		m_dBoilerTemperature = vtof(pRs->GetCollect(_variant_t("锅炉出口过热蒸汽温度")));
		//冷却水温度
		m_dCoolantTemperature = vtof(pRs->GetCollect(_variant_t("冷却水温度")));
	}
	catch (_com_error)
	{	
		return FALSE;
	}
	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/11/14]
// Parameter(s) :
// Return       :
// Remark       :计算热价比主汽价
//------------------------------------------------------------------
void CEditOriginalData::OnButtonCalcRatio() 
{
	UpdateData();
	CString strSQL;	// SQL语句
	double	dPriRatio=0;	// 查表获得的热价比
	if ( gbIsSelTblPrice )	// 查表获得“介质用值系数”
	{
		try
		{
			_RecordsetPtr pRsVol;	// 所有工程的卷册
			pRsVol.CreateInstance( __uuidof( Recordset ) );
			CString strPipeName;	// 当前记录的管道或设备名称
			m_PipeDeviceName.GetWindowText( strPipeName );
			if ( !strPipeName.IsEmpty() )
			{				
				strSQL = " SELECT * FROM [A_VOL] WHERE [VOL_NAME]='"+strPipeName+"' AND [VOL_BAK]='UESoft' ";
				if ( pRsVol->State == adStateOpen )
				{
					pRsVol->Close();
				}
				pRsVol->Open( _variant_t( strSQL ), GetPublicConnect().GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText );
				if ( !pRsVol->adoEOF ||  !pRsVol->BOF)
				{
					pRsVol->MoveFirst();
					dPriRatio = vtof( pRsVol->GetCollect( _variant_t( "vol_price" ) ) );
				}
				else
				{
					strSQL = " SELECT * FROM [A_VOL] WHERE [VOL_NAME]='"+strPipeName+"' AND LEFT(VOL,5)=LEFT('"+m_VolNo+"',5) ";
					if ( pRsVol->State == adStateOpen )
					{
						pRsVol->Close();
					}
					pRsVol->Open( _variant_t( strSQL ), GetPublicConnect().GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText );
					if ( !pRsVol->adoEOF || !pRsVol->BOF )
					{
						pRsVol->MoveFirst();
						dPriRatio = vtof( pRsVol->GetCollect( _variant_t( "vol_price" ) ) );
					}
				}
			}
		}
		catch (_com_error& e)
		{
			AfxMessageBox( e.Description() );
			return ;
		}
		if ( dPriRatio > 0 )	// 在表中找到对应的比值
		{
			m_PriceRatio = dPriRatio;
			UpdateData( FALSE );
			return ;
		}
	}
	//////////////////////////////////////////////////////////////////////////	
	//					利用公式计算介质用值系数

	InitConstant();		// 初始常数
	int RANGE;
	double	dPipePress = m_dPressure+atmospheric_pressure;	//计算水蒸气时取管内表压力+大气压力
	double  dMediumTemp = m_MediumTemperature;	//介质温度
	double  Tw = m_dCoolantTemperature; //冷却水温度
	double	Sw=0;					//冷却水比熵
	double  Hw=0;					//冷却水比焓
	double	Sst=0;					//锅炉出口过热蒸汽比熵
	double  Hst=0;					//锅炉出口过热蒸汽比焓
	double  S=0;					//介质比熵
	double  H=0;					//介质比焓

	CString strMediaName;
	CString strTempMedia;
	m_ctlPipeMedium.GetWindowText(strMediaName);
	CHeatPreCal::GetMediumDensity(strMediaName, &strTempMedia);
	if ( !strTempMedia.IsEmpty() )
	{
		//只能在介质为“水” 和 “蒸汽”时用动态库计算其特性.
		if ( strTempMedia.Find("水") || strTempMedia.Find("蒸汽") )
		{
			if ( 0 == giCalSteanProp )			//      IF97 或 IFC67
				SETSTD_WASP( 97 );
			else
				SETSTD_WASP( 67 );
			
			//冷却水
			PT2S(dPipePress,Tw, &Sw, &RANGE);
			PT2H(dPipePress,Tw, &Hw, &RANGE);
			
			//锅炉出口过热蒸汽
			PT2S(m_dBoilerPressure, m_dBoilerTemperature, &Sst, &RANGE);
			PT2H(m_dBoilerPressure, m_dBoilerTemperature, &Hst, &RANGE);
			
			//介质
			PT2S(dPipePress,dMediumTemp, &S, &RANGE);
			PT2H(dPipePress,dMediumTemp, &H, &RANGE);
			
			double dAe = (H-Hw-(Tw+273)*(S - Sw)) / (Hst-Hw-(Tw+273)*(Sst-Sw));
			
			dAe = (dAe < 0) ? 0 : dAe;
			CString str;
			
			str.Format( "%.2f", dAe );
			m_PriceRatio = strtod(str, NULL);
		}
	}
	
	UpdateData(FALSE);
}

//------------------------------------------------------------------
// DATE         :[2005/11/17]
// Parameter(s) :
// Return       :
// Remark       :从数据库表中根据管道类型名获得介质佣值系数
//------------------------------------------------------------------
BOOL CEditOriginalData::GetAeValue(CString strTypeName, double &dAe)
{
	if (strTypeName.IsEmpty())
	{
		return FALSE;
	}
	CString strSQL;
	
	return TRUE;
}

//------------------------------------------------------------------                
// DATE         : [2006/02/11]
// Author       : ZSY
// Parameter(s) : 
// Return       : 
// Remark       : 创建一个子对话框，
//------------------------------------------------------------------
void CEditOriginalData::AddDynamicDlg(CDialog *pDlg, UINT nID)
{
	//2006.01.07  动态加载一个输入对话框 

	if ( IsWindow(m_pDlgCurChild->GetSafeHwnd()) )
	{
		m_pDlgCurChild->ShowWindow( SW_HIDE );
	}
	if ( pDlg == NULL || !IsWindow( pDlg->GetSafeHwnd() ) )
	{ 
		pDlg->Create( nID, this );	
		DWORD style=::GetWindowLong(pDlg->GetSafeHwnd(),GWL_STYLE);
		if ( style & WS_CAPTION ) 
		{ 
			style &= ~WS_CAPTION; 
		} 
		style &= ~WS_SYSMENU; 
		SetWindowLong( pDlg->GetSafeHwnd(), GWL_STYLE, style );
	}

	CRect rect;
	m_ctlTitleTab.GetWindowRect( &rect );
	this->ScreenToClient( &rect );
	rect.DeflateRect(3,11,3,3);
	pDlg->SetWindowPos( NULL, rect.left, rect.top, rect.Width(), rect.Height(), SWP_SHOWWINDOW|SWP_NOZORDER|SWP_NOACTIVATE );
	// 设置标题
	CString strTitle;
	pDlg->GetWindowText(strTitle);
	m_ctlTitleTab.SetWindowText(strTitle);
 
	m_pDlgCurChild = pDlg;
}


//------------------------------------------------------------------
// DATE         : [2006/02/20]
// Author       : ZSY
// Parameter(s) : VOID
// Return       : VOID
// Remark       : 将埋地保温法必须要的数据输入到数据库中。
//------------------------------------------------------------------
void CEditOriginalData::PutDataToSubterraneanDB(long nID)
{
	if ( !m_ctlTitleTab.IsWindowEnabled() )	//如果没显示新增加的窗口时，将不取出相应的数据。
	{
		return ;
	}

	if ( m_pDlgSubterranean != NULL )
	{
		//
		m_pDlgSubterranean->PutDataToSubterraneanDB(nID);
	}
}

void CEditOriginalData::UpdateControlSubterranean(int nPipeCount)
{
	BOOL bIsEnable = TRUE;
	int  i;
	int  j;
	CString strText;
	long pID[][2] = 
	{
		{IDC_CHECK_THICK1,  IDC_EDIT_THICK1},
		{IDC_CHECK_THICK2,  IDC_EDIT_THICK2},
		{IDC_STATIC_THICK1,	IDC_INSIDE_MATNAME},
		{IDC_STATIC_THICK2,	IDC_OUTSIZE_MATNAME},		
		{-1,-1}
	};
	CString strReplace[][2] = 
	{
		{_T("内"), _T("第二根")},
		{_T("外"), _T("第一根")},
		{_T(""),   _T("")},
	};
	switch(nPipeCount)
	{
	case 1:
	case 2:
		m_bIsCalInThick  = FALSE;
		m_bIsCalPreThick = FALSE;
		for (i = 0; pID[i][0] != -1; i++)
		{
			GetDlgItemText(pID[i][0], strText);
			for (j = 0; !strReplace[j][0].IsEmpty(); j++)
			{
				strText.Replace(strReplace[j][0], strReplace[j][1]);
			}
			SetDlgItemText(pID[i][0], strText);
			if ( i + nPipeCount < 4 )
			{
				bIsEnable = FALSE;
			}
			else
			{
				bIsEnable = TRUE;
			}
			GetDlgItem(pID[i][0])->EnableWindow( bIsEnable );
			GetDlgItem(pID[i][1])->EnableWindow( bIsEnable );
		}
		break;
	default:
		// 恢复到最初的状态
		for (i = 0; pID[i][0] != -1; i++)
		{
			GetDlgItemText(pID[i][0], strText);
			for (j = 0; !strReplace[j][0].IsEmpty(); j++)
			{
				strText.Replace(strReplace[j][1], strReplace[j][0]);
			}
			SetDlgItemText(pID[i][0], strText);
			GetDlgItem(pID[i][0])->EnableWindow( TRUE );
			GetDlgItem(pID[i][1])->EnableWindow( TRUE );
		}
		//内保温厚度
		GetDlgItem(IDC_EDIT_THICK1)->EnableWindow(((CButton*)GetDlgItem(IDC_CHECK_THICK1))->GetCheck()?TRUE:FALSE);
		//外保温厚度
		GetDlgItem(IDC_EDIT_THICK2)->EnableWindow(((CButton*)GetDlgItem(IDC_CHECK_THICK2))->GetCheck()?TRUE:FALSE);
	}

	UpdateData(FALSE);	
}

BOOL CEditOriginalData::NumberSubtractOne(_RecordsetPtr &pRs, long lStartPos)
{
	if (pRs == NULL || pRs->GetState() != adStateOpen || (pRs->adoEOF && pRs->BOF) )
	{
		return FALSE;
	}
	long nID;
	for (pRs->MoveFirst(); !pRs->adoEOF; pRs->MoveNext() )
	{
		nID = vtoi( pRs->GetCollect("ID") );
		if (nID > lStartPos)
		{
			nID--;
			pRs->PutCollect(_variant_t("ID"), _variant_t(nID));
			pRs->Update();
		}
	}
	return TRUE;
}

void CEditOriginalData::OnSelchangePipeMat() 
{
	UpdateMediumMaterialDensity();
}

// 根据介质名称和管道材料称获得其对就的密度，自动写入到防冻计算增加的对话框中。
void CEditOriginalData::UpdateMediumMaterialDensity()
{
	UpdateData(TRUE);

	UINT nMethodIndex = -1;
	CString strMedium;
	CString strMat;
	double  dMediumDensity;
	double  dMatDensity;

	UINT pos;
	pos = m_ctlCalcMethod.GetCurSel();
	if (pos >= 0)
	{
		nMethodIndex = m_ctlCalcMethod.GetItemData(pos);
	}
	if (nMethodIndex == nPreventCongealMethod)
	{
		if (NULL == m_pMatProp)
		{
			m_pMatProp = new GetPropertyofMaterial;
		}
	    m_ctlPipeMedium.GetWindowText(strMedium);
		if (!strMedium.IsEmpty())
		{
			dMediumDensity = CHeatPreCal::GetMediumDensity(strMedium);
			if (dMediumDensity > 0)
			{
				m_pDlgCongeal->m_dMediumDensity = dMediumDensity;
			}
		}

		m_PipeMat.GetWindowText(strMat);
		if (!strMat.IsEmpty())
		{
			dMatDensity = m_pMatProp->GetDensity( strMat,2,theApp.m_pConMaterial );
			if (dMatDensity > 0)
			{
				m_pDlgCongeal->m_dPipeDensity = dMatDensity;
			}
		}
		m_pDlgCongeal->UpdateData(FALSE);	
	}
	
	UpdateData(FALSE);
}
