// DlgOption.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DlgOption.h"
#include "EDIBgbl.h"

#include "MaterialName.h"		//用于名称的替换
#include "heatprecal.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDlgOption dialog
extern CAutoIPEDApp theApp;


CDlgOption::CDlgOption(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgOption::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgOption)
	m_IsColseCompress = bIsCloseCompress;
	m_IsReplOld = bIsReplaceOld;
	m_IsMoveUpdate = bIsMoveUpdate;
	m_IsAutoSelectPre = bIsAutoSelectPre;
	m_bIsHeatLoss =		bIsHeatLoss;
	m_IsAutoAddValve = bIsAutoAddValve;
	m_dMaxRatio = 0.0;
	m_dSurfaceMaxTemp = 0.0;
	m_bIsRunSelectEng= gbIsRunSelectEng;
	m_bIsSeleTbl = gbIsSelTblPrice;
	m_bIsReplaceUnit = gbIsReplaceUnit;
	m_bAutoPaint120 = gbAutoPaint120;
	m_bWithoutProtectionCost = gbWithoutProtectionCost;
	m_bInnerByEconomic = gbInnerByEconomic;

	//}}AFX_DATA_INIT
}


void CDlgOption::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgOption)
	DDX_Control(pDX, IDC_PROGRESS1, m_ReplaceProgress);
	DDX_Check(pDX, IDC_CHECK1, m_IsColseCompress);
	DDX_Check(pDX, IDC_CHECK2, m_IsReplOld);
	DDX_Check(pDX, IDC_CHECK3, m_IsMoveUpdate);
	DDX_Check(pDX, IDC_CHECK_AUTOSELECT_PRE, m_IsAutoSelectPre);
	DDX_Check(pDX, IDC_JUDGEMENT_HEAT_LOSS, m_bIsHeatLoss);
	DDX_Check(pDX, IDC_AUTO_ADD_VALVE, m_IsAutoAddValve);
	DDX_Text(pDX, IDC_EDIT_TEMP_RATIO, m_dMaxRatio);
	DDX_Text(pDX, IDC_EDIT_FACE_MAX_TEMP, m_dSurfaceMaxTemp);
	DDX_Check(pDX, IDC_CHECK_SELECT_ENG, m_bIsRunSelectEng);
	DDX_Check(pDX, IDC_CHECK_PRICE_RATIO, m_bIsSeleTbl);
	DDX_Check(pDX, IDC_CHECK_IMPORT_REPLACE_UNIT, m_bIsReplaceUnit);
	DDX_Check(pDX, IDC_CHECK_AUTOPAINT_120, m_bAutoPaint120);
	DDX_Check(pDX, IDC_CHECK_WITHOUT_PROTECTION_COST, m_bWithoutProtectionCost);
	DDX_Check(pDX, IDC_CHECK_INNER_BY_ECONOMIC, m_bInnerByEconomic);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgOption, CDialog)
	//{{AFX_MSG_MAP(CDlgOption)
	ON_BN_CLICKED(IDC_REPLACE_OLDNAME_NEWNAME, OnReplaceOldnameNewname)
	ON_BN_CLICKED(IDC_REPLACE_CURRENT_OLDNAME_NEWNAME, OnReplaceCurrentOldnameNewname)
	ON_BN_CLICKED(IDC_RADIO_HAND_SET, OnRadioHandSet)
	ON_BN_CLICKED(IDC_RADIO_RENEW_SET, OnRadioRenewSet)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgOption message handlers

void CDlgOption::OnOK() 
{
	UpdateData(TRUE);
	long nCheckedNO;

 	//重新选择保温结构。
	nCheckedNO = GetCheckedRadioButton(IDC_RADIO_HAND_SET,IDC_RADIO_RENEW_SET);
	if( nCheckedNO != 0 )
	{
		giInsulStruct = nCheckedNO - IDC_RADIO_HAND_SET;
	}
	//计算水和蒸汽性质的标准
	nCheckedNO = GetCheckedRadioButton( IDC_RADIO_97, IDC_RADIO_67 );
	if ( nCheckedNO != 0 )
	{
		giCalSteanProp = nCheckedNO - IDC_RADIO_97;
	}

	CHeatPreCal::SetConstantDefine("ConstantDefine","FaceMaxTemp",m_dSurfaceMaxTemp);
	CHeatPreCal::SetConstantDefine("ConstantDefine","Ratio_MaxHotTemp",m_dMaxRatio);

	bIsCloseCompress= m_IsColseCompress;    //关闭时是否压缩数据库 
	bIsReplaceOld	= m_IsReplOld;        	//打开时是否自动替换旧材料名称
	bIsMoveUpdate	= m_IsMoveUpdate;       //编辑原始数据移动记录时是否更新
	bIsAutoSelectPre= m_IsAutoSelectPre;    //编辑原始数据移动记录时自动选择保温材料
	bIsHeatLoss		= m_bIsHeatLoss;		//计算时判断最大散热密度
	bIsAutoAddValve	= m_IsAutoAddValve; 	//编辑原始数据移动记录时在管道上自动增加阀门
	gbIsRunSelectEng= m_bIsRunSelectEng;	//启动时是否弹出选择工程卷册对话框
	gbIsSelTblPrice	= m_bIsSeleTbl;			//编辑原始数据，自动时取表中的热价比主汽价
	gbIsReplaceUnit = m_bIsReplaceUnit;		// 导入保温油漆数据时替换单位
	gbAutoPaint120  = m_bAutoPaint120;		// 统计油漆时，自动加上保温数据介质温度小于120度的记录。
	gbWithoutProtectionCost  = m_bWithoutProtectionCost;		// 计算经济厚度时不包含保护层材料费用，默认为0-包含
	gbInnerByEconomic  = m_bInnerByEconomic;		// 双层异材保温计算经济厚度时内层不按表面温度法计算，默认为0-按表面温度法计算
	//将选项框中的状态值写入到注册表
	theApp.WriteRegedit();

	CDialog::OnOK();
}

/////////////////////////////////////////////////
//
// 响应"替换所有数据库的旧材料的名称"按钮
//
void CDlgOption::OnReplaceOldnameNewname() 
{
	CWaitCursor WaitCursor;
	CMaterialName MaterialName;
	CMaterialName::tagReplaceStruct ReplaceInfo[6];
	int ProgressPos=1;
	int i;
	int pos=1;
	//显示进度条
	m_ReplaceProgress.ShowWindow(SW_NORMAL);

	m_ReplaceProgress.SetRange(1,10);
	m_ReplaceProgress.SetPos(ProgressPos);

	ReplaceInfo[pos].strTableName=_T("E_PREEXP");
	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

//	ReplaceInfo[pos].strTableName=_T("EDIT");				// 该表为一个临时的,不需要处理. [7/8/2005]
//	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

	ReplaceInfo[pos].strTableName=_T("PRE_CALC");
	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

//	ReplaceInfo[pos].strTableName=_T("PRE_TAB");
//	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

	ProgressPos++;
	m_ReplaceProgress.SetPos(ProgressPos);
 
	try
	{
		MaterialName.SetNameRefRecordset(theApp.m_pConMaterial,_T("Material_OLD_NEW"));
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
	}

	try
	{
		for(i=1;i<pos;i++)
		{
			MaterialName.ReplaceNameOldToNew(theApp.m_pConnection,ReplaceInfo[i]);

			ProgressPos++;
			m_ReplaceProgress.SetPos(ProgressPos);
		}
	}
	catch(_com_error &e)
	{
		CString str;

		str.Format(_T("%s表的%d字段出错"),ReplaceInfo[i].strTableName,ReplaceInfo[i].pstrFieldsName);

		Exception::SetAdditiveInfo(str);

		ReportExceptionErrorLV1(e);
	}

//	ReplaceInfo[0].strTableName=_T("A_C09");
//	ReplaceInfo[0].pstrFieldsName=_T("PIPE_MAT");

	ReplaceInfo[0].strTableName=_T("A_PIPE");
	ReplaceInfo[0].pstrFieldsName=_T("CL1");

	try
	{
		for(int i=0;i<1;i++)
		{
			MaterialName.ReplaceNameOldToNew(theApp.m_pConnectionCODE,ReplaceInfo[i]);

			ProgressPos++;
			m_ReplaceProgress.SetPos(ProgressPos);
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
	}
			ProgressPos++;
			m_ReplaceProgress.SetPos(ProgressPos);

	//隐藏进度条
	m_ReplaceProgress.SetPos(10);
	Sleep(10);
	m_ReplaceProgress.ShowWindow(SW_HIDE);
}

/////////////////////////////////////////////////
//
// 响应"替换当前工程的旧材料的名称"按钮
//
void CDlgOption::OnReplaceCurrentOldnameNewname() 
{
	CWaitCursor WaitCursor;
	CMaterialName MaterialName;
	_RecordsetPtr IRecordset;
	CString SQL;
	CMaterialName::tagReplaceStruct ReplaceInfo[6];
	HRESULT hr;
	int ProgressPos=0;
	int pos=1;
	hr=IRecordset.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		ReportExceptionErrorLV1(_com_error(hr));
		return;
	}

	//显示进度条
	m_ReplaceProgress.ShowWindow(SW_NORMAL);

	m_ReplaceProgress.SetRange(1,4);
	m_ReplaceProgress.SetPos(ProgressPos);

	ReplaceInfo[pos].strTableName=_T("E_PREEXP");
	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

//	ReplaceInfo[pos].strTableName=_T("EDIT");
//	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

	ReplaceInfo[pos].strTableName=_T("PRE_CALC");
	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

//	ReplaceInfo[pos].strTableName=_T("PRE_TAB");
//	ReplaceInfo[pos++].pstrFieldsName=_T("C_PI_MAT");

	try
	{
		MaterialName.SetNameRefRecordset(theApp.m_pConMaterial,_T("Material_OLD_NEW"));
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
	}


	try
	{
		for(int i=1;i<pos;i++)
		{
//			MaterialName.ReplaceNameOldToNew(theApp.m_pConnection,ReplaceInfo[i]);
			SQL.Format(_T("SELECT * FROM %s WHERE EnginID='%s'"),ReplaceInfo[i].strTableName,EDIBgbl::SelPrjID);

			IRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)theApp.m_pConnection),
							adOpenDynamic,adLockOptimistic,adCmdText);

			MaterialName.ReplaceNameOldToNew(IRecordset,ReplaceInfo[i]);

			IRecordset->Close();

			ProgressPos++;
			m_ReplaceProgress.SetPos(ProgressPos);
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
	}
	//隐藏进度条
	m_ReplaceProgress.ShowWindow(SW_HIDE);
}
//////////////////////////////
//选择手工确定的记录
void CDlgOption::OnRadioHandSet() 
{
	CheckRadioButton(IDC_RADIO_HAND_SET,IDC_RADIO_RENEW_SET,IDC_RADIO_HAND_SET);
}
///////////////////////////
//选择重算保温结构
void CDlgOption::OnRadioRenewSet() 
{
	CheckRadioButton(IDC_RADIO_HAND_SET,IDC_RADIO_RENEW_SET,IDC_RADIO_RENEW_SET);
}

BOOL CDlgOption::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	CString	strTmp;		//临时变量

 	//计算保温结构。
	CheckRadioButton(IDC_RADIO_HAND_SET,IDC_RADIO_RENEW_SET,IDC_RADIO_HAND_SET+giInsulStruct );

	//计算水和蒸汽性质的标准
	CheckRadioButton( IDC_RADIO_97, IDC_RADIO_67, IDC_RADIO_97 + giCalSteanProp );

	//复合保温结构界面温度与外保温材料的最高使用温度的比值
	if( !CHeatPreCal::GetConstantDefine("ConstantDefine","Ratio_MaxHotTemp",strTmp) )
	{
		m_dMaxRatio = 0.9;
	}
	else
	{
		m_dMaxRatio = strtod(strTmp, NULL);
	}
	//外表面允许最大温度
	if (!CHeatPreCal::GetConstantDefine("ConstantDefine", "FaceMaxTemp", strTmp))
	{
		m_dSurfaceMaxTemp = 50;
	}
	else
	{
		m_dSurfaceMaxTemp = strtod(strTmp,NULL);
	}

	UpdateData(FALSE);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
