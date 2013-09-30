// DlgSelCityWeather.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DlgSelCityWeather.h"
#include "EDIBgbl.h"
#include "vtot.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
extern CAutoIPEDApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CDlgSelCityWeather dialog


CDlgSelCityWeather::CDlgSelCityWeather(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgSelCityWeather::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgSelCityWeather)
	m_strStaticCaption = _T("当前工程的气象参数:");
	//}}AFX_DATA_INIT
}


void CDlgSelCityWeather::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgSelCityWeather)
	DDX_Control(pDX, IDC_COMBO_CITY, m_ctlCity);
	DDX_Control(pDX, IDC_COMBO_PROVINCE, m_ctlProvince);
	DDX_Control(pDX, IDC_DATAGRID1, m_gridWeatherData);
	DDX_Text(pDX, IDC_STATIC_WEATHER, m_strStaticCaption);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgSelCityWeather, CDialog)
	//{{AFX_MSG_MAP(CDlgSelCityWeather)
	ON_CBN_SELCHANGE(IDC_COMBO_PROVINCE, OnSelchangeComboProvince)
	ON_CBN_SELCHANGE(IDC_COMBO_CITY, OnSelchangeComboCity)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgSelCityWeather message handlers
//初始加入所有的省城
bool CDlgSelCityWeather::InitProvinceCombox()
{
	try
	{
		_RecordsetPtr pRsProvince;
		CString strSQL, strTmp;
		pRsProvince.CreateInstance(__uuidof(Recordset));
		
		strSQL = "SELECT * FROM [省索引表] ORDER BY ProvinceIndex";
		pRsProvince->Open(_variant_t(strSQL), m_pConWeather.GetInterfacePtr(), 
					adOpenStatic, adLockOptimistic, adCmdText);
		
		m_ctlProvince.ResetContent();
		for(; !pRsProvince->adoEOF; pRsProvince->MoveNext() )
		{
			strTmp = vtos( pRsProvince->GetCollect(_variant_t("Province")));
			if( !strTmp.IsEmpty() )
			{
				m_ctlProvince.AddString(strTmp);
			}
		}
		if(	!m_strCurProvinceName.IsEmpty() )
		{	//初始定位，从常量库中取出的省名。
			int nPos = m_ctlProvince.FindString(-1, m_strCurProvinceName);
			if( nPos != -1 )
			{
				m_ctlProvince.SetCurSel(nPos);
			}
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}

//连接气象数据库
//用一个成员变量连接。
bool CDlgSelCityWeather::ConnectWeatherDB()
{
	if(	m_pConWeather == NULL )
	{
		m_pConWeather.CreateInstance(__uuidof(Connection));
	}
	try
	{
		CString strSQL, strDBName;
		strDBName = "中国城市气象参数.mdb";
		strSQL = CONNECTSTRING	+ EDIBgbl::sCritPath + strDBName;

		m_pConWeather->Open(_bstr_t(strSQL), "", "", -1 );
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}

	return true;
}

//
//根据当前的省，列出对应的市.
bool CDlgSelCityWeather::UpdateCity()
{
	try
	{
		if( m_pRsCity->State == adStateClosed )
		{	
			//没有打开数据库。
			//重新打开。
			if( !OpenWeatherDB_Table(m_pRsCity, "中国城市气象参数") )
			{
				return false;
			}
		}
		if(m_pRsCity->GetRecordCount() <= 0)
		{
			//没有记录
			return false;
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	CString strProv, strSQL, strCity;
	int nIndex = m_ctlProvince.GetCurSel();
	if( nIndex == -1 )
	{
		return false;
	}
	m_ctlProvince.GetLBText( nIndex, strProv);
	if( strProv.IsEmpty() )
	{
		return false;
	}
	try
	{
		m_ctlCity.ResetContent();
		m_pRsCity->Filter = _variant_t("");
		//过滤其它的。保留当前省的记录。
		strSQL = "Province='"+strProv+"' ";
		m_pRsCity->Filter = _variant_t(strSQL);

		for( ; !m_pRsCity->adoEOF; m_pRsCity->MoveNext() )
		{
			strCity = vtos(m_pRsCity->GetCollect(_variant_t("City")));
			if( !strCity.IsEmpty() )
			{
				m_ctlCity.AddString(strCity);
			}
		}
		if( !m_strCurCityName.IsEmpty() )
		{	
			nIndex = m_ctlCity.FindString(-1, m_strCurCityName);
			if( -1 != nIndex )
			{
				m_ctlCity.SetCurSel(nIndex);
			}
			m_strCurCityName.Empty();
		}
		else
		{
			//初始第一条。
			m_ctlCity.SetCurSel(0);
		}
		//取消过滤
		m_pRsCity->Filter = _variant_t("");
	}
	catch(_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return false;
	}	
	return true;
}

//
//打开气象库中的某个表。
bool CDlgSelCityWeather::OpenWeatherDB_Table(_RecordsetPtr &pRs, CString strFilter)
{
	CString strSQL;
	try
	{
		if( pRs == NULL )
		{
			pRs.CreateInstance(__uuidof(Recordset));
		}
		strSQL = "SELECT * FROM ["+strFilter+"] ";

		if(	pRs->State == adStateOpen )
		{
			pRs->Close();
		}
		pRs->Open(_variant_t(strSQL), m_pConWeather.GetInterfacePtr(), 
			adOpenStatic, adLockOptimistic, adCmdText);

	}
	catch(_com_error& e)
	{
		AfxMessageBox( e.Description());
		return false;
	}
	return true;
}

BOOL CDlgSelCityWeather::OnInitDialog() 
{
	CDialog::OnInitDialog();

	//连接气象数据库
	if( !ConnectWeatherDB() )
	{
		return false;
	}
	//首先获得常量库中的省名，
	GetCurrentCityName();
	//初始省组合框。
	InitProvinceCombox();
	//初始气象表。
	if( !InitDataGridWeather() )
	{
		return false;
	}
	//	打开表。
	OpenWeatherDB_Table( m_pRsCity, "中国城市气象参数");

	//更新城市组合框.
	if( UpdateCity() )
	{
	//	UpdateDataGridWeather();
	}


	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

//
//选择不同的省／自治区
void CDlgSelCityWeather::OnSelchangeComboProvince() 
{
	//更新城市
	if( UpdateCity() )
	{
		UpdateDataGridWeather();
	}
}

//根据选择的城市更新对应气象数据
bool CDlgSelCityWeather::UpdateDataGridWeather()
{
	CString strProvince, strCity;
	int nIndex;
	//当前的省
	nIndex = m_ctlProvince.GetCurSel();
	if( nIndex == -1 )
	{
		return false;
	}
	m_ctlProvince.GetLBText(nIndex, strProvince);
	//城市
	nIndex = m_ctlCity.GetCurSel();
	if( nIndex == -1 )
	{
		return false;
	}
	m_ctlCity.GetLBText(nIndex, strCity);

	//
	CString strSQL, strTmp;
	_variant_t var;

	try
	{
		strSQL = "Province='"+strProvince+"' and City='"+strCity+"' ";
		m_pRsCity->Filter = _variant_t(strSQL);
		if( !m_pRsCity->adoEOF && m_pRsWeather->GetRecordCount() > 0 )
		{
			//将当前城市的气象参数取出来。
			//对表中取出与城市的气象参数中对应的字段名
			for(m_pRsWeather->MoveFirst(); !m_pRsWeather->adoEOF; m_pRsWeather->MoveNext() )
			{
				try
				{
					strTmp = vtos( m_pRsWeather->GetCollect(_variant_t("对应气象参数表中的字段名")));
					if( !strTmp.IsEmpty() )
					{
						var = m_pRsCity->GetCollect(_variant_t(strTmp));
						strTmp = vtos(var);
						m_pRsWeather->PutCollect(_variant_t("值"), _variant_t(strTmp));
					}
					else
					{
						//如果为空。将作另外的处理。
					}

				}
				catch(_com_error)
				{	
				}
			}
			//改变提示语名。
			m_strStaticCaption = "选择城市的气象参数:";
			UpdateData(FALSE);
		}
		m_pRsCity->Filter = _variant_t("");
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}
//
//选择不同的城市
void CDlgSelCityWeather::OnSelchangeComboCity() 
{
	this->UpdateDataGridWeather();	
}

//初始当前工程的常数
bool CDlgSelCityWeather::InitDataGridWeather()
{
	CString strSQL;
	m_strSourTblName = "Ta_Variable";	//存储气象数据的表.
	m_strTempTblName = "temp_table";	//临时表名.
	try
	{
		if( m_pRsWeather == NULL )
		{
			m_pRsWeather.CreateInstance(__uuidof(Recordset));
			m_pRsWeather->CursorLocation = adUseClient;
		}
		//复制Ta_Variable中,当前工程的气象记录到一个临时表.
		try
		{
			theApp.m_pConnectionCODE->Execute(_bstr_t("DROP TABLE "+m_strTempTblName+" "), NULL, adCmdText);
		}
		catch(_com_error& e)
		{
			if( e.Error() != DB_E_NOTABLE )
			{
				return false;
			}
		}
		strSQL = "SELECT * INTO "+m_strTempTblName+" FROM ["+m_strSourTblName+"] ";
		theApp.m_pConnectionCODE->Execute(_bstr_t(strSQL), NULL, adCmdText);

		//启动事务
//		theApp.m_pConnectionCODE->BeginTrans();

		strSQL = "SELECT Index AS 序号, Ta_Variable_DESC AS 名称, Ta_Variable_VALUE AS 值 ,Ta_Variable_FieldName AS 对应气象参数表中的字段名 FROM ["+m_strTempTblName+"] ";
		if( m_pRsWeather->State == adStateOpen )
		{
			m_pRsWeather->Close();
		}
		m_pRsWeather->Open(_variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(), 
					adOpenStatic, adLockOptimistic, adCmdText);

		m_gridWeatherData.SetRefDataSource( m_pRsWeather->GetDataSource());

		//放弃事务
//		theApp.m_pConnectionCODE->RollbackTrans();
		//提交事务
//		theApp.m_pConnectionCODE->CommitTrans();
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;	
}
 
//确定,
//保存当前的设置.
void CDlgSelCityWeather::OnOK() 
{
	CString strSQL;
	try
	{
		try
		{
			
			if( m_gridWeatherData.GetDataSource() != NULL )
			{	//自动移动行,将最后写入的数据提交到数据库中,
				//解决,改变了DataGrid控件的内容但没有换行就退出的情况.
				m_gridWeatherData.SetBookmark(_variant_t((long)1));	
			}
		}
		catch(_com_error)
		{
		}
		//将工程式原来的记录删除.
		strSQL = "DELETE * FROM ["+m_strSourTblName+"] ";
		theApp.m_pConnectionCODE->Execute(_bstr_t(strSQL), NULL, adCmdText);
		//将临时表中的数据插入到对应的工程表中.
		strSQL = "INSERT INTO ["+m_strSourTblName+"] SELECT * FROM ["+m_strTempTblName+"]";
		theApp.m_pConnectionCODE->Execute(_bstr_t(strSQL), NULL, adCmdText);
		//删除临时表。
		strSQL = "DROP TABLE ["+m_strTempTblName+"] ";
		theApp.m_pConnectionCODE->Execute(_bstr_t(strSQL), NULL, adCmdText);
		//将城市名称写入到常数库中。
		CString strProvince, strCity;
		int nCity = m_ctlCity.GetCurSel(), nProvince = m_ctlProvince.GetCurSel();
		if( -1 != nCity && -1 != nProvince )
		{
			m_ctlCity.GetLBText(nCity, strCity);
			m_ctlProvince.GetLBText(nProvince, strProvince);

			strSQL = "UPDATE [Engin] SET 省='"+strProvince+"',市='"+strCity+"' WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
			theApp.m_pConAllPrj->Execute(_bstr_t(strSQL), NULL, adCmdText);
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
	}
	CDialog::OnOK();
}

//功能：
//从常数库中获得省市名称。
bool CDlgSelCityWeather::GetCurrentCityName()
{
	CString strSQL;
	_RecordsetPtr pRsDef;
	pRsDef.CreateInstance(__uuidof(Recordset));
	try
	{
		strSQL = "SELECT * FROM [Engin] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
		pRsDef->Open(_variant_t(strSQL), theApp.m_pConAllPrj.GetInterfacePtr(),
				adOpenStatic, adLockOptimistic, adCmdText);
		if( pRsDef->GetRecordCount() > 0 )
		{
			m_strCurProvinceName = vtos(pRsDef->GetCollect(_variant_t("省")));
			m_strCurCityName	 = vtos(pRsDef->GetCollect(_variant_t("市")));
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}
//------------------------------------------------------------------
// DATE         : [2006/03/21]
// Author       : ZSY
// Parameter(s) : dTa[in] 夏季空气调节室外计算干球温度(环境温度) 
//				: dHumidity[in] 最热月平均相对湿度
//				: pRsPoint[in] 如果传入的记录已经连接则用之，否则重新打开
//				: dReDwePoint[out] 查表获得的露点温度
// Return       : 查表成功返回TRUE 否则返回FALS
// Remark       : 根据夏季空气调节室外计算干球温度和最热月平均相对湿度查表获得露点温度
//------------------------------------------------------------------
BOOL CDlgSelCityWeather::GetDewPointTemp(_RecordsetPtr pRsPoint, double dTa, double dHumidity, double &dReDwePoint)
{
	try
	{
		CString strSQL;	// SQL语句
		if (pRsPoint == NULL || pRsPoint->GetState() == adStateClosed)
		{
			// 重新再打开
			pRsPoint.CreateInstance(__uuidof(Recordset));
			strSQL = "SELECT * FROM [DewPoint]";
			pRsPoint->Open(_variant_t(strSQL), theApp.m_pConWeather.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		}
		if (pRsPoint->adoEOF && pRsPoint->BOF)
		{// 记录为空
			return FALSE;
		}
		strSQL.Format("dTa=%f", dTa);
//		pRsPoint->Find(_bstr_t(strSQL), 0, adSearchForward);
		pRsPoint->PutFilter((long)adFilterNone);
		pRsPoint->PutFilter(_variant_t(strSQL));

		BOOL bFirst = TRUE;	// 没有进行循环 
		double dPreVal = 0;	// 前一次的值
		double dPreKey = 0;	// 前一次的关键字
		double dCurVal;		// 当前的值
		double dCurKey;		// 当前的关键字

		for (; !pRsPoint->adoEOF; pRsPoint->MoveNext())
		{
			dCurVal = vtof(pRsPoint->GetCollect(_variant_t("Val")));
			dCurKey = vtof(pRsPoint->GetCollect(_variant_t("Key")));
			
			if (dCurKey >= dHumidity)
			{
				if (bFirst)
					dReDwePoint = dCurVal;	// 须要查找的关键字小于或等于表中最小的关键值
				else
					dReDwePoint = (dCurVal - dPreVal) / (dCurKey - dPreKey) * (dHumidity - dPreKey) + dPreVal;
				
				return TRUE;
			}
			bFirst = FALSE;
			dPreKey = dCurKey;
			dPreVal = dCurVal;
		}
		if ( bFirst == FALSE && pRsPoint->adoEOF)	// 查找的关键字超过了记录集中最大的
		{
			dReDwePoint = dPreVal;	// 使用最大的值
		}

	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
		
	}
	return TRUE;
}
