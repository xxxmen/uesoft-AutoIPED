// ImportAutoPD.cpp: implementation of the CImportAutoPD class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "ImportAutoPD.h"

#include "vtot.h"
#include "edibgbl.h"
//#include "newproject.h"
#include "projectoperate.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

extern CAutoIPEDApp theApp;
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CImportAutoPD::CImportAutoPD()
{

}

CImportAutoPD::~CImportAutoPD()
{

}

//导入PD中的保温数据
bool CImportAutoPD::ImportAutoPDInsul( CString strCommandLine )
{
	int nk = 0;				//关键字的个数
	int nStart=0;
	CString  strTemp;
	CString  strFindValue;

	typedef	struct 
	{
		LPCTSTR strKey;		//关键字
		CString strValue;	//对应的参数
	}COMMAND1;
	COMMAND1  CommandArr[10];
	CommandArr[nk++].strKey   ="  -DATA_FILENAME ";		 //导入的Excel文件路径

	CommandArr[nk++].strKey   = "  -PROJECT_ID ";			 //工程ID

	CommandArr[nk++].strKey   = "  -PROJECT_ENGVOL ";		  //卷册ID

	CommandArr[nk++].strKey   = "  -RECORD_COUNT ";		//原始数据的行数.

	//解析命令行。
	for(int k=0; k<nk; k++)
	{
		strFindValue = CommandArr[k].strKey;
		nStart = strCommandLine.Find(CommandArr[k].strKey);
		if( nStart == -1 )
		{
			continue;
		}
		else
		{
			nStart += strFindValue.GetLength();
		}
		strTemp = strCommandLine.Mid( nStart );
		strTemp.TrimLeft(" ");
		if( k == nk-1 )
		{
			strTemp.TrimRight(" ");
			CommandArr[k].strValue = strTemp;
		}
		else
		{
			strTemp = strTemp.Mid(0,strTemp.Find(" -")+1);
			strTemp.TrimRight(" ");
			CommandArr[k].strValue = strTemp;
		}
	}
	if( !FileExists(CommandArr[0].strValue) )
	{
		//给定的文件不存在.
		AfxMessageBox("找不到文件 "+CommandArr[0].strValue+"  , 请确认是否存在 !");
		return false;
	}

	//{{如果是新工程, 则初始其它相应的表.
// 	newproject dlg;
// 	dlg.m_engin =EDIBgbl::SelPrjName  = "autoiped";
// 	dlg.m_eng_code = EDIBgbl::SelPrjID =CommandArr[1].strValue; 
// 	EDIBgbl::SelVlmCODE = CommandArr[2].strValue;
// 	dlg.SetAuto( true );
// 	dlg.insertitem();
	theApp.InitializeProjectTable( CommandArr[1].strValue, CommandArr[2].strValue );
	//}}

	//导入保温原始数据到表[pre_calc]
	CProjectOperate   project;
	long	nDataRowCount;

	CProjectOperate::ImportFromXLSStruct ImportTable;
	CProjectOperate::ImportFromXLSElement ImportTableItem[22];
	struct
	{
		LPCTSTR NameCh;		// pre_calc的字段名的中文名称 保温原始数据
		LPCTSTR NameField;	// pre_calc的字段名
	}FieldsName[]=
	{
	//	{_T("序号"),			_T("c_num")},		//	序号要特殊处理
		{_T("卷册号"),			_T("c_vol")},
		{_T("色环"),			_T("c_color")},
		{_T("管道/设备名称"),	_T("c_name1")},
		{_T("管道外径/规格"),	_T("c_size")},
		{_T("管道壁厚"),		_T("c_pi_thi")},
		{_T("介质温度"),		_T("c_temp")},
		{_T("安装地点"),		_T("c_place")},		// 只存放地点， [2005/06/02] 
//		{_T("安装地点"),		_T("c_p_s")},
		{_T("管道材质"),		_T("c_pi_mat")},
		{_T("备注"),			_T("c_mark")},
		{_T("内保温层材料名称"),_T("c_name_in")},
		{_T("外保温层材料名称"),_T("c_name2")},
		{_T("外保温层厚度"),    _T("c_pre_thi")},
		{_T("内保温层厚度"),    _T("c_in_thi")},
		{_T("保护层材料名称"),	_T("c_name3")},
		{_T("风速"),			_T("c_wind")},
		{_T("热价比主汽价"),	_T("c_price")},
		{_T("年运行小时数"),	_T("c_hour")},
		{_T("油管道保温厚"),	_T("c_pre_thi")},
		{_T("数量"),			_T("c_amount")},
		{_T("管内压力"),		_T("c_Pressure")},
		{_T("管内介质"),		_T("c_Medium")},
		
	//	{_T("工程代号"),		_T("EnginID")},
	};	
	for( int i=0, n=sizeof(FieldsName)/sizeof(FieldsName[0]); i<n; i++ )
	{
		ImportTableItem[i].DestinationName =  FieldsName[i].NameField;
		ImportTableItem[i].SourceFieldName =  'B' + i;
	}
	ImportTable.ProjectDBTableName = "pre_calc";
	ImportTable.ProjectID = EDIBgbl::SelPrjID; 
	ImportTable.XLSFilePath = CommandArr[0].strValue;
	ImportTable.XLSTableName = "原始数据";
	ImportTable.BeginRow    = 2;
	ImportTable.RowCount  = nDataRowCount = _tcstol( CommandArr[3].strValue, NULL, 10);

	if( ImportTable.RowCount < 1 )
	{
		AfxMessageBox("原始数据为空!");
		return false;
	}
	ImportTable.pElement = ImportTableItem;
	ImportTable.ElementNum = n;
	project.SetProjectDbConnect( theApp.m_pConnection );

	//自动从Excel中导入数据到表中.
//	project.ImportTblEnginIDXLS(&ImportTable, TRUE );

	//从EXCEL中导入数据到ACCESS不考虑以前导入相同类型的数据.
//	this->ImportExcelToAccess(CommandArr[0].strValue,"原始数据",theApp.m_pConnection, "pre_calc",EDIBgbl::SelPrjID);

	//test zsy 2005.11.5
	//导入数据，考虑相同类型的数据，
	UpdateImportAmount(CommandArr[0].strValue,"原始数据",theApp.m_pConnection, "pre_calc",EDIBgbl::SelPrjID,EDIBgbl::SelVlmCODE, nDataRowCount);

	//对原始数据进一步处理.()
	RefreshData( EDIBgbl::SelPrjID );
 
	return true;
}

//更新指定工程的记录, 由已知的字段,算出其它字段的值.
bool CImportAutoPD::RefreshData(CString EnginID)
{
	if( EnginID.IsEmpty() )
	{ 
		return false;
	}
	CString strSQL;
	_RecordsetPtr  pRs;

	_RecordsetPtr  pRs_Weather;		//参数表.
	pRs_Weather.CreateInstance(__uuidof(Recordset));
	HRESULT hr;
	double  e_rate, con_temp1, con_temp2, e_hours, e_wind, Pro_thi;
	double	e_wind_in;	//室内风速

	double   dWeatherTemp=0;							//在气象表中的环境温度值
	try
	{
		strSQL = "SELECT * FROM [Ta_Variable] ";
		pRs_Weather->Open(_variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(),
						adOpenStatic, adLockOptimistic, adCmdText);
		if (pRs_Weather->GetRecordCount() > 0)
		{
			//环境温度的索引值 为0
			strSQL = "Index=0";
			pRs_Weather->Find(_bstr_t(strSQL), 0, adSearchForward); 
			if (!pRs_Weather->adoEOF)
			{
				//气象参数表中的"室内环境温度"值
				dWeatherTemp = vtof(pRs_Weather->GetCollect(_variant_t("Ta_Variable_VALUE")) );					
			}
		}

		hr = pRs.CreateInstance(__uuidof(Recordset));
		if( FAILED(hr) )
		{
			return false;
		}
		strSQL = "SELECT * FROM [a_config] WHERE EnginID='"+EnginID+"' ";
		if( pRs->State == adStateOpen )
		{
			pRs->Close();
		}
		pRs->PutCursorLocation(adUseClient);
		pRs->Open(_variant_t(strSQL), (IDispatch*)theApp.m_pConnection,
					adOpenStatic, adLockOptimistic, adCmdText);

		if( pRs->GetRecordCount() < 1 )
		{	
			e_rate = con_temp1 = con_temp2 = e_hours = e_wind = 0;
		}
		else
		{
			//取出[a_config]表中的值,
			pRs->MoveFirst();
			e_rate    = vtof( pRs->GetCollect("年费用率") );
			con_temp1 = vtof( pRs->GetCollect("室内温度") );
			con_temp2 = vtof( pRs->GetCollect("室外温度") );
			e_hours   = vtof( pRs->GetCollect("年运行小时") );
			e_wind    = vtof( pRs->GetCollect("室外风速") );
		}
		e_wind_in = e_wind;    //室内风速为室外风速
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	//如果气象库中的环境温度没有值时，则使用常量库中的环境温度
	con_temp1 = (dWeatherTemp <=0)? con_temp1:dWeatherTemp;

	strSQL = "SELECT * FROM [pre_calc] WHERE EnginID='"+EnginID+"' ";
	try
	{
		if (pRs->GetState() == adStateOpen)
		{
			pRs->Close();
		}
		
		pRs->Open(_variant_t(strSQL), theApp.m_pConnection.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);

		if( pRs->GetRecordCount() <= 0 )
		{	//没有记录时
			return true;
		}
		CString tempStr, place, steam;
		pRs->MoveFirst();//
		//循环.对一些字段进行处理.
		//如:
		// 安装地点(c_p_s)字段值为(室外汽管) 将其分为"室外"另存到c_place;
		//											 "汽管"存到c_steam
		while( !pRs->adoEOF )
		{
			//年费用率
			pRs->PutCollect(_variant_t("c_rate"), _variant_t(e_rate));
			//年运行小时
			pRs->PutCollect(_variant_t("c_hour"), _variant_t(e_hours));

			tempStr = vtos( pRs->GetCollect("c_name3") );
			if( !tempStr.IsEmpty() )
			{
				if( tempStr.Find("(") == -1 )
				{
					Pro_thi = 20.0;
				}
				else
				{
					tempStr = tempStr.Mid( tempStr.Find("(")+1 );
					Pro_thi = strtod(tempStr, NULL);
				}
				pRs->PutCollect(_variant_t("c_pro_thi"), _variant_t(Pro_thi) );
			}
			//处理c_place字段
			tempStr = vtos(pRs->GetCollect("c_place"));		// 自动导入PD中的数据时，对安装地点进行处理。 [2005/06/02] 
			tempStr.TrimLeft(); 
			tempStr.TrimRight();
			tempStr.MakeUpper();

			if( !tempStr.IsEmpty() )
			{
				place = tempStr.Left(4);
				tempStr = tempStr.Mid(4);
				steam = tempStr.Right(4);

				//安装地点字段不存储汽管或水管类型
//				pRs->PutCollect(_variant_t("c_place"), _variant_t(place));
//				pRs->PutCollect(_variant_t("c_steam"), _variant_t(steam));

				if( -1 != place.Find("室内") )
				{
					pRs->PutCollect(_variant_t("c_con_temp"), _variant_t(con_temp1));
					//室内风速为0
					pRs->PutCollect(_variant_t("c_wind"), _variant_t(e_wind_in));
				}
				else
				{
					pRs->PutCollect(_variant_t("c_con_temp"), _variant_t(con_temp2));
					//室外风速
					pRs->PutCollect(_variant_t("c_wind"), _variant_t(e_wind));
				}
			}
			pRs->Update();
			pRs->MoveNext();
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
// DATE         :[2005/07/14]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :增加阀门
//------------------------------------------------------------------
BOOL CImportAutoPD::AddValve()
{
	_RecordsetPtr pRs_ValveType;
	CString	strSQL;						//SQL语句
	int		nValveIndex;				//阀门类型的索引
	CString	strValveName;				//阀门名称
	pRs_ValveType.CreateInstance(__uuidof(Recordset));
 	pRs_ValveType->PutCursorLocation(adUseClient);

	try
	{
		strSQL = "SELECT * FROM [ValveType] ";		//打开阀门类型表
		pRs_ValveType->Open(_variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(),
			adOpenStatic, adLockOptimistic, adCmdText);
		if (pRs_ValveType->GetRecordCount() <= 0)
		{
			return FALSE;
		}
		for (; !pRs_ValveType->adoEOF; pRs_ValveType->MoveNext())
		{

			nValveIndex  = vtoi(pRs_ValveType->GetCollect(_variant_t("ValveTypeID")));
			strValveName = vtos(pRs_ValveType->GetCollect(_variant_t("TypeName")));
			if (strValveName.IsEmpty())
			{
				continue;
			}
			if (FileExists(strValveName))
			{
				CFile::Remove(strValveName);
			}		
		}
	}
	catch (_com_error)
	{
		return FALSE;
	}

	return TRUE;
}
 
//------------------------------------------------------------------
// DATE         :[2005/09/30]
// Parameter(s) :
// Return       :
// Remark       :将EXCEL文件导出到ACCESS数据库表中
//------------------------------------------------------------------
BOOL CImportAutoPD::ImportExcelToAccess(CString strExcelFileName, CString strWorksheetName, _ConnectionPtr pConDes, CString strTblName,CString strCurProID,CString KeyFieldName,CString ProFieldName)
{
	CString		strSQL;					//SQL语句
	_ConnectionPtr pConExcel;			//连接EXCEL文件
	_RecordsetPtr  pRsExcel;
	_RecordsetPtr  pRsAccess;
	_RecordsetPtr  pRsTmp;
	pConExcel.CreateInstance(__uuidof(Connection));
	pRsExcel.CreateInstance(__uuidof(Recordset));
	pRsAccess.CreateInstance(__uuidof(Recordset));
	pRsTmp.CreateInstance(__uuidof(Recordset));
	STRUCT_ENG_ID listID[10];		//不同工程的最大序号和工程名称
	int	 nProCount=0;		//不同工程的个数
	CString strTemp;
	int Rj=0;				//工程号在Excel中的列
	int ProjNum=1;
	int ProjIndex=0;
	int	ID;

	if (NULL == pConDes)
	{
		return FALSE;
	}
	try
	{
		EDIBgbl::CAPTION2FIELD* pFieldStruct=NULL;
		_variant_t varTmp;
		//获得EXCEL中的字段名和ACCESS中的字段的对应值,返回字段个数
		int nFieldCount = GetField2Caption(pFieldStruct);
		if (nFieldCount <= 0)
		{
			return FALSE;
		}
		
		strSQL = CONNECTSTRING_EXCEL + strExcelFileName;//连接EXCEL文件
		pConExcel->Open(_bstr_t(strSQL), "", "", -1);

		//打开Excel工作表，加一个符号$如果出错再重试不加$再打开一次.
		strSQL = "SELECT * FROM ["+strWorksheetName+"$]";
		try
		{
			pRsExcel->Open(_variant_t(strSQL), pConExcel.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		}
		catch (_com_error)
		{
			strSQL = "SELECT * FROM ["+strWorksheetName+"]";
			pRsExcel->Open(_variant_t(strSQL), pConExcel.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		}

		if (pRsExcel->adoEOF && pRsExcel->BOF)
		{
			AfxMessageBox("文件中没有记录！");
			return FALSE;
		}
		//打开Access表
		strSQL = "SELECT * FROM ["+strTblName+"] WHERE "+ProFieldName+"='"+strCurProID+"' ORDER BY "+KeyFieldName+" ";
		
		pRsAccess->Open(_variant_t(strSQL), pConDes.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		
		if (pRsAccess->adoEOF && pRsAccess->BOF)
		{
			ID=0;
		}
		else
		{
			pRsAccess->MoveLast();
			ID = vtoi(pRsAccess->GetCollect(_variant_t(KeyFieldName)));
		}
		ID++;
		listID[0].ID = ID;		//当前工程最大的序号
		listID[0].EnginID = strCurProID;

		for(;Rj < nFieldCount;Rj++)
		{
			if( !ProFieldName.CompareNoCase(pFieldStruct[Rj].strField) )
			{
				break;
			}
		}
		_RecordsetPtr pRsTmp;
		pRsTmp.CreateInstance(__uuidof(Recordset));

		CString strGroup;		//作为关键字的所有字段组合
		//strGroup = "[管道外径/规格], [管道壁厚], [管内压力], [介质温度], [管道材质]";

/*		strGroup = "[卷册号],[色环],[管道/设备名称],[管道外径/规格],[管道壁厚],[介质温度],[安装地点],[管道材质],[备注],\
					[内保温层材料名称],[外保温层材料名称],[保护层材料名称],[风速],[热价比主汽价],[年运行小时数],\
					[油管道保温厚],[管内压力],[管内介质],[工程代号]";
		
		strSQL = "SELECT "+strGroup+" FROM ["+strTblName+"] WHERE "+ProFieldName+"="+strCurProID+" ";
		pRsTmp->Open(_variant_t(strSQL), pConDes.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		if (pRsTmp->adoEOF && pRsTmp->BOF)
		{
			variant_t var = pRsTmp->GetCollect(_variant_t("enginid"));
		}
*/		//复制数据到ACCESS数据库中。 
		while (!pRsExcel->adoEOF)
		{
			pRsAccess->AddNew();
			for (int i=0; i < nFieldCount; i++)
			{
				try
				{
					if (pFieldStruct[i].strCaption.IsEmpty())
					{
						continue;
					}
					varTmp = pRsExcel->GetCollect(_variant_t(pFieldStruct[i].strCaption));
				}catch (_com_error& e) 
				{
					if (e.Error() == -2146825023)
					{
						strTemp = "原始数据表中没有字段 '"+pFieldStruct[i].strCaption+"' 。";
						AfxMessageBox(strTemp);
						pFieldStruct[i].strCaption.Empty();
						continue;
					}
					AfxMessageBox(e.Description());
					return FALSE;
				}
				
				if (Rj < nFieldCount && i == Rj)	//工程名称的字段
				{
					strTemp = vtos(varTmp);
					for(int c=0; c<ProjNum; c++)
					{
						if( !strTemp.CompareNoCase(listID[c].EnginID) || (c==0 && strTemp.IsEmpty()))
						{
							ProjIndex = c;
							break;
						}
					}
					if( c == ProjNum )		//不同的工程找出最大的序号.
					{
						listID[c].EnginID = strTemp;
						strSQL = "SELECT * FROM ["+strTblName+"] \
							WHERE "+ProFieldName+"='"+listID[c].EnginID+"' ORDER BY "+KeyFieldName+" ";
						if(pRsTmp->State == adStateOpen)
							pRsTmp->Close();
						pRsTmp->Open(_variant_t(strSQL), pConDes.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
						if( pRsTmp->adoEOF && pRsTmp->BOF )
						{
							ID = 0;
						}
						else
						{
							pRsTmp->MoveLast();
							varTmp = pRsTmp->GetCollect(_variant_t(KeyFieldName));
							ID = vtoi(varTmp);
						}
						//最大的序号,将记录加到末尾.
						listID[c].ID = ++ID;
						ProjIndex = c;
						ProjNum++; 
					}
				}
				
				pRsAccess->PutCollect(_variant_t(pFieldStruct[i].strField),varTmp);
			}
			//设置关键值和工程ID
			pRsAccess->PutCollect(_variant_t(KeyFieldName), _variant_t((long)listID[ProjIndex].ID++));
			pRsAccess->PutCollect(_variant_t(ProFieldName),_variant_t(listID[ProjIndex].EnginID));
			pRsAccess->PutCollect(_variant_t("c_bImport"),_variant_t((short)1));	//导入标志
			pRsAccess->PutCollect(_variant_t("c_CalInThi"),_variant_t((short)1));	//是否自动计算内保温的厚度 1:不计算
			pRsAccess->PutCollect(_variant_t("c_CalPreThi"),_variant_t((short)1));	//是否自动计算外保温的厚度


			pRsAccess->Update();
			pRsExcel->MoveNext();
		}

		if (NULL != pFieldStruct)
		{
			delete [] pFieldStruct;
		}
	}
	catch (_com_error& e) 
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}

	return TRUE;
}
                                              
//-------------------  -----------------------------------------------
// DATE         :[2005/10/08]
// Parameter(s) :EDIBgbl::CAPTION2FIELD* &pFieldStruct	从数据库中取出的字段对应关系存于结构数组中。
//				：const BOOL bFlag				取字段的过滤条件的标记 默认为0，为1时用于导入PD数据的关键字段。
// Return       :返回字段个数
// Remark       :从IPEDsort.mdb的PD2IPED表中取出EXCEL中的字段名和ACCESS中的字段的对应值
//------------------------------------------------------------------
int CImportAutoPD::GetField2Caption(EDIBgbl::CAPTION2FIELD* &pFieldStruct, const BOOL bFlag)
{
	_RecordsetPtr pRs;
	CString strSQL;
	int		nFieldCount;	//字段的个数    
	pRs.CreateInstance(__uuidof(Recordset));
	try
	{
		switch(bFlag) {
		case 0:strSQL = "SELECT * FROM [PD2IPED]";
			break;
		case 1:strSQL = "SELECT * FROM [PD2IPED] WHERE CADFieldSeq IS NOT NULL";
			break;
		default:strSQL = "SELECT * FROM [PD2IPED]";
		}
		
		pRs->Open(_variant_t(strSQL), theApp.m_pConIPEDsort.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		nFieldCount = pRs->GetRecordCount();
		if (nFieldCount <= 0)
		{
			return FALSE;
		} 
		pFieldStruct = new EDIBgbl::CAPTION2FIELD[nFieldCount];
		for (int i=0; !pRs->adoEOF && i < nFieldCount; pRs->MoveNext(),i++)
		{
			pFieldStruct[i].strCaption = vtos(pRs->GetCollect("LocalCaption"));		//EXCEL中的字段描述
			pFieldStruct[i].strField   = vtos(pRs->GetCollect("FieldName"));		//ACCESS中的字段名称
			pFieldStruct[i].nFieldType = vtoi(pRs->GetCollect("FieldType"));		//字段类型
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return nFieldCount;
}

//------------------------------------------------------------------
// DATE         :[2005/10/27]
// Parameter(s) :
// Return       :
// Remark       :从PD导入数据到IPED时，如果同一种管道增加或减少长度时，则只改变数量字段的值，
//				:其它字段的数据还是用IPED中可能修改了的值。
//------------------------------------------------------------------
/*BOOL CImportAutoPD::UpdateImportAmount1(CString strExcelFileName, CString strWorksheetName,
					_ConnectionPtr pConDes, CString strTblName,CString strCurProID,CString strCurVol,
					long nExcelRecCount,CString KeyFieldName,CString ProFieldName)
*/
BOOL CImportAutoPD::UpdateImportAmount(const CString strExcelFileName, CString strWorksheetName, 
					const _ConnectionPtr pConDes, const CString strTblName, const CString strCurProID,
					CString strCurVol, long nExcelRecCount, const CString KeyFieldName, const CString ProFieldName)

{
	CString strSQL;			//SQL语句
	short	nGroupFieldCount;
	short   nAllFieldCount;
	_ConnectionPtr pConExcel;
	_RecordsetPtr pRsExcel;
	_RecordsetPtr pRsAccess;
	pRsExcel.CreateInstance(__uuidof(Recordset));
	pRsAccess.CreateInstance(__uuidof(Recordset));
	try
	{
		if (pConDes == NULL)
		{
			return FALSE;
		}
		EDIBgbl::CAPTION2FIELD * pAllFieldStruct = NULL;
		EDIBgbl::CAPTION2FIELD * pGroupFieldStruct = NULL;
		nGroupFieldCount = GetField2Caption(pGroupFieldStruct, 1);		//获得组合关键字段对应名称
		nAllFieldCount	 = GetField2Caption(pAllFieldStruct);			//需要导入的所有字段名称
		if (nGroupFieldCount <= 0 || nAllFieldCount <= 0)
		{
			return FALSE;
		}
		
		//连接EXCEL文件
		if ( !OpenExcelTable( pRsExcel, strWorksheetName, strExcelFileName ) )
		{
			return FALSE;
		}
		if (pRsExcel->adoEOF && pRsExcel->BOF)
		{
			return TRUE;
		}
		//打开ACCESS原始数据表
		strSQL = "SELECT * FROM ["+strTblName+"] WHERE "+ProFieldName+"='"+strCurProID+"' AND c_vol LIKE '_"+strCurVol+"_' AND c_bImport=1 ";
		pRsAccess->Open(_variant_t(strSQL), pConDes.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		for (; !pRsAccess->adoEOF; pRsAccess->MoveNext())
		{
			pRsAccess->PutCollect(_variant_t("c_bImport"), _variant_t((short)14));
		}

		CString strAccessVal;
		CString strExcelVal;
		CString strTmp;
		CString strWhere;
		_variant_t varTmp;
		CString strIDField = "ID";
		int			  nAmountFieldCount=1;		//关键字段的个数.
		EDIBgbl::CAPTION2FIELD amountFieldName[1];			//数量字段名
		amountFieldName[0].nFieldType=1;
		amountFieldName[0].strCaption="数量";
		amountFieldName[0].strField="c_amount";
		if (nExcelRecCount <= 0)
		{
			nExcelRecCount = pRsExcel->GetRecordCount();
		}
		//在ACCESS原始表中处理以前导入过的同一类型的记录。
		for (; !pRsExcel->adoEOF && nExcelRecCount > 0; pRsExcel->MoveNext(), nExcelRecCount--)
		{
			strWhere = "";
			for (int i=0; i < nGroupFieldCount; i++)
			{
				strExcelVal = vtos(pRsExcel->GetCollect(_variant_t(pGroupFieldStruct[i].strCaption)));
				switch(pGroupFieldStruct[i].nFieldType)
				{
				case 0:		//字段类型为字符型
					strTmp = "='"+strExcelVal+"'";
					break;
				case 1:		//字段类型为数字型
					strTmp = "="+strExcelVal+" ";
					break;
				default: strTmp="IS NOT NULL";
				}
				if (i < nGroupFieldCount-1)  //并列满足的条件
				{
					strTmp += " AND ";
				}
				strWhere += "[" + pGroupFieldStruct[i].strField + "]" + strTmp;
			}

			try
			{
				pRsAccess->PutFilter(_variant_t((long)adFilterNone));
				pRsAccess->PutFilter(_variant_t(strWhere));
			}catch (_com_error)
			{
				//在EXCEL中判断记录的条数可能出错
				continue;
			}
			//改变相同类型的数量
			if (!(pRsAccess->adoEOF && pRsAccess->BOF))
			{
				for (int j=0; j < nAmountFieldCount; j++)
				{
					varTmp = pRsExcel->GetCollect(_variant_t(amountFieldName[j].strCaption));
					pRsAccess->PutCollect(_variant_t(amountFieldName[j].strField), varTmp);
					pRsAccess->PutCollect(_variant_t("c_bImport"), _variant_t((short)1));
					pRsAccess->Update();
				}
			}else
			{
				pRsAccess->AddNew();
				for (int i=0; i < nAllFieldCount; i++)
				{
					try
					{
						//如果字段名为空时会出错 
						if (pAllFieldStruct[i].strCaption.IsEmpty() || pAllFieldStruct[i].strField.IsEmpty())
						{
							continue;
						}
						varTmp = pRsExcel->GetCollect(_variant_t(pAllFieldStruct[i].strCaption));
					}catch (_com_error& e) 
					{
						if (e.Error() == -2146825023)
						{
							strTmp = "原始数据表中没有字段 '"+pAllFieldStruct[i].strCaption+"' 。";
							AfxMessageBox(strTmp);
							pAllFieldStruct[i].strCaption.Empty();
							continue;
						}
						AfxMessageBox(e.Description());
						return FALSE;
					}
					
					pRsAccess->PutCollect(_variant_t(pAllFieldStruct[i].strField),varTmp);
				}
				//设置关键值和工程ID
				pRsAccess->PutCollect(_variant_t(ProFieldName), _variant_t(strCurProID));
				pRsAccess->PutCollect(_variant_t("c_bImport"),_variant_t((short)1));	//导入标志
				pRsAccess->PutCollect(_variant_t("c_CalInThi"),_variant_t((short)1));	//是否自动计算内保温的厚度 1:不计算
				pRsAccess->PutCollect(_variant_t("c_CalPreThi"),_variant_t((short)1));	//是否自动计算外保温的厚度
				
				
				pRsAccess->Update();
			}//
		}
		//删除IPED中导入的但修改了类型的管道记录
		strSQL = "DELETE FROM ["+strTblName+"] WHERE EnginID='"+strCurProID+"' AND c_vol LIKE '_"+strCurVol+"_' AND c_bImport=14";
		pConDes->Execute(_bstr_t(strSQL),NULL, -1);

		//对当前工程的记录重新编号
		strSQL = "SELECT * FROM ["+strTblName+"] WHERE EnginID='"+strCurProID+"' ORDER BY c_bImport desc, c_vol asc ";
		if (pRsAccess->State == adStateOpen)
		{
			pRsAccess->Close();
		}
		pRsAccess->PutFilter(_variant_t((long)adFilterNone));
		pRsAccess->Open(_variant_t(strSQL), pConDes.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		if (!(pRsAccess->adoEOF && pRsAccess->BOF))
		{
			pRsAccess->MoveFirst();
		}
		for (long NO=1; !pRsAccess->adoEOF; pRsAccess->MoveNext(),NO++)
		{
			pRsAccess->PutCollect(_variant_t(strIDField), _variant_t(NO));
		}
		if (NULL != pGroupFieldStruct)
		{
			delete [] pGroupFieldStruct;
		}
		if (NULL != pAllFieldStruct)
		{
			delete [] pAllFieldStruct;
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE;
}


//------------------------------------------------------------------
// DATE         :[2005/12/05]
// Parameter(s) :
// Return       :
// Remark       :用ADO的方式，联接EXCEL文件.
//------------------------------------------------------------------
BOOL CImportAutoPD::ConnectExcelFile(const CString strExcelName, _ConnectionPtr &pConExcel)
{
	CString strSQL;
	try
	{
		if (pConExcel == NULL)
		{
			pConExcel.CreateInstance(__uuidof(Connection));
		}
		if ( pConExcel->State == adStateOpen )
		{
			pConExcel->Close();
		}

		strSQL = CONNECTSTRING_EXCEL + strExcelName;
		pConExcel->Open(_bstr_t(strSQL), "", "", adCmdUnspecified);
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE; 

}
//------------------------------------------------------------------
// DATE         :[2005/12/05]
// Parameter(s) :CString  strExcelFileName   EXCEL文件全名
//				:CString& strSheetName	     在EXCEL中的须要打开的表名
//				:_RecordsetPtr& pRs			 返回打开表的记录集
// Return       :成功为TRUE，返则返回FALSE
// Remark       :用ADO记录集的方式打开，EXCEL文件中的表。
//------------------------------------------------------------------
BOOL CImportAutoPD::OpenExcelTable(_RecordsetPtr& pRsTbl, CString& strSheetName, CString strExcelFileName)
{ 
	CString strSQL;
	try
	{ 
		if ( m_pConExcel == NULL || m_pConExcel->State == adStateClosed )
		{
			if ( !ConnectExcelFile( strExcelFileName, m_pConExcel ) )
			{
				return FALSE;
			}
		}
		if (pRsTbl == NULL)
		{
			pRsTbl.CreateInstance(__uuidof(Recordset));
		}
		try
		{
			if (pRsTbl->State == adStateOpen)
			{
				pRsTbl->Close();
			}
			strSQL = "SELECT * FROM ["+strSheetName+"$] ";		//在EXCEL中的表名加上$
			pRsTbl->Open(_variant_t(strSQL), m_pConExcel.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
			strSheetName += "$";								//在调用该函数之后的程序中作为EXCEL表名
		}
		catch (_com_error)
		{
			strSQL = "SELECT * FROM ["+strSheetName+"] ";
			pRsTbl->Open(_variant_t(strSQL), m_pConExcel.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		} 
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE;
}
