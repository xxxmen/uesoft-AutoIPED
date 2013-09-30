// EDIBgbl.cpp: implementation of the EDIBgbl class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "EDIBgbl.h"
#include "AutoIPED.h"
#include "vtot.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
extern CAutoIPEDApp theApp;

CString EDIBgbl::SelPrjID = "";  //工程代号
CString EDIBgbl::SelPrjName = "";//工程名称
CString EDIBgbl::SelVlmCODE = "";//卷册代号

long EDIBgbl::SelSpecID = 3;  //热机
long EDIBgbl::SelDsgnID = 4;  //施工图
long EDIBgbl::SelHyID =0;	  //电力
long EDIBgbl::SelVlmID = 0;
const double EDIBgbl::kgf2N= 9.80665;

CString EDIBgbl::sMaterialPath;
CString EDIBgbl::sCur_MaterialCodeName;
CString EDIBgbl::sCritPath;
CString EDIBgbl::sCur_CritDbName;
CString EDIBgbl::sCur_CodeName;
long EDIBgbl::iCur_CodeKey;
CString EDIBgbl::sCur_CodeNO;

//CString EDIBgbl::strCur_MediumDBPath;	//介质数据库的路径

_RecordsetPtr EDIBgbl::pRsCalc;		//指向原始数据表[pre_calc]的指针。

EDIBgbl::EDIBgbl()
{

}

EDIBgbl::~EDIBgbl()
{

}

bool EDIBgbl::GetCurDBName()
{
	CString strTmpPath = gsIPEDInsDir + "IPED.ini";
	if( !FileExists( strTmpPath ) )
	{
		//信息文件不存在，设置默认的数据库名
		WritePrivateProfileString("specification", "lastCriterion","IPEDcode.mdb", strTmpPath);
		WritePrivateProfileString("specification", "lastCriterionName","电力行业", strTmpPath);//行业
		WritePrivateProfileString("specification", "lastMaterial","Material.mdb", strTmpPath);
		WritePrivateProfileString("specification", "lastCodeNO", "DL/T5054-1996", strTmpPath ); // 规范代号
		WritePrivateProfileString("specification", "lastCodeKey","", strTmpPath ); // 规范标志号
	}
	char cszRet[0x100];
	memset(cszRet, '\0', sizeof(char)*256);  // 缓冲区清零
	//获得前一次选择的规范规范数据库名称
	GetPrivateProfileString("specification","lastCriterion","***.!!!",cszRet,256,strTmpPath);
	EDIBgbl::sCur_CritDbName = cszRet;

	memset(cszRet, '\0', sizeof(char)*256);
	GetPrivateProfileString("specification","lastCriterionName","",cszRet,256,strTmpPath);
	EDIBgbl::sCur_CodeName = cszRet;

	memset(cszRet, '\0', sizeof(char)*256);
	//获得前一次选择的材料数据库名称
	GetPrivateProfileString( "specification","MaterialCodeName","***.!!!",cszRet,256,strTmpPath );
	EDIBgbl::sCur_MaterialCodeName = cszRet;

	memset( cszRet, '\0', sizeof(char)*256 );
	GetPrivateProfileString( "specification", "lastCodeNO", "", cszRet, 256, strTmpPath );
	EDIBgbl::sCur_CodeNO = cszRet;

	EDIBgbl::iCur_CodeKey =GetPrivateProfileInt("specification", "lastCodeKey", nKey_CODE_DL5072_1997,  strTmpPath );
	
	//读取失败，设置缺省值
	if( EDIBgbl::sCur_CritDbName=="" || EDIBgbl::sCur_CritDbName=="***.!!!" )
	{
		EDIBgbl::sCur_CritDbName = "IPEDcode.mdb";
	}
	if( EDIBgbl::sCur_MaterialCodeName=="" || EDIBgbl::sCur_MaterialCodeName=="***.!!!" )
	{
		EDIBgbl::sCur_MaterialCodeName = "ALL CODE";
	}

	return true;
}
bool EDIBgbl::SetCurDBName()
{
	CString strTmpPath;
	strTmpPath = gsIPEDInsDir + "IPED.ini";
	//保存选择的标准库名称
	if( EDIBgbl::sCur_CritDbName != "" )
	{
		// 规范数据库名称
		WritePrivateProfileString("specification", "lastCriterion",	EDIBgbl::sCur_CritDbName,strTmpPath);
		
		// 中文行业名称
		WritePrivateProfileString("specification", "lastCriterionName",	EDIBgbl::sCur_CodeName, strTmpPath);

		// 规范代号 
		WritePrivateProfileString( _T("specification"), _T("lastCodeNO"), EDIBgbl::sCur_CodeNO, strTmpPath );

		// 规范标志号
		CString strTmp;
		strTmp.Format("%d",EDIBgbl::iCur_CodeKey);
		WritePrivateProfileString("specification", "lastCodeKey",strTmp, strTmpPath ); 
	}
	else
	{

		AfxMessageBox("部分文件被破坏,请重新安装 AutoIPED8.0 !");
		return false;
	}
	//保存选择的材料库名称
	if( EDIBgbl::sCur_MaterialCodeName != "" )
	{
		WritePrivateProfileString("specification", "MaterialCodeName",
			EDIBgbl::sCur_MaterialCodeName, strTmpPath);
	}
	else 
	{
		AfxMessageBox("部分文件被破坏,请重新安装 AutoIPED8.0 !");
		return false;
	}
	return true;
}
//
//在数据库pCon中,对指定的表进行操作.将默认的记录插入到指定的工程中
bool EDIBgbl::InsertDefTOEng(_ConnectionPtr &pCon, CString strTblName, CString strPrjID, CString strDefKey)
{
	CString			strSQL;
	CString strTmpTblName;
	strTmpTblName = "ppp";		//临时表名
	try
	{
		try
		{	//删除临时表
			pCon->Execute(_bstr_t("drop table "+strTmpTblName+" "), NULL, adCmdText);
		}
		catch(_com_error& e)
		{
			if( e.Error() != DB_E_NOTABLE )
			{
				ReportExceptionError(e);
				return false;
			}
		}
		//将默认的记录先插入到临时表中.WHERE EnginID='"+strDefKey+"'
		strSQL = "SELECT * INTO ["+strTmpTblName+"] FROM ["+strTblName+"] ";
		if(	strDefKey.IsEmpty() )
		{
			strSQL += "	WHERE EnginID IS NULL ";
		}
		else
		{
			strSQL += " WHERE EnginID='"+strDefKey+"' ";
		}
		pCon->Execute(_bstr_t(strSQL), NULL, adCmdText);

		//改变工程号.
		strSQL = "UPDATE ["+strTmpTblName+"] SET EnginID='"+strPrjID+"' ";
		pCon->Execute(_bstr_t(strSQL), NULL, adCmdText);

		//将其加入到表中,对应了指定的工程.
		strSQL = "INSERT INTO ["+strTblName+"] SELECT * FROM ["+strTmpTblName+"] WHERE EnginID='"+strPrjID+"' ";
		pCon->Execute(_bstr_t(strSQL), NULL, adCmdText);

		//删除临时表.
		pCon->Execute(_bstr_t("drop table "+strTmpTblName+" "), NULL, adCmdText);

	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return false;
	}
	return true;
}

//功能：
//新建一个临时的，运算计算公式的表。
//如果已经存在则先删除。
//该表存于工作项目数据库中。(workprj.mdb)
int EDIBgbl::NewCalcAlfaTable()
{
	CString sSQL;
	CString strTblName = "tmpCalcInsulThick";		//计算放热系数的表名.

	try
	{
		theApp.m_pConWorkPrj->Execute(_bstr_t("DROP TABLE "+strTblName+" "), NULL, adCmdText);
	}
	catch(_com_error& e)
	{	//不存在该表。
		if(e.Error() != DB_E_NOTABLE )
		{
			return 0;
		}
	}
	//创建一个运行计算公式的表；

	sSQL = "CREATE TABLE ["+strTblName+"] (ALPHA double, ALPHAn double, ALPHAc double,\
			Epsilon double, ts double, ta double, W double,	B double,D0 double, D1 double, Dmax double ";
	//以上为计算放热系数要使用到的字段名。
	//下面计算保温厚时新增加的字段。
	sSQL += " , lamda double, Ph double, Ae double, t double,tao double, P1 double, P2 double, P3 double,\
			S double, tb double, D2 double, Q double, pi double, Kr double,	\
			taofr double, tfr double, V double, ro double, C double, Vp double, rop double, \
			Cp	double,	Hfr double, delta double, delta1 double, delta2 double	";
	//计算外表面温度用到的字段名.
	sSQL += ", lamda1 double, lamda2 double 	\
			) ";
	try
	{
		theApp.m_pConWorkPrj->Execute(_bstr_t(sSQL), NULL, adCmdText);
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	return 1;

}

//------------------------------------------------------------------                
// Parameter(s) : strEnginID[in] 指定的工程或所有工程
// Return       : 成功返回TRUE，否则返回FALSE
// Remark       : 替换数据库中一些字段的值
//------------------------------------------------------------------
BOOL EDIBgbl::ReplaceFieldValue(CString strEnginID)
{
	CString strWhereEngin;
	if (strEnginID.IsEmpty())
		strWhereEngin = _T(" AND EnginID IS NOT NULL ");
	else
		strWhereEngin = _T(" AND EnginID='"+strEnginID+"' ");

	CString SQL;
	_RecordsetPtr pRsInfo;
	_RecordsetPtr pRsFormat;
	_RecordsetPtr pRsUnit;
	pRsFormat.CreateInstance(__uuidof(Recordset));
	pRsInfo.CreateInstance(__uuidof(Recordset));
	pRsUnit.CreateInstance(__uuidof(Recordset));
	CString strDataTblName;
	CString strFormatTblName;
	CString strFieldName;
	CString strOldUnit;
	CString strNewUnit;

	try
	{
		SQL = _T("SELECT * FROM [A_UNIT] ");
		pRsUnit->Open(_variant_t(SQL), theApp.m_pConnectionCODE.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		if (pRsUnit->GetadoEOF() && pRsUnit->GetBOF())
		{
			return FALSE;
		}
		SQL = _T("SELECT * FROM [TableINFO] WHERE [ReplaceUnit]=1");
		pRsInfo->Open(_variant_t(SQL), theApp.m_pConIPEDsort.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);		
		if (pRsInfo->GetadoEOF() && pRsInfo->GetBOF())
		{
			return FALSE;
		}

		for (pRsInfo->MoveFirst(); !pRsInfo->GetadoEOF(); pRsInfo->MoveNext())
		{
			strDataTblName	 = vtos(pRsInfo->GetCollect(_variant_t("ManuTBName")));
			strFormatTblName = vtos(pRsInfo->GetCollect(_variant_t("StructTblName")));
			if (strDataTblName.IsEmpty() || strFormatTblName.IsEmpty())
			{
				continue;
			}
			SQL = _T("SELECT * FROM ["+strFormatTblName+"] WHERE [ReplaceVal]=1");
			if (pRsFormat->GetState() == adStateOpen)
				pRsFormat->Close();
			
			pRsFormat->Open(_variant_t(SQL), theApp.m_pConIPEDsort.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
			if (pRsFormat->GetadoEOF() && pRsFormat->GetBOF())
			{
				continue;
			}
			for (pRsFormat->MoveFirst(); !pRsFormat->GetadoEOF(); pRsFormat->MoveNext())
			{
				strFieldName = vtos(pRsFormat->GetCollect(_variant_t("FieldName")));
				if (strFieldName.IsEmpty())
					continue;

				for (pRsUnit->MoveFirst(); !pRsUnit->GetadoEOF(); pRsUnit->MoveNext())
				{
					strOldUnit	 = vtos(pRsUnit->GetCollect(_variant_t("OldUnit")));
					strNewUnit   = vtos(pRsUnit->GetCollect(_variant_t("NewUnit")));
					if (strOldUnit.IsEmpty())
						continue;

					SQL = _T("UPDATE ["+strDataTblName+"] SET "+strFieldName+"= '"+strNewUnit+"' WHERE "+strFieldName+"='"+strOldUnit+"' ");
					SQL += strWhereEngin;
					theApp.m_pConnection->Execute(_bstr_t(SQL), NULL, adCmdText);
				}
			}
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE;
}

CString EDIBgbl::GetShareDbPath()
{
	HKEY hKey;
	char sv[256];
	unsigned long vlen = 256;
	unsigned long nType = REG_SZ;
	CString strDbPath = _T("");
	CString strSubKey = "SOFTWARE\\长沙优易软件开发有限公司\\shareMDB\\";
	if(RegOpenKey(HKEY_LOCAL_MACHINE,strSubKey,&hKey)!=ERROR_SUCCESS)
	{
		return strDbPath;
	}

	if(RegQueryValueEx(hKey,_T("shareMDB"),NULL,&nType,(unsigned char*)sv,&vlen)!=ERROR_SUCCESS)
	{
		RegCloseKey(hKey);
		return strDbPath;
	}

	strDbPath.Format("%s",sv);
	strDbPath.TrimLeft();
	strDbPath.TrimRight();

	if( strDbPath.Right(1) != "\\" )
		strDbPath += "\\";

	RegCloseKey(hKey);

	return strDbPath;
}
