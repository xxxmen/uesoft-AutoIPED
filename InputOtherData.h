// InputOtherData.h: interface for the InputOtherData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_INPUTOTHERDATA_H__B79B579A_6C80_41B6_8B3B_D17935461C7C__INCLUDED_)
#define AFX_INPUTOTHERDATA_H__B79B579A_6C80_41B6_8B3B_D17935461C7C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "odbcinst.h"

class InputOtherData  
{
public:
	bool MeetRequirement(CString AllPath);
	bool Have_Table(_ConnectionPtr m_pCon, CString str);
	struct CONNFEATURE {
		CString strCONN;
		CString strEnginTable;
		CString strEnginIDFieldName;
		CString strEnginName;
		CString strFILENAME;
		CString strFilter;
		CString strExtension;
	} strConnFeature[5];
	int cancel;
	CString *Pname;//工程名称
	CString *ProID;//工程代号
	int count;
	bool ImportTable_a_dir(_ConnectionPtr pcon_source,CString strProID);
	bool ImportTable_abnormal(_ConnectionPtr pcon_source,CString strTableName, CString strProID);
	CString strFileName;   //导入的工程路径，带DATA文件夹名
	CString strPathName;  //导入的工程路径，不带工程代号
	CString strPathProID;  //导入的工程代号，带路径
	bool InputData(CString strBWPath,CString strPrjID);    //导入数据
	bool SelectPro();    //选择要导入的工程
	void SelectFile(int);   //选择要导入的工程路径
	InputOtherData();
	virtual ~InputOtherData();

private:
	long GetMaxValue(CString SQL,_ConnectionPtr pConn,CString strFieldName);
	bool PutFieldValue(_RecordsetPtr pSource,_RecordsetPtr pTarget,CString strFieldName);
	int iCONN;

	//更新Volume表，throw(_com_error);
	void UpdateVolumeTable(LPCTSTR szProjectName);
protected:
	bool GetImportTblNames(CString* &pTableName,int &nTblCount,CString pNotTbl[],int nNotCount=0);
	enum nIPEDFilterIndex{
	nDBFileterIndex_CUSTOM=0,//用户自定义过滤器
	nDBFileterIndex_UESOFT=1,
	nDBFileterIndex_VFP_SW=2,
	nDBFileterIndex_DB3_JS=3
	};
};

#endif // !defined(AFX_INPUTOTHERDATA_H__B79B579A_6C80_41B6_8B3B_D17935461C7C__INCLUDED_)
