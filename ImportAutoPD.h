// ImportAutoPD.h: interface for the CImportAutoPD class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_IMPORTAUTOPD_H__9A4548CC_8F91_4CF8_93E5_C38DC77400D1__INCLUDED_)
#define AFX_IMPORTAUTOPD_H__9A4548CC_8F91_4CF8_93E5_C38DC77400D1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "edibgbl.h"

class CImportAutoPD  
{
public:

	typedef struct STRUCT_ENG_ID
	{
		int ID;
		CString EnginID;
	};
 
	int GetField2Caption(EDIBgbl::CAPTION2FIELD* &pFieldStruct, const BOOL bFlag=0); 
	BOOL ImportExcelToAccess(const CString strExcelFileName,const CString strWorksheetName,const _ConnectionPtr pConDes,const CString strTblName,const CString strCurProID,const CString KeyFieldName="ID",const CString ProFieldName="EnginID");
	CImportAutoPD();
	virtual ~CImportAutoPD();

	//导入PD中的保温数据
	bool ImportAutoPDInsul( CString strCommandLine );

	//更新指定工程的记录, 由已知的字段,算出其它字段的值.
	static bool RefreshData(CString EnginID);


protected:
	BOOL OpenExcelTable(_RecordsetPtr& pRsTbl, CString& strSheetName, CString strExcelFileName);
	BOOL ConnectExcelFile(const CString strExcelName, _ConnectionPtr &pConExcel);

	BOOL AddValve();
private:
	_ConnectionPtr m_pConExcel;
	BOOL UpdateImportAmount(const CString strExcelFileName,CString strWorksheetName,const _ConnectionPtr pConDes,const CString strTblName,const CString strCurProID,CString strCurVol,long nExcelRecCount,const CString KeyFieldName="ID",const CString ProFieldName="EnginID");
};

#endif // !defined(AFX_IMPORTAUTOPD_H__9A4548CC_8F91_4CF8_93E5_C38DC77400D1__INCLUDED_)
