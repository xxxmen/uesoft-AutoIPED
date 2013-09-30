// ProjectOperate.h: interface for the CProjectOperate class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_PROJECTOPERATE_H__C62D5FD8_0560_4AAB_8D04_051D6425E8AC__INCLUDED_)
#define AFX_PROJECTOPERATE_H__C62D5FD8_0560_4AAB_8D04_051D6425E8AC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

////////////////////////////////////////////////
//
// 对工程的操作类
//
////////////////////////////////////////////////

class CProjectOperate  
{
public:
	CProjectOperate();
	virtual ~CProjectOperate();

	// 从EXCEL文件导入数据的子结构
	struct ImportFromXLSElement		
	{
		CString SourceFieldName;	// Excel的起始列号
		int		ExcelColNum;		// Excel的列数.
		CString DestinationName;	// 对应到目的数据库的字段名
	};
	typedef struct
	{
		int ID;
		CString EnginID;
	}LIST_ID_ENGIN;

	// 描述EXCEL文件导如数据的结构
	struct ImportFromXLSStruct
	{
		CString XLSFilePath;		// EXCEL文件的路径
		CString	XLSTableName;		// EXCEL的工作薄名
		int		BeginRow;			// 开始导入的行号
		int		RowCount;			// 需要的导入的记录数

		CString ProjectDBTableName;	// 导入到的目的工程表名
		CString ProjectID;			// 导入到的目的工程名

		ImportFromXLSElement *pElement;	// 指向子单元的结构
		int ElementNum;	//子单元的结构数目
	};

	class CImportFromXLSStruct:public ImportFromXLSStruct
	{
	public:
		ImportFromXLSElement* FindInElement(LPCTSTR szSource,LPCTSTR szDest);

	};

public:
	_ConnectionPtr GetProjectDbConnect();
	void SetProjectDbConnect(_ConnectionPtr IConnection);

	CString GetProjectID();
	void SetProjectID(LPCTSTR strProjectID);

public:
	bool ImportOriginalData(CString strDataPath, CString strManageDB, int nIndex);
	bool ConnectExcelFile(CString strFileName);
	bool ImportAutoPHSPaint(CString strFilePath);
	bool ImportTblEnginIDXLS(ImportFromXLSStruct *pImportStruct,BOOL IsAutoMakeID,LPCTSTR KeyField=_T("ID"), CString EnginID=_T("EnginID"));
	void ImportA_DirFromXLS(ImportFromXLSStruct *pImportStruct);
	void ImportPre_CalcFromXLS(ImportFromXLSStruct *pImportStruct,BOOL IsAutoMakeID,LPCTSTR KeyField=_T("ID"));
	void ImportFromXLS(ImportFromXLSStruct *pImportStruct);


private:
	bool m_bAuto;			//自动导入标示
	CString m_strProjectID;				// 工程的编号
	_ConnectionPtr m_ProjectDbConnect;	// 工程相关的数据库

protected:
	bool GetFieldAsCol(_ConnectionPtr& pCon,ImportFromXLSElement element[],int& elementNum,CString& strTblName);
	bool ConnectDB(_ConnectionPtr& pCon,CString& strManageTbl,CString& strExcelTbl,CString& strDataTbl,int& nStart,int& nRowNum, CString strDBFile,int nIndex);
	CDatabase m_db;
};

#endif // !defined(AFX_PROJECTOPERATE_H__C62D5FD8_0560_4AAB_8D04_051D6425E8AC__INCLUDED_)
