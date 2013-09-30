#if !defined(AFX_IMPORTFROMXLSDLG_H__B50DF418_09BE_4479_9AD3_1EB20C186C24__INCLUDED_)
#define AFX_IMPORTFROMXLSDLG_H__B50DF418_09BE_4479_9AD3_1EB20C186C24__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ImportFromXLSDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// 
// 从外部导入原始数据对话框的基类
//
// 调用次对话框前需要调用SetProjectID，SetProjectDbConnect
// 设置工程的编号与连接
//
/////////////////////////////////////////////////////////////////////////
class _Application;

#include "MyButton.h"
#include "PropertyWnd.h"	// 属性窗口类
#include "excel9.h"
#include "edibgbl.h"


class CImportFromXLSDlg : public CDialog
{
// Construction
public:
	CImportFromXLSDlg(CWnd* pParent = NULL);   // standard constructor
	~CImportFromXLSDlg();

public:
	// 从数据库中取出默认的Excel列号
	void SetDefaultExcelNO(LPCTSTR lpItemTblName);
	
	void SetDataTblName(LPCTSTR lpDataTblName);
	// 设置子控件的标题和 控制子控件的表名
	void SetItemTblName(LPCTSTR lpTitleName, LPCTSTR lpTblName);
	_ConnectionPtr GetIPEDSortDBConnect();
	void SetIPEDSortDBConnect(_ConnectionPtr pConSort);
	CPropertyWnd* GetPropertyWnd();
	CString GetHintInformation();
	void SetHintInformation(LPCTSTR szHint);

	CString GetRegSubKey();
	void SetRegSubKey(LPCTSTR szRegSubkey);

	_ConnectionPtr GetProjectDbConnect();
	void SetProjectDbConnect(_ConnectionPtr IConnection);

	CString GetProjectID();
	void SetProjectID(LPCTSTR szProjectID);


// Dialog Data
	//{{AFX_DATA(CImportFromXLSDlg)
	enum { IDD = IDD_IMPORT_FROM_XLS };
	CMyButton	m_OpenFileDlgButton;
	CString	m_XLSFilePath;
	CString	m_HintInformation;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CImportFromXLSDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

protected:
	// 将信息保存到注册表内
	virtual void SaveSetToReg();
	// 从注册表获得数据初始化
	virtual void InitFromReg();
	
	// 初始化属性窗口
	virtual BOOL InitPropertyWnd();

	// 在导入时验证输入数据的有效性
	virtual BOOL ValidateData();

	// 执行导入数据的操作
	virtual void BeginImport();
		
	// 将当前对应的Excel列号,写入到数据库中
	BOOL WriteExcelNoToTbl(const CString strTblName);

	// 从数据表中取出属性名，动态增加到窗口中
	BOOL AddPropertyWndFromTbl(LPCTSTR lpTitleName, const CString strTblName);

protected:
	CPropertyWnd m_PropertyWnd;		// 属性窗口控件

private:
	// 从IPEDsort.mdb的PD2IPED表中取出EXCEL中的字段名和ACCESS中的字段的对应值
	int GetField2Caption(EDIBgbl::CAPTION2FIELD *&pFieldStruct, const CString strTblName, const CString strDefField = "ExcelColNO");

	CString m_strProjectID;				// 工程的编号
	_ConnectionPtr m_ProjectDbConnect;	// 工程相关的数据库
	_ConnectionPtr m_pConIpedSort;		// 属性控制数据库

	CString	m_RegSubKey;	// 用于存放信息的子键名称
	CString m_strTitleName;	// 子控件的标题
	CString m_strItemTblName;	// 控制子控件的表名
	CString m_strDataTblName;	// 导入的目的数据表的表名
	CString m_strPrecFilePath;

protected:

	// Generated message map functions
	//{{AFX_MSG(CImportFromXLSDlg)
	virtual BOOL OnInitDialog();
	virtual afx_msg void OnOpenFiledlg();
	afx_msg void OnDestroy();
	virtual afx_msg void OnBeginImport();
	afx_msg void OnSetDefaultNo();
	//}}AFX_MSG
	virtual afx_msg	void OnSelectChange(NMHDR *pNMHDR,LRESULT *pResult);
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_IMPORTFROMXLSDLG_H__B50DF418_09BE_4479_9AD3_1EB20C186C24__INCLUDED_)
