//{{AFX_INCLUDES()
#include "DataGridEx.h"
#include "datagrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_DATASHOWDLG_H__D0159595_D370_4E3C_8092_F04831CD8155__INCLUDED_)
#define AFX_DATASHOWDLG_H__D0159595_D370_4E3C_8092_F04831CD8155__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DataShowDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
//
// 数据显示对话框
//
// 在对话框创建之前，需要调用SetRecordset设置记录集的连接
// 打开记录集之前需  设置CursorLocation = adUseClient;
//
///////////////////////////////////////////////////////////////////////////// 


class CDataShowDlg : public CDialog
{
// Construction
public:
	CDataShowDlg(CWnd* pParent = NULL);   // standard constructor
	~CDataShowDlg();

	enum
	{
		HideFirstColumn=-1,
		HideLastColumn=-2,
		EndHideArray=-4
	};

	struct tag_DefaultFieldValue	//当在DataGrid控件中添加记录时，字段的默认值 
	{
		CString strFieldName;		// 默认字段名
		_variant_t varFieldValue;	// 字段的值
	};

public:
	//设置在DataGrid控件显示数据时，需要隐藏的列throw(COleDispatchException)	
	void SetHideColumns(int *pIndex,DWORD IndexSize);	
	//设置在DataGrid控件显示数据时，需要隐藏的列throw(COleDispatchException)
	void SetHideColumns(LPCTSTR szFields);		
	
	BOOL UpdateBrowWindow();	//如果是重新设置了记录集，则需更新窗口的字段名，标题，默认值等

	void SetAllowUpdate(BOOL IsAllow);				// 设置是否允许更新
	BOOL GetAllowUpdate();							// 返回是否允许更新

	void SetAllowAddNew(BOOL IsAllow);				// 设置是否允许增加记录
	BOOL GetAllowAddNew();							// 返回是否允许新增加记录

	void SetAllowDelete(BOOL IsAllow);				// 设置是否允许删除记录
	BOOL GetAllowDelete();							// 返回是否允许删除记录

	void SetDataGridCaption(LPCTSTR pCaption);		// 设置DataGrid控件的标题
	CString GetDataGridCaption();					// 返回DataGrid控件的标题	

	void SetDlgCaption(LPCTSTR pCaption);			// 设置对话框的标题
	CString GetDlgCaption();						// 返回对话框的标题

	void SetRecordset(_RecordsetPtr &IRecordset, CDataGridEx* pCtlDataGrid=NULL );	// 设置需要显示的记录集
	_RecordsetPtr& GetRecordset();					// 返回需要显示的记录集

	// 设置在添加记录时，字段的默认值
	void SetDefaultFieldValue(tag_DefaultFieldValue *pDefault,DWORD dwLength);

	// 设置在非模态中是否在窗口销毁时是否自动删除对象
	void SetAutoDelete(BOOL bAutoDelete);

	//设置指定的字段赋一个特定的值.
	void SetPassField(CString strField);

protected:
	//隐藏指定的列
	void HideColumns();		//throw(COleDispatchException)

	// 隐藏指定字段名称的列
	void HideColumnsByFieldsName(LPCTSTR szFields); // throw(COleDispatchException);

	//定位光标到指定的列.
	bool FindDataGridCol(short colIndex);

public:
	BOOL SetFieldCaption(CString strTblName);//设置字段的标题和宽度
	void SetCursorCol(short startCol);	//设置光标到指定的列.
	

// Dialog Data
	//{{AFX_DATA(CDataShowDlg)
	enum { IDD = IDD_DATA_SHOW_DLG };
	CDataGridEx	m_ResultBrowseControl;
	CDataGridEx	m_BrowseGrid;

	CDataGridEx* pCurDataGrid;		//当前处理的DataGrid控件.
	//}}AFX_DATA

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDataShowDlg)
	public:
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	short m_nCursorCol;
	CString m_strCol_Pass;
	HICON m_hIcon;
	virtual void OnCancel();
	virtual void OnOK();

	// Generated message map functions
	//{{AFX_MSG(CDataShowDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnOnAddNewResultBrowse();
	afx_msg void OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI);
	afx_msg void OnDestroy();
	afx_msg void OnRowColChangeResultBrowse(VARIANT FAR* LastRow, short LastCol);
	afx_msg void OnClickResultBrowse();
	afx_msg void OnColResizeResultBrowse(short ColIndex, short FAR* Cancel);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	

	BOOL m_IsAutoDelete;			// 在非模态中是否在窗口销毁时自动删除对象
	tag_DefaultFieldValue* m_pDefaultFieldValueStruct;

	CString m_strFieldsNameToHide;	//需要隐藏列的字段名字符串
	int* m_pIndexArrayToHide;		//需要隐藏列的列索引

	BOOL m_IsAllowDel;				// 是否允许删除
	BOOL m_IsAllowAddNew;			// 是否允许新增
	BOOL m_IsAllowUpdate;			// 是否允许更新
	CString m_DataGridCaption;		// DataGrid控件的标题
	CString m_DlgCaption;			// 对话框的标题
	_RecordsetPtr m_IRecordset;		// 需显示的记录集的智能指针

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DATASHOWDLG_H__D0159595_D370_4E3C_8092_F04831CD8155__INCLUDED_)


extern	BOOL GetStructTblName(_ConnectionPtr& pCon, CString strDataTblName,CString& strStructTblName);
extern	BOOL SetFieldCaptionAndWidth(CDataGridEx* pDataGrid, _ConnectionPtr& pConSort, CString strTblName);

extern	BOOL SaveDataGridFieldWidth(CDataGridEx* pDataGrid, _ConnectionPtr& pConSort, CString strTblName);
extern	BOOL GetScreenRatio(float& xRatio, float& yRatio);