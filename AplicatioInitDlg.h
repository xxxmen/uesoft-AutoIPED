#if !defined(AFX_APLICATIOINITDLG_H__F6E86972_6E3C_4895_B6E4_37DA22C82A90__INCLUDED_)
#define AFX_APLICATIOINITDLG_H__F6E86972_6E3C_4895_B6E4_37DA22C82A90__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AplicatioInitDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// 
// 程序运行时初始化所用的对话框
//
// 主要是完成文件的COPY，从模板目录到各个工作目录。
//
// 当各个工作目录不存在或没有所需的文件时，应该用此
// 对话框。
//
// 当初始条件满足，再调用此对话框时，对话框并不会显示。
//
/////////////////////////////////////////////////////////////////////////////

#define WM_INITIALIZEWORK	(WM_USER+10)

class CAplicatioInitDlg : public CDialog
{
// Construction
public:
	CAplicatioInitDlg(CWnd* pParent = NULL);   // standard constructor

public:
	// CopyFileEx的回调函数
	static DWORD CALLBACK CopyProgressRoutine(
												LARGE_INTEGER TotalFileSize,
												LARGE_INTEGER TotalBytesThransferred,
												LARGE_INTEGER StreamSize,
												LARGE_INTEGER StreamBytesTransferred,
												DWORD dwStreamNumber,
												DWORD dwCallbackReason,
												HANDLE hSourceFile,
												HANDLE hDestinatonFile,
												LPVOID lpData);

protected:
	// 格式化使路径合乎程序要求
	void FormatDirPath(CString &DirPath);

	short UpgradeDB(CString strDestDB, CString strSourceDB);

private:
	// 拷贝文件到各个指定的目录
	BOOL CopyFileToEveryDir();

// Dialog Data
	//{{AFX_DATA(CAplicatioInitDlg)
	enum { IDD = IDD_APPINIT_DLG };
	CProgressCtrl	m_InitProgress;
	CString	m_InitializeInfo;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAplicatioInitDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation


protected:
	short UpgradeDatabase();
	CStringList* GetFieldNameList(CDaoTableDef *pTDef);
	BOOL StringInList(CString strTableName, CStringList *strList);
	CStringList* GetTableNameList(CDaoDatabase *pDB);
	// Generated message map functions
	//{{AFX_MSG(CAplicatioInitDlg)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	afx_msg LRESULT OnInitializeWorkFunc(WPARAM wParam,LPARAM lParam);
	DECLARE_MESSAGE_MAP()
private:
	CString ReadCommonMdbPath();

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_APLICATIOINITDLG_H__F6E86972_6E3C_4895_B6E4_37DA22C82A90__INCLUDED_)
