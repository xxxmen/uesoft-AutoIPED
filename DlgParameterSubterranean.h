#if !defined(AFX_DLGPARAMETERSUBTERRANEAN_H__91E34FAA_AB90_4032_8926_12260348C5E3__INCLUDED_)
#define AFX_DLGPARAMETERSUBTERRANEAN_H__91E34FAA_AB90_4032_8926_12260348C5E3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgParameterSubterranean.h : header file
//
#include "XDialog.h" 

/////////////////////////////////////////////////////////////////////////////
// CDlgParameterSubterranean dialog

class CDlgParameterSubterranean : public CXDialog 
{
// Construction 
public:  
	void PutCurrentNO(long nCurNO);
	short PutDataToSubterraneanDB(long nID=-1);	// 将对话框中的数据写入到数据库中
	void SetProjectID(LPCTSTR lpstrProjectID);	// 设置当前的工程ID号
	short GetDataToSubterraneanControl(long nID);		// 从数据库中读取原始数据 
	void SetCurrentProjectConnect(_ConnectionPtr pCon);	//  设置当前防冻原始表的数据库连接

	void SetCurrentRes(_RecordsetPtr pRsCur);			// 设置数据记录集

	CDlgParameterSubterranean(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgParameterSubterranean)
	enum { IDD = IDD_PARAMETER_SUBTERRANEAN };
	CComboBox	m_ctlPipeCount;
	CComboBox	m_comboPipeLay;
	double	m_dPipeSubDepth;
	double	m_dEdaphicTemp;
	double	m_dPipeSpan;
	double	m_dEdaphicLamda;
	int		m_nPipeLay;
	double	m_dSupportK;
	double	m_dPipeDag;
	double	m_dCanalDepth;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgParameterSubterranean)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CDlgParameterSubterranean)
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeComboLay();
	afx_msg BOOL OnHelpInfo(HELPINFO* pHelpInfo);
	afx_msg void OnSelchangeComboPipeCount();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	CToolTipCtrl m_pctlTip;
private:
	void UpdateControlState();
	void InitPipeLayList();
	void InitToolTip();	// 初始输入框的提示信息

	long	m_nID;		// 当前记录的序号
	CString m_strProjectID;
	_RecordsetPtr m_pRsSubterranean;
	_ConnectionPtr m_pConSubter;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPARAMETERSUBTERRANEAN_H__91E34FAA_AB90_4032_8926_12260348C5E3__INCLUDED_)
