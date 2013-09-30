#if !defined(AFX_DLGPARAMETERCONGEAL_H__6293FA5A_FFE5_4794_A9C9_875143CB8825__INCLUDED_)
#define AFX_DLGPARAMETERCONGEAL_H__6293FA5A_FFE5_4794_A9C9_875143CB8825__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgParameterCongeal.h : header file
//
#include "XDialog.h"

/////////////////////////////////////////////////////////////////////////////
// CDlgParameterCongeal dialog

class CDlgParameterCongeal : public CXDialog
{
// Construction
public:
	short PutDataToPreventCongealDB( long nID = -1 );
	void SetCurrentRes(_RecordsetPtr pCurRs);
	void SetProjectID(LPCTSTR ProjectID);
	void SetCurrentProjectConnect(_ConnectionPtr pCon);
	CDlgParameterCongeal(CWnd* pParent = NULL);   // standard constructor
	short GetDataToPreventCongealControl( long nID );	// 从防冻记录集中取出数据到控件

// Dialog Data 
	//{{AFX_DATA(CDlgParameterCongeal)
	enum { IDD = IDD_PARAMETER_CONGEAL };
	CComboBox	m_ctlHeatLossData;
	double	m_dFlux;
	double	m_dMediumCongealTemp;
	double	m_dMediumHeat;
	double	m_dMediumMeltHeat;
	double	m_dPipeDensity;
	double	m_dPipeHeat;
	double	m_dStopTime;
	double	m_dMediumDensity;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgParameterCongeal)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL PreTranslateMessage(MSG* pMsg) ;

	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CDlgParameterCongeal)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	CToolTipCtrl m_pctlTip;
private:
	void InitToolTip();
	void OpenCongealDBTbl();
	_ConnectionPtr m_pConCongeal;	// 当前防冻原始表的数据库连接
	_RecordsetPtr m_pRsCongealData;	// 防冻计算的记录集
	CString	m_strProjectID;		// 当前工程的ID
	long	m_ID;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPARAMETERCONGEAL_H__6293FA5A_FFE5_4794_A9C9_875143CB8825__INCLUDED_)
