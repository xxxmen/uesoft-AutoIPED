//{{AFX_INCLUDES()
#include "datagrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_DLGSELCITYWEATHER_H__A2676A65_BAF9_431C_BF2F_E6ED38F7C9AE__INCLUDED_)
#define AFX_DLGSELCITYWEATHER_H__A2676A65_BAF9_431C_BF2F_E6ED38F7C9AE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgSelCityWeather.h : header file
//
#include "DataGridEx.h"
/////////////////////////////////////////////////////////////////////////////
// CDlgSelCityWeather dialog

class CDlgSelCityWeather : public CDialog
{
// Construction
public:
	static BOOL GetDewPointTemp(_RecordsetPtr pRsPoint, double dTa, double dHumidity, double& dReDwePoint);
	CDlgSelCityWeather(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgSelCityWeather)
	enum { IDD = IDD_SEL_CITY_WEATHER };
	CComboBox	m_ctlCity;
	CComboBox	m_ctlProvince;
	CDataGridEx	m_gridWeatherData;
	CString	m_strStaticCaption;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgSelCityWeather)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	bool GetCurrentCityName();
	CString m_strTempTblName;
	CString m_strSourTblName;
	CString m_strCurCityName;			//当前工程的市名。
	CString m_strCurProvinceName;		//当前工程的省名。

	bool InitDataGridWeather();
	_RecordsetPtr m_pRsWeather;
	bool UpdateDataGridWeather();
	bool OpenWeatherDB_Table(_RecordsetPtr& pRs, CString strFilter);
	_RecordsetPtr m_pRsCity;
	bool UpdateCity();
	_ConnectionPtr m_pConWeather;
	bool ConnectWeatherDB();
	bool InitProvinceCombox();

	// Generated message map functions
	//{{AFX_MSG(CDlgSelCityWeather)
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeComboProvince();
	afx_msg void OnSelchangeComboCity();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGSELCITYWEATHER_H__A2676A65_BAF9_431C_BF2F_E6ED38F7C9AE__INCLUDED_)
