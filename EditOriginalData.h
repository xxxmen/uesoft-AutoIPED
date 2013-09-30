#if !defined(AFX_EDITORIGINALDATA_H__FC6BD122_A151_4D62_A00D_9FB54575FFAC__INCLUDED_)
#define AFX_EDITORIGINALDATA_H__FC6BD122_A151_4D62_A00D_9FB54575FFAC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// EditOriginalData.h : header file
//

#include "MyButton.h"
#include "Foxbase.h"
#include "DlgParameterCongeal.h"
#include "DlgParameterSubterranean.h"
#include "GetPropertyofMaterial2.h"

/////////////////////////////////////////////////////////////////////////////
// CEditOriginalData dialog

//////////////////////////////////////////////
//
// "保温原始数据编辑"对话框
//
// 未用新材料代替旧材料
//
//
// 创建对话框前需要调用
// SetCurrentProjectConnect	设置工程数据库
// SetPublicConnect			设置公共参考数据库
// SetProjectID				设置工程名
//
////////////////////////////////////////////////

class CEditOriginalData : public CDialog,public CFoxBase
{
// Construction 
public:
	CEditOriginalData(CWnd* pParent = NULL);   // standard constructor

public: 
	_ConnectionPtr m_IPtrShareMaterial;
	void UpdateControlSubterranean(int nPipeCount);
	virtual  ~CEditOriginalData();
	BOOL AddValveData();
	bool AutoSelAllMat();
	// 设置在滚动时是否自动选材
	void SetIsAutoSelectMatOnRoll(BOOL bSelect);
	// 返回是否在滚动时自动选材
	BOOL GetIsAutoSelectMatOnRoll();

	//设置所选工程的数据库的连接
	void SetCurrentProjectConnect(_ConnectionPtr &IConnection);
	//返回数据库连接的智能指针的引用
	_ConnectionPtr& GetCurrentProjectConnect();

	//设置公共的数据库的连接
	void SetPublicConnect(_ConnectionPtr& IConnection);
	// 返回公共数据库连接的智能指针的引用
	_ConnectionPtr& GetPublicConnect();

	//设置所选工程的ID号
	void SetProjectID(LPCTSTR ProjectID);
	// 返回所选工程的ID号
	CString& GetProjectID();

// Dialog Data
	//{{AFX_DATA(CEditOriginalData)
	enum { IDD = IDD_EDIT_ORIGINAL_DATA };
	CButton	m_ctlTitleTab;
	CComboBox	m_ctlPipeMedium;
	CComboBox	m_ctlCalcMethod;
	CComboBox	m_ctlHeatTran;		//放热系数取值的索引.
	CComboBox	m_ctlEnvTemp;		//环境温度的取值索引
	CMyButton	m_DeleteAll;
	CComboBox	m_ColorRing;
	CMyButton	m_SaveCurrent;
	CMyButton	m_DelCurrent;
	CMyButton	m_AddNew;
	CMyButton	m_MoveLast;
	CMyButton	m_MoveSubsequent;
	CMyButton	m_MovePrevious;
	CMyButton	m_MoveForefront;
	CComboBox	m_SafeguardMatName;
	CComboBox	m_OutSizeMatName;
	CComboBox	m_InsideMatName;
	CListCtrl	m_HeatPreservationTypeList;
	CListCtrl	m_VolList;
	CComboBox	m_PipeMat;
	CComboBox	m_BuidInPlace;
	CComboBox	m_PipeThick;
	CComboBox	m_PipeSize;
	CComboBox	m_PipeDeviceName;
	CString	m_VolNo;
	int		m_ID;
	CString	m_ReMark;
	BOOL	m_IsVolListEnable;
	BOOL	m_IsHeatPreservationTypeEnable;
	double	m_WideSpeed;
	double	m_PriceRatio;
	double	m_RunPerHour;
	double	m_MediumTemperature;
	BOOL	m_IsUpdateByRoll;
	double	m_Amount;
	BOOL	m_bIsAutoSelectMatOnRoll;
	CString	m_dHeatLoss;
	CString	m_dHeatDensity;
	double	m_dThick1;		//内层保温厚
	double	m_dThick2;		//外层保温厚
	double	m_dTa;			//环境温度
	double	m_dTs;			//外表面温度
	double	m_dPressure;	//管内温度
	BOOL	m_bIsCalInThick; 
	BOOL	m_bIsCalPreThick;
	double	m_dWindPlaimWidth;
	double	m_dBoilerPressure;		//获得锅炉出口过热蒸汽的压力
	double  m_dBoilerTemperature;	//锅炉出口过热蒸汽温度
	double	m_dCoolantTemperature;	//冷却水温度
	//}}AFX_DATA

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEditOriginalData)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	DECLARE_DYNCREATE(CEditOriginalData)

	BOOL InitToolTip();
	BOOL RenewNumberFind(_RecordsetPtr& pRs,long ID,long FindNO);
	short PutDataToPreventCongealDB(long nID=-1);
	short GetDataToPreventCongealControl();
	short ShowWindowRect();
	short AutoChangeValue();
	bool CheckDataValidity();
	short GetSiteIndex(_RecordsetPtr& pRs,short nIndexv, short& SiteIndex);
	short UpdateSite(short SiteIndex);
	short InitPipeMedium();
	short InitCalcMethod();
	int GetConditionTemp(double& dConTemp, int & nMethodIndex,int nIndex=-1);
	bool UpdateHeatTransferCoef(CString strFilter="");			
	bool UpdateEnvironmentComBox(CString strFilter="");		
	virtual void OnCancel(void);

	// Generated message map functions
	//{{AFX_MSG(CEditOriginalData)
	virtual BOOL OnInitDialog();
	afx_msg void OnForefront();
	afx_msg void OnPrevious();
	afx_msg void OnSubsequent();
	afx_msg void OnLast();
	afx_msg void OnAddNew();
	afx_msg void OnDelCurrent();
	afx_msg void OnSaveCurrent();
	afx_msg void OnDropdownColor();
	afx_msg void OnDropdownPipeDeviceName();
	afx_msg void OnDropdownPipeSize();
	afx_msg void OnDropdownPipeThick();
	afx_msg void OnDropdownInsideMatname();
	afx_msg void OnDropdownOutsizeMatname();
	afx_msg void OnDropdownSafeguardMatname(); 
	afx_msg void OnItemchangedHeatPreservationType(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnItemchangedVollist(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnUpdateAllCheckBox();
	afx_msg void OnChangeVolno();
	afx_msg void OnAutoSelectmat();
	afx_msg void OnChangeId();
	afx_msg void OnSortByVol();
	afx_msg void OnExit();
	afx_msg void OnDelAllInsu();
	afx_msg void OnCheckMethod();
	afx_msg void OnCheckTa();
	afx_msg void OnCheckThick1();
	afx_msg void OnCheckThick2();
	afx_msg void OnCheckTs();
	afx_msg void OnSelchangeEnvironTemp();
	afx_msg void OnSelchangeHeatTransferCoef();
	afx_msg void OnCheckEnvTemp();
	afx_msg void OnCheckAlpha();
	afx_msg void OnSelchangeEditMethod();
	afx_msg void OnSelchangeBuildinPlace();
	afx_msg void OnSelchangePipeMedium();
	afx_msg void OnButtonAddValve();
	afx_msg void OnButtonCalcRatio();
	afx_msg void OnSelchangePipeMat();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
		
private:
	BOOL InitCurrentProjectRecordset();

private:
	void AUTO(CString &v,CString &s,CString &t,CString &pre,CString &pro,CString &p,CString &pre_in);
	void cond(_RecordsetPtr &IRecordset, CString &mat,CString &v,CString &p,CString t,CString d,CString h=_T(""));
	void compound(CString &mat_in,CString &mat);
	BOOL Refresh();

	//初始化位图按扭
	BOOL InitBitmapButton();

	//初始化保温类型LIST控件 
	BOOL InitHeatPreservationTypeList();

	//初始化卷册LIST控件 
	BOOL InitVolListControl();

	//从以打开的记录集的当前游标所在的位置取数据到各个控件
	BOOL PutDataToEveryControl();

	//从各个控件取数据到数据库
	BOOL PutDataToDatabaseFromControl(long nID=-1);

	//更新所有组合框的内容
	BOOL UpdateEveryComboBox();

	//更新“色环”组合框
	BOOL UpdateColorComBox();

	//更新“管道设备名称”组合框
	BOOL UpdatePipeDeviceNameComBox();

	//更新“管道外径规格”组合框
	BOOL UpdatePipeSizeComBox();

	//更新“管道壁厚”组合框
	BOOL UpdatePipeThickComBox();

	//更新“安装地点”组合框
	BOOL UpdateBuildPlaceComBox();

	//更新“内保温层材料名称”组合框
	BOOL UpdateInsideMatNameCombox();

	//更新"保护层材料名称"组合框
	BOOL UpdateSafeguardMatNameComBox();

	//更新“外保温层材料名称”组合框
	BOOL UpdateOutSizeMatNameComBox();

	//当在a_vol中找不到对话框中的卷册号时,更新a_vol表
	BOOL UpdateVolTable();

	//更新各个控件的状态
	void UpdateControlsState();


private:
	void UpdateMediumMaterialDensity();
	BOOL NumberSubtractOne(_RecordsetPtr& pRs, long lStartPos);
	void PutDataToSubterraneanDB(long nID=-1);
	void AddDynamicDlg(CDialog *pDlg, UINT nID);
	BOOL GetAeValue(CString strTypeName, double &dAe);
	double m_dValveSpace;
	CDlgParameterCongeal* m_pDlgCongeal; // 保冷的参数输入对话框
	CDlgParameterSubterranean* m_pDlgSubterranean; // 埋地管道保温计算的参数的输入对话框
	CDialog* m_pDlgCurChild;		// 指向当前扩展的一个子对话框。

	GetPropertyofMaterial* m_pMatProp;

	double m_dMaxD0;					//管道的最大外径。
	BOOL InitConstant();
	BOOL m_bIsPlane;					//当前记录的设备是否是平壁
//	BOOL m_bIsAutoSelectMatOnRoll;
	CToolTipCtrl m_pctlTip;				//提示信息.
 
	CString m_ProjectID;							//工程的ID号
	_RecordsetPtr m_ICurrentProjectRecordset;		//原始数据表.
	_RecordsetPtr m_pRsCongealData;					//防冻结计算的原始数据表
	_RecordsetPtr m_pRsSubterranean;				// 埋地管道的原始数据表
	_RecordsetPtr m_pRsEnvironment;					//环境温度取值索引表
	_RecordsetPtr m_pRsHeat;						//放热系数取值索引表
 
	_ConnectionPtr m_IPublicConnection; 				//公共数据库的连接
	_ConnectionPtr m_ICurrentProjectConnection;		//所选工程的数据库的连接

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDITORIGINALDATA_H__FC6BD122_A151_4D62_A00D_9FB54575FFAC__INCLUDED_)
