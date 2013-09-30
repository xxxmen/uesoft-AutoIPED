// AutoIPED.h : main header file for the AUTOIPED application
//

#if !defined(AFX_AUTOIPED_H__8F862085_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
#define AFX_AUTOIPED_H__8F862085_B060_11D7_9BCC_000AE616B8F1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
//全局变量

extern  const _TCHAR szSoftwareKey[];//引用注册表位置；

/////////////////////////////////////////////////////////////////////////////
// CAutoIPEDApp:
// See AutoIPED.cpp for the implementation of this class
//

class CAutoIPEDApp : public CWinApp
{
public:
	CAutoIPEDApp();

public:
	// 判断数据库中是否没有一个工程
	BOOL IsNoOneEngin();

	bool WriteRegedit();

	CString GetRegKey(LPCTSTR pszKey, LPCTSTR pszName,const CString Default);

    void SetRegValue(CString vKey, CString vName, CString vValue);

	BOOL ReadInitKey();

	BOOL InitializeProjectTable( const CString& strPrjID, const CString& strVolCode, 
										  const CString& strPrjName = _T("AutoIPED") );

protected:
	void CopyFromTemplateDirToPrjDir();

public:
    _ConnectionPtr m_pConnection;		// 项目数据库(AutoIPED.mdb)
	_ConnectionPtr m_pConnectionCODE;	// 标准数据库
	_ConnectionPtr m_pConAllPrj;        // 共享的工程数据库,里面包含了选择工程卷册所要用到的表 

	_ConnectionPtr m_pConMaterial;		// 材料库指针.
	_ConnectionPtr m_pConMaterialName;  // 材料名数据库指针
	_ConnectionPtr m_pConIPEDsort;		// 管理数据库

	_ConnectionPtr m_pConWorkPrj;		// 项目临时库。
	_ConnectionPtr m_pConWeather;		// 城市气象参数库

	_ConnectionPtr m_pConMedium;		// 介质数据库.
	_ConnectionPtr m_pConRefInfo;		// 行业管理数据库
	
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAutoIPEDApp)
	public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementation
	//{{AFX_MSG(CAutoIPEDApp)
	afx_msg void OnAppAbout();
	afx_msg void OnFileNew();
	afx_msg void OnReplaceCurrentOldnameNewname();
	afx_msg void OnUpdateReplaceCurrentOldnameNewname(CCmdUI* pCmdUI);
	//}}AFX_MSG

public:
//	bool ImportAutoPD( CString strCommandLine );
	DECLARE_MESSAGE_MAP()

//更新指定工程的记录, 由已知的字段,算出其它字段的值.
//	bool RefreshData(CString EnginID);
};

extern CAutoIPEDApp theApp;

/////////////////////////////////////////////////////////////////////////////
extern BOOL FileExists(LPCTSTR lpszPathName);
extern BOOL DirExists(LPCTSTR lpszDir);
extern CString _GetFileName(LPCTSTR lpszPathName);
extern BOOL IsTableExists(_ConnectionPtr pCon, CString tb);

extern CString gsIPEDInsDir;    //AutoIPED安装目录
extern CString gsProjectDBDir;  //AutoIPED数据库安装目录
extern CString gsProjectDir;	//本地用户配置数据库目录
extern CString gsTemplateDir;	//本地用户配置数据库目录
extern CString gsShareMdbDir;  //共享数据库的存放目录  xl 10.19.2007

extern BOOL bIsCloseCompress;   //关闭时是否压缩数据库 
extern BOOL bIsReplaceOld;		//打开时是否自动替换旧材料名称
extern BOOL bIsMoveUpdate;		//编辑原始数据移动记录时是否更新
extern BOOL bIsAutoSelectPre;   //编辑原始数据移动记录时自动选择保温材料
extern BOOL bIsHeatLoss;		//计算时判断最大散热密度
extern BOOL bIsAutoAddValve;	//编辑原始数据移动记录时在管道上自动增加阀门
extern BOOL gbIsStopCalc;	 	//是否计算.
extern BOOL gbIsRunSelectEng;	//启动时是否弹出选择工程卷册对话框.
extern BOOL gbIsSelTblPrice;	//选择表中的热价比主汽价值
extern BOOL gbIsReplaceUnit;	//导入保温油漆数据时替换单位

extern int  giInsulStruct;		//计算时，保温结构与材料的选择。0，手工选择。     1，重新自动选择。
extern int  giCalSteanProp;		//计算水和蒸汽性质的标准:		0, IAPWS-IF97。   1, IFC67
extern BOOL gbAutoPaint120; 

extern BOOL gbWithoutProtectionCost;    //计算经济厚度时不包含保护层材料费用，默认为0-包含
extern BOOL gbInnerByEconomic;    //双层异材保温计算经济厚度时内层不按表面温度法计算，默认为0-按表面温度法计算

//extern int	giSupportOldCode;	//计算支吊架间距的标准.			0,新管规(国家电力行业标准,DL/T 5054-1996)  1,老管规(原国家电力部标准,DLGJ 23-81).
//extern int  giSupportRigidity;	//计算支吊架间距的刚度条件.		0,按强度条件  1, 按刚度条件	
//extern int  giHCrSteel;			//高铬钢,

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AUTOIPED_H__8F862085_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
