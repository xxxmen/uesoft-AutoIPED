// AutoIPEDView.h : interface of the CAutoIPEDView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_AUTOIPEDVIEW_H__8F86208D_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
#define AFX_AUTOIPEDVIEW_H__8F86208D_B060_11D7_9BCC_000AE616B8F1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "AutoIPEDDoc.h"

#include "VtoT.h"

#include "FoxBase.h"
#include "DlgAutoTotal.h"	// Added by ClassView
#include "uematerial.h"

class CAutoIPEDView : public CEditView,public CFoxBase
{
protected: // create from serialization only
	CAutoIPEDView();
	virtual ~CAutoIPEDView();

	DECLARE_DYNCREATE(CAutoIPEDView)

// Attributes
public:
	CAutoIPEDDoc* GetDocument();

// Operations
public:


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAutoIPEDView)
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual void OnInitialUpdate();
	virtual void OnUpdate(CView* pSender, LPARAM lHint, CObject* pHint);
	protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);
	//}}AFX_VIRTUAL

// Implementation

protected:
	BOOL OpenPassTable();

	void passModiTotal(CString str1,CString str2,CString str3);

private:
	// 从一张map中获得指定的标题的对话框对象
	CWnd* GetModelessDlgInMap(CString strDlgCaption);

	// 打开一个记录集，并显示
	void OpenATableToShow(CString SQL,_ConnectionPtr &IConnection,CString strDlgCaption=_T(""),
						  CString strDataGridCaption=_T(""),BOOL IsAllowUpdate=TRUE,
						  BOOL IsAllowAdd=TRUE,BOOL IsAllowDel=TRUE);

public:
	short DisplayRemarksInfo(CString strRemarks );

	bool CloseAllBrowseWindow();
	bool m_stateAutoTotal;
	BOOL bIsExplainPass;
	BOOL MinimizeAllWindow();

#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

// Generated message map functions
protected:
	BOOL m_passState51;
	_RecordsetPtr m_passSet;
	CUeMaterial*  m_pMaterialDlg;

private:
	BOOL CheckMaterialExists(_RecordsetPtr& pRsA_MAT,_RecordsetPtr& pRsCheck,CString strField,short nFieldCount, CString& strMessage,_RecordsetPtr pRsA_Com=NULL);
//	CString m_strRemarks;
	CMapStringToOb m_ModelessDlgMap;		//用于记录非模态对话框对象

protected:
	bool CheckFieldRepeat(CString strTblName, CString strTblCaption, CString strFieldName,CString strFieldCaption,CString& strRemarks);
	CString m_strMessage;
	CDlgAutoTotal m_autoTotal;
	//{{AFX_MSG(CAutoIPEDView)
	afx_msg void OnMenuitem61();
	afx_msg void OnEditOriginalDataMenu();
	afx_msg void OnMenuitem33();
	afx_msg void OnMenuitem34();
	afx_msg void OnMenuitem36();
	afx_msg void OnMenuitem41();
	afx_msg void OnMenuitem42();
	afx_msg void OnMenuitem43();
	afx_msg void OnMenuitem44();
	afx_msg void OnMenuitem45();
	afx_msg void OnMenuitem46();
	afx_msg void OnMenuitem51();
	afx_msg void OnMenuitem510();
	afx_msg void OnMenuitem511();
	afx_msg void OnMenuitem52();
	afx_msg void OnMenuitem53();
	afx_msg void OnMenuitem54();
	afx_msg void OnMenuitem55();
	afx_msg void OnMenuitem56();
	afx_msg void OnMenuitem57();
	afx_msg void OnMenuitem58();
	afx_msg void OnMenuitem59();
	afx_msg void OnMenuitem610();
	afx_msg void OnMenuitem611();
	afx_msg void OnMenuitem612();
	afx_msg void OnMenuitem62();
	afx_msg void OnMenuitem63();
	afx_msg void OnMenuitem64();
	afx_msg void OnMenuitem65();
	afx_msg void OnMenuitem66();
	afx_msg void OnMenuitem67();
	afx_msg void OnMenuitem68();
	afx_msg void OnMenuitem69();
	afx_msg void OnMenuitem71();
	afx_msg void OnMenuitem72();
	afx_msg void OnMenuitem73();
	afx_msg void OnMenuitem74();
	afx_msg void OnMenuitem613();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSelPrjVolume();
	afx_msg void OnCopyFromExistPrj();
	afx_msg void OnFileSeldir();
	afx_msg void OnOption();
	afx_msg void OnImportOtherIpedUE();
	afx_msg void OnImportOtherIpedSw();
	afx_msg void OnImportOtherIpedJs();
	afx_msg void OnSelectCriterionDatabase();
	afx_msg void OnUpdateMenuitem51(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem55(CCmdUI* pCmdUI);
	afx_msg void OnUpdateProjectConnectMenu(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem59(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem510(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem511(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem56(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem57(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem58(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem52(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem53(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem54(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem33(CCmdUI* pCmdUI);
	afx_msg void OnUpdateExplainTblPass(CCmdUI* pCmdUI);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnButtonAutoTotal();
	afx_msg void OnImportOriginalData();
	afx_msg void OnDisplayCalErrorInfo();
	afx_msg void OnImportDirData();
	afx_msg void OnModiRule1();
	afx_msg void OnMaterialDatabase();
	afx_msg void OnCloseAllWindow();
	afx_msg void OnImportPaintData();
	afx_msg void OnUpdateImportPaintData(CCmdUI* pCmdUI);
	afx_msg void OnNewOldMaterial();
	afx_msg void OnOutView();
	afx_msg void OnCheckCoherence();
	afx_msg void OnThicknessRegular();
	afx_msg void OnShowAllTable();
	afx_msg void OnCityWeather();
	afx_msg void OnCalcSupportSpan();
	afx_msg void OnMenuiRingName();
	afx_msg void OnUpdateOutView(CCmdUI* pCmdUI);
	afx_msg void OnAllMinimize();
	afx_msg void OnUpdateAllMinimize(CCmdUI* pCmdUI);
	afx_msg void OnMenuReplaceUnit();
	afx_msg void OnNewOldUnit();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // debug version in AutoIPEDView.cpp
inline CAutoIPEDDoc* CAutoIPEDView::GetDocument()
   { return (CAutoIPEDDoc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AUTOIPEDVIEW_H__8F86208D_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
