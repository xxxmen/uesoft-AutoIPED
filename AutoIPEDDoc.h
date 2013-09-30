// AutoIPEDDoc.h : interface of the CAutoIPEDDoc class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_AUTOIPEDDOC_H__8F86208B_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
#define AFX_AUTOIPEDDOC_H__8F86208B_B060_11D7_9BCC_000AE616B8F1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "HeatPreCal.h"
#include "CalThread.h"

class CAutoIPEDDoc : public CDocument,public CHeatPreCal
{
protected: // create from serialization only
	CAutoIPEDDoc();
	virtual ~CAutoIPEDDoc();

	DECLARE_DYNCREATE(CAutoIPEDDoc)

// Attributes
public:
	// 判断是否正在计算
	bool Operation();

// Operations
public:
	void OnEndCalculate();		// 计算完毕时将被调用

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAutoIPEDDoc)
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);
	//}}AFX_VIRTUAL

// Implementation
public:
	CString GetAllCalErrorInfo();
    CString m_Result;
	BOOL IsStopCalc;

#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

public:
	virtual void MinimizeAllWindow();

	afx_msg void OnButtonStopcal();

protected:
	//{{AFX_MSG(CAutoIPEDDoc)
	afx_msg void OnMenuitem_calc();
	afx_msg void OnUpdateButtonStopcal(CCmdUI* pCmdUI);
	afx_msg void OnUpdateMenuitem_calc(CCmdUI* pCmdUI);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

protected:
	// 重载CFoxBase的ExceptionInfo
	virtual void ExceptionInfo(LPCTSTR pErrorInfo);

	virtual void SelectToCal(_RecordsetPtr &IRecordset,int *pIsStop);
	virtual BOOL RangeDlgshow(long &Start,long &End);
	virtual void DisplayRes(LPCTSTR pStr);
	virtual void InputD(LPCTSTR TempStr,double &stres);
	virtual void Say(LPCTSTR pStr);
	void Cancel(int *pState);

private:
	CCalThread* m_pCalThread;

	CString m_ErrorInfo;	// 用于记录在一次计算过程中所有的错误
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_AUTOIPEDDOC_H__8F86208B_B060_11D7_9BCC_000AE616B8F1__INCLUDED_)
