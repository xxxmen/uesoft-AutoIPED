#if !defined(AFX_CALTHREAD_H__196B2990_5327_469D_9527_02267584E35B__INCLUDED_)
#define AFX_CALTHREAD_H__196B2990_5327_469D_9527_02267584E35B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// CalThread.h : header file
//


/////////////////////////////////////////////////////////////////////////////
// CCalThread thread
class CAutoIPEDDoc;

class CCalThread : public CWinThread
{
	DECLARE_DYNCREATE(CCalThread)
public:
	CCalThread();           // protected constructor used by dynamic creation
	virtual ~CCalThread();

// Attributes
public:

// Operations
public:
	LPCTSTR GetProjectName();
	void SetProjectName(LPCTSTR ProjectName);
	CAutoIPEDDoc* m_pDoc;
	IStream* m_IStreamCODE;
	IStream* m_IStream;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCalThread)
	public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CCalThread)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
private:
	CString m_ProjectName;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CALTHREAD_H__196B2990_5327_469D_9527_02267584E35B__INCLUDED_)
