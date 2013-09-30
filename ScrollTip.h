#if !defined(AFX_SCROLLTIP_H__7191B0EA_D916_4173_9933_EB73E7B385AA__INCLUDED_)
#define AFX_SCROLLTIP_H__7191B0EA_D916_4173_9933_EB73E7B385AA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ScrollTip.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CScrollTip window

class CScrollTip : public CStatic
{
// Construction
public:
	CScrollTip();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CScrollTip)
	//}}AFX_VIRTUAL

// Implementation
public:
	void SetBkColor(COLORREF crBk);
	void SetTextColor(COLORREF crText);
	virtual ~CScrollTip();

	// Generated message map functions
protected:
	//{{AFX_MSG(CScrollTip)
	afx_msg void OnPaint();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
private:
	COLORREF m_crTextColor;
	COLORREF m_crBkColor;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SCROLLTIP_H__7191B0EA_D916_4173_9933_EB73E7B385AA__INCLUDED_)
