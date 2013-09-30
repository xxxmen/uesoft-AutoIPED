#if !defined(AFX_INPUT_H__0FD822B5_D5F2_4191_A04B_F04DFC209892__INCLUDED_)
#define AFX_INPUT_H__0FD822B5_D5F2_4191_A04B_F04DFC209892__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Input.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// Input dialog

class Input : public CDialog
{
// Construction
public:
	bool GetStringInput();
	void SetStringInput(bool IsStringInput);
	Input(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(Input)
	enum { IDD = IDD_DIALOG_input };
	CString	m_static;
	double	m_input;
	CString	m_StringInput;
	//}}AFX_DATA


private:
	bool m_IsStringInput;	// 是否输入的是字符串

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(Input)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(Input)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_INPUT_H__0FD822B5_D5F2_4191_A04B_F04DFC209892__INCLUDED_)
