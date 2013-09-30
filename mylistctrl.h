#if !defined(AFX_MYLISTCTRL_H__2CBBF377_59A4_11D3_860E_C141D5317B49__INCLUDED_)
#define AFX_MYLISTCTRL_H__2CBBF377_59A4_11D3_860E_C141D5317B49__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// MyListCtrl.h : header file
//
///////////////////////////////////////////////////////////////////

// CMyListCtrl window

class CMyListCtrl : public CListCtrl
{
public:
   CMyListCtrl();
   virtual ~CMyListCtrl();
   

public: 
	//向列表框中添加一列
	BOOL AddColumn(
		LPCTSTR strItem,int nItem,int nWidth=-1,int nSubItem = -1,
		int nMask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM,
		int nFmt = LVCFMT_LEFT);

	//向列表框中添加一行
	BOOL AddItem(int nItem,int nSubItem,LPCTSTR strItem,int nImageIndex = -1);

	BOOL   IsItem(int nItem) const;//所选行是否存在
	BOOL   IsColumn(int nCol) const;//所选列是否存在
	int    GetSelectedItem(int nStartItem = -1) const;//得到所选行
	BOOL   SelectItem(int nItem);//选择指定行
	BOOL   SelectAll();//选择所有行

public:
	long get_ColumnCount();
	void DeleteAllColumns();
private:

// Overrides
   // ClassWizard generated virtual function overrides
   //{{AFX_VIRTUAL(CMyListCtrl)
	//}}AFX_VIRTUAL

//   virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);//重载函数,绘制列表框中内容

// Generated message map functions
protected:
   //{{AFX_MSG(CMyListCtrl)
	//}}AFX_MSG
   DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MYLISTCTRL_H__2CBBF377_59A4_11D3_860E_C141D5317B49__INCLUDED_)
