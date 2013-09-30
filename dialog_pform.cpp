// Dialog_pform.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "Dialog_pform.h"
#include "InputOtherData.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern CAutoIPEDApp theApp;

/////////////////////////////////////////////////////////////////////////////
// Dialog_pform dialog


Dialog_pform::Dialog_pform(CWnd* pParent /*=NULL*/)
	: CDialog(Dialog_pform::IDD, pParent)
{
	//{{AFX_DATA_INIT(Dialog_pform)
	//}}AFX_DATA_INIT
	engin=NULL;
	engin_name=NULL;
}


Dialog_pform::~Dialog_pform()
{
	if(engin)
		delete[] engin;

	if(engin_name)
		delete[] engin_name;
}

void Dialog_pform::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(Dialog_pform)
	DDX_Control(pDX, IDC_LIST2, c_list);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(Dialog_pform, CDialog)
	//{{AFX_MSG_MAP(Dialog_pform)
	ON_BN_CLICKED(IDC_BUTTONselect, OnBUTTONselect)
	ON_NOTIFY(NM_CLICK, IDC_LIST2, OnClickList)
	ON_BN_CLICKED(IDC_BUTTON_cancel, OnBUTTONcancel)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Dialog_pform message handlers

BOOL Dialog_pform::OnInitDialog() 
{
	_variant_t var;
	CString a,b;

	CDialog::OnInitDialog();
	
	///////////////////////////////////////////////////
	//   对listctrl控件附初始值
	/////////////////////////////////////////////////

	b=a=_T("");

	// 在ADO操作中建议语句中要常用try...catch()来捕获错误信息，
	// 因为它有时会经常出现一些想不到的错误。
	try
	{
		if(!m_pRecordset->BOF)
			m_pRecordset->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
		}
        
		CString str;
        LVITEM  lvItem;
	

		str="";
		c_list.InsertColumn( 0, (LPCTSTR)str, LVCFMT_LEFT, 60);

		str="";
		c_list.InsertColumn( 1, (LPCTSTR)str, LVCFMT_LEFT, 240);

		// 读入库中各字段并加入列表框中
		int s;
		s=0;

		while(!m_pRecordset->adoEOF)
		{   
			if(index==0) 
				var = m_pRecordset->GetCollect(_variant_t(long(0)));
			else 
				var = m_pRecordset->GetCollect(_variant_t(long(1)));

			if(var.vt != VT_NULL)
				a = (LPCSTR)_bstr_t(var);
			else
				a=_T("");

			if(index==0) 
				var = m_pRecordset->GetCollect(_variant_t(long(1)));
			else 
				var = m_pRecordset->GetCollect(_variant_t(long(0)));

			if(var.vt != VT_NULL)
				b = (LPCSTR)_bstr_t(var);
			else
				b=_T("");

			c_list.InsertItem(s, (LPCTSTR)a);

			lvItem.mask = LVIF_TEXT ;
			lvItem.iItem = s;
			lvItem.iSubItem =1;
			lvItem.pszText = b.GetBuffer(240);
			c_list.SetItem(&lvItem);
			b.ReleaseBuffer();

			m_pRecordset->MoveNext();
			s++;
		}

		// 默认列表指向第一项，同时移动记录指针并显示
		//m_list.SetCurSel(0);
	
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
	}

    //设置整行选择样式
	DWORD dStyleEx = c_list.GetExtendedStyle();
	dStyleEx |= LVS_EX_FULLROWSELECT;
	c_list.SetExtendedStyle( dStyleEx );

	if(index==0)  //设置单选样式
	{
		long winstyle;
		winstyle=::GetWindowLong(c_list.GetSafeHwnd(),GWL_STYLE);
		winstyle|=LVS_SINGLESEL;
		::SetWindowLong(c_list.GetSafeHwnd(),GWL_STYLE,winstyle);
	}
	
	m_pform="";
	i=0;
	cancel=0;
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void Dialog_pform::OnBUTTONselect() 
{
	//这里要增加单选多选的代码，根据index区分，index=1为多选，index=0为单选
	if(index==0)
	{
		if(m_pform=="") m_pform=old;

		UpdateData(FALSE);
		OnOK();
	}
	if(index==1)
	{
		POSITION pos;

		CListCtrl* pList = (CListCtrl*) GetDlgItem(IDC_LIST2);

		i=c_list.GetSelectedCount();
		engin=new CString[i];
		engin_name=new CString[i];
		pos=c_list.GetFirstSelectedItemPosition();

		int nItem,k;

		k=0;
		if (pos == NULL)
			TRACE0("No items were selected!\n");
		else
		{
			while(pos!=NULL && k<i)
			{
				CString str,str1;

				nItem= pList->GetNextSelectedItem(pos);

	            pList->GetItemText(nItem, 0, str.GetBufferSetLength(255),255);
	            pList->GetItemText(nItem, 1, str1.GetBufferSetLength(255),255);
	            str.ReleaseBuffer();
				str1.ReleaseBuffer();
			    engin[k]=str;
				engin_name[k]=str1;
				k++;
			}
		}
		OnOK();
	}
	
}

void Dialog_pform::OnClickList(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CListCtrl* pList = (CListCtrl*) GetDlgItem(IDC_LIST2);
	
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
    CPoint pt = pNMListView->ptAction;

	POSITION pos = pList->GetFirstSelectedItemPosition();
	if (pos == NULL)
		TRACE0("No items were selected!\n");
	else
	{
		while (pos)
		{
			int nItem = pList->GetNextSelectedItem(pos);
			if(nItem == pList->HitTest(pt, NULL))
			{
				CString str,str1;

                pList->GetItemText(nItem, 0, str.GetBufferSetLength(255),255);
                pList->GetItemText(nItem, 1, str1.GetBufferSetLength(255),255);
				str.ReleaseBuffer();
				str1.ReleaseBuffer();

				m_pform=str;
				m_pform1=str1;
			}
		}
	}
	*pResult = 0;
}

void Dialog_pform::OnBUTTONcancel() 
{
	m_pform=old;
	m_pform1=old1;
	cancel=1;

	OnCancel();
}


