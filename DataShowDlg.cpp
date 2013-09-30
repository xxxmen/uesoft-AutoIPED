// DataShowDlg.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "DataShowDlg.h"

#include "Column.h"
#include "Columns.h"
#include "vtot.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDataShowDlg dialog
extern CAutoIPEDApp theApp;

CDataShowDlg::CDataShowDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDataShowDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDataShowDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	m_nCursorCol = -1;

	m_strCol_Pass = _T("");
	m_DlgCaption=_T("");

	m_IsAllowDel=FALSE;
	m_IsAllowAddNew=FALSE;
	m_IsAllowUpdate=FALSE;

	m_pIndexArrayToHide=NULL;

	m_pDefaultFieldValueStruct=NULL;

	m_IsAutoDelete=FALSE;

	m_hIcon=AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

CDataShowDlg::~CDataShowDlg()
{
	if(m_pIndexArrayToHide)
		delete[] m_pIndexArrayToHide;

	if(m_pDefaultFieldValueStruct)
		delete[] m_pDefaultFieldValueStruct;
}

void CDataShowDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDataShowDlg)
	DDX_Control(pDX, IDC_RESULT_BROWSE, m_ResultBrowseControl);
	DDX_Control(pDX, IDC_RESULT_BROWSE2, m_BrowseGrid);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDataShowDlg, CDialog)
	//{{AFX_MSG_MAP(CDataShowDlg)
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

///////////////////////////////////////////////////////////////
//
// 设置需要显示的记录集
//
// IRecordset[in]已经打开的记录集
//
void CDataShowDlg::SetRecordset(_RecordsetPtr &IRecordset, CDataGridEx* pCtlDataGrid)
{
	m_IRecordset=IRecordset;
	//ZSYTEST
	if( pCtlDataGrid == NULL )
	{
		//默认第一个为当前要处理的DataGrid控件
		pCurDataGrid = &m_ResultBrowseControl;
	}
	else
	{
		//设置为指定的DataGrid控件
		pCurDataGrid = pCtlDataGrid;
	}
	if( IsWindow( pCurDataGrid->GetSafeHwnd()) )
	{
		pCurDataGrid->SetRefDataSource((LPUNKNOWN)IRecordset);
	}

}

///////////////////////////////////////////////////////////
//
// 返回需要显示的记录集
//
_RecordsetPtr& CDataShowDlg::GetRecordset()
{
	return m_IRecordset;
}

///////////////////////////////////////////////////////
//
// 设置对话框的标题
//
void CDataShowDlg::SetDlgCaption(LPCTSTR pCaption)
{
	if(pCaption==NULL)
		m_DlgCaption=_T("");
	else
		m_DlgCaption=pCaption;

	if(IsWindow(this->m_hWnd))
	{
		SetWindowText(m_DlgCaption);
	}
}

//////////////////////////////////////////////////////
//
// 返回对话框的标题
//
CString CDataShowDlg::GetDlgCaption()
{
	return m_DlgCaption; 
}

//////////////////////////////////////////////////////////////
//
// 设置DataGrid控件的标题
//
void CDataShowDlg::SetDataGridCaption(LPCTSTR pCaption)
{
	if(pCaption==NULL)
		m_DataGridCaption=_T("");
	else
		m_DataGridCaption=pCaption;

	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		//取消在DataGria控件上显示表名。
	//	m_ResultBrowseControl.SetCaption(m_DataGridCaption);
	}
}

/////////////////////////////////////////////////////////////
//
// 返回设置DataGrid控件的标题
//
CString CDataShowDlg::GetDataGridCaption()
{ 
	return m_DataGridCaption; 
}


///////////////////////////////////////////////////////////
//
// 设置是否允许更新
//
// IsAllow[in]	当为TRUE时允许更新，当为FALSE时不能更新
//
void CDataShowDlg::SetAllowUpdate(BOOL IsAllow)
{
	m_IsAllowUpdate=IsAllow;

	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		m_ResultBrowseControl.SetAllowUpdate(IsAllow);
	}
}

//////////////////////////////////////////////////////
//
// 返回是否允许更新
//
BOOL CDataShowDlg::GetAllowUpdate()
{
	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		m_IsAllowUpdate=m_ResultBrowseControl.GetAllowUpdate();
	}

	return m_IsAllowUpdate; 
}



///////////////////////////////////////////////////////////
//
// 设置是否允许增加记录
//
// IsAllow[in]	当为TRUE时允许，当为FALSE时不能增加记录
//
void CDataShowDlg::SetAllowAddNew(BOOL IsAllow)
{
	m_IsAllowAddNew=IsAllow;

	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		m_ResultBrowseControl.SetAllowAddNew(IsAllow);
	}
}

//////////////////////////////////////////////////////
//
// 返回是否允许新增加记录
//
BOOL CDataShowDlg::GetAllowAddNew()
{
	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		m_IsAllowAddNew=m_ResultBrowseControl.GetAllowAddNew();
	}

	return m_IsAllowAddNew;
}



///////////////////////////////////////////////////////////
//
// 设置是否允许删除记录
//
// IsAllow[in]	当为TRUE时允许，当为FALSE时不允许
//
void CDataShowDlg::SetAllowDelete(BOOL IsAllow)
{
	m_IsAllowDel=IsAllow;

	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		m_ResultBrowseControl.SetAllowDelete(IsAllow);
	}
}

//////////////////////////////////////////////////////
//
// 返回是否允许删除记录
//
BOOL CDataShowDlg::GetAllowDelete()
{
	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		m_IsAllowDel=m_ResultBrowseControl.GetAllowDelete();
	}

	return m_IsAllowDel;
}

//////////////////////////////////////////////////////////////////
//
// 设置在DataGrid控件显示数据时，需要隐藏的列
//
// pIndex[in]		需要隐藏列的索引数组，从0开始
// IndexSize[in]	pIndex的大小
//
// throw(COleDispatchException)
//
void CDataShowDlg::SetHideColumns(int *pIndex, DWORD IndexSize)
{
	if(pIndex==NULL || IndexSize<=0)
		return;

	if(m_pIndexArrayToHide)
		delete[] m_pIndexArrayToHide;

	m_strFieldsNameToHide=_T("");

	m_pIndexArrayToHide=new int[IndexSize+1];

	for(DWORD i=0;i<IndexSize;i++)
	{
		m_pIndexArrayToHide[i]=pIndex[i];
	}

	m_pIndexArrayToHide[IndexSize]=EndHideArray;

	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		try
		{
			HideColumns();
		}
		catch(COleDispatchException *e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}

	}
}

//////////////////////////////////////////////////////////////////
//
// 设置在DataGrid控件显示数据时，需要隐藏的列
//
// szFields[in]		需要隐藏列字段名，有多个字段需要隐藏时，各个
//					字段用“|”隔开
//
// throw(COleDispatchException)
//
void CDataShowDlg::SetHideColumns(LPCTSTR szFields)
{
	if(szFields==NULL)
		return;

	m_strFieldsNameToHide=szFields;
	
	if(m_pIndexArrayToHide)
	{
		delete[] m_pIndexArrayToHide;
		m_pIndexArrayToHide=NULL;
	}

	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
	{
		try
		{
			HideColumns();
		}
		catch(COleDispatchException *e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}

	}
}

//////////////////////////////////////////////
//
// 隐藏指定的列
//
// throw(COleDispatchException)
//
void CDataShowDlg::HideColumns()
{
	int *pIndex;
	long iIndex;
	CColumn Col;

	if(!::IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
		return;

	if(!m_strFieldsNameToHide.IsEmpty())
	{
		try
		{
			HideColumnsByFieldsName(m_strFieldsNameToHide);
		}
		catch(COleDispatchException *e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}

		return;
	}

	if(m_pIndexArrayToHide==NULL)
		return;

	for(pIndex=m_pIndexArrayToHide; *pIndex!=EndHideArray; pIndex++)
	{
		if(*pIndex == HideFirstColumn)
		{
			iIndex=0;
		}
		else if(*pIndex == HideLastColumn)
		{
			iIndex=m_ResultBrowseControl.GetColumns().GetCount()-1;
		}
		else
		{
			iIndex=*pIndex;
		}
	}

	try
	{
		Col=m_ResultBrowseControl.GetColumns().GetItem(_variant_t(iIndex));
		Col.SetVisible(FALSE);
	}
	catch(COleDispatchException *e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

}

/////////////////////////////////////////////////////////////
//
// 隐藏指定字段名称的列
//
// szFields[in]	需要隐藏列的字段名
//
// throw(COleDispatchException);
//
void CDataShowDlg::HideColumnsByFieldsName(LPCTSTR szFields)
{
	CColumn Col;
	CString strFields,strField;
	int nPrev,nCur=0;

	if(!IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
		return;

	if(szFields==NULL)
		return;

	strFields=szFields;

	nCur=0;
	nPrev=-1;

	while(TRUE)
	{
		nPrev=nCur;

		nCur=strFields.Find(_T('|'),nCur);

		if(nCur==-1)
		{
			strField=strFields.Mid(nPrev);
		}
		else
		{
			strField=strFields.Mid(nPrev,nCur-nPrev);
		}

		try
		{
			Col=m_ResultBrowseControl.GetColumns().GetItem(_variant_t(strField));

			Col.SetVisible(FALSE);
		}
		catch(COleDispatchException *e)
		{
		//	ReportExceptionErrorLV2(e); //zsy 2005.3.16
		//	throw;
			e->Delete();
		}

		if(nCur==-1)
			break;
		
		nCur++;
		
	}

}


//////////////////////////////////////////////////////////////////////////////////////////
//
// 设置在添加记录时，字段的默认值
//
// pDefault[in]		tag_DefaultFieldValue结构的指针
// dwLength			pDefault的数量
//
void CDataShowDlg::SetDefaultFieldValue(tag_DefaultFieldValue *pDefault, DWORD dwLength)
{
	if(pDefault==NULL || dwLength<=0)
		return;

	if(m_pDefaultFieldValueStruct)
	{
		delete []m_pDefaultFieldValueStruct;
	}

	m_pDefaultFieldValueStruct=new tag_DefaultFieldValue[dwLength+1];

	for(DWORD i=0;i<dwLength;i++)
	{
		m_pDefaultFieldValueStruct[i].strFieldName=pDefault[i].strFieldName;
		m_pDefaultFieldValueStruct[i].varFieldValue=pDefault[i].varFieldValue;
	}

	m_pDefaultFieldValueStruct[dwLength].strFieldName=_T("");
}

////////////////////////////////////////////////////
//
// 设置在非模态中是否在窗口销毁时是否自动删除对象
//
// bAutoDelete[in]	当为TRUE时在窗口销毁时自动删除对象
//					当为FALSE时需要自己销毁对象
//
void CDataShowDlg::SetAutoDelete(BOOL bAutoDelete)
{
	m_IsAutoDelete=bAutoDelete;
}

/////////////////////////////////////////////////////////////////////////////
// CDataShowDlg message handlers
//

BOOL CDataShowDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
//	if(m_IRecordset==NULL)
//	{
//		ReportExceptionError(_T("未指定连接"));
//		OnCancel();
//		return TRUE;
//	}
	pCurDataGrid = &m_ResultBrowseControl;		//默认第一个控件为当前要处理的.

	m_ResultBrowseControl.SetRefDataSource((LPUNKNOWN)m_IRecordset);	

	if(!GetDlgCaption().IsEmpty())
		SetWindowText(GetDlgCaption());

	//取消 DataGrid控件中显示的标题,该标题名称原来是数据库中表的名称. [2005/06/20]

//	if(!GetDataGridCaption().IsEmpty())
//		m_ResultBrowseControl.SetCaption(GetDataGridCaption());
	
	this->ShowWindow(SW_MAXIMIZE);
	RECT rect;
	GetClientRect(&rect);
	m_ResultBrowseControl.MoveWindow(0,0,rect.right-rect.left,rect.bottom-rect.top);

	m_ResultBrowseControl.SetAllowAddNew(m_IsAllowAddNew);
	m_ResultBrowseControl.SetAllowUpdate(m_IsAllowUpdate);
	m_ResultBrowseControl.SetAllowDelete(m_IsAllowDel);

	//设置字段的宽度和标题
	this->SetFieldCaption( GetDataGridCaption() );

	try
	{
		HideColumns();
	}
	catch(COleDispatchException *e)
	{
		ReportExceptionErrorLV1(e);
		e->Delete();
	}

	m_ResultBrowseControl.InitMouseWheel();
	m_BrowseGrid.InitMouseWheel();

	//初始光标定位到指定的列.
	if( m_nCursorCol >= 0 )
	{
		this->FindDataGridCol(m_nCursorCol);
	}

	SetIcon(m_hIcon,TRUE);
	SetIcon(m_hIcon,FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


void CDataShowDlg::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	
	if(IsWindow(m_ResultBrowseControl.GetSafeHwnd()))
		m_ResultBrowseControl.MoveWindow(0,0,cx,cy);

}



BEGIN_EVENTSINK_MAP(CDataShowDlg, CDialog)
    //{{AFX_EVENTSINK_MAP(CDataShowDlg)
	ON_EVENT(CDataShowDlg, IDC_RESULT_BROWSE, 217 /* OnAddNew */, OnOnAddNewResultBrowse, VTS_NONE)
	ON_EVENT(CDataShowDlg, IDC_RESULT_BROWSE, 218 /* RowColChange */, OnRowColChangeResultBrowse, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDataShowDlg, IDC_RESULT_BROWSE, -600 /* Click */, OnClickResultBrowse, VTS_NONE)
	ON_EVENT(CDataShowDlg, IDC_RESULT_BROWSE, 213 /* ColResize */, OnColResizeResultBrowse, VTS_I2 VTS_PI2)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

///////////////////////////////////////////////
//
// 响应DataGrid的AddNew事件
//
// 在此加如默认的字段值
//
void CDataShowDlg::OnOnAddNewResultBrowse() 
{
	tag_DefaultFieldValue *pTemp;
	_RecordsetPtr IRecordset;
	LPUNKNOWN pUnknown;
	HRESULT hr;

	if(m_pDefaultFieldValueStruct==NULL)
		return;

	try
	{
		pUnknown=m_ResultBrowseControl.GetDataSource();

		hr=pUnknown->QueryInterface(__uuidof(_RecordsetPtr),(void**)&IRecordset);

		if(FAILED(hr))
		{
			pUnknown->Release();
			ReportExceptionError(_com_error(hr));
			return;
		}

		pUnknown->Release();
	}
	catch(COleDispatchException *e)
	{
		ReportExceptionErrorLV1(e);
		e->Delete();
	}
	FieldPtr col;
	for(pTemp=m_pDefaultFieldValueStruct ; !pTemp->strFieldName.IsEmpty() ; pTemp++)
	{
		try
		{
			try
			{
				col = IRecordset->GetFields()->GetItem(_variant_t(pTemp->strFieldName));
			}
			catch(_com_error)
			{
				continue;	//没有该列
			}

			IRecordset->PutCollect(_variant_t(pTemp->strFieldName),pTemp->varFieldValue);
		}
		catch(_com_error &e)
		{
			ReportExceptionErrorLV1(e);
		}
	}
	
}


void CDataShowDlg::OnOK()
{
	m_ResultBrowseControl.UnInitMouseWheel();
	m_BrowseGrid.UnInitMouseWheel();

	EndDialog(0);

	if(IsWindow(this->GetSafeHwnd()))
		DestroyWindow();

	if(m_IsAutoDelete)
		delete this;
}

void CDataShowDlg::OnCancel()
{
	m_ResultBrowseControl.UnInitMouseWheel();
	m_BrowseGrid.UnInitMouseWheel();

	EndDialog(0);

	if(IsWindow(this->GetSafeHwnd()))
		DestroyWindow();

	if(m_IsAutoDelete)
		delete this;
}

void CDataShowDlg::OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI) 
{
	CRect rect;
	POINT pt;
	CWnd *pParent;
	CWnd *pActiveWnd;

	pParent=GetParent();

	if(pParent==NULL)
	{
		CDialog::OnGetMinMaxInfo(lpMMI);
		return;
	}

	if(pParent->IsKindOf(RUNTIME_CLASS(CFrameWnd)))
	{
		pActiveWnd=((CFrameWnd*)pParent)->GetActiveView();
	}
	else
	{
		pActiveWnd=pParent;
	}

	if(pActiveWnd && IsWindow(pActiveWnd->GetSafeHwnd()))
	{
		pActiveWnd->GetClientRect(&rect);

		pt.x=rect.left;
		pt.y=rect.top;

		pActiveWnd->ClientToScreen(&pt);

		lpMMI->ptMaxPosition.x=pt.x;
		lpMMI->ptMaxPosition.y=pt.y;

		lpMMI->ptMaxSize.x=rect.Width();
		lpMMI->ptMaxSize.y=rect.Height();
	}

	CDialog::OnGetMinMaxInfo(lpMMI);
}

void CDataShowDlg::OnDestroy() 
{
	m_ResultBrowseControl.UnInitMouseWheel();
	m_BrowseGrid.UnInitMouseWheel();

	CDialog::OnDestroy();	
}

void CDataShowDlg::OnRowColChangeResultBrowse(VARIANT FAR* LastRow, short LastCol) 
{
	if( !m_strCol_Pass.IsEmpty() )  //只对指定的字段进行处理.()
	{
		CString strText;
		short Col = 0 ; 
		Col = m_ResultBrowseControl.GetCol();
		if( Col != -1 )
		{
			//strText = m_ResultBrowseControl.GetColumns().GetItem(_variant_t((short)Col)).GetCaption();
			strText = (LPCTSTR)m_IRecordset->GetFields()->GetItem(Col)->GetName();

			if( !strText.CompareNoCase(m_strCol_Pass) )
			{
				strText = m_ResultBrowseControl.GetText();
				if( !strText.CompareNoCase("Y") )
				{
				//	m_ResultBrowseControl.SetText("");
					m_ResultBrowseControl.GetColumns().GetItem(_variant_t((short)Col)).SetText("");
				}
				else
				{
					m_ResultBrowseControl.GetColumns().GetItem(_variant_t((short)Col)).SetText("Y");
				}
			}
		}
	}

}
//设置指定的字段赋一个特定的值
void CDataShowDlg::SetPassField(CString strField)
{
	m_strCol_Pass = strField;
}

//单击.
void CDataShowDlg::OnClickResultBrowse() 
{
/*	int col = m_ResultBrowseControl.GetCol();
	CString strTmp;
	if( col != -1 )
	{
		strTmp = m_ResultBrowseControl.GetColumns().GetItem(_variant_t((long)col)).GetCaption();
		if( !strTmp.CompareNoCase(m_strCol_Pass) )
		{
			strTmp = m_ResultBrowseControl.GetColumns().GetItem(_variant_t((long)col)).GetText();
			if( !strTmp.CompareNoCase("Y") )
			{
				m_ResultBrowseControl.GetColumns().GetItem(_variant_t((long)col)).SetText("");
			}
			else
			{
				m_ResultBrowseControl.GetColumns().GetItem(_variant_t((long)col)).SetText("Y");
			}
		}
	}
*/
}
////////////////////////////////////////////////
//定位光标到指定的列.
//
bool CDataShowDlg::FindDataGridCol(short colIndex)
{
	try
	{
		if( IsWindow(pCurDataGrid->GetSafeHwnd()) )
		{
			pCurDataGrid->SetCol(colIndex);
		}
	}
	catch(_com_error)
	{
		return false;
	}
	return true;
}
//设置光标到指定的列.
void CDataShowDlg::SetCursorCol(short startCol)
{
	m_nCursorCol = startCol;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/16]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :设置字段的标题和宽度.
//------------------------------------------------------------------
BOOL CDataShowDlg::SetFieldCaption(CString strTblName)
{
	if( !SetFieldCaptionAndWidth(pCurDataGrid,theApp.m_pConIPEDsort, strTblName) )
	{
		return FALSE;
	}
	return TRUE;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/15]
// Author       :ZSY
// Parameter(s) :
// Return       :全局的
// Remark       :设置字段的标题和宽度.在IPEDsort.MDB数据库中可以找到对应关系
//------------------------------------------------------------------
BOOL SetFieldCaptionAndWidth(CDataGridEx* pDataGrid, _ConnectionPtr& pConSort, CString strTblName)
{
	if (pConSort == NULL || strTblName.IsEmpty())
	{
		return FALSE;
	}
	if ( !IsWindow(pDataGrid->GetSafeHwnd()) )
	{
		return FALSE;
	}

	try
	{
		_RecordsetPtr pRsStruct;		//不同数据表与之对应的结构表.

		CString		  strSQL;			//SQL语句
		CString		  strStructTblName;	//数据表对应的结构表名.
		CString		  strFieldName;		//字段名
		CString		  strChineseCaption;//对应的中文描述

		long		nFieldCount;		//显示的字段个数
		float		fFieldWidth;		//字字的宽度
		long		fWindowWidth;		//DataGrid控件的宽度
		
		RECT		rect;
		BOOL		bIsVisible;			//字段是否可见.

		float	xRatio = 1;		//屏幕分辨率的比列.横向
		float	yRatio = 1;
		GetScreenRatio(xRatio,yRatio);

		pRsStruct.CreateInstance(__uuidof(Recordset));
		
		if (!GetStructTblName(pConSort,strTblName,strStructTblName))
		{
			return FALSE;
		}
		
		pDataGrid->GetWindowRect(&rect);
		fWindowWidth = rect.right-rect.left;

		//打开结构表。
		strSQL = "SELECT * FROM ["+strStructTblName+"]";
		pRsStruct->Open(_bstr_t(strSQL), pConSort.GetInterfacePtr(), 
					adOpenStatic, adLockOptimistic, adCmdText);
		if (pRsStruct->GetRecordCount() <= 0)
		{
			return FALSE;
		}
		nFieldCount = pDataGrid->GetColumns().GetCount();
		for (short i=0; i < nFieldCount; i++)
		{
			strFieldName = pDataGrid->GetColumns().GetItem(_variant_t(i)).GetCaption();
			strSQL = "FieldName='"+strFieldName+"'";
			//定位到当前字段名对应的记录上。
			pRsStruct->MoveFirst();
			pRsStruct->Find(_bstr_t(strSQL), 0, adSearchForward);
			if (pRsStruct->adoEOF)
			{
				
				//当前字段没有对应的描述记录，不进行处理 ??须要隐藏吗??
				pDataGrid->GetColumns().GetItem(_variant_t(i)).SetVisible(FALSE);
				continue;
			}
			//字段名的本地描述字符.
			strChineseCaption = vtos(pRsStruct->GetCollect(_variant_t("LocalCaption")));
			if (!strChineseCaption.IsEmpty())
			{
				//设置字段名为本地描述的字符
				pDataGrid->GetColumns().GetItem(_variant_t(i)).SetCaption(strChineseCaption);
			}
			fFieldWidth = (float)vtof(pRsStruct->GetCollect(_variant_t("width")));
			if (fWindowWidth > 0 && fFieldWidth > fWindowWidth)
			{
				fFieldWidth = (float)fWindowWidth*2/30;		//取总宽度的2/30
			}
			fFieldWidth = fFieldWidth > 0 ? fFieldWidth:60;

			fFieldWidth *= xRatio;
			//设置字段的宽度
			pDataGrid->GetColumns().GetItem(_variant_t(i)).SetWidth(fFieldWidth);
			//字段是否可见
			bIsVisible = vtob(pRsStruct->GetCollect(_variant_t("Visible")));
			pDataGrid->GetColumns().GetItem(_variant_t(i)).SetVisible(bIsVisible);				
		}

	}
	catch (_com_error)
	{
		return FALSE;
	}
	return TRUE;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/15]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       : 根据数据表,获得相对应的结构表名.
//------------------------------------------------------------------
BOOL GetStructTblName(_ConnectionPtr &pCon, CString strDataTblName, CString& strStructTblName)
{ 
	try
	{
		_RecordsetPtr pRsTblINFO;	//管理表.
		pRsTblINFO.CreateInstance(__uuidof(Recordset));
		CString strSQL;

		strSQL = "SELECT * FROM [TableINFO] WHERE ManuTBName='"+strDataTblName+"' ";
		pRsTblINFO->Open(_bstr_t(strSQL), pCon.GetInterfacePtr(),
						adOpenStatic, adLockOptimistic, adCmdText);
		if (pRsTblINFO->GetRecordCount() <= 0)
		{
			return FALSE;
		}
		strStructTblName = vtos( pRsTblINFO->GetCollect(_variant_t("StructTblName")) );
		if (strStructTblName.IsEmpty())		//没有结构表。
		{
			return FALSE;
		}
	}
	catch (_com_error)
	{
		return FALSE;
	}

	return TRUE;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/17]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :保存DataGrid控件的字段的宽度
//------------------------------------------------------------------
BOOL SaveDataGridFieldWidth(CDataGridEx* pDataGrid, _ConnectionPtr& pConSort, CString strTblName)
{
	if (pConSort == NULL || strTblName.IsEmpty())
	{
		return FALSE;
	}
	if ( !IsWindow(pDataGrid->GetSafeHwnd()) )
	{
		return FALSE;
	}

	try
	{
		_RecordsetPtr pRsStruct;		//不同数据表与之对应的结构表.

		CString		  strSQL;			//SQL语句
		CString		  strStructTblName;	//数据表对应的结构表名.
		CString		  strFieldName;		//字段名
		CString		  strChineseCaption;//对应的中文描述

		long		nFieldCount;		//显示的字段个数
		float		fFieldWidth;		//字字的宽度

		float	xRatio = 1;		//屏幕分辨率的比列.横向
		float	yRatio = 1;
		GetScreenRatio(xRatio,yRatio);

		
		pRsStruct.CreateInstance(__uuidof(Recordset));
		
		if (!GetStructTblName(pConSort,strTblName,strStructTblName))
		{
			return FALSE;
		}
		//打开结构表。
		strSQL = "SELECT * FROM ["+strStructTblName+"]";
		pRsStruct->Open(_bstr_t(strSQL), pConSort.GetInterfacePtr(), 
					adOpenStatic, adLockOptimistic, adCmdText);
		if (pRsStruct->GetRecordCount() <= 0)
		{
			return FALSE;
		}
		nFieldCount = pDataGrid->GetColumns().GetCount();
		for (short i=0; i < nFieldCount; i++)
		{
			strFieldName = pDataGrid->GetColumns().GetItem(_variant_t(i)).GetCaption();
			strSQL = "LocalCaption='"+strFieldName+"'";
			//定位到当前字段名对应的记录上。
			pRsStruct->MoveFirst();
			pRsStruct->Find(_bstr_t(strSQL), 0, adSearchForward);
			if (!pRsStruct->adoEOF)
			{
				
				//获得当前字段的宽度。
				fFieldWidth = pDataGrid->GetColumns().GetItem(_variant_t(i)).GetWidth();

				//考虑屏幕分辨率.
				fFieldWidth /= xRatio;
				//将当前字段的宽度保存到数据库中。
				pRsStruct->PutCollect(_variant_t("width"), _variant_t(fFieldWidth));
				pRsStruct->Update();
			}
		}

	}
	catch (_com_error)
	{
		return FALSE;
	}
	return TRUE;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/17]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :改变列的大小
//------------------------------------------------------------------
void CDataShowDlg::OnColResizeResultBrowse(short ColIndex, short FAR* Cancel) 
{
	SaveDataGridFieldWidth(pCurDataGrid, theApp.m_pConIPEDsort, GetDataGridCaption());
}

//------------------------------------------------------------------                
// DATE         :[2005/06/20]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :获得屏幕的分辨率比。以800 * 600为1
//------------------------------------------------------------------
BOOL GetScreenRatio(float& xRatio, float& yRatio)
{
	CSize sizeScreen;		//屏幕大小
	sizeScreen.cx = GetSystemMetrics(SM_CXSCREEN);
	sizeScreen.cy = GetSystemMetrics(SM_CYSCREEN);

	double ratio;		//比
	double ratio1;
	//数据库中以分辨率800*600为标准.
	ratio=8.0/5.0;
	ratio1=89.5/59.0;
	
	switch(sizeScreen.cx*sizeScreen.cy)
	{
 	case 640*480:ratio=2.0;
 			     ratio1=112.0/59.0;
 		break;
 	case 720*480:ratio=24.0/14.0;
 		         ratio1=98.0/58.0;
 				 break;
 	case 720*576:ratio=24.0/14.0;
 		          ratio1=98.0/58.0;
 				 break;
 	case 848*480:
 		 ratio=24.0/16.0;
 		 ratio1=98.0/68.5;
 		 break;
  	case 800*600:
 			ratio=8.0/5.0;
			ratio1=89.5/59.0;
 		break;
	case 1024*768:ratio=6.0/5.0;
 		          ratio1=69.7/59.0;
 		break;
 	case 1152*864:ratio=24.0/22.0;
 		          ratio1=98.0/93.0;
 		break;
 	case 1280*1024:ratio=1.0;
 		           ratio1=98.0/103.5;
 		break;
	}
	//数据库中以分辨率800*600为标准.

	xRatio = (float)((8.0/5.0) / ratio);
	yRatio = (float)((89.5/59.0) / ratio1);
	
	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/11/23]
// Parameter(s) :
// Return       :
// Remark       :刷新浏览结果的对话框
//------------------------------------------------------------------
BOOL CDataShowDlg::UpdateBrowWindow()
{
	if (!IsWindow(pCurDataGrid->GetSafeHwnd()))
	{
		return FALSE;
	}
	pCurDataGrid->SetRefDataSource((LPUNKNOWN)m_IRecordset);	

	if(!GetDlgCaption().IsEmpty())
		SetWindowText(GetDlgCaption());

	//取消 DataGrid控件中显示的标题,该标题名称原来是数据库中表的名称. [2005/06/20]

//	if(!GetDataGridCaption().IsEmpty())
//		m_ResultBrowseControl.SetCaption(GetDataGridCaption());
	
	this->ShowWindow(SW_MAXIMIZE);
//	RECT rect;
//	GetClientRect(&rect);
//	pCurDataGrid->MoveWindow(0,0,rect.right-rect.left,rect.bottom-rect.top);

	pCurDataGrid->SetAllowAddNew(m_IsAllowAddNew);
	pCurDataGrid->SetAllowUpdate(m_IsAllowUpdate);
	pCurDataGrid->SetAllowDelete(m_IsAllowDel);

	//设置字段的宽度和标题
	this->SetFieldCaption( GetDataGridCaption() );

	try
	{
		HideColumns();
	}
	catch(COleDispatchException *e)
	{
		ReportExceptionErrorLV1(e);
		e->Delete();
	}

	pCurDataGrid->InitMouseWheel();

	//初始光标定位到指定的列.
	if( m_nCursorCol >= 0 )
	{
		this->FindDataGridCol(m_nCursorCol);
	}
	return TRUE;
}
