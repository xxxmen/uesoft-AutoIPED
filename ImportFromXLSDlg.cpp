// ImportFromXLSDlg.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "ImportFromXLSDlg.h"
#include "vtot.h"
#include "projectoperate.h"
#include "importautopd.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CImportFromXLSDlg dialog

extern CAutoIPEDApp theApp;

CImportFromXLSDlg::CImportFromXLSDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CImportFromXLSDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CImportFromXLSDlg)
	m_XLSFilePath = _T("");
	m_HintInformation = _T("");
	//}}AFX_DATA_INIT

	m_strItemTblName = _T("");
	m_strTitleName = _T("");
	//设置默认的子键的名称
	m_RegSubKey=_T("ImportFromXLSDlg");
}

CImportFromXLSDlg::~CImportFromXLSDlg()
{
	CPropertyWnd::ElementDef Element;

	//
	// 释放用于保存提示信息的内存
	//

	for(int i=0;i<this->m_PropertyWnd.GetElementCount();i++)
	{
		this->m_PropertyWnd.GetElementAt(&Element,i);

		if(Element.pVoid)
		{
			delete ((CString*)Element.pVoid);
			Element.pVoid=NULL;
		}
	}
}

void CImportFromXLSDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CImportFromXLSDlg)
	DDX_Control(pDX, IDC_OPEN_FILEDLG, m_OpenFileDlgButton);
	DDX_Text(pDX, IDC_FILE_FULLPATH, m_XLSFilePath);
	DDX_Text(pDX, IDC_HINT_INFORMATION, m_HintInformation);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CImportFromXLSDlg, CDialog)
	//{{AFX_MSG_MAP(CImportFromXLSDlg)
	ON_BN_CLICKED(IDC_OPEN_FILEDLG, OnOpenFiledlg)
	ON_WM_DESTROY()
	ON_BN_CLICKED(IDC_BEGIN_IMPORT, OnBeginImport)
	ON_BN_CLICKED(IDC_SET_DEFAULT_NO, OnSetDefaultNo)
	//}}AFX_MSG_MAP
	ON_NOTIFY(PWN_SELECTCHANGE,IDC_PROPERTY_WND, OnSelectChange)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CImportFromXLSDlg message handlers

BOOL CImportFromXLSDlg::OnInitDialog() 
{
	if ( m_pConIpedSort == NULL )
	{
		AfxMessageBox(" 没有设置数据库连接 ！");
		return FALSE;
	}
	CBitmap Bitmap;

	CDialog::OnInitDialog();

	Bitmap.LoadBitmap(IDB_OPENFILE);
	this->m_OpenFileDlgButton.SetBitmap((HBITMAP)Bitmap.Detach());


	CRect Rect;

	//
	// 以隐藏的Group控件来定位属性控件
	//
	GetDlgItem(IDC_PROP_WND_RECT)->GetWindowRect(&Rect);

	this->ScreenToClient(&Rect);

	m_PropertyWnd.Create(NULL,NULL,WS_CHILD|WS_VISIBLE|WS_BORDER|WS_TABSTOP,Rect,this,IDC_PROPERTY_WND);

	//初始化属性控件
	InitPropertyWnd();
	
	//从注册表获得初始信息
	InitFromReg();

	UpdateData( FALSE );

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

////////////////////////////////////////////////
//
// 响应“选择文件按钮”
//
void CImportFromXLSDlg::OnOpenFiledlg() 
{
	BOOL bRet;
	LPDISPATCH pDispatch;
	_Application Application;
	Workbooks workbooks;
	_Workbook workbook;
	Worksheets worksheets;
	_Worksheet worksheet;

	if( m_strPrecFilePath.IsEmpty() )
	{
		m_strPrecFilePath = theApp.GetRegKey(_T("Directory"),
			_T("EDinBox_templateDir"), CString(_T("")));		//默认为模板文件夹的路径.
	}

	CFileDialog  FileDlg(TRUE,NULL,NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
						 _T("Microsoft Excel 文件 (*.xls)|*.xls|All Files (*.*)|*.*||"));

	FileDlg.m_ofn.lpstrInitialDir = m_strPrecFilePath;			 //05/1/4
	// 如果选择了XLS文件就打开它，并获得工作表名
	//
	if(FileDlg.DoModal()==IDOK)
	{
		CWaitCursor WaitCursor;

		this->m_XLSFilePath=FileDlg.GetPathName();
		this->UpdateData(FALSE);

		bRet=Application.CreateDispatch(_T("Excel.Application"));

		if(!bRet)
		{
			ReportExceptionErrorLV1(_T("请先安装Excel!"));
			return;
		}

		//获得工作薄的集合
		pDispatch=Application.GetWorkbooks();
		workbooks.AttachDispatch(pDispatch);

		try
		{
			//打开指定的EXCEL文件
			pDispatch=workbooks.Open(this->m_XLSFilePath,vtMissing,vtMissing,vtMissing,vtMissing,
									 vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
									 vtMissing);
			workbook.AttachDispatch(pDispatch);

			//获得指定工作簿
			pDispatch=workbook.GetWorksheets();
			worksheets.AttachDispatch(pDispatch);

			long iCount;
			int iRet;
			CString strTemp, strSheetName;
			int iPos;
			CPropertyWnd::ElementDef ElementInfo;	//用以返回属性控件单元的信息

			iCount=worksheets.GetCount();


			//返回"Excel工作表名"单元的信息
			iPos=this->m_PropertyWnd.FindElement(0,_T("Excel工作表名"));
			this->m_PropertyWnd.CreateElementControl(iPos);
			this->m_PropertyWnd.GetElementAt(&ElementInfo,iPos);
			//获得当前状态的工作表名   1/11     
			strSheetName = ElementInfo.RightElementText;
			//

			do
			{
				iRet=((CComboBox*)ElementInfo.pControlWnd)->DeleteString(0);
			}while(iRet!=CB_ERR);

			for(long i=1;i<=iCount;i++)
			{
				pDispatch=worksheets.GetItem(_variant_t(i));
				worksheet.AttachDispatch(pDispatch);

				strTemp=worksheet.GetName();

				worksheet.ReleaseDispatch();

				((CComboBox*)ElementInfo.pControlWnd)->AddString(strTemp);
			}
			///////获得指定的工作表名      05/1/4	

			int pos = ((CComboBox*)ElementInfo.pControlWnd)->FindString(0, strSheetName);
			pos = (pos>0)?pos:0;
			((CComboBox*)ElementInfo.pControlWnd)->SetCurSel(pos);
			//////05/1/4

			//重绘属性控件的窗口
			this->m_PropertyWnd.Invalidate();
		}
		catch(COleDispatchException *e)
		{
			Application.Quit();

			ReportExceptionErrorLV2(e);
			throw;
		}
		catch(COleException *e)
		{
			Application.Quit();

			_com_error e2(e->m_sc);
			e->Delete();
			ReportExceptionErrorLV2(e2);
			throw e2;
		}

		Application.Quit();
		///当前的05/1/4
		m_strPrecFilePath = m_XLSFilePath;
	}
}

void CImportFromXLSDlg::OnDestroy() 
{
	CDialog::OnDestroy();

	//关闭窗口时保存数据到注册表
	SaveSetToReg();
}

//////////////////////////////////////////////////////////////

///////////////////////////////////////////
//
// 初始化属性窗口
//
// 函数成功返回TRUE，否则返回FALSE
//
// 此函数应该在WM_INITDIALOG消息中调用
//
BOOL CImportFromXLSDlg::InitPropertyWnd()
{
	//
	// 初始化属性控件
	//
	CPropertyWnd::ElementDef ElementDef;

	struct
	{
		LPCTSTR ElementName;	// 单元的名称
		UINT	Style;			// 显示样式
		LPCTSTR HintInfo;		// 与单元相关的提示信息
	}ElementSet[]=
	{
		{_T("Excel"),			CPropertyWnd::TitleElement,NULL},								//0

		{_T("Excel工作表名"),	CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择将哪张表导入进来")},	//1

		{_T("导入数据开始行号"),CPropertyWnd::ChildElement|CPropertyWnd::EditElement,NULL},		//2
		{_T("导入多少条数据"),	CPropertyWnd::ChildElement|CPropertyWnd::EditElement,NULL},		//3
	};


	for(int i=0;i<sizeof(ElementSet)/sizeof(ElementSet[0]);i++)
	{
		ElementDef.szElementName=ElementSet[i].ElementName;
		ElementDef.Style=ElementSet[i].Style;

		ElementDef.pVoid=NULL;

		if(ElementSet[i].HintInfo!=NULL)
		{			
			ElementDef.pVoid=new CString(ElementSet[i].HintInfo);
		}

		m_PropertyWnd.InsertElement(&ElementDef);

	}

	m_PropertyWnd.RefreshData();

	if ( !m_strItemTblName.IsEmpty() )
	{	
		// 从数据表中取出属性名，动态增加到窗口中
		AddPropertyWndFromTbl( m_strTitleName, m_strItemTblName );
	}
	return TRUE;
}

///////////////////////////////////////////
//
// 从注册表获得数据初始化
//
void CImportFromXLSDlg::InitFromReg()
{
	LONG lRet;
	DWORD dwDisposition;
	HKEY hKey;
	CString SubKey;
	CPropertyWnd::ElementDef ElementDef;
	TCHAR szTemp[256];
	DWORD dwData;

	SubKey=szSoftwareKey;
	SubKey+=m_RegSubKey;

	//打开注册表
	lRet=::RegCreateKeyEx(HKEY_LOCAL_MACHINE,SubKey,0,NULL,REG_OPTION_NON_VOLATILE,KEY_READ,NULL,&hKey,&dwDisposition);
	
	if(lRet!=ERROR_SUCCESS)
	{
		ReportExceptionErrorLV1(_T("打开注册表失败"));
		return;
	}
	// 从注册表获得数据放入属性控件
	//
//	for(int i=0;i<this->m_PropertyWnd.GetElementCount();i++)
	int nPos = this->m_PropertyWnd.FindElement(0, "导入多少条数据");
	if ( nPos < 0 )
		nPos = 3;
	
	for( int i = 0; i <= nPos; i++ )
	{
		this->m_PropertyWnd.CreateElementControl(i);
		this->m_PropertyWnd.GetElementAt(&ElementDef,i);

		dwData=256;

		lRet=RegQueryValueEx(hKey,ElementDef.szElementName,NULL,NULL,(BYTE*)szTemp,&dwData);

		if(lRet==ERROR_SUCCESS)
		{
			if(ElementDef.pControlWnd && IsWindow(ElementDef.pControlWnd->GetSafeHwnd()))
			{
				if(dwData>0)
					ElementDef.pControlWnd->SetWindowText(szTemp);
			}
		}
	}
	/////////  初始Excel 文件路径
	dwData=256;
	lRet=RegQueryValueEx(hKey,"ExcelFilePath",NULL,NULL,(BYTE*)szTemp,&dwData);  

	if(lRet==ERROR_SUCCESS)
	{
		m_strPrecFilePath = szTemp;
		m_XLSFilePath = m_strPrecFilePath;
	}
	//////
	::RegCloseKey(hKey);
}

////////////////////////////////////////////
//
// 将信息保存到注册表内
//
void CImportFromXLSDlg::SaveSetToReg()
{
	UpdateData();
	
	LONG lRet;
	DWORD dwDisposition;
	HKEY hKey;
	CString SubKey;
	CPropertyWnd::ElementDef ElementDef;

	SubKey=szSoftwareKey;
	SubKey+=m_RegSubKey;

	// 打开注册表
	lRet=::RegCreateKeyEx(HKEY_LOCAL_MACHINE,SubKey,0,NULL,REG_OPTION_NON_VOLATILE,KEY_WRITE,NULL,&hKey,&dwDisposition);
	
	if(lRet!=ERROR_SUCCESS)
	{
		ReportExceptionErrorLV1(_T("打开注册表失败"));
		return;
	}

	//
	// 将属性控件的内容存如注册表
	//
//	for(int i=0;i<this->m_PropertyWnd.GetElementCount();i++)
	int nPos = this->m_PropertyWnd.FindElement(0, "导入多少条数据");
	if ( nPos < 0 )
		nPos = 3;

	for( int i = 0; i <= nPos; i++ )
	{
		this->m_PropertyWnd.GetElementAt(&ElementDef,i);

		if(!ElementDef.RightElementText.IsEmpty())
		{
			lRet=RegSetValueEx(hKey,ElementDef.szElementName,NULL,REG_SZ,(BYTE*)(LPCTSTR)ElementDef.RightElementText,
						  strlen(ElementDef.RightElementText)+sizeof(TCHAR));
		}
		else
		{
			lRet=RegSetValueEx(hKey,ElementDef.szElementName,NULL,REG_SZ,(BYTE*)NULL,
						  0);
		}
	}
	///////  保存路径
	if ( !m_XLSFilePath.IsEmpty() )
	{
		lRet=RegSetValueEx(hKey,"ExcelFilePath",NULL,REG_SZ,(BYTE*)(LPCTSTR)m_strPrecFilePath,
				  strlen(m_XLSFilePath)+sizeof(TCHAR));
	}
	//////
	::RegCloseKey(hKey);
	
	if ( !m_strItemTblName.IsEmpty() )
	{
		// 将子控件对应的EXCEL列号写入到数据库中
		WriteExcelNoToTbl( m_strItemTblName );
	}
}


/////////////////////////////////////////////////////////////////////
//
// 响应属性控件选中的通知
//
void CImportFromXLSDlg::OnSelectChange(NMHDR *pNMHDR, LRESULT *pResult)
{
	CPropertyWnd::PWNSelectChangeStruct *pSelect;

	pSelect=(CPropertyWnd::PWNSelectChangeStruct*)pNMHDR;

	this->m_HintInformation=(pSelect->szElementName);

	if(pSelect->pVoid)
	{
		this->m_HintInformation+=_T(":\r\n");

		this->m_HintInformation+=*((CString*)pSelect->pVoid);
	}

	this->UpdateData(FALSE);
	*pResult=0;
}

///////////////////////////////////////////
//
// 响应开始”导入“
//
void CImportFromXLSDlg::OnBeginImport() 
{
	BOOL bRet;

	this->UpdateData();

	// 验证数据输入 是否有效
	bRet=ValidateData();

	if(!bRet)
	{
		return;
	}

	BeginImport();

	// 对导入的记录进行处理
	CImportAutoPD::RefreshData( m_strProjectID );
}

/////////////////////////////////////////////
//
// 验证输入数据的有效性
//
BOOL CImportFromXLSDlg::ValidateData()
{
	CPropertyWnd::ElementDef ElementDef;
	CString strTemp;

	this->UpdateData();

	if(GetProjectID().IsEmpty())
	{
		ReportExceptionErrorLV1(_T("未选择工程"));

		return FALSE;
	}

	//判断文件路径是否为空
	if(this->m_XLSFilePath.IsEmpty())
	{
		ReportExceptionErrorLV1(_T("文件路径不能为空"));
		return FALSE;
	}

	//
	// 检查属性控件中指定的单元是否为空
	//
	for(int i=0;i<this->m_PropertyWnd.GetElementCount();i++)
	{
		LPCTSTR szTableName[]=
		{
			_T("Excel工作表名"),
			_T("导入数据开始行号"),
			_T("导入多少条数据"),
		};

		this->m_PropertyWnd.GetElementAt(&ElementDef,i);

		for(int j=0;j<sizeof(szTableName)/sizeof(szTableName[0]);j++)
		{
			if(ElementDef.szElementName==szTableName[j] && ElementDef.RightElementText.IsEmpty())
			{
				strTemp.Format(_T("%s不能为空"),szTableName[j]); 
				ReportExceptionErrorLV1(strTemp); 
				return FALSE; 
			}
		}

	}

	return TRUE;
}

//////////////////////////////////////////////////////////
//
// 设置工程编号
//
// szProjectID[in]	工程编号
//
void CImportFromXLSDlg::SetProjectID(LPCTSTR szProjectID)
{
	if(szProjectID==NULL)
	{
		m_strProjectID=_T("");
	}
	else
	{
		m_strProjectID=szProjectID;
	}
}

//////////////////////////////////////////////////////
//
// 返回工程的编号
//
CString CImportFromXLSDlg::GetProjectID()
{
	return m_strProjectID;
}

//------------------------------------------------------------------
// DATE         :[2005/12/27]
// Parameter(s) :数据库的联接
// Return       :
// Remark       :设置控制属性的数据库
//------------------------------------------------------------------
void CImportFromXLSDlg::SetIPEDSortDBConnect(_ConnectionPtr pConSort)
{
	this->m_pConIpedSort = pConSort;
}

//------------------------------------------------------------------
// DATE         :[2005/12/27]
// Parameter(s) :
// Return       :
// Remark       :获得控制属性的数据库
//------------------------------------------------------------------
_ConnectionPtr CImportFromXLSDlg::GetIPEDSortDBConnect()
{
	return this->m_pConIpedSort;
}

//////////////////////////////////////////////////////////////////////
//
// 设置与工程相关数据库的连接
//
// IConnection[in]	数据库的连接
//
void CImportFromXLSDlg::SetProjectDbConnect(_ConnectionPtr IConnection)
{
	this->m_ProjectDbConnect=IConnection;
}

////////////////////////////////////////////////////////////////////
//
// 返回与数据库的连接
//
_ConnectionPtr CImportFromXLSDlg::GetProjectDbConnect()
{
	return this->m_ProjectDbConnect;
}

//------------------------------------------------------------------
// DATE         :[2005/12/28]
// Parameter(s) :lpTitleName[in] 子控件的标题, lpTblName[in] 控制子控件的表名
// Return       :
// Remark       :设置子控件的标题和 控制子控件的表名
//------------------------------------------------------------------
void CImportFromXLSDlg::SetItemTblName(LPCTSTR lpTitleName, LPCTSTR lpTblName)
{
	m_strTitleName = lpTitleName;
	m_strItemTblName = lpTblName;
}

//------------------------------------------------------------------
// DATE         :[2005/12/29]
// Parameter(s) :
// Return       :
// Remark       :导入的目的数据表名
//------------------------------------------------------------------
void CImportFromXLSDlg::SetDataTblName(LPCTSTR lpDataTblName)
{
	m_strDataTblName = lpDataTblName;
}

///////////////////////////////////////////////////////////
//
// 设置用于存放信息的子键名
//
// szRegSubKey[in]	用于存放信息的子键名
//
void CImportFromXLSDlg::SetRegSubKey(LPCTSTR szRegSubKey)
{
	if(szRegSubKey)
		m_RegSubKey = szRegSubKey;
}

////////////////////////////////////////////////////////
//
// 返回用于存放信息的子键名
//
CString CImportFromXLSDlg::GetRegSubKey()
{
	return m_RegSubKey;
}

///////////////////////////////////////////////////////////
//
// 设置需在对话框下部提示信息区显示的内容
//
// szHint[in]	在对话框下部提示信息区显示的内容
//
void CImportFromXLSDlg::SetHintInformation(LPCTSTR szHint)
{
	if(szHint)
		m_HintInformation=szHint;
	else
		m_HintInformation=_T("");

	if(IsWindow(GetSafeHwnd()))
	{
		UpdateData(FALSE);
	}
}

////////////////////////////////////////////////////////
//
// 返回需在对话框下部提示信息区显示的内容
//
CString CImportFromXLSDlg::GetHintInformation()
{
	return m_HintInformation;
}

////////////////////////////////////////////////////////
//
// 获得属性控件的窗口
//
CPropertyWnd* CImportFromXLSDlg::GetPropertyWnd()
{
	return &m_PropertyWnd;
}

//////////////////////////////////////////////////
//
// 执行导入数据的操作
//
void CImportFromXLSDlg::BeginImport()
{
	int iPos;
	CPropertyWnd::ElementDef ElementDef;
	CProjectOperate::ImportFromXLSStruct ImportTable;
	CProjectOperate::ImportFromXLSElement ImportTableItem[60];
	int TableItemCount;
	CProjectOperate Import;

	CWaitCursor WaitCursor;
	BOOL	bIsAutoNO = FALSE;		// 对记录自动编号

	this->m_HintInformation=_T("开始导入数据");
	this->UpdateData(FALSE);
	this->UpdateWindow();

	Import.SetProjectDbConnect(this->GetProjectDbConnect());

	//
	// 填写ImportFromXLSStruct结构信息
	//
	if ( m_strDataTblName.IsEmpty() )
	{
		AfxMessageBox("请指定数据表名！");
		return ;
	}
	ImportTable.ProjectDBTableName = m_strDataTblName;

	ImportTable.ProjectID=this->GetProjectID();

	ImportTable.XLSFilePath=this->m_XLSFilePath;

	iPos=this->m_PropertyWnd.FindElement(0,_T("导入数据开始行号"));
	this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);
	ImportTable.BeginRow=_ttoi(ElementDef.RightElementText);

	iPos=this->m_PropertyWnd.FindElement(0,_T("Excel工作表名"));
	this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);
	ImportTable.XLSTableName=ElementDef.RightElementText;

	iPos=this->m_PropertyWnd.FindElement(0,_T("导入多少条数据"));
	this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);
	ImportTable.RowCount=_ttoi(ElementDef.RightElementText);
	//
	// 将属性控件中的内容添入ImportTableItem结构
	//
	iPos++;
	TableItemCount = 0;
	while(TRUE)
	{
		if(iPos>=this->m_PropertyWnd.GetElementCount())
			break;

		this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);
		iPos++;

		if ( ElementDef.szElementName == _T("序号") && ElementDef.RightElementText == _T("自动") )
		{
			bIsAutoNO = TRUE;
			continue;
		}
		
		// 当右单元不为空时才添入
		if( !ElementDef.RightElementText.IsEmpty() )
		{
			ImportTableItem[TableItemCount].SourceFieldName = ElementDef.RightElementText;
			ImportTableItem[TableItemCount].DestinationName = ElementDef.strFieldName;
			TableItemCount++;
		}
	}

	ImportTable.pElement=ImportTableItem;
	ImportTable.ElementNum=TableItemCount;
	try
	{
		//
		// 如果在”序号“里添了”自动“则ID号自动编写
		// 对于一个表没有ID字段,则不能用下面的函数
		//Import.ImportPre_CalcFromXLS( &ImportTable, bIsAutoNO );
		Import.ImportTblEnginIDXLS( &ImportTable, bIsAutoNO );
	}
	catch(_com_error &e)
	{
		this->m_HintInformation=_T("数据导入失败");
		this->UpdateData(FALSE);
		ReportExceptionErrorLV1(e);		
		return;
	}
	catch(COleDispatchException *e)
	{
		this->m_HintInformation=_T("数据导入失败");
		this->UpdateData(FALSE);
		ReportExceptionErrorLV1(e);
		e->Delete();
		return;
	}

	this->m_HintInformation=_T("数据导入成功");
	this->UpdateData(FALSE);
}

//------------------------------------------------------------------
// DATE         :[2005/12/27]
// Parameter(s) :const CString strTblName, 在IPEDsort.mdb库中丰放不同属性的表名
// Return       :调用成功返回TRUE，否则返回FALSE
// Remark       :从数据表中取出属性名，动态增加到窗口中
//------------------------------------------------------------------
BOOL CImportFromXLSDlg::AddPropertyWndFromTbl(LPCTSTR lpTitleName, const CString strTblName)
{
	if ( strTblName.IsEmpty() )
		return	FALSE;

	CString strSQL;		// SQL语句	
	CString strTmp;
	_RecordsetPtr pRsInfo;		// 存入属性的记录集
	CPropertyWnd::ElementDef ElementDef;	// 子控件
	CHAR ColumeName[10];
	int  iPos;
	pRsInfo.CreateInstance( __uuidof( Recordset ) );
	EDIBgbl::CAPTION2FIELD * pAllFieldStruct = NULL;	// 子窗口的属性数组
	int nWndCount = GetField2Caption( pAllFieldStruct, strTblName );	// 子窗口的属性个数
	if ( nWndCount <= 0 )
	{
		return FALSE;
	}
	// 首先创建子控件的标题
	if ( NULL != lpTitleName )
	{
		ElementDef.szElementName = lpTitleName;
		ElementDef.Style = CPropertyWnd::TitleElement;
		this->m_PropertyWnd.InsertElement( &ElementDef );
	}
	//创建子控件
	for ( int i = 0; i < nWndCount; i++ )
	{
		//原始数据表中对应的字段名称		
		if ( pAllFieldStruct[i].strField.IsEmpty() )
			continue;
		ElementDef.strFieldName = pAllFieldStruct[i].strField;

		//字段描述
		if ( pAllFieldStruct[i].strCaption.IsEmpty() )
			continue;
		ElementDef.szElementName = pAllFieldStruct[i].strCaption;
		//提示信息
		if ( pAllFieldStruct[i].strRemark.IsEmpty() )
			strTmp.Format(IDS_IMPORT_ITEM_REMARK, pAllFieldStruct[i].strCaption);
		else
			strTmp = pAllFieldStruct[i].strRemark;
		ElementDef.pVoid = new CString(strTmp);
		//风格
		ElementDef.Style = CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement;			
		
		//默认的EXCEL列号
		ElementDef.RightElementText = pAllFieldStruct[i].strExcelNO;
		
		//插入一项
		iPos = m_PropertyWnd.InsertElement(&ElementDef);
		//创建子控件
		if ( m_PropertyWnd.CreateElementControl(iPos) )
		{
			this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);

			if ( !ElementDef.szElementName.Compare("序号") )
			{
				((CComboBox*)ElementDef.pControlWnd)->AddString("自动");
			}
			
			//右边提供选的EXCEL列号
			for(int k = 0; k < 100; k++)
			{
				int j,m,m2;
				for(j=k,m=1;j/26!=0;m++,j/=26);
				ColumeName[m]=_T('\0');
				m--;
				j=k;
				m2=m;
				while(m>=0)
				{
					ColumeName[m]=j%26+_T('A');					
					j=j/26;					
					if(m2!=m)
						ColumeName[m]--;					
					m--;
				}
				((CComboBox*)ElementDef.pControlWnd)->AddString(ColumeName);
			}	
		}
	}
	// 重新编号
	this->m_PropertyWnd.RefreshData();
	
	if ( NULL != pAllFieldStruct )
	{
		delete [] pAllFieldStruct;
	}

	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/12/27]
// Parameter(s) :EDIBgbl::CAPTION2FIELD* &pFieldStruct[out]	从数据库中取出的字段对应关系存于结构数组中。
//				:const CString strTblName[in]	保存子控件的结构表名
//				:const CString strDefField[in]	保存Excel默认列号的字段落名
// Return       :返回字段个数
// Remark       :从IPEDsort.mdb的PD2IPED表中取出EXCEL中的字段名和ACCESS中的字段的对应值
//------------------------------------------------------------------
int CImportFromXLSDlg::GetField2Caption(EDIBgbl::CAPTION2FIELD *&pFieldStruct, const CString strTblName, const CString strDefField)
{
	_RecordsetPtr pRs;
	CString strSQL;
	int		nFieldCount;	//字段的个数    
	pRs.CreateInstance(__uuidof(Recordset));
	try
	{        
		strSQL = " SELECT * FROM ["+strTblName+"] WHERE SEQ IS NOT NULL ORDER BY SEQ";
		
		pRs->Open(_variant_t(strSQL), m_pConIpedSort.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		nFieldCount = pRs->GetRecordCount();
		if (nFieldCount <= 0)
		{
			return FALSE;
		} 
		pFieldStruct = new EDIBgbl::CAPTION2FIELD[nFieldCount];
		for (int i=0; !pRs->adoEOF && i < nFieldCount; pRs->MoveNext(),i++)
		{
			pFieldStruct[i].strCaption = vtos(pRs->GetCollect("LocalCaption"));		//EXCEL中的字段描述
			pFieldStruct[i].strField   = vtos(pRs->GetCollect("FieldName"));		//ACCESS中的字段名称
			pFieldStruct[i].strExcelNO = vtos(pRs->GetCollect(_variant_t(strDefField)));		// 对应Excel的列号
			pFieldStruct[i].strRemark  = vtos(pRs->GetCollect("Remark"));			// 备注信息
			pFieldStruct[i].nFieldType = vtoi(pRs->GetCollect("FieldType"));		//字段类型
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return nFieldCount;
}
//------------------------------------------------------------------
// DATE         :[2005/12/28]
// Parameter(s) :数据库的表名
// Return       :
// Remark       :将当前对应的Excel列号,写入到数据库中
//------------------------------------------------------------------
BOOL CImportFromXLSDlg::WriteExcelNoToTbl(const CString strTblName)
{
	if ( strTblName.IsEmpty() )
		return FALSE;
	CString strSQL;			// SQL语句
	CString strTmp;			// 临时字符串
	int iPos;
	CPropertyWnd::ElementDef ElementDef;
	_RecordsetPtr pRsInfo;	// 存入属性的记录集
	pRsInfo.CreateInstance( __uuidof( Recordset ) );

	try
	{
		strSQL = "SELECT * FROM ["+strTblName+"] ";
		pRsInfo->Open( _variant_t( strSQL ), m_pConIpedSort.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText );
		if ( pRsInfo->adoEOF && pRsInfo->BOF )
		{
			return FALSE;
		}
		for ( pRsInfo->MoveFirst(); !pRsInfo->adoEOF; pRsInfo->MoveNext() )
		{
			strTmp = vtos( pRsInfo->GetCollect( _variant_t("LocalCaption") ) );
			if ( strTmp.IsEmpty() )
				continue;

			iPos = this->m_PropertyWnd.FindElement( 0, strTmp );
			if ( iPos == -1 )
				continue;
			
			this->m_PropertyWnd.GetElementAt( &ElementDef, iPos );

			//写入到数据库
			pRsInfo->PutCollect( _variant_t("ExcelColNO"), _variant_t(ElementDef.RightElementText) );
			pRsInfo->Update();
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return FALSE;
	}
	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/12/30]
// Parameter(s) :
// Return       :
// Remark       :
//------------------------------------------------------------------
void CImportFromXLSDlg::SetDefaultExcelNO(LPCTSTR lpItemTblName)
{
	if ( NULL == lpItemTblName )
		return;
	
	CPropertyWnd::ElementDef ElementDef;
	EDIBgbl::CAPTION2FIELD* pFieldStruct;	// 子控件的结构
	int nFieldCount = GetField2Caption( pFieldStruct, lpItemTblName, "DefExcelColNO" );
	int iPos;

	for ( int i = 0; i < nFieldCount ; i++ )
	{
		iPos = this->m_PropertyWnd.FindElement( 0, pFieldStruct[i].strCaption );
		if ( -1 != iPos )
		{	
			m_PropertyWnd.SetRightElementContent( iPos, pFieldStruct[i].strExcelNO );
		}
	}

	if ( NULL != pFieldStruct )
	{
		delete [] pFieldStruct;
	}
	UpdateData( FALSE );
}

void CImportFromXLSDlg::OnSetDefaultNo() 
{
	SetDefaultExcelNO( m_strItemTblName );
}
