// ImportDirFromXlsDlg.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "ImportDirFromXlsDlg.h"

#include "ProjectOperate.h"	// 工程操作类,用于导入工程


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CImportDirFromXlsDlg dialog


CImportDirFromXlsDlg::CImportDirFromXlsDlg(CWnd* pParent /*=NULL*/)
	: CImportFromXLSDlg(pParent)
{
	//{{AFX_DATA_INIT(CImportDirFromXlsDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT

	this->SetRegSubKey(_T("ImportDirFromXlsDlg"));

}


void CImportDirFromXlsDlg::DoDataExchange(CDataExchange* pDX)
{
	CImportFromXLSDlg::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CImportDirFromXlsDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CImportDirFromXlsDlg, CImportFromXLSDlg)
	//{{AFX_MSG_MAP(CImportDirFromXlsDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CImportDirFromXlsDlg message handlers

BOOL CImportDirFromXlsDlg::InitPropertyWnd()
{
	CImportFromXLSDlg::InitPropertyWnd();

	
	return TRUE;
	//  [2005/12/28]
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
		{_T("卷册原始数据"),	CPropertyWnd::TitleElement,NULL},

		{_T("卷册号"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与卷册号对应的Excel列作为输入")},

		{_T("卷册名称"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与卷册名称对应的Excel列作为输入")},

		{_T("行业ID"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与行业ID对应的Excel列作为输入\r\n注意此列应该输入行业的代号而不是名称")
		 _T("比如\"电力\"的代号是0，此项可不填写默认为0")},

		{_T("设计阶段ID"),	CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与设计阶段ID对应的Excel列作为输入\r\n注意此列应该输入行业的代号而不是名称")
		 _T("比如\"施工图\"的代号是4，此项可不填写默认为4")},

		{_T("专业ID"),	CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与专业ID对应的Excel列作为输入\r\n注意此列应该输入行业的代号而不是名称")
		 _T("比如\"热机\"的代号是3，此项可不填写默认为3")},
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

	//
	// 从“保温原始数据”以后的单元项中添加EXCEL列名的信息
	//
	TCHAR ColumeName[128];
	int iPos;

	for(i=1;i<sizeof(ElementSet)/sizeof(ElementSet[0]);i++)
	{
		iPos=this->m_PropertyWnd.FindElement(0,ElementSet[i].ElementName);

		this->m_PropertyWnd.CreateElementControl(iPos);

		this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);

		for(int k=0;k<100;k++)
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
				{
					ColumeName[m]--;
				}

				m--;
			}

			((CComboBox*)ElementDef.pControlWnd)->AddString(ColumeName);
		}

	}

	return TRUE;
}

/////////////////////////////////////////////////////////
//
// 验证输入的数据是否有效
//
// 如果有效返回TRUE，否则返回FALSE此时不会继续导入数据
//
BOOL CImportDirFromXlsDlg::ValidateData()
{
	BOOL bRet;
	CPropertyWnd::ElementDef ElementDef;
	CString strTemp;
	int iPos;

	bRet=CImportFromXLSDlg::ValidateData();

	if(!bRet)
		return FALSE;

	LPCTSTR szFields[]=
	{
		_T("卷册号"),
		_T("卷册名称")
	};

	for(int i=0;i<sizeof(szFields)/sizeof(szFields[0]);i++)
	{
		iPos=this->m_PropertyWnd.FindElement(0,szFields[i]);
		this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);

		if(ElementDef.RightElementText.IsEmpty())
		{
			strTemp.Format(_T("%s不能为空"),szFields[i]);
			ReportExceptionErrorLV1(strTemp);
			return FALSE;
		}
	}

	return TRUE;
}

/////////////////////////////////////////////////
//
// 开始导入数据
//
// 但此之前ValidateData必须返回TRUE
//
void CImportDirFromXlsDlg::BeginImport()
{
	int iPos;
	CPropertyWnd::ElementDef ElementDef;
	CProjectOperate::ImportFromXLSStruct ImportTable;
	CProjectOperate::ImportFromXLSElement ImportTableItem[50];
	int TableItemCount;
	CProjectOperate Import;

	CWaitCursor WaitCursor;

	//
	// 中文名称与字段的对照表
	//
/*	struct
	{
		LPCTSTR NameCh;		// pre_calc的字段名的中文名称
		LPCTSTR NameField;	// pre_calc的字段名
	}FieldsName[]=
	{
		{_T("卷册号"),			_T("jcdm")},
		{_T("卷册名称"),		_T("jcmc")},
		{_T("行业ID"),			_T("SJHYID")},
		{_T("设计阶段ID"),		_T("SJJDID")},
		{_T("专业ID"),			_T("ZYID")},
	};
*/
	this->SetHintInformation(_T("开始导入数据"));

	Import.SetProjectDbConnect(this->GetProjectDbConnect());

	//
	// 填写ImportFromXLSStruct结构信息
	//
	ImportTable.ProjectDBTableName=_T("volume");

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
	TableItemCount = 0;		// 子控件的个数
	while(TRUE)
	{
		if(iPos>=this->m_PropertyWnd.GetElementCount())
			break;

		this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);
		iPos++;

		// 当右单元不为空时才添入
		if(!ElementDef.RightElementText.IsEmpty())
		{
			ImportTableItem[TableItemCount].SourceFieldName = ElementDef.RightElementText;

			//
			// 通过FieldsName表找到对应的pre_calc的字段名
			//
/*			for(int i=0;i<sizeof(FieldsName)/sizeof(FieldsName[0]);i++)
			{
				if(ElementDef.szElementName==_T(FieldsName[i].NameCh))
					break;
			}

			//如果没找到继续
			if(i==sizeof(FieldsName)/sizeof(FieldsName[0]))
				continue;

			ImportTableItem[TableItemCount].DestinationName=FieldsName[i].NameField;
*/
			ImportTableItem[TableItemCount].DestinationName = ElementDef.strFieldName;
			TableItemCount++;
		}
	}

	ImportTable.pElement=ImportTableItem;
	ImportTable.ElementNum=TableItemCount;

	//
	// 开始导入数据
	//
	try
	{
		Import.ImportA_DirFromXLS(&ImportTable);
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

	this->SetHintInformation(_T("数据导入成功"));
}

BOOL CImportDirFromXlsDlg::OnInitDialog() 
{
	CImportFromXLSDlg::OnInitDialog();
	
	this->SetWindowText(_T("导入卷册目录"));
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
