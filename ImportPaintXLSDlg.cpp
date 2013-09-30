// ImportPaintXLSDlg.cpp: implementation of the CImportPaintXLSDlg class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "ImportPaintXLSDlg.h"
#include "ProjectOperate.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CImportPaintXLSDlg::CImportPaintXLSDlg()
{
	this->SetRegSubKey("ImportPaintXLSDlg");
}

CImportPaintXLSDlg::~CImportPaintXLSDlg()
{

}
BOOL CImportPaintXLSDlg::OnInitDialog()
{
	CImportFromXLSDlg::OnInitDialog();
	this->SetWindowText("从Excel导入油漆原始数据");

	return TRUE;
}

BOOL CImportPaintXLSDlg::InitPropertyWnd()
{
	CImportFromXLSDlg::InitPropertyWnd();
	
	//  zsy
	return TRUE;
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
		{_T("油漆原始数据"),	CPropertyWnd::TitleElement,NULL},								//4

		{_T("序号"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("如果选择自动，导入的数据将会在原始数据表的末尾加入\r\n否则需要选择在Excel对应的列的数据作为序号")
		 _T("但此时如果原始数据表内已有相同的序号的记录将会被覆盖")},	//5

		{_T("卷册号"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与卷册号对应的Excel列作为输入")},

		{_T("管道/设备名称"),	CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与管道/设备名称对应的Excel列作为输入")},

		{_T("管道外径"),	CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与管道外径对应的Excel列作为输入")},

		{_T("设备外表面积"),	CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与设备外表面积对应的Excel列作为输入")},

		{_T("管道长度/设备数"),	CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与管道长度/设备数对应的Excel列作为输入")},

		{_T("油漆颜色"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与油漆颜色对应的Excel列作为输入")},

		{_T("油漆类型"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与油漆类型对应的Excel列作为输入")},

		{_T("备注"),			CPropertyWnd::ChildElement|CPropertyWnd::ComBoxElement,
		 _T("请选择与备注对应的Excel列作为输入")},

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
	// 从“油漆原始数据”以后的单元项中添加EXCEL列名的信息
	//
	TCHAR ColumeName[128];
	int iPos;

	for(i=1;i<sizeof(ElementSet)/sizeof(ElementSet[0]);i++)
	{
		iPos=this->m_PropertyWnd.FindElement(0,ElementSet[i].ElementName);

		this->m_PropertyWnd.CreateElementControl(iPos);

		this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);

		if(ElementDef.szElementName==_T("序号"))
		{
			((CComboBox*)ElementDef.pControlWnd)->AddString(_T("自动"));
		}

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


void CImportPaintXLSDlg::BeginImport()
{
	int iPos;
	CPropertyWnd::ElementDef ElementDef;
	CProjectOperate::ImportFromXLSStruct ImportTable;
	CProjectOperate::ImportFromXLSElement ImportTableItem[50];
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
	ImportTable.ProjectDBTableName=_T("paint");

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
		if(!ElementDef.RightElementText.IsEmpty())
		{
			ImportTableItem[TableItemCount].SourceFieldName = ElementDef.RightElementText;

			ImportTableItem[TableItemCount].DestinationName = ElementDef.strFieldName;
			TableItemCount++;
		}
	}

	ImportTable.pElement = ImportTableItem;
	ImportTable.ElementNum = TableItemCount;


	try
	{
		//
		// 如果在”序号“里添了”自动“则ID号自动编写
		// 考虑了不同的工程代号
//		Import.ImportPre_CalcFromXLS( &ImportTable, bIsAutoNO );

		Import.ImportTblEnginIDXLS( &ImportTable, bIsAutoNO );

/*		iPos=this->m_PropertyWnd.FindElement(0,_T("序号"));
		this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);
		if(ElementDef.RightElementText==_T("自动"))
		{
			Import.ImportPre_CalcFromXLS(&ImportTable,TRUE);
		}
		else
		{
			if(!ElementDef.RightElementText.IsEmpty())
			{
				ImportTableItem[TableItemCount].SourceFieldName=ElementDef.RightElementText;
				ImportTableItem[TableItemCount].DestinationName=_T("id");
				TableItemCount++;

				ImportTable.ElementNum=TableItemCount;

				Import.ImportPre_CalcFromXLS(&ImportTable,FALSE);
			}
		}
*/
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
/////////////////////////////////////////////
//
// 验证输入数据的有效性
//
BOOL CImportPaintXLSDlg::ValidateData()
{
	BOOL bRet;
	CPropertyWnd::ElementDef ElementDef;
	CString strTemp;
	int iPos;

	bRet=CImportFromXLSDlg::ValidateData();

	if(!bRet)
		return FALSE;

	iPos=this->m_PropertyWnd.FindElement(0,_T("序号"));
	this->m_PropertyWnd.GetElementAt(&ElementDef,iPos);

	if(ElementDef.RightElementText.IsEmpty())
	{
		strTemp.Format(_T("序号不能为空"));
		ReportExceptionErrorLV1(strTemp);
		return FALSE;
	}

	return TRUE;

}
