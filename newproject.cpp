// newproject.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "newproject.h"
#include "math.h"
#include "stdlib.h"
#include "edibgbl.h"
#include "mainfrm.h"
#include "vtot.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern CAutoIPEDApp theApp;

/////////////////////////////////////////////////////////////////////////////
// newproject dialog


newproject::newproject(CWnd* pParent /*=NULL*/)
	: CDialog(newproject::IDD, pParent)
{
	//{{AFX_DATA_INIT(newproject)
	m_eng_code = _T("");
	m_engin = _T("");
	m_strCaption = _T("");
	//}}AFX_DATA_INIT
	m_bAuto = false;
}


void newproject::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(newproject)
	DDX_Control(pDX, IDC_PROGRESS1, c_progress);
	DDX_Control(pDX, IDC_LIST1, c_list);
	DDX_Text(pDX, IDC_EDIT_eng_code, m_eng_code);
	DDX_Text(pDX, IDC_EDIT_engin, m_engin);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(newproject, CDialog)
	//{{AFX_MSG_MAP(newproject)
	ON_BN_CLICKED(IDC_BUTTON_new, OnBUTTONnew)
	ON_NOTIFY(NM_CLICK, IDC_LIST1, OnClickList1)
	ON_BN_CLICKED(IDC_ADDNEWPROJECT, OnAddnewproject)
	ON_BN_CLICKED(IDC_DELETESELECTPROJECT, OnDeleteSelectproject)
	ON_EN_SETFOCUS(IDC_EDIT_eng_code, OnSetfocusEDITengcode)
	ON_EN_SETFOCUS(IDC_EDIT_engin, OnSetfocusEDITengin)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// newproject message handlers

BOOL newproject::OnInitDialog() 
{
	_RecordsetPtr	IRecordset;
	CString str;
    LVITEM  lvItem;
	_variant_t var;
	CString a,b;
	int s;
	HRESULT hr;
	
	CDialog::OnInitDialog();
	if( !m_strCaption.IsEmpty() )
	{
		SetWindowText(m_strCaption);
	}

	//*******************************************88
	//     设置listctrl控件选择一整行的属性
	//*********************************************8
	DWORD dStyleEx = c_list.GetExtendedStyle();
	dStyleEx |= LVS_EX_FULLROWSELECT;
	c_list.SetExtendedStyle( dStyleEx );

	str=_T("工程代号");
	c_list.InsertColumn( 0, (LPCTSTR)str, LVCFMT_LEFT, 80);
	str=_T("工程名称");
	c_list.InsertColumn( 1, (LPCTSTR)str, LVCFMT_LEFT, 220);

	//*********************************************
	//     给listctrl设置初始植
	//***********************************************
	hr=IRecordset.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		OnCancel();
		return TRUE;
	}

	try
	{
		IRecordset->Open("select enginid,gcmc from engin ",                
							theApp.m_pConAllPrj.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
							adOpenDynamic,
							adLockOptimistic,
							adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		OnCancel();
		return TRUE;
	}

	// 读入库中各字段并加入列表框中
	try
	{
		s=0;
		if(IRecordset->adoEOF && IRecordset->BOF)
		{
			if( m_strCaption.IsEmpty() )
			{
				SetWindowText("新建工程");
			}
			m_str4=_T("");
			OnUpdateState();
			return TRUE;
		}

		IRecordset->MoveFirst();
		while(!IRecordset->adoEOF)
		{   
			var = IRecordset->GetCollect("enginID");
			if(var.vt != VT_NULL)
				a = (LPCSTR)_bstr_t(var);
			else
				a=_T("");

			var = IRecordset->GetCollect("gcmc");
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

			IRecordset->MoveNext();
			s++;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		OnCancel();
		return TRUE;
	}

		//**************************************************
	m_str4=_T("");

	OnUpdateState();

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

//
//建立新工程
//
void newproject::OnAddnewproject() 
{
	BOOL bRet;
	CString OldPrjID;

	UpdateData(TRUE);

	m_str4=_T("");
	newpro=1;

	GetDlgItem(IDC_BUTTON_new)->EnableWindow(FALSE);
	GetDlgItem(IDC_DELETESELECTPROJECT)->EnableWindow(FALSE);

	OldPrjID=EDIBgbl::SelPrjID;
	EDIBgbl::SelPrjID=m_eng_code;   //把EDIBgbl::SelPrjID的值付成新建工程的工程编号

	UpdateData(TRUE);//zsy pd

	bRet=insertitem();    //调用函数，插入新工程的数据
	OnUpdateState();

	//如果新建工程成功关闭对话框
	if(bRet)
	{
		//在标题栏上显示当前的工程名
		((CMainFrame*)::AfxGetMainWnd())->ShowCurrentProjectName();
		OnOK();
	}
	else
	{
		//在标题栏上显示当前的工程名
		EDIBgbl::SelPrjID=OldPrjID;
		((CMainFrame*)::AfxGetMainWnd())->ShowCurrentProjectName();
	}
}

//
//套用工程
//
void newproject::OnBUTTONnew() 
{
	BOOL bRet;
	CString OldPrjID;

	m_str4.TrimLeft();
	if(m_str4==_T(""))   
	{
		return;   //说明没有选择可套用的工程返回
	}

	newpro=0;

	UpdateData(TRUE);

	OldPrjID=EDIBgbl::SelPrjID;
	EDIBgbl::SelPrjID=m_eng_code;   //把EDIBgbl::SelPrjID的值付成新建工程的工程编号

	UpdateData(TRUE);  //zsy pd

	bRet=insertitem();    //调用函数，插入新工程的数据
	OnUpdateState();

	//如果新建工程成功关闭对话框
	if(bRet)
	{
		//在标题栏上显示当前的工程名
		((CMainFrame*)::AfxGetMainWnd())->ShowCurrentProjectName();
		OnOK();
	}
	else
	{
		//在标题栏上显示当前的工程名
		((CMainFrame*)::AfxGetMainWnd())->ShowCurrentProjectName();
		EDIBgbl::SelPrjID=OldPrjID;
	}
}    

//
// 删除所选择的工程
//
void newproject::OnDeleteSelectproject() 
{
	POSITION pos;
	CString ProjectId;
	int SelectItem;
	BOOL bRet;

	CString *pTableNameToDel;
	int     nTblCount;
	if( !GetImportTblNames(pTableNameToDel, nTblCount, NULL, 0) )
	{
		return ;
	}
/*
	LPCTSTR pTableNameToDel[]=
	{
		_T("Volume"),
		_T("A_CONFIG"),
		_T("PAINT"),
		_T("PRE_CALC"),
		_T("Engin"),
		_T("A_ARROW"),
		_T("A_COLOR"),
		_T("A_COMPND"),
		_T("A_LOC"),
		_T("A_MAT"),
		_T("A_PAI"),
		_T("A_PRE"),
		_T("A_PRO"),
		_T("A_MATAST"),
		_T("A_VERT"),
		_T("A_WIRE"),
		_T("A_YL"),
		_T("PASS"),
		_T("E_AST"),
		_T("EPREAST"),
		_T("T120"),


		_T("E_PAIAST"),
		_T("E_PAICOL"),
		_T("E_PAIEXP"),
		_T("E_PREAST"),
		_T("E_PRECOL"),
		_T("E_PREEXP"),
		_T("E_PRESIZ"),
		_T("EDIT"),
		_T("PAINT_C"),
		_T("PRE_COLR"),
		_T("PRE_TAB"),
		_T("A_DIR")
	};	

*/
	pos=c_list.GetFirstSelectedItemPosition();

	c_progress.SetRange(1,nTblCount/*sizeof(pTableNameToDel)/sizeof(LPCTSTR)*/);

	while(pos)
	{
		SelectItem=c_list.GetNextSelectedItem(pos);
		ProjectId=c_list.GetItemText(SelectItem,0);

		c_progress.SetPos(1);

		for(int i=0;i<nTblCount/*sizeof(pTableNameToDel)/sizeof(LPCTSTR)*/;i++)
		{
			bRet=DeleteProjectInTable(pTableNameToDel[i],ProjectId);
			if(!bRet)
			{
				c_progress.SetPos(1);

				delete [] pTableNameToDel;
				return;
			}

			c_progress.SetPos(i+1);
		}

		//
		// 如果删除的是当前工程,清空当前的确工程ID
		//
		if(ProjectId==EDIBgbl::SelPrjID)
		{
			EDIBgbl::SelSpecID=0;
			EDIBgbl::SelPrjID=_T("");
			EDIBgbl::SelDsgnID=0;
			EDIBgbl::SelHyID=0;
			EDIBgbl::SelVlmID=0;

			//在标题栏上显示当前的工程名
			((CMainFrame*)::AfxGetMainWnd())->ShowCurrentProjectName();
		}

	}

	c_progress.SetPos(1);
	//
	// 删除ListControl中的多选项
	//
	while(TRUE)
	{
		pos=c_list.GetFirstSelectedItemPosition();
		if(pos==NULL)
			break;

		SelectItem=c_list.GetNextSelectedItem(pos);
		c_list.DeleteItem(SelectItem);
	}

	OnUpdateState();
	delete [] pTableNameToDel;

}

BOOL newproject::insertitem()
{
	
	CString tablename;
	_RecordsetPtr	IRecordset;
	HRESULT hr;
	BOOL bRet;

	hr=IRecordset.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		ReportExceptionError(_com_error(hr));
		return FALSE;
	}

 	if(m_eng_code!=_T("") && m_engin!=_T(""))
	{
		CString q;
		q.Format(_T("select * from engin where EnginID='%s'"),m_eng_code);

		try
		{
			// 获取库接库的IDispatch指针
			IRecordset->Open(_variant_t(q),
								theApp.m_pConAllPrj.GetInterfacePtr(),	 
								adOpenDynamic,
								adLockOptimistic,
								adCmdText);
		}
		catch(_com_error &e)
		{
			ReportExceptionError(e);
			return FALSE;
		}
		 //如果没有重名的工程编号
		if(IRecordset->BOF && IRecordset->adoEOF || m_bAuto/*zsypd*/) 
		{
			//_RecordsetPtr pRsTblINFO;
			CString sql, strTmpTbl;
			int pos, nTblCount;
			CString *pTableName;

			CString lpNotTbl[] =  //列出作另外处理的表名
			{
				_T("Engin"),
				_T("Volume")
			};
			//将要处理的表名从数据库中读取.
			if( !GetImportTblNames(pTableName,nTblCount,lpNotTbl,sizeof(lpNotTbl)/sizeof(lpNotTbl[0])) )
			{
				return FALSE;
			}

/*			pRsTblINFO.CreateInstance(__uuidof(Recordset));
			try
			{
				sql = "SELECT * FROM [ImportTblINFO] WHERE IPED=TRUE";
				pRsTblINFO->Open(_bstr_t(sql), theApp.m_pConIPEDsort.GetInterfacePtr(),
						adOpenStatic, adLockOptimistic, adCmdText);
				nTblCount = pRsTblINFO->GetRecordCount();
				if( nTblCount <= 0 )
				{
					return FALSE;		//没有要处理的表。
				}
				pTableName = new CString[nTblCount];

				//循环取出各表名。如果不是作另外处理的则加入到pTableName;
				for(nTblCount=0; !pRsTblINFO->adoEOF; pRsTblINFO->MoveNext() )
				{
					strTmpTbl = vtos(pRsTblINFO->GetCollect(_variant_t("TableName")));
					bTblExis = false;		
					for(int i=0; i<sizeof(lpNotTbl)/sizeof(lpNotTbl[0]); i++)
					{
						if( !strTmpTbl.CompareNoCase(lpNotTbl[i]) )
						{
							bTblExis = true;
							break;
						}
					}
					if( !bTblExis )		
					{
						//表不须要做另外处理。
						pTableName[nTblCount] = strTmpTbl;
						nTblCount++;			//计数器加1.
					}
				}
			}
			catch(_com_error& e)
			{
				ReportExceptionError(e);
				return FALSE;
			}
*/
/*//zsy 
			LPCTSTR pTableName[]=
			{
//				_T("Volume"),
				_T("A_CONFIG"),
				_T("PAINT"),
				_T("PRE_CALC"),
//				_T("Engin"),
				_T("A_ARROW"),
				_T("A_COLOR"),
				_T("A_COMPND"),
				_T("A_LOC"),
				_T("A_MAT"),
				_T("A_PAI"),
				_T("A_PRE"),
				_T("A_PRO"),
				_T("A_MATAST"),
				_T("A_VERT"),
				_T("A_WIRE"),
				_T("A_YL"),
				_T("PASS"),
				_T("E_AST"),
				_T("EPREAST"),
				_T("T120"),
				_T("A_DIR")
			};

*/			try
			{
				//启动事务处理
				theApp.m_pConAllPrj->BeginTrans();	

				//*******************************************************
				//                  在表engin中注册新的工程
				//*******************************************************
				
				sql.Format(_T("insert into engin (EnginID,gcmc) values('%s','%s')"),m_eng_code,m_engin);
				theApp.m_pConAllPrj->Execute(_bstr_t(sql),NULL,adCmdText);
			}
			catch(_com_error)
			{
			}
			BeginWaitCursor();    //因为对数据库操作的时间很长，所以让鼠标显示漏斗状 

			if( m_bAuto )    
			{
				newpro = 1;  //自动新建时建立
			}
			else
			{	//当为自动新建时,不显示对话框.所以不能设置进度条
				c_progress.SetRange(1,/*sizeof(pTableName)/sizeof(pTableName[0])*/nTblCount+1);
				pos=1;
				c_progress.SetPos(pos);
			}
			//*******************************************************
			//                  在表Volume中添加新工程的记录集
			//*******************************************************
/*
			bRet=addto_volume();
			if(!bRet)
			{
				theApp.m_pConnection->RollbackTrans();
				c_progress.SetPos(1);
				return FALSE;
			}
			*/


			//
			// 如果是新建工程，Volume表应该用AddToValume添加记录
			//

			for(int i=0;i<nTblCount/*sizeof(pTableName)/sizeof(pTableName[0])*/;i++)
			{
				tablename=pTableName[i];
				if( 1 == m_bAuto || 1 == newpro)
				{	//当为自动导入时,或新建工程时.不对pre_calc表进行处理。它将从Excel文件中读取数据。
					if( !tablename.CompareNoCase("pre_calc") || !tablename.CompareNoCase("Paint") )
					{
						continue;
					}
				}
				bRet=addto_table(tablename);
				if(!bRet)
				{
					theApp.m_pConAllPrj->RollbackTrans();
					if( !m_bAuto )
					{
						c_progress.SetPos(1);
					}
					delete [] pTableName;
					return FALSE;
				}
				if( !m_bAuto )
				{
					pos++;
					c_progress.SetPos(pos);
				}
			}

			//为自动导入时,在卷册表中只增加当前卷册代号的记录.
			if( !m_bAuto )
			{
				// 更新了A_DIR表以后向Valume表写数据
				//
				try
				{
					if(newpro==1)
						AddToValumeOnNew();
					else
						AddToValumeOnApply();
				}
				catch(_com_error &e)
				{
					theApp.m_pConAllPrj->RollbackTrans();
					if( !m_bAuto )
					c_progress.SetPos(1);

					ReportExceptionErrorLV1(e);
					return FALSE;
				}
				c_progress.SetPos(1);
			}
			else
			{
				//在卷册表中查找当前工程的当前卷册是否存在,不存在则增加.
				this->AddCurrentVolume();
			}

			delete [] pTableName;
			EndWaitCursor(); 
			theApp.m_pConAllPrj->CommitTrans();
		}
		else 
		{
			MessageBox("工程编号有重复，请重新填写！");
			return FALSE;
		}
	}
	else if(m_eng_code==_T(""))
	{
		MessageBox(_T("工程编号不能为空！"));
		return FALSE;
	}
	else if(m_engin==_T("")) 
	{
		MessageBox(_T("工程名称不能为空！"));
		return FALSE;
	}

	return TRUE;
}

void newproject::OnClickList1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString str;
	int x;

	str=_T("");
	x = c_list.GetNextItem(-1, LVNI_SELECTED);

	if(x >= 0)
	{   	
		str = c_list.GetItemText(x, 0);
	}

    m_str4=str;
	*pResult = 0;

	OnUpdateState();
}


BOOL newproject::addto_table(CString tablename)
{
	CString sql;
	_RecordsetPtr pRs;
	pRs.CreateInstance(__uuidof(Recordset));

	try
	{
		theApp.m_pConnection->Execute("drop table temp",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionError(e);
			return FALSE;
		}
	}
	try
	{
		//当为自动增加时（从导入数据），如果表中当前工程已经有记录。不作处理。
		if( m_bAuto )  
		{
			sql.Format("SELECT * FROM ["+tablename+"] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ");
			pRs->Open(_bstr_t(sql), (IDispatch*)theApp.m_pConnection, 
						adOpenStatic, adLockOptimistic, adCmdText);
			if( pRs->GetRecordCount() > 0 )
			{
				return true;
			}
		}
		else
		{
			try
			{
				//新建工程之前，清空已经存在的记录。
				sql.Format("DELETE FROM ["+tablename+"] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ");
				theApp.m_pConnection->Execute(_bstr_t(sql), NULL, adCmdText);
			}
			catch(_com_error)
			{
				//当没有数据时,有可能出错.但不须要做处理.
			}
		}

 		if(newpro==1)
		{
			//新建工程
			sql.Format(_T("select * into temp from %s where EnginID is null"),tablename);
		}
 		else
 		{
			//套用工程
			sql.Format(_T("select * into temp from %s where EnginID='%s'"),tablename,m_str4);
		}

		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	
		//更新eng_code字段为选择的工程
		sql.Format(_T("UPDATE temp SET EnginID='%s'"),EDIBgbl::SelPrjID);
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);

		sql.Format(_T("insert into %s select * from temp"),tablename);
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);

		sql=_T("drop table temp");
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}

BOOL newproject::addto_volume()
{
	CString sql;
	_RecordsetPtr	IRecordset;
	_RecordsetPtr	IRecordset1;

	try
	{
		theApp.m_pConAllPrj->Execute("drop table volume1",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionError(e);
			return FALSE;
		}
	}

 	if(newpro==1)
	{
		//新建工程
		sql=_T("select * into volume1 from Volume where EnginID is null");
	}					
 	else
 	{
		//套用工程
		sql.Format(_T("select * into volume1 from Volume where EnginID='%s'"),m_str4);
	}

	try
	{
		theApp.m_pConAllPrj->Execute(_bstr_t(sql),NULL,adCmdText);
		
		//更新eng_code字段为选择的工程
		sql.Format(_T("UPDATE Volume1 SET EnginID='%s'"),EDIBgbl::SelPrjID);
		theApp.m_pConAllPrj->Execute(_bstr_t(sql),NULL,adCmdText);

		IRecordset.CreateInstance(__uuidof(Recordset));
		IRecordset->Open("select * from volume1",
						   theApp.m_pConAllPrj.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenDynamic,
								adLockOptimistic,
								adCmdText);

		IRecordset1.CreateInstance(__uuidof(Recordset));
		IRecordset1->Open("select * from volume ORDER BY VolumeID ASC",
						   theApp.m_pConAllPrj.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenDynamic,
								adLockOptimistic,
								adCmdText);

		long id;

		IRecordset1->MoveLast();

		id=IRecordset1->GetCollect("VolumeID");
		id++;

		IRecordset->MoveFirst();
		while(!IRecordset->adoEOF)
		{
			IRecordset->PutCollect("VolumeID",_variant_t(id));
			IRecordset->Update();
			IRecordset->MoveNext();
			id++;
		}

		sql="insert into Volume select * from Volume1";
		theApp.m_pConAllPrj->Execute(_bstr_t(sql),NULL,adCmdText);

		IRecordset->Close();
		IRecordset1->Close();

		sql="drop table Volume1";
		theApp.m_pConAllPrj->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}
	
	return TRUE;
}


//////////////////////////////////
//
// 更新各个控件的状态
//
void newproject::OnUpdateState()
{
	POSITION pos;
	pos=c_list.GetFirstSelectedItemPosition();

	if(pos==NULL)
	{
		GetDlgItem(IDC_BUTTON_new)->EnableWindow(FALSE);
		GetDlgItem(IDC_DELETESELECTPROJECT)->EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(IDC_BUTTON_new)->EnableWindow(TRUE);
		GetDlgItem(IDC_DELETESELECTPROJECT)->EnableWindow(TRUE);
	}

	c_list.SetFocus();
}

////////////////////////////////////////////////////////////////////////
//
// 从删除指定的工程
//
// TableName[in]	表名
// ProjectID[in]	工程名
//
// 函数成功返回TRUE，否则返回FALSE
//
BOOL newproject::DeleteProjectInTable(LPCTSTR TableName,LPCTSTR ProjectID)
{
	CString SQL;
	_RecordsetPtr IRecordset;
	HRESULT hr;
	hr=IRecordset.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		ReportExceptionError(_com_error(hr));
		return FALSE;
	}

	//
	_ConnectionPtr pTemp;
	if ( tableExists( theApp.m_pConnection, TableName ) )
	{
		pTemp = theApp.m_pConnection;	
	}
	else
	{
		pTemp = theApp.m_pConAllPrj;
	}

	if(TableName==NULL || ProjectID==NULL)
	{
		return FALSE;
	}


	try
	{
		SQL.Format(_T("SELECT * FROM %s WHERE EnginID='%s'"),TableName,ProjectID);

		IRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)pTemp),
						 adOpenDynamic,adLockOptimistic,adCmdText);

		if(IRecordset->adoEOF && IRecordset->BOF)
		{
			return TRUE;
		}

		IRecordset->Close();

		SQL.Format(_T("DELETE * FROM %s WHERE EnginID='%s'"),TableName,ProjectID);
		pTemp->Execute(_bstr_t(SQL),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return FALSE;
	}

	return TRUE;
}

///////////////////////////////////////////////
//
// 当新建工程时向Valume加数据
//
// 次函数应该在A_DIR表数据新建后调用
// throw(_com_error)
//
void newproject::AddToValumeOnNew()
{
	_RecordsetPtr IRecordset_Valume;
	_RecordsetPtr IRecordset_Dir;
	CString SQL,strTemp;
	HRESULT hr;
	_variant_t varTemp;
	long iVolumeID;

	if(m_eng_code.IsEmpty())
	{
		_com_error e(E_FAIL);
		ReportExceptionErrorLV2(e);
		throw e;
	}
	
	hr=IRecordset_Valume.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	hr=IRecordset_Dir.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	//
	// 取得新增加记录用的VolumeID
	//
	try
	{
		iVolumeID=GetNextVolumeID();
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	//
	// 打开a_dir表
	//
	SQL.Format(_T("SELECT * FROM A_DIR WHERE EnginID='%s'"),m_eng_code);

	try
	{
		IRecordset_Dir->Open(_variant_t(SQL),_variant_t((IDispatch*)theApp.m_pConnection),
							 adOpenDynamic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	try
	{
		SQL.Format(_T("SELECT * FROM Volume ORDER BY VolumeID ASC"));

		IRecordset_Valume->Open(_variant_t(SQL),_variant_t((IDispatch*)theApp.m_pConAllPrj),
								adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Dir->adoEOF)
		{
			varTemp=IRecordset_Dir->GetCollect(_variant_t("DIR_VOL"));
			if( varTemp.vt==VT_EMPTY || varTemp.vt==VT_NULL )
			{
				IRecordset_Dir->MoveNext();
				continue;
			}			
			strTemp=(LPCTSTR)(_bstr_t)varTemp;

			strTemp.TrimLeft();

			strTemp=strTemp.Mid(1);

			strTemp.TrimRight();

			IRecordset_Valume->AddNew();
			IRecordset_Valume->PutCollect(_variant_t("VolumeID"),_variant_t(iVolumeID));
			IRecordset_Valume->PutCollect(_variant_t("EnginID"),_variant_t(m_eng_code));
			IRecordset_Valume->PutCollect(_variant_t("jcdm"),_variant_t(strTemp));

			varTemp=IRecordset_Dir->GetCollect(_variant_t("DIR_NAME"));
			IRecordset_Valume->PutCollect(_variant_t("jcmc"),varTemp);

			IRecordset_Valume->PutCollect(_variant_t("SJHYID"),_variant_t((long)0));
			IRecordset_Valume->PutCollect(_variant_t("SJJDID"),_variant_t((long)4));
			IRecordset_Valume->PutCollect(_variant_t("ZYID"),_variant_t((long)3));

			iVolumeID++;
			IRecordset_Dir->MoveNext();

		}

		IRecordset_Valume->Update();
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

}

//////////////////////////////////////////////
//
// 当套用工程时向Valume加数据
//
// 次函数应该在A_DIR表数据套用后调用
// throw(_com_error)
//
void newproject::AddToValumeOnApply()
{
	CString SQL;
	_RecordsetPtr IRecordset_Temp;
	HRESULT hr;
	long iVolumeID;

	if(m_eng_code.IsEmpty() || m_str4.IsEmpty())
	{
		_com_error e(E_FAIL);
		ReportExceptionErrorLV2(e);
		throw e;
	}


	hr=IRecordset_Temp.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}


	try
	{
		theApp.m_pConAllPrj->Execute(_bstr_t("DROP TABLE VOLUME1"),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}
	}

	SQL.Format(_T("SELECT * INTO VOLUME1 from Volume where EnginID='%s'"),m_str4);

	theApp.m_pConAllPrj->Execute(_bstr_t(SQL),NULL,adCmdText);

	SQL.Format(_T("UPDATE VOLUME1 SET EnginID='%s'"),m_eng_code);

	theApp.m_pConAllPrj->Execute(_bstr_t(SQL),NULL,adCmdText);

	//
	// 取得新增加记录用的VolumeID
	//
	try
	{
		iVolumeID=GetNextVolumeID();
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	SQL.Format(_T("SELECT * FROM Volume1 ORDER BY VolumeID ASC"));

	try
	{
		IRecordset_Temp->Open(_variant_t(SQL),_variant_t((IDispatch*)theApp.m_pConAllPrj),
							  adOpenDynamic,adLockOptimistic,adCmdText);

		while(!IRecordset_Temp->adoEOF)
		{
			IRecordset_Temp->PutCollect(_variant_t("VolumeID"),_variant_t(iVolumeID));

			iVolumeID++;

			IRecordset_Temp->MoveNext();
		}

		IRecordset_Temp->Close();
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	try
	{
		theApp.m_pConAllPrj->Execute(_bstr_t("INSERT INTO VOLUME SELECT * FROM VOLUME1"),NULL,adCmdText);

		theApp.m_pConAllPrj->Execute(_bstr_t("DROP TABLE VOLUME1"),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}


}

///////////////////////////////////////////
//
// 返回Volume表中的下一个可用的VolumeID号
//
// throw(_com_error)
//
long newproject::GetNextVolumeID()
{
	_RecordsetPtr IRecordset_Valume;
	HRESULT hr;
	CString SQL;
	long iVolumeID;
	_variant_t varTemp;

	hr=IRecordset_Valume.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	SQL.Format(_T("SELECT * FROM Volume ORDER BY VolumeID ASC"));

	try
	{
		IRecordset_Valume->Open(_variant_t(SQL),_variant_t((IDispatch*)theApp.m_pConAllPrj),
								adOpenDynamic,adLockOptimistic,adCmdText);


		if(IRecordset_Valume->adoEOF && IRecordset_Valume->BOF)
		{
			iVolumeID=1;
		}
		else
		{
			IRecordset_Valume->MoveLast();
			varTemp=IRecordset_Valume->GetCollect(_variant_t("VolumeID"));
			iVolumeID=varTemp;

			iVolumeID++;
		}

		IRecordset_Valume->Close();

	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	return iVolumeID;
}

////////////////////////////////////////
//
// 设置回车时默认的ID
//
void newproject::OnSetfocusEDITengcode() 
{
	SetDefID(IDC_EDIT_eng_code);
}

void newproject::OnSetfocusEDITengin() 
{
	SetDefID(IDC_EDIT_engin);	
}

//自动新建工程.设置状态.
bool newproject::SetAuto( bool bAuto )
{
	m_bAuto = bAuto;
	return true;
}
//在当前工程中查找给定的卷册,如果不存在则增加.
void newproject::AddCurrentVolume()
{
	_RecordsetPtr  pRs;
	CString strSQL;
	try
	{
		if ( EDIBgbl::SelPrjID.IsEmpty() || EDIBgbl::SelVlmCODE.IsEmpty() )
		{//如果工程代号或者卷册代号为空时,将不考虑.
			return;
		}
		pRs.CreateInstance(__uuidof(Recordset));
		pRs->CursorLocation = adUseClient;

		strSQL = "SELECT * FROM [Volume] WHERE EnginID = '"+EDIBgbl::SelPrjID+"' AND JCDM = '"+EDIBgbl::SelVlmCODE+"' ";
		pRs->Open(_bstr_t(strSQL), theApp.m_pConAllPrj.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);

		if( pRs->adoEOF && pRs->BOF )
		{
			//没有记录,增加卷册.
			pRs->Close();
			strSQL = "Select * FROM [Volume] ORDER BY VolumeID ";
			pRs->Open(_bstr_t(strSQL),theApp.m_pConAllPrj.GetInterfacePtr(),
					adOpenDynamic, adLockOptimistic, adCmdText);
			pRs->MoveLast();
			_variant_t var = pRs->GetCollect("VolumeID");
			long VlmID = vtoi(var);
			VlmID++;
			pRs->AddNew();
			pRs->PutCollect(_variant_t("VolumeID"),_variant_t(VlmID));
			pRs->PutCollect(_variant_t("EnginID"),_variant_t(EDIBgbl::SelPrjID) );
			pRs->PutCollect(_variant_t("JCDM"),_variant_t(EDIBgbl::SelVlmCODE));
			pRs->PutCollect(_variant_t("JCMC"),_variant_t("") ); //无
			pRs->PutCollect(_variant_t("SJHYID"),_variant_t(EDIBgbl::SelHyID));
			pRs->PutCollect(_variant_t("SJJDID"),_variant_t(EDIBgbl::SelDsgnID) );
			pRs->PutCollect(_variant_t("ZYID"),_variant_t(EDIBgbl::SelSpecID) );
			pRs->Update();

		}
	}
	catch(_com_error& e)
	{
		ReportExceptionErrorLV2(e);
	}

}
//////////////////////////////////
//从数据库中读取将要处理的表名.
//
bool newproject::GetImportTblNames(CString *&pTableName, int &nTblCount, CString pNotTbl[], int nNotCount)
{
	_RecordsetPtr pRsTblINFO;
	CString sql, strTmpTbl;
	bool	bTblExis=false;
	pRsTblINFO.CreateInstance(__uuidof(Recordset));
	try
	{
		//将要处理的表名从数据库中读取.
		sql = "SELECT * FROM [ImportTblINFO] WHERE IPED=TRUE ORDER BY TableName DESC";
		pRsTblINFO->Open(_bstr_t(sql), theApp.m_pConIPEDsort.GetInterfacePtr(),
				adOpenStatic, adLockOptimistic, adCmdText);
		nTblCount = pRsTblINFO->GetRecordCount();
		if( nTblCount <= 0 )
		{
			return FALSE;		//没有要处理的表。
		}
		pTableName = new CString[nTblCount];

		//循环取出各表名。如果不是作另外处理的则加入到pTableName;
		for(nTblCount=0; !pRsTblINFO->adoEOF; pRsTblINFO->MoveNext() )
		{
			strTmpTbl = vtos(pRsTblINFO->GetCollect(_variant_t("TableName")));
			bTblExis = false;		
			for(int i=0; i<nNotCount; i++)
			{
				if( !strTmpTbl.CompareNoCase(pNotTbl[i]) )
				{
					bTblExis = true;
					break;
				}
			}
			if( !bTblExis )		
			{
				//表不须要做另外处理。
				pTableName[nTblCount] = strTmpTbl;
				nTblCount++;			//计数器加1.
			}
		}
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return FALSE;
	}
	return TRUE;
}

BOOL newproject::tableExists(_ConnectionPtr pCon, CString tbn)
{
	if(pCon==NULL || tbn=="")
		return false;
	_RecordsetPtr rs;
	if(tbn.Left(1)!="[")
		tbn="["+tbn;
	if(tbn.Right(1)!="]")
		tbn+="]";
	try{
		rs=pCon->Execute(_bstr_t(tbn),NULL,adCmdTable);
		rs->Close();
		return true;
	}
	catch(_com_error e)
	{
		return false;
	}

}