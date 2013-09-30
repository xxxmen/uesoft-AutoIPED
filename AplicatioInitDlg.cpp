// AplicatioInitDlg.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "AplicatioInitDlg.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#include "autoiped.h"
extern CAutoIPEDApp theApp;
/////////////////////////////////////////////////////////////////////////////
// CAplicatioInitDlg dialog


CAplicatioInitDlg::CAplicatioInitDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAplicatioInitDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAplicatioInitDlg)
	m_InitializeInfo = _T("");
	//}}AFX_DATA_INIT
}


void CAplicatioInitDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAplicatioInitDlg)
	DDX_Control(pDX, IDC_INIT_PROGRESS, m_InitProgress);
	DDX_Text(pDX, IDC_STATIC_INFO, m_InitializeInfo);
	//}}AFX_DATA_MAP
}



BEGIN_MESSAGE_MAP(CAplicatioInitDlg, CDialog)
	//{{AFX_MSG_MAP(CAplicatioInitDlg)
	//}}AFX_MSG_MAP
	ON_MESSAGE(WM_INITIALIZEWORK,OnInitializeWorkFunc)
END_MESSAGE_MAP()


//////////////////////////////////////////////////////////////////////////////
//
// 响应自定义的消息的响应函数
//
// 完成对本工程的初始化
//
LRESULT CAplicatioInitDlg::OnInitializeWorkFunc(WPARAM wParam, LPARAM lParam)
{
	//
	if(!CopyFileToEveryDir())
	{
		OnCancel();
		return 0;
	}
	//升级数据库
	if ( !UpgradeDatabase() )
	{
		OnCancel();
		return 0;
	}

	OnOK();
	return 0;
}
//////////////////////////////////////
//读取公共数据库存放目录
CString CAplicatioInitDlg::ReadCommonMdbPath()
{
	HKEY hKey;
	DWORD szSize = 255;
	CString strCommon = "";

	if( RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("SOFTWARE\\UESoft Shared\\AutoPDS3.0"),0, KEY_ALL_ACCESS, &hKey) != ERROR_SUCCESS )
	{
		AfxMessageBox( "打开注册表失败!" );
		return "";
	}

	if( RegQueryValueEx(hKey, "ArxIniPath", NULL, NULL, (LPBYTE)strCommon.GetBufferSetLength(255), &szSize) != ERROR_SUCCESS )
	{
		AfxMessageBox( "打开注册表失败!" );
		return "";
	}
	
	strCommon.ReleaseBuffer();
	strCommon.TrimRight("\\");
	strCommon = strCommon + _T("\\Common Mdb");
	return strCommon;

}
/////////////////////////////////////////////////////////////
//
// 拷贝文件到各个指定的目录
//
// 如果函数成功返回TRUE，否则返回FALSE
//
BOOL CAplicatioInitDlg::CopyFileToEveryDir()
{
	CString strFrom,strTo;
	BOOL bRet;
	DWORD dwRet;
	UINT uRet;

	if( gsProjectDBDir.IsEmpty()	||
		gsProjectDir.IsEmpty()		||
		gsTemplateDir.IsEmpty()		||
		gsShareMdbDir.IsEmpty() )
		
	{
		ReportExceptionErrorLV1(_T("注册表信息不完整"));

		return FALSE;
	}


	bRet=::CreateDirectory(gsProjectDBDir,NULL);

	if(!bRet)
	{
		dwRet=GetLastError();

		if(dwRet!=ERROR_ALREADY_EXISTS)
		{
			uRet=GetDriveType(gsProjectDBDir);

			if(uRet==DRIVE_UNKNOWN || uRet==DRIVE_NO_ROOT_DIR)
			{
				ReportExceptionErrorLV1(_T("创建文件夹失败"));
				return FALSE;
			}
		}
	}

	bRet=::CreateDirectory(gsProjectDir,NULL);

	if(!bRet)
	{
		dwRet=GetLastError();

		if(dwRet!=ERROR_ALREADY_EXISTS)
		{
			uRet=GetDriveType(gsProjectDBDir);

			if(uRet==DRIVE_UNKNOWN || uRet==DRIVE_NO_ROOT_DIR)
			{
				ReportExceptionErrorLV1(_T("创建文件夹失败"));
				return FALSE;
			}
		}
	}

	//
	// 从gsTemplateDir到gsProjectDBDir
	//

	LPCTSTR szCopyFileInfo[]=
	{
		_T("AUTOIPED.MDB"),
	};

	for(int i=0;i<sizeof(szCopyFileInfo)/sizeof(szCopyFileInfo[0]);i++)
	{
		strFrom=gsTemplateDir;
		strTo=gsProjectDBDir;

		FormatDirPath(strFrom);
		FormatDirPath(strTo);

		strFrom+=szCopyFileInfo[i];
		strTo+=szCopyFileInfo[i];

		if(!FileExists(strTo))
		{
			ShowWindow(SW_NORMAL);
			UpdateWindow();
		}
		else
		{
			continue;
		}

		m_InitializeInfo.Format(_T("拷贝文件从%s到%s"),strFrom,strTo);

		UpdateData(FALSE);
		UpdateWindow();
		
		bRet=::CopyFileEx(strFrom,strTo,CopyProgressRoutine,this,NULL,COPY_FILE_FAIL_IF_EXISTS);

		if(!bRet)
		{
			dwRet=GetLastError();

			if(dwRet!=ERROR_FILE_EXISTS)
			{
				//
				// 如果调用CopyFileEx失败有可能是因为操作系统在98或以下
				// 将用CopyFile重试
				//
				bRet=CopyFile(strFrom,strTo,TRUE);

				if(!bRet)
				{
					dwRet=GetLastError();

					if(dwRet!=ERROR_FILE_EXISTS)
					{
						ReportExceptionErrorLV1(_T("拷贝文件失败:\n请确认文件‘"+strFrom+"’是否存在。"));
						return FALSE;
					}
				}

			}
		}

	}

	//
	// 从gsTemplateDir到gsProjectDir
	//

	LPCTSTR szCopyFileInfo2[]=
	{
		_T("WORKPRJ.MDB")
	};


	for(i=0;i<sizeof(szCopyFileInfo2)/sizeof(szCopyFileInfo2[0]);i++)
	{
		strFrom=gsTemplateDir;
		strTo=gsProjectDir;

		FormatDirPath(strFrom);
		FormatDirPath(strTo);

		strFrom+=szCopyFileInfo2[i];
		strTo+=szCopyFileInfo2[i];

		if(!FileExists(strTo))
		{
			ShowWindow(SW_NORMAL);
			UpdateWindow();
		}
		else
		{
			continue;
		}

		m_InitializeInfo.Format(_T("拷贝文件从%s到%s"),strFrom,strTo);

		UpdateData(FALSE);
		UpdateWindow();

		bRet=::CopyFileEx(strFrom,strTo,CopyProgressRoutine,this,NULL,COPY_FILE_FAIL_IF_EXISTS);

		if(!bRet)
		{
			dwRet=GetLastError();

			if(dwRet!=ERROR_FILE_EXISTS)
			{
				//
				// 如果调用CopyFileEx失败有可能是因为操作系统在95或以下
				// 将用CopyFile重试
				//
				bRet=CopyFile(strFrom,strTo,TRUE);

				if(!bRet)
				{
					dwRet=GetLastError();

					if(dwRet!=ERROR_FILE_EXISTS)
					{
						ReportExceptionErrorLV1(_T("拷贝文件失败"));
						return FALSE;
					}
				}
			}
		}

	}
//拷贝数据库到共享文件夹下面

	LPCTSTR szCopyFileInfo3[]=
	{
		_T("TableFormat.MDB"),
		_T("DrawingSize.MDB"),
		_T("Material.MDB"),
		_T("Medium.mdb"),
		_T("MaterialName.mdb"),
		_T("ALLPRJDB.MDB")
	};


	for(i=0;i<sizeof(szCopyFileInfo3)/sizeof(szCopyFileInfo3[0]);i++)
	{
		strFrom=ReadCommonMdbPath();
		if ( "" == strFrom )
		{
			return FALSE;
		}

		strTo=gsShareMdbDir;

		FormatDirPath(strFrom);
		FormatDirPath(strTo);

		strFrom += szCopyFileInfo3[i];
		strTo += szCopyFileInfo3[i];

		if(!FileExists(strTo))
		{
			ShowWindow(SW_NORMAL);
			UpdateWindow();
		}
		else
		{
			continue;
		}

		m_InitializeInfo.Format(_T("拷贝文件从%s到%s"),strFrom,strTo);

		UpdateData(FALSE);
		UpdateWindow();

		bRet=::CopyFileEx(strFrom,strTo,CopyProgressRoutine,this,NULL,COPY_FILE_FAIL_IF_EXISTS);

		if(!bRet)
		{
			dwRet=GetLastError();

			if(dwRet!=ERROR_FILE_EXISTS)
			{
				//
				// 如果调用CopyFileEx失败有可能是因为操作系统在95或以下
				// 将用CopyFile重试
				//
				bRet=CopyFile(strFrom,strTo,TRUE);

				if(!bRet)
				{
					dwRet=GetLastError();

					if(dwRet!=ERROR_FILE_EXISTS)
					{
						ReportExceptionErrorLV1(_T("拷贝文件失败"));
						return FALSE;
					}
				}
			}
		}

	}

	return TRUE;
}

////////////////////////////////////////////////////////////////////////////////////
//
// CopyFileEx的回调函数
//
DWORD CALLBACK CAplicatioInitDlg::CopyProgressRoutine(
  													   LARGE_INTEGER TotalFileSize,
													   LARGE_INTEGER TotalBytesThransferred,
													   LARGE_INTEGER StreamSize,
													   LARGE_INTEGER StreamBytesTransferred,
													   DWORD dwStreamNumber,
													   DWORD dwCallbackReason,
													   HANDLE hSourceFile,
													   HANDLE hDestinatonFile,
													   LPVOID lpData)
{
	CAplicatioInitDlg *pDlg;
	double fPercent;

	pDlg=(CAplicatioInitDlg*)lpData;

	fPercent=(double)TotalBytesThransferred.LowPart/(double)TotalFileSize.LowPart*100.0;
 
  
	pDlg->m_InitProgress.SetPos((int)fPercent);

	pDlg->UpdateWindow();
	
	return PROGRESS_CONTINUE;
}

//////////////////////////////////////////////////////////////////
//
// 格式化使路径合乎程序要求
//
// DirPath[in/out]	需要格式化的字符串，返回已经格式化的字符串
//
// 次函数将路径左右多余的空格去除，如果最末尾没有“\”并加上
//
void CAplicatioInitDlg::FormatDirPath(CString &DirPath)
{
	DirPath.TrimLeft();
	DirPath.TrimRight();

	if(DirPath.IsEmpty())
	{
		return;
	}

	if(DirPath.Right(1)!=_T("\\"))
	{
		DirPath+=_T("\\");
	}
}


/////////////////////////////////////////////////////////////////////////////
// CAplicatioInitDlg message handlers

BOOL CAplicatioInitDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();

	m_InitProgress.SetRange(1,100);
	m_InitProgress.SetPos(0);
	
	this->PostMessage(WM_INITIALIZEWORK,0,0);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}



//------------------------------------------------------------------                
// DATE         :[2005/06/03]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :对指定的数据库进行升级.
//------------------------------------------------------------------
short CAplicatioInitDlg::UpgradeDB(CString strDestDB, CString strSourceDB)
{

	 _ConnectionPtr pConAlter;
	 pConAlter.CreateInstance(__uuidof(Connection));

	CDaoDatabase	sDB,dDB;
	CDaoTableDef	sTDef(&sDB);		//数据表
	CDaoTableDef	dTDef(&dDB);

	CDaoFieldInfo	fieldInfo;			//字段信息
	CDaoFieldInfo	fieldInfoS;
	CDaoFieldInfo	fieldInfoD;
	CString			strSQL;				//SQL语句
	try
	{
		//Microsoft.Jet.OLEDB.4.0;	的驱动. 用于修改字段的大小.
		pConAlter->Open(_bstr_t(CONNECTSTRING+strDestDB),"","",-1);

		sDB.Open(strSourceDB);			//打开数据库
		dDB.Open(strDestDB);
		
		CStringList *sListTableName;	//所有表名的列表.
		CStringList *dListTableName;
		CStringList *sListFieldName;	//指定表中所有字段名的列表。
		CStringList *dListFieldName;
		
		CString		strTableName;		//当前处理的表名
		CString		strFieldName;		//当前处理的字段名

		sListTableName = GetTableNameList(&sDB);		//获得来源数据库中的各表的名称
		dListTableName = GetTableNameList(&dDB);		//目的数据库表名.
		
		for (POSITION pos=sListTableName->GetHeadPosition(); pos != NULL ; )
		{
			strTableName = sListTableName->GetNext(pos);

			if ( StringInList(strTableName, dListTableName) )
			{
				//两个数据库中都存在相同的表, 
				sTDef.Open(strTableName);		//打开
				dTDef.Open(strTableName);
				
				sListFieldName = GetFieldNameList(&sTDef);		//获得来源数据库中的指定表的所有字段名.
				dListFieldName = GetFieldNameList(&dTDef);		//目的库中表的字段.

				for ( POSITION pos=sListFieldName->GetHeadPosition(); pos != NULL; )
				{
					strFieldName = sListFieldName->GetNext(pos);
					if ( StringInList(strFieldName, dListFieldName) )
					{
						//两个表中有相同的字段
						sTDef.GetFieldInfo(strFieldName, fieldInfoS);
						dTDef.GetFieldInfo(strFieldName, fieldInfoD);

						if (fieldInfoD.m_lSize != fieldInfoS.m_lSize)
						{
							//字段大小不一样.改变目的表中的字段的大小,为原数据库表中的字段
							strSQL.Format("ALTER TABLE ["+strTableName+"] ALTER COLUMN "+strFieldName+" VARCHAR (%d) ", fieldInfoS.m_lSize);
							pConAlter->Execute(_bstr_t(strSQL),NULL,adCmdText);
						}

					}
					else
					{
						//两个表中没有相同的字段
						sTDef.GetFieldInfo(strFieldName, fieldInfo);
						fieldInfo.m_bAllowZeroLength = FALSE;
						fieldInfo.m_bRequired		 = FALSE;

						dTDef.CreateField(fieldInfo);			//在目的数据表中增加一个字段.

						//增加新的字段时,连同数据一起添加.
				//		strSQL.Format("UPDATE %s SET %s.%s=T2.%s FROM (SELECT * FROM [%s] IN \'%s\') AS T2 ",
				//				strTableName, strTableName,strFieldName,strFieldName,strTableName,sDB.GetName());
				//		dDB.Execute(strSQL);
				//		pConAlter->Execute(_bstr_t(strSQL),NULL,adCmdText);
					}
				}
				sListFieldName->RemoveAll();	//清空所有字段名的列表. 并释放内存
				dListFieldName->RemoveAll();
				delete sListFieldName;
				delete dListFieldName;
				sTDef.Close();
				dTDef.Close();
			}
			else
			{	//在目的数据库中没有源数据库指定的表。	则插入表.
				strSQL.Format("SELECT * INTO %s FROM [%s] IN \"%s\" ",
							strTableName, strTableName, sDB.GetName());
				dDB.Execute(strSQL);
			}
		}

		sListTableName->RemoveAll();		//清空所有表名的列表. 并释放内存
		dListTableName->RemoveAll();
		delete sListTableName;
		delete dListTableName;
		sDB.Close();
		dDB.Close();

		pConAlter->Close();		
	}catch (CException* e)
	{
		e->ReportError();
		return 0 ;
	}	
	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/03]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :返回指定数据库中的所有表的名称. 返回的列表指针都要进行释放内存空间
//------------------------------------------------------------------
CStringList* CAplicatioInitDlg::GetTableNameList(CDaoDatabase *pDB)
{
	ASSERT_VALID(pDB);
	CStringList *tableNameList;
	CDaoTableDefInfo tDefInfo;
	tableNameList = new CStringList();
	int iCount = pDB->GetTableDefCount();
	for(int i=0; i<iCount; i++)
	{
		pDB->GetTableDefInfo(i, tDefInfo);
		CString strTemp = tDefInfo.m_strName;
		if((tDefInfo.m_lAttributes & dbSystemObject) == 0)
		{
			tableNameList->AddTail(tDefInfo.m_strName);
		}
	}
	return tableNameList;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/03]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :获得指定表的所有字段名称. 返回的列表指针都要进行释放内存空间
//------------------------------------------------------------------
CStringList* CAplicatioInitDlg::GetFieldNameList(CDaoTableDef *pTDef)
{
	ASSERT_VALID(pTDef);
	CStringList	*	fieldNameList;
	CDaoFieldInfo	tDefInfo;
	
	fieldNameList = new CStringList();
	int nCount = pTDef->GetFieldCount();
	for (int i=0; i<nCount; i++)
	{
		pTDef->GetFieldInfo(i, tDefInfo);

		fieldNameList->AddTail(tDefInfo.m_strName);
	}
	return fieldNameList;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/03]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :叛断一个指定的字符串是否在字符列表中，存在返回TRUE，否则返回FALSEL；
//------------------------------------------------------------------
BOOL CAplicatioInitDlg::StringInList(CString strTableName, CStringList *strList)
{
	CString		strTmp;

	for (POSITION pos=strList->GetHeadPosition(); pos != NULL; )
	{
		strTmp = strList->GetNext(pos);
		if ( !strTmp.CompareNoCase(strTableName) )
		{
			return TRUE;
		}
	}
	return FALSE;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/03]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :如果应用程序是第一次运行时,对数据库进行升级
//------------------------------------------------------------------
short CAplicatioInitDlg::UpgradeDatabase()
{
	if (_T("0") == theApp.GetRegKey(_T("Status"), _T("Install"), CString(_T("0"))) ) 
	{
		try
		{		
			CString strTo;
			CString strFrom;
			LPCTSTR szCopyFileInfo[]=
			{
				_T("AUTOIPED.MDB"),		//problem      xl
//				_T("IPEDsort.MDB"),
			};
			
			for(int i=0;i<sizeof(szCopyFileInfo)/sizeof(szCopyFileInfo[0]);i++)
			{
				strFrom=gsTemplateDir;
				strTo=gsProjectDBDir;
				
				FormatDirPath(strFrom);
				FormatDirPath(strTo);
				
				strFrom+=szCopyFileInfo[i];
				strTo+=szCopyFileInfo[i];
				
				UpgradeDB(strTo,strFrom);		//升级库
			}
			//
			// 从gsTemplateDir到gsProjectDir
			//
			
			LPCTSTR szCopyFileInfo2[]=
			{
				_T("WORKPRJ.MDB")
			};		
			for(i=0;i<sizeof(szCopyFileInfo2)/sizeof(szCopyFileInfo2[0]);i++)
			{
				strFrom=gsTemplateDir;
				strTo=gsProjectDir;
				
				FormatDirPath(strFrom);
				FormatDirPath(strTo);
				
				strFrom+=szCopyFileInfo2[i];
				strTo+=szCopyFileInfo2[i];

				UpgradeDB(strTo,strFrom);		//升级库				
			}
			//

			LPCTSTR szCopyFileInfo3[]=
			{
				_T("TableFormat.MDB"),
				_T("DrawingSize.MDB"),
				_T("Material.MDB"),
				_T("Medium.mdb"),
				_T("MaterialName.mdb"),
				_T("ALLPRJDB.MDB")
			};		
			for(i=0;i<sizeof(szCopyFileInfo3)/sizeof(szCopyFileInfo3[0]);i++)
			{
				strFrom=ReadCommonMdbPath();
				if ( "" == strFrom )
				{
					return FALSE;
				}
				strTo=gsShareMdbDir;

				FormatDirPath(strFrom);
				FormatDirPath(strTo);
				
				strFrom+=szCopyFileInfo3[i];
				strTo+=szCopyFileInfo3[i];

				UpgradeDB(strTo,strFrom);		//升级库				
			}
			//
			theApp.SetRegValue(_T("Status"), _T("Install"), CString(_T("1")));
		}
		catch (_com_error& e) 
		{
			AfxMessageBox(e.Description());
			return 0;
		}
	}
	return 1;
}
