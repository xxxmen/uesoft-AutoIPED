// AutoIPED.cpp : Defines the class behaviors for the application.
//

#include "HyperLink.h"	// Added by ClassView
#include "stdafx.h"
#include "AutoIPED.h"

#include "MainFrm.h"
#include "AutoIPEDDoc.h"
#include "AutoIPEDView.h"
#include "SelEngVolDll.h"
#include "EDIBgbl.h"
#include "Newproject.h"
#include "HyperLink.h"

#include "AplicatioInitDlg.h"	//软件初始化对话框

//#include "ImportFromPreCalcXLSDlg"
#include "ImportFromPreCalcXLSDlg.h"
#include "ProjectOperate.h"
#include "mainfrm.h"
#include "ImportAutoPD.h"

#include "CoolControlsManager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//////////////////////////////////////////////////////////////////////////
// CAutoIPEDApp

BEGIN_MESSAGE_MAP(CAutoIPEDApp, CWinApp)
	//{{AFX_MSG_MAP(CAutoIPEDApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
	ON_COMMAND(ID_FILE_NEW, OnFileNew)
//	ON_COMMAND(IDM_REPLACE_CURRENT_OLDNAME_NEWNAME, OnReplaceCurrentOldnameNewname)
//	ON_UPDATE_COMMAND_UI(IDM_REPLACE_CURRENT_OLDNAME_NEWNAME, OnUpdateReplaceCurrentOldnameNewname)
	//}}AFX_MSG_MAP
	// Standard file based document commands
//	ON_COMMAND(ID_FILE_NEW, CWinApp::OnFileNew)
	ON_COMMAND(ID_FILE_OPEN, CWinApp::OnFileOpen)
	// Standard print setup command
	ON_COMMAND(ID_FILE_PRINT_SETUP, CWinApp::OnFilePrintSetup)
END_MESSAGE_MAP()

//////////////////////////////////////////////////////////////////////////////
//全局变量

const _TCHAR szSoftwareKey[]=_T("Software\\长沙优易软件开发有限公司\\AutoIPED\\8.0\\");
CString gsIPEDInsDir;	//AutoIPED安装目录
CString gsProjectDBDir;	//AutoIPED数据库安装目录
CString gsProjectDir;	//本地用户配置数据库目录
CString gsTemplateDir;	//本地用户配置数据库目录
CString gsShareMdbDir;  //共享数据库的存放目录  xl 10.19.2007
CString gsHelpFilePath; //帮助文档路径  Luorj   2008.06.18


BOOL bIsCloseCompress;  //关闭时是否压缩数据库 
BOOL bIsReplaceOld;		//打开时是否自动替换旧材料名称
BOOL bIsMoveUpdate;		//编辑原始数据移动记录时是否更新
BOOL bIsAutoSelectPre;  //编辑原始数据移动记录时自动选择保温材料
BOOL bIsHeatLoss;		//计算时判断最大散热密度
BOOL gbIsStopCalc;		//是否计算.
BOOL bIsAutoAddValve;	//编辑原始数据移动记录时在管道上自动增加阀门
BOOL gbIsRunSelectEng;	//启动时是否弹出选择工程卷册对话框.
BOOL gbIsSelTblPrice;	//选择表中的热价比主汽价值
int  giInsulStruct;		//计算时，保温结构与材料的选择。0，手工选择。     1, 重新自动选择。
int  giCalSteanProp;	//计算水和蒸汽性质的标准:		0, IAPWS-IF97。   1, IFC67
//int Logo;				//mark_hubb_标识导入何种数据库
BOOL gbIsReplaceUnit;	//导入保温油漆数据时替换单位
BOOL gbAutoPaint120;    //统计油漆时，自动加上保温数据介质温度小于120度的记录。
BOOL gbWithoutProtectionCost;    //计算经济厚度时不包含保护层材料费用，默认为0-包含
BOOL gbInnerByEconomic;    //双层异材保温计算经济厚度时内层不按表面温度法计算，默认为0-按表面温度法计算

//int	 giSupportOldCode;	//计算支吊架间距的标准.			0,新管规(国家电力行业标准,DL/T 5054-1996)  1,老管规(原国家电力部标准,DLGJ 23-81).
//int  giSupportRigidity;	//计算支吊架间距的刚度条件.		0,按强度条件  1, 按刚度条件	
//int  giHCrSteel;    	//高铬钢,


/////////////////////////////////////////////////////////////////////////////
// CAutoIPEDApp construction

CAutoIPEDApp::CAutoIPEDApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CAutoIPEDApp object

CAutoIPEDApp theApp;

//
// 全局共享的变量，保存第一个主窗口
// 防止多个程序的运行
//
#pragma data_seg(".IpedSeg")
HWND gShareMainWnd=NULL;		
#pragma data_seg()

#pragma comment(linker,"/SECTION:.IpedSeg,RWS")

/////////////////////////////////////////////////////////////////////////////
// CAutoIPEDApp initialization

BOOL CAutoIPEDApp::InitInstance()
{
	BOOL bRet;

	AfxEnableControlContainer();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	//
	// 如果已经有个程序在运行就停止运行此程序,
	// 并显示先前的程序
	//
/*	//对于运行多个程序不设限制, 考虑到期AutoPD导出数据到IPED时不能关闭手动打开的程序. [2005/07/22]
	if(gShareMainWnd && IsWindow(gShareMainWnd))
	{
		if(::IsIconic(gShareMainWnd))
		{
			::ShowWindow(gShareMainWnd,SW_RESTORE);
		}
		::SetWindowPos(gShareMainWnd,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_ASYNCWINDOWPOS);
		::SetWindowPos(gShareMainWnd,HWND_NOTOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_ASYNCWINDOWPOS);

		return FALSE;
	}
*/
	// Change the registry key under which our settings are stored.
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization.
	free((void*)m_pszRegistryKey);
	CString str;
	str.LoadString(IDS_UESOFTCO);
	SetRegistryKey(str);
	if(m_pszProfileName!=NULL)
		free((void*)m_pszProfileName);
	m_pszProfileName=_tcsdup(_T("AutoIPED\\8.0"));
	LoadStdProfileSettings(8);  // Load standard INI file options (including MRU)

	ReadInitKey();//读入所有注册表值

	//对本工程的数据库进行初始化,拷贝或升级数据库.
	CopyFromTemplateDirToPrjDir();

	CString strHelpPath=gsHelpFilePath+_T("AutoIPED.chm");
	// Register the application's document templates.  Document templates
	//  serve as the connection between documents, frame windows and views.

	free((void*)this->m_pszHelpFilePath);
	m_pszHelpFilePath=NULL;
	WriteProfileString(_T("Directory"),_T("HelpFilePath"),strHelpPath);
	m_pszHelpFilePath=_tcsdup(strHelpPath);


	strHelpPath=this->GetProfileString(_T("Directory"),_T("HelpFilePath"),strHelpPath);
	if(FileExists(strHelpPath))
	{
		if(m_pszHelpFilePath != NULL)
		{
			free((void*)this->m_pszHelpFilePath);
			m_pszHelpFilePath = NULL;
		}
		this->m_pszHelpFilePath=_tcsdup(strHelpPath);
		this->WriteProfileString(_T("Directory"),_T("HelpFilePath"),strHelpPath);
	}

	::CoInitialize(NULL);
	AfxOleInit();
	m_pConnection.CreateInstance(__uuidof(Connection));
	m_pConnectionCODE.CreateInstance(__uuidof(Connection));
	m_pConAllPrj.CreateInstance( __uuidof(Connection) );

	m_pConMaterial.CreateInstance(__uuidof(Connection));
	m_pConMaterialName.CreateInstance( __uuidof(Connection) );
	m_pConIPEDsort.CreateInstance(__uuidof(Connection));

	m_pConWorkPrj.CreateInstance(__uuidof(Connection));
	m_pConWeather.CreateInstance(__uuidof(Connection));

	m_pConMedium.CreateInstance(__uuidof(Connection));
	m_pConRefInfo.CreateInstance(__uuidof(Connection));

	// 在ADO操作中建议语句中要常用try...catch()来捕获错误信息，
	// 因为它有时会经常出现一些想不到的错误。
	CString strConn;
//	strConnCODE = CONNECTSTRING + gsProjectDBDir + "IPEDcode.mdb";	
//	strConn.Format(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s%s"),gsProjectDBDir,"AutoIPED.mdb");	
//	strConnCODE.Format(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s%s"),gsProjectDBDir,"IPEDcode.mdb");	
	try                 
	{	
		// 打开本地Access库bw.mdb
		strConn = CONNECTSTRING + gsProjectDBDir + "AutoIPED.mdb";
		m_pConnection->Open(_bstr_t(strConn),"","",-1);	

		//打开保温管理库
		strConn = CONNECTSTRING + gsTemplateDir + "IPEDsort.mdb";
		m_pConIPEDsort->Open(_bstr_t(strConn),"","",-1);	

		//打开项目临时库
		strConn = CONNECTSTRING + gsProjectDir + "WorkPrj.mdb";
		m_pConWorkPrj->Open(_bstr_t(strConn),"","",-1);	

		//打开介质数据库,该数据库放在共享路径下
		//strConn = CONNECTSTRING + EDIBgbl::strCur_MediumDBPath + "Medium.mdb";
		strConn = CONNECTSTRING + gsShareMdbDir + "Medium.mdb";
		m_pConMedium->Open(_bstr_t(strConn),"","",-1);
		
		// 行业管理数据库
		strConn = CONNECTSTRING + gsTemplateDir + "RefInfo.mdb";
		m_pConRefInfo->Open( _bstr_t( strConn ), _T(""), _T(""), -1 );

		// 工程数据库
		strConn = CONNECTSTRING + gsShareMdbDir + "AllPrjDB.mdb";
		m_pConAllPrj->Open(_bstr_t(strConn),"","",-1);
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	//打开标准库和材料库
	try
	{
		strConn = CONNECTSTRING + EDIBgbl::sCritPath + EDIBgbl::sCur_CritDbName;
		m_pConnectionCODE->Open(_bstr_t(strConn),"","",-1);	

		//材料库
		strConn = CONNECTSTRING + EDIBgbl::sMaterialPath + _T("Material.mdb");
		m_pConMaterial->Open(_bstr_t(strConn), "", "", -1 );

		//材料名数据库
		strConn = CONNECTSTRING + EDIBgbl::sMaterialPath + "MaterialName.mdb";
		m_pConMaterialName->Open( _bstr_t(strConn), "", "", -1 );

		// "中国城市气象参数.mdb";
		strConn = CONNECTSTRING	+ EDIBgbl::sCritPath + "中国城市气象参数.mdb";
		m_pConWeather->Open(_bstr_t(strConn), "", "", -1 );
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}

	//新建一个临时表，用于计算放热系数。
	EDIBgbl::NewCalcAlfaTable() ;

	// Register the application's document templates.  Document templates
	//  serve as the connection between documents, frame windows and views.

	CSingleDocTemplate* pDocTemplate;
	pDocTemplate = new CSingleDocTemplate(
		IDR_MAINFRAME,
		RUNTIME_CLASS(CAutoIPEDDoc),
		RUNTIME_CLASS(CMainFrame),       // main SDI frame window
		RUNTIME_CLASS(CAutoIPEDView));
	AddDocTemplate(pDocTemplate);

	// Parse command line for standard shell commands, DDE, file open
	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);

	//当命令行传入路径名出现空格时.自动解析时,不能将整个文件名解析出来.
	//一律都用新建文件.
	cmdInfo.m_nShellCommand=CCommandLineInfo::FileNew;	//zsy

	if (!ProcessShellCommand(cmdInfo))
		return FALSE;

	// The one and only window has been initialized, so show and update it.
	m_pMainWnd->ShowWindow(SW_SHOWMAXIMIZED);
	m_pMainWnd->UpdateWindow();

	//如果一个工程也没有就不调用对话框
	try
	{
		bRet=IsNoOneEngin();
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return FALSE;
	}

	CString strCommandLine;
	strCommandLine = ::GetCommandLine();
	//test 
//	strCommandLine = "C:\\Program Files\\长沙优易软件开发有限公司\\AutoIPED7.0\\AutoIPED.exe  -DATA_FILENAME C:\\Program Files\\长沙优易软件开发有限公司\\AutoIPED7.0\\Template\\导入Excel原始数据例题.xls  -PROJECT_ID 567  -PROJECT_ENGVOL 56  -RECORD_COUNT 10";
///	CImportAutoPD  test;
///	test.ImportAutoPDInsul(strCommandLine);
	
	//功能.  从AutoPD中导入保温数据到AutoIPED中. 
	//实现   pd生成Excel中间文件, 再由AutoIPED读取该中间文件.
	if( strCommandLine.Find(" -DATA_FILENAME ", 0 ) != -1 )
	{
 		CDocument *pDocument;
  		pDocument=((CFrameWnd*)AfxGetMainWnd())->GetActiveDocument();
 		if(pDocument!=NULL)
		{
			//清空新建的文档.
			((CAutoIPEDDoc*)pDocument)->m_Result=_T("");
			((CAutoIPEDDoc*)pDocument)->UpdateAllViews(NULL);
		}
		CImportAutoPD inPDInsul;
		//导入实现
		if( inPDInsul.ImportAutoPDInsul(strCommandLine) )
		{
			bRet = true;  //不再弹出下面的选择工程对话框
		}
	}
	
	if(!bRet)
	{
		if ( gbIsRunSelectEng )	//在选项对话框中的一个选项值。（启动时弹出“选择工程卷册”对话框）
		{
			//现在使用彭范庚开发的动态库选择工程卷册
			CSelEngVolDll dlg;

			CString AllPrjDBPathName=gsShareMdbDir+"AllPrjDb.mdb";
			CString sortPathName=gsShareMdbDir+"DrawingSize.mdb";
			CString workprjPathName=gsProjectDir+"workprj.mdb";
			LPSTR cAllPrjDBPathName = (LPSTR)(LPCTSTR) AllPrjDBPathName;
			LPSTR csortPathName = (LPSTR)(LPCTSTR) sortPathName;
			LPSTR cworkprjPathName = (LPSTR)(LPCTSTR) workprjPathName;
			
			if(dlg.ShowEngVolDlg( cworkprjPathName, cAllPrjDBPathName, csortPathName )==false)
			{
				// 在标题栏内显示当前的工程名称
				((CMainFrame*)m_pMainWnd)->ShowCurrentProjectName();
				return TRUE;
			}
						
			EDIBgbl::SelSpecID=dlg.m_SelSpecID;
			EDIBgbl::SelPrjID=dlg.m_SelPrjID;
			EDIBgbl::SelPrjName=dlg.m_SelPrjName;//工程名称
			EDIBgbl::SelDsgnID=dlg.m_SelDsgnID;
			EDIBgbl::SelHyID=dlg.m_SelHyID;
			EDIBgbl::SelVlmID=dlg.m_SelVlmID;
			EDIBgbl::SelVlmCODE=dlg.m_SelVlmCODE;
			

			//EDIBgbl::SelPrjID="12";
			//EDIBgbl::SelPrjName="12";

			// 
			InitializeProjectTable( EDIBgbl::SelPrjID, EDIBgbl::SelVlmCODE, EDIBgbl::SelPrjName );
		}
	}
	

	//实现控件的浮动效果
//	GetCtrlManager().InstallHook();

	// 在标题栏内显示当前的工程名称
	((CMainFrame*)m_pMainWnd)->ShowCurrentProjectName();

	gShareMainWnd=m_pMainWnd->m_hWnd;

	return TRUE;
}


/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CHyperLink myURL;
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	CHyperLink	m_address;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}
 
void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	DDX_Control(pDX, IDC_STATIC_address, m_address);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

// App command to run the dialog
void CAutoIPEDApp::OnAppAbout()
{
	CAboutDlg aboutDlg;
	aboutDlg.DoModal();
}

/////////////////////////////////////////////////////////////////////////////
//自定义函数
CString CAutoIPEDApp::GetRegKey(LPCTSTR pszKey, LPCTSTR pszName,const CString Default)
{
	HKEY hKey;
	_TCHAR sv[256];
	unsigned long vtype;
	unsigned long vlen;
	CString subKey=szSoftwareKey;
	subKey+=pszKey;
	if(RegOpenKey(HKEY_LOCAL_MACHINE,subKey,&hKey)!=ERROR_SUCCESS)
	{
		return Default;
	}

	vlen=256*sizeof(TCHAR);
	if(RegQueryValueEx(hKey,pszName,NULL,&vtype,(LPBYTE)sv,&vlen)!=ERROR_SUCCESS)
	{
		RegSetValueEx(hKey,pszName,NULL,REG_SZ,(LPBYTE)(LPCTSTR)Default,(Default.GetLength()+1)*sizeof(_TCHAR));
		RegCloseKey(hKey);
		return Default;
	}

	CString ret=Default;
	if(vtype==REG_SZ)
		ret=sv;
	ret.TrimLeft();
	ret.TrimRight();
	if(ret==_T(""))
	{
		RegSetValueEx(hKey,pszName,NULL,REG_SZ,(LPBYTE)(LPCTSTR)Default,(Default.GetLength()+1)*sizeof(_TCHAR));
		ret = Default;
	}
	if(vtype==REG_DWORD)
	{
		DWORD dwVal=*((DWORD*)(void*)sv);
		ret.Format(_T("%d"),dwVal);
	}

	RegCloseKey(hKey);
	return ret;
}

void CAutoIPEDApp::SetRegValue(CString vKey, CString vName, CString vValue)
{
	HKEY hKey;

	CString subKey=CString(szSoftwareKey+vKey);
	if(RegOpenKey(HKEY_LOCAL_MACHINE,subKey,&hKey)!=ERROR_SUCCESS)
	{
		RegCreateKey(HKEY_LOCAL_MACHINE,subKey,&hKey);
	}
	RegSetValueEx(hKey,vName,NULL,REG_SZ,(unsigned char*)LPCSTR(vValue),vValue.GetLength()+1);
	RegCloseKey(hKey);
}


BOOL CAutoIPEDApp::ReadInitKey()
{
	//功能：读入所有注册表值，在程序启动时调用
	//default directory is Appdir + "\\Example"
	char szPath[MAX_PATH];
	::GetCurrentDirectory( sizeof(szPath) , szPath );
	CString strDefault = "";

	//软件安装路径
	gsIPEDInsDir = GetRegKey(_T("Directory"),_T("EDinBox_InsDir"),strDefault);
	gsIPEDInsDir.TrimLeft();
	gsIPEDInsDir.TrimRight();
	if (gsIPEDInsDir.IsEmpty())
	{
		gsIPEDInsDir.Format("%s\\" , szPath);
		SetRegValue(_T("Directory"),_T("EDinBox_InsDir"),gsIPEDInsDir);
	}
	if( gsIPEDInsDir.Right(1) != "\\" )
	{
		gsIPEDInsDir += "\\";
	}
	HKEY key;
	DWORD szSize = 255;
	
	if( RegOpenKeyEx( HKEY_LOCAL_MACHINE, _T("SOFTWARE\\长沙优易软件开发有限公司\\shareMDB"),0, KEY_ALL_ACCESS, &key) != ERROR_SUCCESS )
	{
		AfxMessageBox( "打开注册表失败!" );
		return FALSE;
	}
    gsShareMdbDir = "";
	if( RegQueryValueEx(key, "shareMDB", NULL, NULL, (LPBYTE)/*(LPTSTR)(LPCTSTR)*/gsShareMdbDir.GetBufferSetLength(255), &szSize) != ERROR_SUCCESS )
	{
		AfxMessageBox( "打开注册表失败!" );
		return FALSE;
	}
	
	RegCloseKey( key );
	gsShareMdbDir.TrimLeft();
	gsShareMdbDir.TrimRight();
	gsShareMdbDir.ReleaseBuffer();
	if( gsShareMdbDir.Right(1) != "\\" )
	{
		gsShareMdbDir += "\\";
	}

	//获得标准数据库路径
	EDIBgbl::sCritPath = GetRegKey(_T("Directory"),_T("EDinBox_CodeDBDir"), "");
	EDIBgbl::sCritPath.TrimLeft();
	EDIBgbl::sCritPath.TrimRight();
	if( EDIBgbl::sCritPath.IsEmpty() )
	{
		EDIBgbl::sCritPath = gsIPEDInsDir + "Code\\";
		SetRegValue(_T("Directory"), _T("EDinBox_CodeDBDir"), EDIBgbl::sCritPath);
	}
	if( EDIBgbl::sCritPath.Right(1) != "\\" )
	{
		EDIBgbl::sCritPath += "\\";
	}
	//获得材料数据库路径
	//改为与AutoPDMS3.0共享目录一致 ligb 2010.08.11
	EDIBgbl::sMaterialPath = gsShareMdbDir;
	if( EDIBgbl::sMaterialPath.IsEmpty() )
	{
		AfxMessageBox(IDS_SHAREDIR_ISEMPTY,MB_OK); 
	}
	//项目数据库路径
	strDefault = "";	
	gsProjectDBDir=GetRegKey(_T("Directory"),_T("EDInBox_prjDbDir"),strDefault);
	gsProjectDBDir.TrimLeft();
	gsProjectDBDir.TrimRight();
	if (gsProjectDBDir.IsEmpty())
	{
		gsProjectDBDir.Format("%s" ,_T("c:\\IPEDprjdb8.0\\"));
		SetRegValue(_T("Directory"),_T("EDInBox_prjDbDir"),gsProjectDBDir);
	}
	if( gsProjectDBDir.Right(1) != "\\" )
	{
		gsProjectDBDir += "\\";
	}
	//临时数据库路径
	strDefault = "";	
	gsProjectDir=GetRegKey(_T("Directory"),_T("EDInBox_prjDir"),strDefault);
	gsProjectDir.TrimLeft();
	gsProjectDir.TrimRight();
	if (gsProjectDir.IsEmpty())
	{
		gsProjectDir.Format("%s" ,_T("c:\\IPEDprj8.0\\"));
		SetRegValue(_T("Directory"),_T("EDInBox_prjDir"),gsProjectDir);
	}
	if( gsProjectDir.Right(1) != "\\" )
	{
		gsProjectDir += "\\";
	}
	//模板路径
	strDefault = "";	
	gsTemplateDir=GetRegKey(_T("Directory"),_T("EDinBox_templateDir"),strDefault);
	gsTemplateDir.TrimLeft();
	gsTemplateDir.TrimRight();
	if (gsTemplateDir.IsEmpty())
	{
		gsTemplateDir = gsIPEDInsDir + "Templater\\";
		SetRegValue(_T("Directory"),_T("EDinBox_templateDir"),gsTemplateDir);
	}
	if( gsTemplateDir.Right(1) != "\\" )
	{
		gsTemplateDir += "\\";
	}
	//帮助文档路径
	strDefault = "";	
	gsHelpFilePath=GetRegKey(_T("Directory"),_T("HelpFilePath"),strDefault);
	gsHelpFilePath.TrimLeft();
	gsHelpFilePath.TrimRight();
	if (gsHelpFilePath.IsEmpty())
	{
		gsHelpFilePath = gsIPEDInsDir + "Help\\";
		SetRegValue(_T("Directory"),_T("HelpFilePath"),gsHelpFilePath);
	}
	if( gsHelpFilePath.Right(1) != "\\" )
	{
		gsHelpFilePath += "\\";
	}

	//介质数据库路径
	// 放到项目数据库中  [2005/06/13]
/*	EDIBgbl::strCur_MediumDBPath = GetRegKey(_T("Directory"),_T("EDINBox_MediumDBDir"),"");
	EDIBgbl::strCur_MediumDBPath.TrimLeft();
	EDIBgbl::strCur_MediumDBPath.TrimRight();
	if(EDIBgbl::strCur_MediumDBPath.IsEmpty())
	{
		EDIBgbl::strCur_MediumDBPath = gsProjectDBDir;
		SetRegValue(("Directory"),_T("EDINBox_MediumDBDir"),EDIBgbl::strCur_MediumDBPath);
	}
	if (EDIBgbl::strCur_MediumDBPath.Right(1) != "\\" )
	{
		EDIBgbl::strCur_MediumDBPath += "\\";
	}
*/
	
	/////////////////////////选项对话框中的状态数据。
	CString str = GetRegKey(_T("Settings"),_T("CompactDBatClose"),_T(""));
	str.TrimLeft();
	str.TrimRight();
	if(str.IsEmpty())
	{
		SetRegValue(_T("Settings"),_T("CompactDBatClose"),_T("1"));
		str = _T("1");
	}
	if(!str.Compare("1"))
	{
		bIsCloseCompress = TRUE;
	}else 
		bIsCloseCompress = FALSE; 

	str = GetRegKey(_T("Settings"),_T("ReplaceOldMaterialWithNewMaterial"),_T("1"));
	if(!str.Compare("1"))
	{
		bIsReplaceOld = TRUE;
	}else
		bIsReplaceOld = FALSE;

	str = GetRegKey(_T("Settings"),_T("UpdateRecordAfterMove"),_T("1"));
	if(!str.Compare("1"))
	{
		bIsMoveUpdate = TRUE;
	}else
		bIsMoveUpdate = FALSE;

	str = GetRegKey(_T("Settings"),_T("AutoSelectPre"),_T("1"));
	if(!str.Compare("1"))
	{
		bIsAutoSelectPre = TRUE;
	}else
		bIsAutoSelectPre = FALSE;
	
 	//编辑原始数据移动记录时在管道上自动增加阀门
	str = GetRegKey(_T("Settings"),_T("AutoAddValve"),_T("0"));
	if(!str.Compare("1"))
	{
		bIsAutoAddValve = TRUE;
	}else
		bIsAutoAddValve = FALSE;

	//计算时判断最大散热密度
	str = GetRegKey(_T("Settings"),_T("IsHeatLoss"),_T("1"));
	str.TrimLeft();
	str.TrimRight();
	if(!str.Compare("1"))
	{
		bIsHeatLoss = TRUE;
	}else
		bIsHeatLoss = FALSE;
	//启动时是否弹出选择工程卷册对话框.
	str = GetRegKey( _T( "Settings" ), _T( "IsRunSelEng" ), _T( "1" ));
	gbIsRunSelectEng = !str.CompareNoCase( "1" ) ? TRUE : FALSE;

 	//编辑原始数据，自动时取表中的热价比主汽价
	str = GetRegKey( _T( "Settings" ), _T( "IsSelPrice" ), _T( "0" ));
	gbIsSelTblPrice = !str.CompareNoCase( "1" ) ? TRUE : FALSE;
	
 	//导入保温油漆数据时替换单位
	str = GetRegKey( _T( "Settings" ), _T( "IsReplaceUnit" ), _T( "0" ));
	gbIsReplaceUnit = !str.CompareNoCase( "1" ) ? TRUE : FALSE;
	
 	//统计油漆时，自动加上保温数据介质温度小于120度的记录。
	str = GetRegKey( _T( "Settings" ), _T( "AutoPaint120" ), _T( "1" ));
	gbAutoPaint120 = !str.CompareNoCase( "1" ) ? TRUE : FALSE;

	//计算保温结构。
	str = GetRegKey(_T("Settings"),_T("CalcInsulStruct"),_T("0"));
	giInsulStruct = _tcstol(str,NULL,10);

	//计算水和蒸汽性质的标准
	str = GetRegKey( _T("Settings"), _T("CalcSteanProp"), _T("0") );
	giCalSteanProp = _tcstol(str, NULL, 10);

 	//计算经济厚度时不包含保护层材料费用
	str = GetRegKey( _T( "Settings" ), _T( "WithoutProtectionCost" ), _T( "0" ));
	gbWithoutProtectionCost = !str.CompareNoCase( "1" ) ? TRUE : FALSE;

 	//双层异材保温计算经济厚度时内层不按表面温度法计算
	str = GetRegKey( _T( "Settings" ), _T( "InnerByEconomic" ), _T( "0" ));
	gbInnerByEconomic = !str.CompareNoCase( "1" ) ? TRUE : FALSE;

	//读取存于IPED.ini文件中的数据库名称
	EDIBgbl::GetCurDBName();

	return true;
}

BOOL FileExists(LPCTSTR lpszPathName)
{
	DWORD att=::GetFileAttributes(lpszPathName);
	if(att==0xFFFFFFFF || ((att & FILE_ATTRIBUTE_DIRECTORY)!=0) ) return FALSE;
	return TRUE;

}
BOOL DirExists(LPCTSTR lpszDir)
{
	DWORD att=::GetFileAttributes(lpszDir);
	if(att!=0xFFFFFFFF && ((att & FILE_ATTRIBUTE_DIRECTORY)!=0) ) return TRUE;
	return FALSE;
}

//判断指定表是否存在pCon数据库中。
BOOL IsTableExists(_ConnectionPtr pCon, CString tb)
{
	if(pCon==NULL || tb=="")
		return FALSE;
	_RecordsetPtr rs;
	if(tb.Left(1)!="[")
		tb="["+tb;
	if(tb.Right(1)!="]")
		tb+="]";
	try{
		rs=pCon->Execute(_bstr_t(tb), NULL, adCmdTable);
		rs->Close();
		return TRUE;
	}
	catch(_com_error e)
	{
		return FALSE;
	}
}

CString _GetFileName(LPCTSTR lpszPathName)
{
	CString strFile;
	::GetFileTitle(lpszPathName,strFile.GetBuffer(MAX_PATH),MAX_PATH);
	strFile.ReleaseBuffer();
	return strFile;
}

/////////////////////////////////////////////////////////////////////////////
// CAutoIPEDApp message handlers


int CAutoIPEDApp::ExitInstance() 
{
	WriteRegedit();
	
	::CoUninitialize();	
	return CWinApp::ExitInstance();
}

BOOL CAboutDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	myURL.SetURL("http://www.uesoft.com");
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

bool CAutoIPEDApp::WriteRegedit()
{
	CString strTmp;
	//选项对话框中的状态 
	if(bIsCloseCompress)
	{
		theApp.SetRegValue("Settings", "CompactDBatClose", "1");
	}else
		theApp.SetRegValue("Settings", "CompactDBatClose", "0");

	if(bIsReplaceOld)
	{
		theApp.SetRegValue("Settings", "ReplaceOldMaterialWithNewMaterial", "1");
	}else
		theApp.SetRegValue("Settings", "ReplaceOldMaterialWithNewMaterial", "0");
	
	if(bIsMoveUpdate)
	{
		theApp.SetRegValue("Settings", "UpdateRecordAfterMove", "1");
	}else
		theApp.SetRegValue("Settings", "UpdateRecordAfterMove", "0");

	if(bIsAutoSelectPre)
	{
		theApp.SetRegValue("Settings", "AutoSelectPre", "1");
	}else
		theApp.SetRegValue("Settings", "AutoSelectPre", "0");

 	//编辑原始数据移动记录时在管道上自动增加阀门
	if (bIsAutoAddValve)
	{
		theApp.SetRegValue("Settings", "AutoAddValve", "1");
	}else
		theApp.SetRegValue("Settings", "AutoAddValve", "0");
 
	//计算时判断最大散热密度
	if(bIsHeatLoss)
	{
		theApp.SetRegValue("Settings", "IsHeatLoss", "1");
	}else
		theApp.SetRegValue("Settings", "IsHeatLoss", "0");
	
	//启动时是否弹出选择工程卷册对话框
	strTmp = ( gbIsRunSelectEng == 0 ) ? "0" : "1";
	theApp.SetRegValue( "Settings", "IsRunSelEng", strTmp );
	
 	//编辑原始数据，自动时取表中的热价比主汽价
	strTmp = ( gbIsSelTblPrice == 0 ) ? "0" : "1";
	theApp.SetRegValue( "Settings", "IsSelPrice", strTmp );

 	// 导入保温油漆数据时替换单位
	strTmp = ( gbIsReplaceUnit == 0 ) ? "0" : "1";
	theApp.SetRegValue( "Settings", "IsReplaceUnit", strTmp );

 	// 统计油漆时，自动加上保温数据介质温度小于120度的记录。
	strTmp = ( gbAutoPaint120 == 0 ) ? "0" : "1";
	theApp.SetRegValue( "Settings", "AutoPaint120", strTmp );

 	// 计算经济厚度时不包含保护层材料费用，默认为0-包含
	strTmp = ( gbWithoutProtectionCost == 0 ) ? "0" : "1";
	theApp.SetRegValue( "Settings", "WithoutProtectionCost", strTmp );

 	// 双层异材保温计算经济厚度时内层不按表面温度法计算，默认为0-按表面温度法计算
	strTmp = ( gbInnerByEconomic == 0 ) ? "0" : "1";
	theApp.SetRegValue( "Settings", "InnerByEconomic", strTmp );

	//重新选择保温结构。
	strTmp.Format("%d",giInsulStruct);
	theApp.SetRegValue("Settings", "CalcInsulStruct", strTmp);
	//计算水和蒸汽性质的标准
	strTmp.Format("%d", giCalSteanProp );
	theApp.SetRegValue("Settings", "CalcSteanProp", strTmp);
	
	return true;
}

/////////////////////////////////////////////////////
//
// 拷贝数据库文件从模板目录到工程目录
//
void CAutoIPEDApp::CopyFromTemplateDirToPrjDir()
{
	CAplicatioInitDlg InitializeDlg;

	InitializeDlg.DoModal();
	
}

////////////////////////////////////////
//
// 判断数据库中是否没有一个工程
//
// 如果一个工程也没有将返回TRUE，否则返回FALSE
//
// throw(_com_error)
//
BOOL CAutoIPEDApp::IsNoOneEngin()
{
	_RecordsetPtr IRecordset;
	HRESULT hr;

	if(this->m_pConnection==NULL)
	{
		_com_error e(E_POINTER);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	hr=IRecordset.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	try
	{
		IRecordset->Open(_variant_t("Engin"),_variant_t((IDispatch*)m_pConAllPrj),
						 adOpenStatic,adLockOptimistic,adCmdTable);


		if(IRecordset->adoEOF && IRecordset->BOF)
		{
			return TRUE;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	return FALSE;
}

void CAutoIPEDApp::OnFileNew() 
{
	CWinApp::OnFileNew();

	CDocument *pDocument;

	pDocument=((CFrameWnd*)AfxGetMainWnd())->GetActiveDocument();

	if(pDocument==NULL)
		return;

	((CAutoIPEDDoc*)pDocument)->m_Result=_T("");

	((CAutoIPEDDoc*)pDocument)->UpdateAllViews(NULL);

	((CMainFrame*)::AfxGetMainWnd())->ShowCurrentProjectName();
}

/////////////////////////////////////////////////////
//
// 响应“替换当前工程中所有旧的材料名称”
//
void CAutoIPEDApp::OnReplaceCurrentOldnameNewname() 
{
	CMainFrame *pMainFrame;
	
	pMainFrame=(CMainFrame*)AfxGetMainWnd();

	pMainFrame->m_wndStatusBar.GetProgressCtrl().SetRange(0,100);

//	pMainFrame->m_wndStatusBar.OnProgress(10);
}

void CAutoIPEDApp::OnUpdateReplaceCurrentOldnameNewname(CCmdUI* pCmdUI) 
{
	// TODO: Add your command update UI handler code here
	
}

BOOL CAutoIPEDApp::InitializeProjectTable( const CString& strPrjID, const CString& strVolCode, 
										  const CString& strPrjName )
{
	newproject dlg;

	dlg.m_engin = EDIBgbl::SelPrjName  = strPrjName;
	dlg.m_eng_code = EDIBgbl::SelPrjID = strPrjID; 
	EDIBgbl::SelVlmCODE = strVolCode;
	dlg.SetAuto( true );

	return dlg.insertitem();
}