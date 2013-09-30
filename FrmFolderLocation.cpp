// FrmFolderLocation.cpp : implementation file
//

#include "stdafx.h"
#include "autoiped.h"
#include "FrmFolderLocation.h"
#include "BrowseForFolerModule.h"
#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
#include "EDIBgbl.h"
/////////////////////////////////////////////////////////////////////////////
// CFrmFolderLocation dialog
//
CString gstrEDIBdir[4];

CFrmFolderLocation::CFrmFolderLocation(CWnd* pParent /*=NULL*/)
	: CDialog(CFrmFolderLocation::IDD, pParent)
{
	//{{AFX_DATA_INIT(CFrmFolderLocation)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CFrmFolderLocation::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CFrmFolderLocation)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CFrmFolderLocation, CDialog)
	//{{AFX_MSG_MAP(CFrmFolderLocation)
	ON_BN_CLICKED(IDC_SF1, OnSf1)
	ON_BN_CLICKED(IDC_SF2, OnSf2)
	ON_BN_CLICKED(IDC_SF3, OnSf3)
	ON_BN_CLICKED(IDC_SF4, OnSf4)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CFrmFolderLocation message handlers

BOOL CFrmFolderLocation::OnInitDialog() 
{
	CDialog::OnInitDialog();
    SetWindowCenter(this->m_hWnd);
	CString strDir;
	LoadCaption();

	InitgstrEDIBdir();
	CString strCmp;
	for (int i=0;i<4;i++)      //设置标题和路径。
	{
		strCmp.LoadString(IDS_EDIBDir0+i);
		SetDlgItemText(IDC_DIR1+i, strCmp);

		strDir = GetDir(gstrEDIBdir[i]);
		while(strDir.Right(1)=="\\")     //消除最后的'\'
		{
			strDir.Delete(strDir.GetLength()-1);
		}
		SetDlgItemText(IDC_FL1+i, strDir);
		m_strFormDir[i] = strDir;    //保存原来的路径名。
	}

    //iEDIBinstallDir=6
	((CEdit*)GetDlgItem(IDC_FL1))->SetReadOnly();
	GetDlgItem(IDC_SF1)->EnableWindow(false);

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

//设置对话框标题。
void CFrmFolderLocation::LoadCaption()
{
	CString str;
	str.LoadString(IDS_frmFolderLocation_frmFolderLocation);
	SetWindowText(str);

	str.LoadString(IDS_WarnDontRandomEDITthisDirectoryString);
	SetDlgItemText(IDC_CAUTION,str);

}
//窗口位置
void CFrmFolderLocation::SetWindowCenter(HWND hWnd)
{
	long SW,SH;
	if(hWnd==NULL)
		return;
	SW=::GetSystemMetrics(SM_CXSCREEN);
	SH=::GetSystemMetrics(SM_CYSCREEN);
	CRect rc;
	::GetWindowRect(hWnd,&rc);
	long x=(SW-rc.Width()) / 2;
	long y=(SH-rc.Height()) / 2;
	::SetWindowPos(hWnd,NULL,x,y,0,0,SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
}
//初始化，保存在注册表中的路径的，
void CFrmFolderLocation::InitgstrEDIBdir()
{
	gstrEDIBdir[0] = "EDInBox_InsDir";				//应用程序目录。

	gstrEDIBdir[1] = "EDInBox_TemplateDir";			//模板文件目录

	gstrEDIBdir[2] = "EDInBox_prjDbDir";              //项目数据库存放目录。

    gstrEDIBdir[3] = "EDInBox_prjDir";				//项目临时数据库存放目录。

}

CString CFrmFolderLocation::GetDir(CString key)
{
	CString sDir("");
	sDir=GetRegKey(_T("directory"),key,CString(_T("")));
	sDir.MakeLower();
	return sDir;

}

CString CFrmFolderLocation::GetRegKey(LPCTSTR pszKey, LPCTSTR pszName, const CString Default)
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
	ret.TrimLeft();ret.TrimRight();
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
	//ret= Default;
	RegCloseKey(hKey);
	return ret;

}

void CFrmFolderLocation::SetRegValue(LPCTSTR pszKey, LPCTSTR pszName, const CString vValue)
{
		HKEY hKey;
	CString subKey=szSoftwareKey;
	subKey+=pszKey;
	if(RegOpenKey(HKEY_LOCAL_MACHINE,subKey,&hKey)!=ERROR_SUCCESS)
	{
		RegCreateKey(HKEY_LOCAL_MACHINE,subKey,&hKey);
	}
	if(RegSetValueEx(hKey,pszName,NULL,REG_SZ,(LPBYTE)(LPCTSTR)vValue,(vValue.GetLength()+1)*sizeof(TCHAR))!=ERROR_SUCCESS)
	{
		if(::RegDeleteValue(hKey,pszName)==ERROR_SUCCESS)
			RegSetValueEx(hKey,pszName,NULL,REG_SZ,(LPBYTE)(LPCTSTR)vValue,(vValue.GetLength()+1)*sizeof(TCHAR));
	}
	RegCloseKey(hKey);

}
//选择对话框。
void CFrmFolderLocation::OnSf(UINT uID)
{
	CString strStart, strSelect;
	CFrmFolderLocation::GetDlgItemText(uID, strStart);
	while(strStart.Right(1)=="\\")
	{
		strStart.Delete(strStart.GetLength()-1);
	}
	BrowseForFolerModule::BrowseForFoldersFromPathStart(this->GetSafeHwnd(),
		           strStart, strSelect);
	if(!strSelect.IsEmpty())
	{
		CFrmFolderLocation::SetDlgItemText(uID, strSelect);
	}

}

void CFrmFolderLocation::OnSf1() 
{
	OnSf(IDC_FL1);	
}

void CFrmFolderLocation::OnSf2() 
{
	OnSf(IDC_FL2);	
	
}

void CFrmFolderLocation::OnSf3() 
{
	OnSf(IDC_FL3);	
	
}

void CFrmFolderLocation::OnSf4() 
{
	OnSf(IDC_FL4);	
	
}


void CFrmFolderLocation::OnOK() 
{
	CString strCap, strMess, str;
	bool PrjDBDirChanged = false;//项目数据库目录/或项目临时数据库目录， 被改变，需要升级
	bool PrjXDirChanged = false;
	for(int i=0; i<4; i++)
	{
		CFrmFolderLocation::GetDlgItemText(IDC_FL1+i, strCap);
		strCap.TrimLeft( );
		strCap.TrimRight( );
		while(strCap.Right(1)=="\\" || (!strCap.IsEmpty() && strCap.Right(1)==" "))       
		{
			strCap.Delete(strCap.GetLength()-1);
		}

		if( !IsDirExists(strCap + "\\") )//判断文件夹是否存在。
		{
			strMess.Format("' %s '不是一个有效的路径，请重新设置！",strCap);
			::MessageBox(NULL, strMess, "AutoIPED", MB_OK|MB_ICONEXCLAMATION|MB_SYSTEMMODAL);
			this->BringWindowToTop();
			return;
		}
		if( strCap.IsEmpty() )
		{
			str.LoadString(IDS_EDIBDir0+i);
			strMess = ("'"+str+"'不能为空！");
			AfxMessageBox(strMess);
			return;
		}
		SetRegValue("directory", gstrEDIBdir[i], strCap+"\\");   //写入注册表中加'\'
		
		if( strCap != m_strFormDir[i] )    //原来的路径是否改变。
		{
			PrjXDirChanged=true;
			if( i == 2/*prjDBDir*/|| i == 3/*prjDir*/)        //数据库目录改变。
			{
				PrjDBDirChanged=true;
			}
		}

	}
	if (PrjXDirChanged)
	{
		if (IDYES==(::MessageBox(this->m_hWnd,_T("重要的目录数据被修改，需要重新启动！"),
			_T("AutoIPED"),MB_YESNO|MB_ICONEXCLAMATION|MB_SYSTEMMODAL)))
		{
			((CMainFrame*)AfxGetApp()->m_pMainWnd)->m_bIsExitMsg=false;
			//新路径下的数据库可能需要升级。
			if (PrjDBDirChanged)
			{
				SetRegValue(_T("Status"), _T("Install"), CString(_T("0")));
			}
			::PostMessage(AfxGetMainWnd()->m_hWnd,WM_CLOSE,0,0);

		}
	}

	CDialog::OnOK();
}

bool CFrmFolderLocation::IsDirExists(CString Directory)
{
	DWORD code;
	code = GetFileAttributes(Directory);
	if(code!=-1 && (code & FILE_ATTRIBUTE_DIRECTORY)!=0)
		return TRUE;
	return FALSE;
}
