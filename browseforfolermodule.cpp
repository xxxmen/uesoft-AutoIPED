// BrowseForFolerModule.cpp
//

#include "stdafx.h"
#include "BrowseForFolerModule.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
//int BrowseForFolerModule::WM_USER = (int) &H400;
//long BrowseForFolerModule::BFFM_SETSELECTIONA = (long) (WM_USER + 102);
//long BrowseForFolerModule::BFFM_SETSELECTIONW = (long) (WM_USER + 103);
//int BrowseForFolerModule::LPTR = (int) (&H0|&H40);


//void BrowseForFolerModule::CoTaskMemFree(Long ByVal pv)
//{
//}

//Long BrowseForFolerModule::SendMessage(Long ByVal hwnd, Long ByVal wMsg, Long ByVal wParam, Any lParam)
//{
//    return (Long)0;
//}

//void BrowseForFolerModule::CopyMemory(Any pDest, Any pSource, Long ByVal dwLength)
//{
//}

//Long BrowseForFolerModule::LocalAlloc(Long ByVal uFlags, Long ByVal uBytes)
//{
//    return (Long)0;
//}

//Long BrowseForFolerModule::LocalFree(Long ByVal hMem)
//{
//    return (Long)0;
//}

//Long BrowseForFolerModule::BrowseCallbackProcStr(Long ByVal hwnd, Long ByVal uMsg, Long ByVal lParam, Long ByVal lpData)
//{
//    return (Long)0;
//}

//Long BrowseForFolerModule::FunctionPointer(Long FunctionAddress)
//{
//    return (Long)0;
//}

//String BrowseForFolerModule::BrowseForFolder(Long hwnd, String selectedPath)
//{
//}
int CALLBACK BrowseCallbackProc(HWND hwnd,UINT uMsg,LPARAM lParam,LPARAM lpData)
{
	if (uMsg==1){
		SendMessage(hwnd,BFFM_SETSELECTION,true,lpData);
	}
	return 0;
}

void BrowseForFolerModule::BrowseForFoldersFromPathStart(HWND hwnd, CString PathStart, CString &PathSelected)
{
	_TCHAR szTmp[MAX_PATH];//用于转换成PathSelected
	
	BROWSEINFO info;
	memset(&info,0,sizeof(BROWSEINFO));
	info.lpfn=BrowseCallbackProc;
	info.hwndOwner=hwnd;
    CString str;
	str.LoadString(IDS_SELECT_NEED_FILE);
	info.lpszTitle = str;//_T("请选择一个AutoPHS需要的文件夹");
	//Remove _T('\') at the end of the path if user added it
	while(PathStart.Right(1)==_T("\\")) 
		PathStart.Delete(PathStart.GetLength()-1,1);
	info.lParam=(DWORD)(LPCTSTR)PathStart;

	SHGetPathFromIDList(SHBrowseForFolder(&info),szTmp);
	if (szTmp==_T("")){CString str;
	     str.LoadString(IDS_NOSELECT_ANYFILE);
		MessageBox(hwnd,str,_T("AutoPHS"),MB_OK|MB_TASKMODAL|MB_TOPMOST);
		PathSelected=PathStart;
	}
	else PathSelected=szTmp;

	//PathStart.ReleaseBuffer();	
}


