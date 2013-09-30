
#include "stdafx.h"
#include "ExceptionInfoHandle.h"
#include "mainfrm.h"

namespace Exception
{
	TCHAR gAdditiveInfo[1024]=_T("");	//用于附加的信息

	void ReportErrorInfo(LPCTSTR szErrorInfo)
	{
		if(lstrlen(gAdditiveInfo)>0)
		{
			lstrcat(gAdditiveInfo,_T("\r\n"));
			lstrcat(gAdditiveInfo,szErrorInfo);
		//	MessageBox(NULL,gAdditiveInfo,"Error",MB_OK|MB_ICONEXCLAMATION);
			MessageBox(((CMainFrame*)::AfxGetMainWnd())->m_hWnd,gAdditiveInfo,"Error",MB_OK|MB_ICONEXCLAMATION);
			lstrcpy(gAdditiveInfo,_T(""));
			return;
		}
		
//		MessageBox(NULL,szErrorInfo,"Error",MB_OK|MB_ICONEXCLAMATION);
		MessageBox(((CMainFrame*)::AfxGetMainWnd())->m_hWnd,szErrorInfo,"Error",MB_OK|MB_ICONEXCLAMATION);

		PutErrorInfoToFile(szErrorInfo);
	}

	void ReportErrorInfo(LPCTSTR szErrorInfo,LPCTSTR szErrorFile,int iLine)
	{
//		TCHAR tsErrorInfo[1024];
//
//		::_stprintf(tsErrorInfo,_T(" %s\r\n File:%s\r\n Line:%d"),szErrorInfo,szErrorFile,iLine);

		CString strError;
		strError.Format(_T(" %s"),szErrorInfo);
#ifdef _DEBUG
		CString strTmp;
		strTmp.Format(_T("\r\n File:  %s\r\n Line: %d"),szErrorFile,iLine);
		strError += strTmp;
#endif

		ReportErrorInfo(strError);
	}

	void ReportErrorInfo(_com_error &e,LPCTSTR szErrorFile,int iLine)
	{
		TCHAR tsErrorInfo[1024];
		TCHAR *pStr;
		int iRet;

		pStr=tsErrorInfo;

		iRet=_stprintf(pStr,_T("Error\r\n"));
		pStr+=iRet;

		iRet=_stprintf(pStr,_T("Code=%lx\r\n"),e.Error());
		pStr+=iRet;

		if(e.ErrorMessage()!=NULL)
		{
			iRet=_stprintf(pStr,_T("Code meaning=%s\r\n"),e.ErrorMessage());
			pStr+=iRet;
		}

		if(e.Source().length()>0)
		{
			iRet=_stprintf(pStr,_T("Source=%s\r\n"),(LPCTSTR)e.Source());
			pStr+=iRet;
		}

		if(e.Description().length()>0)
		{
			iRet=_stprintf(pStr,_T("Description=%s\r\n"),(LPCTSTR)e.Description());
			pStr+=iRet;
		}

		ReportErrorInfo(tsErrorInfo,szErrorFile,iLine);
	}

	void ReportErrorInfo(COleDispatchException *e,LPCTSTR szErrorFile,int iLine)
	{
		TCHAR tsErrorInfo[1024];
		TCHAR *pStr;
		int iRet;
		
		pStr=tsErrorInfo;

		iRet=_stprintf(pStr,_T("Error\r\n"));
		pStr+=iRet;

		iRet=_stprintf(pStr,_T("error code: %d\r\n"),e->m_wCode);
		pStr+=iRet;

		if(!e->m_strSource.IsEmpty())
		{
			iRet=_stprintf(pStr,_T("Source=%s\r\n"),e->m_strSource);
			pStr+=iRet;
		}

		if(!e->m_strDescription.IsEmpty())
		{
			iRet=_stprintf(pStr,_T("Description:%s\r\n"),e->m_strDescription);
			pStr+=iRet;
		}

		if(!e->m_strHelpFile.IsEmpty())
		{
			iRet=_stprintf(pStr,_T("Help File:\r\n"),e->m_strHelpFile);
			pStr+=iRet;

			iRet=_stprintf(pStr,_T("Help Context:%d"),e->m_dwHelpContext);
			pStr+=iRet;
		}

		ReportErrorInfo(tsErrorInfo,szErrorFile,iLine);
	}

	/////////////////////////////////////////
	//
	// 用于在ReportExceptionError输出附加的信息
	//
	// szInfo[in]	需要输出的信息,如果为NULL将回删
	//				除以前的所有输入信息。
	//
	// 主要用于在抛出异常时与异常一起附加显示的信息
	// 当调用ReportExceptionError等函数时，回显示并清空
	// 信息。当在之前调用多次调用该函数输入的信息回，不会相互
	// 覆盖。
	//
	void SetAdditiveInfo(LPCTSTR szInfo)
	{
		if(szInfo==NULL)
		{
			gAdditiveInfo[0]=_T('\0');
		}
		else
		{
			lstrcat(gAdditiveInfo,szInfo);
			lstrcat(gAdditiveInfo,_T("\r\n"));
		}
	}

	////////////////////////////////////////////
	//
	// 返回在ReportExceptionError输出附加的信息
	//
	LPCTSTR GetAdditiveInfo()
	{
		if(lstrlen(gAdditiveInfo)>0)
			return gAdditiveInfo;
		else
			return NULL;
	}

///////////////////////

	void ReportErrorInfoLV2(LPCTSTR szErrorInfo)
	{
		#ifdef _DEBUG
//		MessageBox(NULL,szErrorInfo,"Error",MB_OK);
		MessageBox(((CMainFrame*)::AfxGetMainWnd())->m_hWnd,szErrorInfo,"Error",MB_OK|MB_ICONEXCLAMATION);
		#endif

		PutErrorInfoToFile(szErrorInfo);
	}

	void ReportErrorInfoLV2(LPCTSTR szErrorInfo,LPCTSTR szErrorFile,int iLine)
	{
//		TCHAR tsErrorInfo[1024];
//
//		::_stprintf(tsErrorInfo,_T("ErrorLV2\r\n %s\r\n File:%s\r\n Line:%d"),szErrorInfo,szErrorFile,iLine);

		CString strError;
		strError.Format(_T(" %s"),szErrorInfo);
#ifdef _DEBUG
		CString strTmp;
		strTmp.Format(_T("\r\n File:  %s\r\n Line: %d"),szErrorFile,iLine);
		strError += strTmp;
#endif

//		ReportErrorInfoLV2(tsErrorInfo);
		ReportErrorInfoLV2(strError);
	}

	void ReportErrorInfoLV2(_com_error &e,LPCTSTR szErrorFile,int iLine)
	{
		TCHAR tsErrorInfo[1024];
		TCHAR *pStr;
		int iRet;

		pStr=tsErrorInfo;

		iRet=_stprintf(pStr,_T("ErrorLV2\r\n"));
		pStr+=iRet;

		iRet=_stprintf(pStr,_T("Code=%lx\r\n"),e.Error());
		pStr+=iRet;

		if(e.ErrorMessage()!=NULL)
		{
			iRet=_stprintf(pStr,_T("Code meaning=%s\r\n"),e.ErrorMessage());
			pStr+=iRet;
		}

		if(e.Source().length()>0)
		{
			iRet=_stprintf(pStr,_T("Source=%s\r\n"),(LPCTSTR)e.Source());
			pStr+=iRet;
		}

		if(e.Description().length()>0)
		{
			iRet=_stprintf(pStr,_T("Description=%s\r\n"),(LPCTSTR)e.Description());
			pStr+=iRet;
		}

		ReportErrorInfoLV2(tsErrorInfo,szErrorFile,iLine);
	}

	void ReportErrorInfoLV2(COleDispatchException *e,LPCTSTR szErrorFile,int iLine)
	{
		TCHAR tsErrorInfo[1024];
		TCHAR *pStr;
		int iRet;

		pStr=tsErrorInfo;

		iRet=_stprintf(pStr,_T("ErrorLV2\r\n"));
		pStr+=iRet;

		iRet=_stprintf(pStr,_T("error code: %d\r\n"),e->m_wCode);
		pStr+=iRet;

		if(!e->m_strSource.IsEmpty())
		{
			iRet=_stprintf(pStr,_T("Source=%s\r\n"),e->m_strSource);
			pStr+=iRet;
		}

		if(!e->m_strDescription.IsEmpty())
		{
			iRet=_stprintf(pStr,_T("Description:%s\r\n"),e->m_strDescription);
			pStr+=iRet;
		}

		if(!e->m_strHelpFile.IsEmpty())
		{
			iRet=_stprintf(pStr,_T("Help File:\r\n"),e->m_strHelpFile);
			pStr+=iRet;

			iRet=_stprintf(pStr,_T("Help Context:%d"),e->m_dwHelpContext);
			pStr+=iRet;
		}

		ReportErrorInfoLV2(tsErrorInfo,szErrorFile,iLine);
	}

	///////////////////////////////////////////
	//
	// 输出错误信息到文件中
	//
	// szInfo[in]	错误信息的字符串
	//
	void PutErrorInfoToFile(LPCTSTR szErrorInfo)
	{
		SYSTEMTIME CurrentTime;
		HANDLE hFile;
		TCHAR szInfo[1024*4];
		TCHAR *pStr;
		int iRet;
		DWORD dwRet;

		if(szInfo==NULL)
			return;

		//获得当前的时间
		GetLocalTime(&CurrentTime);

		pStr=szInfo;

		iRet=_stprintf(pStr,_T("\r\n\r\n错误描述:\r\n"));
		pStr+=iRet;

		iRet=_stprintf(pStr,_T("%d-%d-%d\r\n%d:%d:%d\r\n"),CurrentTime.wYear,CurrentTime.wMonth,CurrentTime.wDay,
					   CurrentTime.wHour,CurrentTime.wMinute,CurrentTime.wSecond);
		pStr+=iRet;

		_stprintf(pStr,_T("%s\r\n"),szErrorInfo);

		hFile=CreateFile(_T("ErrorDescription.txt"),GENERIC_WRITE,FILE_SHARE_WRITE,
						 NULL,OPEN_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL);

		SetFilePointer(hFile,0,NULL,FILE_END);

		WriteFile(hFile,szInfo,lstrlen(szInfo)*sizeof(TCHAR),&dwRet,NULL);

		CloseHandle(hFile);


	}
}