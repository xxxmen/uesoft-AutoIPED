
#pragma once 

#define ReportExceptionError(ERRORINFO)							\
	Exception::ReportErrorInfo((ERRORINFO),__FILE__,__LINE__);	\

#define ReportExceptionErrorLV1(ERRORINFO)							\
	Exception::ReportErrorInfo((ERRORINFO),__FILE__,__LINE__);		\

#define ReportExceptionErrorLV2(ERRORINFO)							\
	Exception::ReportErrorInfoLV2((ERRORINFO),__FILE__,__LINE__);		\

namespace Exception
{
	void ReportErrorInfo(LPCTSTR szErrorInfo);

	void ReportErrorInfo(LPCTSTR szErrorInfo,LPCTSTR szErrorFile,int iLine);

	void ReportErrorInfo(_com_error &e,LPCTSTR szErrorFile,int iLine);

	void ReportErrorInfo(COleDispatchException *e,LPCTSTR szErrorFile,int iLine);

	void ReportErrorInfoLV2(LPCTSTR szErrorInfo);

	void ReportErrorInfoLV2(LPCTSTR szErrorInfo,LPCTSTR szErrorFile,int iLine);

	void ReportErrorInfoLV2(_com_error &e,LPCTSTR szErrorFile,int iLine);

	void ReportErrorInfoLV2(COleDispatchException *e,LPCTSTR szErrorFile,int iLine);

	void SetAdditiveInfo(LPCTSTR szInfo);

	void PutErrorInfoToFile(LPCTSTR szErrorInfo);

	LPCTSTR GetAdditiveInfo();
}