// FillExcelDll.h : main header file for the FILLEXCELDLL DLL
//

#if !defined(AFX_FILLEXCELDLL_H__BE0DAF8C_EF88_4FED_8040_48C0E9A215B1__INCLUDED_)
#define AFX_FILLEXCELDLL_H__BE0DAF8C_EF88_4FED_8040_48C0E9A215B1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
#include "excel9.h"
#include "mobject.h"
/////////////////////////////////////////////////////////////////////////////
// CFillExcelDllApp
// See FillExcelDll.cpp for the implementation of this class
//

class CFillExcelDllApp : public CWinApp
{
public:
	CFillExcelDllApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFillExcelDllApp)
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CFillExcelDllApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

bool DLFileExists(LPCTSTR filename);

int  DLvtoi(COleVariant & v);
int  DLvtoi(_variant_t & v);

CString DLvtos(COleVariant& v);
CString DLvtos(_variant_t& v);

BOOL DLvtob( _variant_t &v);
BOOL DLvtob(COleVariant& v);

extern "C" __declspec(dllexport) bool DLFillExcelContent(CString& strSFileName,CString &strDFileName,
					                                   const int nTableId,const _RecordsetPtr& pRs,
					                                   const CString &strDbPath, const bool bAddTable=false);
bool DLGetInfo(int nTableId,CMapStringToString &strFormatInfo,CStringArray &strSortInfo,
			                                  CStringArray &saWhere,CString &strWorksheetName,
			                                  CString &strTableCaption,int &nRowsPerPage,int &nColsPerPage,bool &bDownToUp,
			                                  const CString &strDbPath,CString &strExcelTemplateName);
bool DLConnectSortMdb(_ConnectionPtr &pConnection,const CString &strDataPath);

bool DLGetTableInfo(_ConnectionPtr &pConnection,int TableId,CString &strTableName,
				  CString &strWorksheetName,CString &strTableCaption,int &nRowsPerPage,
				  int &nColsPerPage,bool &bDownToUp,CString &strExcelTemplateName);

//获得与设置Excel表格相关的信息:包括精度设置，在Excel表格中的Field
//Field在Excel表格中的位置
bool DLGetFieldFormatInfo(_ConnectionPtr &pConnection,CMapStringToString &strFormatInfo,
						const CString &strTableName,CStringArray &strSortInfo,
						CStringArray &saWhere);

bool DLStartExcelAndLoadTemplate(const CString &strSFileName,const CString &strDFileName,
							     Worksheets &objWorksheets, Worksheets& templateSheets);


bool DLCopyTemplate(const int nStartRow,const int nEndRow,const int nColsPerPage,
				    const int nRowsPerPage,const int nPageNum, const bool bDownToUp,
					_Worksheet &workTable,  _Worksheet& templateTable, const bool bAddTable);

void DLGetCells(CString &strStartCell,CString &strEndCell,const int nStartCol,
			                                   const int nCurRow,const int nIndex,const CStringArray &saWhere);

void DLProcessItem(CString &strItem,const CMapStringToString &strFormatInfo,
				                                  const CString &strFieldName,float &fSum);

void DLSaveAndGetControl(Worksheets objWorksheets,const CString &strDefaultPath);

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FILLEXCELDLL_H__BE0DAF8C_EF88_4FED_8040_48C0E9A215B1__INCLUDED_)
