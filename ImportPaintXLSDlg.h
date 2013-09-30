// ImportPaintXLSDlg.h: interface for the CImportPaintXLSDlg class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_IMPORTPAINTXLSDLG_H__B5FEAE8A_AFA4_4B5A_9410_6D0C8540CB77__INCLUDED_)
#define AFX_IMPORTPAINTXLSDLG_H__B5FEAE8A_AFA4_4B5A_9410_6D0C8540CB77__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ImportFromXLSDlg.h"

class CImportPaintXLSDlg : public CImportFromXLSDlg  
{
public:
	CImportPaintXLSDlg();
	virtual ~CImportPaintXLSDlg();

protected:
	virtual BOOL ValidateData();
	virtual void BeginImport();
	virtual BOOL OnInitDialog();
	virtual BOOL InitPropertyWnd();
};

#endif // !defined(AFX_IMPORTPAINTXLSDLG_H__B5FEAE8A_AFA4_4B5A_9410_6D0C8540CB77__INCLUDED_)
