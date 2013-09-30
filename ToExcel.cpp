// ToExcel.cpp: implementation of the ToExcel class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "ToExcel.h"

#include "EDIBgbl.h"
#include "FoxBase.h"

#include "SelEngVolDll.h"
#include "FrmFolderLocation.h"

#include "excel9.h"
#include <comutil.h>
#include <afxwin.h>

#include "MainFrm.h"
#include "vtot.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//#define setmyprogress(pos) ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);

extern CAutoIPEDApp theApp;

CString c[100];
_variant_t var[100];
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

extern "C" __declspec(dllexport) bool DLFillExcelContent(char* cSFileName,char* cDFileName,
					                                   const int nTableId,const _RecordsetPtr& pRs,
					                                   const char* cDbPath, const bool bAddTable=false);
ToExcel::ToExcel()
{

}

ToExcel::~ToExcel()
{

}


//结果浏览表
bool ToExcel::Menu34()
{
	CWaitCursor wait;
    int pos1,pos;
	_RecordsetPtr	m_pRecordset;

	pos=1;
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	CString sql;

	try
	{
		theApp.m_pConnection->Execute("drop table p",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionError(e);
			throw;
		}
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

    sql.Format(_T("select * into p from pre_calc where EnginID='%s' order by c_vol,c_temp,c_size desc"),
			   EDIBgbl::SelPrjID);

	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
    pos++;


	m_pRecordset.CreateInstance(__uuidof(Recordset));
	
	sql=_T("select * from p");

	try
	{
		m_pRecordset->Open(_variant_t(sql),
							theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
							adOpenStatic,
							adLockOptimistic,
							adCmdText);

		if(m_pRecordset->adoEOF && m_pRecordset->BOF) 
		{
			AfxMessageBox("没有数据可以导出!");
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);

			return false;
		}

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		pos++;
		pos1=0;

		while(!m_pRecordset->adoEOF)
		{
			pos1++;
 			if((pos1%20)==0 && pos<30) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;
			}
			ReplNext(m_pRecordset,_variant_t("c_num"),_variant_t(RecNo(m_pRecordset)),1);
			m_pRecordset->MoveNext();
		}

		m_pRecordset->Close();

		theApp.m_pConnection->Execute("UPDATE p SET c_size=0 where c_size>=2000",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

		sql=_T("select * from p ORDER BY ID");
	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

    _RecordsetPtr	m_pRecordset1;

	m_pRecordset1.CreateInstance(__uuidof(Recordset));

	sql.Format(_T("select * from Volume where EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		m_pRecordset1->Open(_variant_t(sql),
						   theApp.m_pConAllPrj.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	_variant_t var_vol;
    CString p_vol,p_vol1,vol,p_vol_name;
	BOOL locate;
	int after;

	p_vol=_T("      ");

	try
	{
		while(!m_pRecordset->adoEOF)
		{
			pos1++;

 			if((pos1%15)==0 && pos<50) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;
			}

			var_vol=m_pRecordset->GetCollect("c_vol");
			vol=_T("");

			if(var_vol.vt!=VT_NULL && var_vol.vt != VT_EMPTY)
			{
				vol=(LPCSTR)_bstr_t(var_vol);
			}

			if(vol.Left(5)!=p_vol.Left(5))
			{
				p_vol=vol.Left(5);
				p_vol1=p_vol.Right(4);
				locate=LocateFor(m_pRecordset1,_variant_t("jcdm"),CFoxBase::EQUAL,
						  _variant_t(p_vol1));

				if(locate) 
				{
					p_vol_name=(LPCSTR)_bstr_t(m_pRecordset1->GetCollect("jcmc"));
					
					after=0;
					InsertNew(m_pRecordset,after);

					m_pRecordset->PutCollect("c_name1",_variant_t(p_vol_name));
					m_pRecordset->PutCollect("c_mark",_variant_t(p_vol));

					m_pRecordset->Update();
					m_pRecordset->MoveNext();
				}
				else
				{
					m_pRecordset->MoveNext();
				}
			}
			else
			{
				m_pRecordset->MoveNext();
			}
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}

    CString name,thi,c_name_in;
	double c_ts,wei,t,c_in_thi,c_in_wei;
	_variant_t var1,varTemp;

	try
	{
		m_pRecordset->MoveFirst();

		while(!m_pRecordset->adoEOF)
		{
			pos1++;

 			if((pos1%15)==0 && pos<75) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;
			}

			var1=m_pRecordset->GetCollect("c_name_in");
			if	( vtos(var1).IsEmpty() )
			{
				m_pRecordset->MoveNext();
			}
			else
			{
				//////////////////////////////////////////////////////////
				// 看以下字段是否为空~~出错~~
				//////////////////////////////////////////////////////
//				name=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_name2"));
//				thi=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_pre_thi"));
//				wei=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_pre_wei"));
//				t=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_tb3"));

				GetTbValue(m_pRecordset,_variant_t("c_name2"),name);
				GetTbValue(m_pRecordset,_variant_t("c_pre_thi"),thi);
				GetTbValue(m_pRecordset,_variant_t("c_pre_wei"),wei);
				GetTbValue(m_pRecordset,_variant_t("c_tb3"),t);

				after=1;

				InsertNew(m_pRecordset,after);

				m_pRecordset->PutCollect("c_name2",_variant_t(name));
				m_pRecordset->PutCollect("c_pre_thi",_variant_t(thi));
				m_pRecordset->PutCollect("c_pre_wei",_variant_t(wei));
				m_pRecordset->PutCollect("c_tb3",_variant_t(t));

				m_pRecordset->Update();

				m_pRecordset->MovePrevious();

//				c_name_in=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_name_in"));
//				c_in_thi=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_in_thi"));
//				c_in_wei=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_in_wei"));

				GetTbValue(m_pRecordset,_variant_t("c_name_in"),c_name_in);
				GetTbValue(m_pRecordset,_variant_t("c_in_thi"),c_in_thi);
				GetTbValue(m_pRecordset,_variant_t("c_in_wei"),c_in_wei);

//				c_ts=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_ts"));

				GetTbValue(m_pRecordset,_variant_t("c_ts"),c_ts);				

				m_pRecordset->PutCollect("c_name2",_variant_t(c_name_in));
				m_pRecordset->PutCollect("c_pre_thi",_variant_t(c_in_thi));
				m_pRecordset->PutCollect("c_pre_wei",_variant_t(c_in_wei));
				m_pRecordset->PutCollect("c_tb3",_variant_t(c_ts));

				m_pRecordset->Update();

				m_pRecordset->MoveNext();
			}
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}

	CString name3;
	_variant_t vname3;
	int a;

	try
	{
		m_pRecordset->MoveFirst();

		while(!m_pRecordset->adoEOF)
		{
			pos1++;

 			if((pos1%20)==0 && pos<95) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;
			}

			//name3=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("c_name3"));
			vname3=m_pRecordset->GetCollect("c_name3");

			if(vname3.vt!=VT_NULL && vname3.vt != VT_EMPTY)
			{   
				name3=(LPCSTR)_bstr_t(vname3);
				a=name3.Find(_T("("),0);
				name3=name3.Left(a);
				m_pRecordset->PutCollect("c_name3",_variant_t(name3));
				m_pRecordset->Update();
			}

			m_pRecordset->MoveNext();
		}
			
		m_pRecordset->MoveFirst();
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}
//	bool bool2;
//	CopyFile(LPCTSTR("c:\\bw\\bwx1.exe"),LPCTSTR("d:\\bwx1.exe"),bool2);

    while(pos<98)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		Sleep(20);
		pos++;
	}

	CString strPathSrc;
	CString strPathDest;

	strPathSrc.Format(_T("%s%s"),gsTemplateDir,"insu.xls");
	strPathDest.Format(_T("%s%s"),gsProjectDBDir,"insu.xls");

	if( !::FileExists(strPathDest) )
	{
		CopyFile(strPathSrc,strPathDest,FALSE);//强制覆盖文件
	}

	int r;
	r=4;
	int fieldnum;
	fieldnum=19;

	c[0]="A";
	c[1]="B";
	c[2]="C";
	c[3]="D";
	c[4]="E";
	c[5]="F";
	c[6]="G";
	c[7]="H";
	c[8]="I";
	c[9]="J";
	c[10]="K";
	c[11]="L";
	c[12]="M";
	c[13]="N";
	c[14]="O";
	c[15]="P";
	c[16]="Q";
	c[17]="R";
	c[18]="S";
	var[0]=long(0);
	var[1]=long(3);
	var[2]=long(6);
	var[3]=long(4);
	var[4]=long(5);
	var[5]=long(13);
	var[6]=long(8);
	var[7]=long(10);
	var[8]=long(18);
	var[9]=long(9);
	var[10]=long(12);
	var[11]=long(19);
	var[12]=long(20);
	var[13]=long(21);
	var[14]=long(22);
	var[15]=long(23);
	var[16]=long(36);
	var[17]=long(30);
	var[18]=long(26);

//	CFile file;
//	file.Open("c:\\insu.xls",CFile::modeReadWrite,NULL);
//		CString sheetname;
//		sheetname="bwjg";

	try
	{
		CString strTmpPath = gsTemplateDir;
		CString strPrj = gsProjectDBDir;
		CString strTmpName = gsShareMdbDir +"TableFormat.mdb";

		if( !DLFillExcelContent( strTmpPath.GetBuffer( 255 ), strPrj.GetBuffer( 255 ),
			     37, m_pRecordset, strTmpName.GetBuffer( 255 )) )
		{
     		return false;
		}

	//	OutoExcel(m_pRecordset,strPathDest,"bwjg",r,fieldnum);
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
	}
	catch(_com_error &e)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
		ReportExceptionError(e);
		throw;
	}

//	AfxMessageBox(_T("数据已导出！"));
	return true;

}

void ToExcel::OutoExcel(_RecordsetPtr m_pRecordset, CString strPathDest, CString Sheetname, int startrow, int fieldnum)
{

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	BOOL bRet;
	LPDISPATCH pDispatch;
	Workbooks workbooks;
	_Workbook workbook;
	Worksheets worksheets;
	_Worksheet worksheet;
	Range range,range1, rangeDelete;
	_variant_t null;
	_Application Application;
	int pos = 1, pos1=1;

	null.vt=VT_ERROR;
	null.scode=DISP_E_PARAMNOTFOUND;

	bRet=Application.CreateDispatch(_T("Excel.Application"));

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(1);

	if(bRet)
	{

//		Application.SetVisible(true);
	
		pDispatch=Application.GetWorkbooks();
	   	workbooks.AttachDispatch(pDispatch);

//		pDispatch=workbooks.Add(null);
//		workbook.AttachDispatch(pDispatch);
		pDispatch = workbooks.Open(strPathDest,    
		  covOptional, covOptional, covOptional, covOptional, covOptional,
		  covOptional, covOptional, covOptional, covOptional, covOptional,
		  covOptional, covOptional );  

		ASSERT(pDispatch);  

            workbook.AttachDispatch(pDispatch);
		pDispatch=workbook.GetWorksheets();
		worksheets.AttachDispatch(pDispatch);
        
		pDispatch=worksheets.GetItem(_variant_t(Sheetname));

		worksheet.AttachDispatch(pDispatch);
	      worksheet.Activate();
		try
		{
			//首先清空原来的数据。
			CString strStartrow;
			strStartrow.Format("A%d",startrow);
			range1 = worksheet.GetRange(_variant_t(strStartrow),_variant_t("Z1000"));
			range1.Clear();
			//
			m_pRecordset->MoveFirst();
			while(!m_pRecordset->adoEOF)
			{

				for(int k=0;k<fieldnum;k++)
				{
					_variant_t var1;
					var1=m_pRecordset->GetCollect(var[k]);
					//var1.vt=VT_NULL;
					CString cell;
					cell.Format(_T("%s%d"),c[k],startrow);
					pDispatch=worksheet.GetRange(_variant_t(cell),null);

					range.AttachDispatch(pDispatch);
					range.SetValue(var1);
					range1 = range.GetEntireColumn();
					range1.AutoFit();			
				}
				if( pos1++ ==2 && pos<60 )
				{
					pos1=1;
					((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
				}
				if( pos >=60 && pos1++ >= 10)
				{
					pos1=1;
					((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
				}
				startrow++;
				m_pRecordset->MoveNext();			
			}
		}
		catch(_com_error &e)
		{
			ReportExceptionError(e);
			throw;
		}
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(100);
		Application.SetVisible(true);
		Application.SetShowWindowsInTaskbar(true);
		Application.SetUserControl(true);
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
	}
	else
	{
		AfxMessageBox(_T("出错！没有安装excel，或excel已损坏！"));
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
	}
}

//保温说明表
bool ToExcel::Menu43()
{
	CWaitCursor wait;
	int pos1,pos=1;
	CString sql, strMar;
	_variant_t var_Mar;

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
	try
	{
		theApp.m_pConnection->Execute(_T("drop table temp_print"),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionError(e);
			return false;
		}
	}

      ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	sql.Format(_T("select * into temp_print from e_preexp where EnginID='%s' order by c_vol asc, 温度t desc,规格 desc"),
			   EDIBgbl::SelPrjID);

	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
 	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
      pos++;

	_RecordsetPtr	m_pRecordset;
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	
	sql=_T("select * from temp_print");
	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);

		if(m_pRecordset->adoEOF && m_pRecordset->BOF)
		{
			AfxMessageBox(_T("没有数据可以导出!"));

			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);

			return false;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
      pos++;
      pos1=0;
	try
	{
		m_pRecordset->MoveFirst();
		for(int n, i=1; !m_pRecordset->adoEOF; i++, pos1++)
		{
 			if((pos1%20)==0 && pos<25) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
			}
			var_Mar = m_pRecordset->GetCollect("保护材料");
			strMar = (var_Mar.vt==VT_NULL || var_Mar.vt==VT_EMPTY)?_T(""):(CString)var_Mar.bstrVal;
			n = strMar.Find("(",0);
			strMar = strMar.Left(n);

			m_pRecordset->PutCollect("保护材料",_variant_t(strMar));
			m_pRecordset->PutCollect(_variant_t("序号"), (short)i);
			m_pRecordset->Update();
		//	ReplNext(m_pRecordset,_variant_t("序号"),_variant_t(RecNo(m_pRecordset)),1);
			m_pRecordset->MoveNext();
		}

		m_pRecordset->Close();
		theApp.m_pConnection->Execute("UPDATE temp_print SET 规格=0 where 规格>=2000",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
      ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

      _RecordsetPtr	m_pRecordset1;
	m_pRecordset1.CreateInstance(__uuidof(Recordset));

	sql.Format(_T("select * from Volume where EnginID='%s'"),EDIBgbl::SelPrjID);
	try
	{
		m_pRecordset1->Open(_variant_t(sql),
						   theApp.m_pConAllPrj.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	_variant_t var_vol;
      CString p_vol,p_vol1,vol,p_vol_name;
	p_vol=_T("      ");

	try
	{
		while(!m_pRecordset->adoEOF)
		{
			pos1++;
 			if((pos1%15)==0 && pos<65) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
			}
			var_vol=m_pRecordset->GetCollect("c_vol");
			vol = (var_vol.vt == VT_NULL || var_vol.vt == VT_EMPTY)?_T(""):(CString)var_vol.bstrVal;

			if(vol.Left(5) != p_vol.Left(5))
			{
				p_vol=vol.Left(5);
				p_vol1=p_vol.Right(4);
				if( m_pRecordset1->adoEOF && m_pRecordset1->BOF)
				{
					break;
				}
				m_pRecordset1->MoveFirst();
				m_pRecordset1->Find(_bstr_t("jcdm='"+p_vol1+"'"), 0, adSearchForward);
			//	locate=LocateFor(m_pRecordset1,_variant_t("jcdm"),CFoxBase::EQUAL,
			//			  _variant_t(p_vol1));
				if(!m_pRecordset1->adoEOF) 
				{
//					p_vol_name=(LPCSTR)_bstr_t(m_pRecordset1->GetCollect("jcmc"));
					GetTbValue(m_pRecordset1,_variant_t("jcmc"),p_vol_name);

					if( !InsertNew(m_pRecordset,0) )
					{
						if( !m_pRecordset->adoEOF )
						{ 
							m_pRecordset->MoveNext();
							continue;
						}
						break;
					}
					m_pRecordset->PutCollect("名称",_variant_t(p_vol_name));
					m_pRecordset->PutCollect("备注",_variant_t(p_vol));
					m_pRecordset->Update();
				}
			}
			m_pRecordset->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

      CString name,thi,v,vt,t,c_name_in,c_in_thi,c_in_wei,c_ts,strBw;
	_variant_t var1;
	try
	{
		m_pRecordset->MoveFirst();
		while(!m_pRecordset->adoEOF)
		{
			pos1++;
 			if((pos1%15)==0 && pos<95) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
			}
			var1=m_pRecordset->GetCollect("内保温材料");
			strBw = (var1.vt==VT_NULL||var1.vt==VT_EMPTY)?"":(CString)var1.bstrVal;
			strBw.TrimLeft();
			strBw.TrimRight();
			if(strBw == "")
			{
				m_pRecordset->MoveNext();
			}
			else
			{
				var1 = m_pRecordset->GetCollect("外保温材料");
				name = (var1.vt==VT_NULL || var1.vt==VT_EMPTY)?_T(""):(CString)var1.bstrVal;
				var1 = m_pRecordset->GetCollect("外保温厚");
				thi.Format("%f",(var1.vt==VT_NULL || var1.vt==VT_EMPTY)?0.0: var1.dblVal);
				var1 = m_pRecordset->GetCollect("外单体积");
				v.Format("%f",(var1.vt==VT_NULL || var1.vt==VT_EMPTY)?0.0: var1.dblVal);
				var1 = m_pRecordset->GetCollect("外全体积");
				vt.Format("%f",(var1.vt==VT_NULL || var1.vt==VT_EMPTY)?0.0: var1.dblVal);
				var1 = m_pRecordset->GetCollect("外表面温度");
				t.Format("%f",(var1.vt==VT_NULL || var1.vt==VT_EMPTY)?0.0: var1.dblVal);

			//	name=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("外保温材料"));
			//	thi=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("外保温厚"));
			//	v=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("外单体积"));
			//	vt=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("外全体积"));
			//	t=(LPCSTR)_bstr_t(m_pRecordset->GetCollect("外表面温度"));

				InsertNew(m_pRecordset, 1);

				m_pRecordset->PutCollect("外保温材料",_variant_t(name));
				m_pRecordset->PutCollect("外保温厚",_variant_t(thi));
				m_pRecordset->PutCollect("外单体积",_variant_t(v));
				m_pRecordset->PutCollect("外全体积",_variant_t(vt));
				m_pRecordset->PutCollect("外表面温度",_variant_t(t));

				m_pRecordset->Update();
				m_pRecordset->MovePrevious();

				m_pRecordset->PutCollect("外保温材料",m_pRecordset->GetCollect("内保温材料"));
				m_pRecordset->PutCollect("外保温厚",m_pRecordset->GetCollect("内保温厚"));
				m_pRecordset->PutCollect("外单体积",m_pRecordset->GetCollect("内单体积"));
				m_pRecordset->PutCollect("外全体积",m_pRecordset->GetCollect("内全体积"));
				m_pRecordset->PutCollect("外表面温度",m_pRecordset->GetCollect("c_ts"));

				m_pRecordset->Update();
				m_pRecordset->MoveNext();
			}
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
/*
go top
do while .not.eof()
	repl 保护材料 with SUBS(保护材料,1,AT("(",保护材料)-1)
	skip 1
enddo
//*/
/*	CString name3;
	_variant_t vname3;
	int a;

	try
	{
		m_pRecordset->MoveFirst();

		while(!m_pRecordset->adoEOF)
		{
			
			pos1++;

 			if((pos1%20)==0 && pos<95) 
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;

			}vname3=m_pRecordset->GetCollect("保护材料");

			if(vname3.vt!=VT_NULL)
			{   
				name3=(LPCSTR)_bstr_t(vname3);
				a=name3.Find("(",0);
				name3=name3.Left(a);
				m_pRecordset->PutCollect("保护材料",_variant_t(name3));
				m_pRecordset->Update();
			}

			m_pRecordset->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		throw;
	}
		
    while(pos<98)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		Sleep(20);
		pos++;
	}

*/
	CString strPathSrc;
	CString strPathDest;

	strPathSrc.Format(_T("%s%s"),gsTemplateDir,"insu.xls");
	strPathDest.Format(_T("%s%s"),gsProjectDBDir,"insu.xls");
	if( !::FileExists(strPathDest) )
	{
		CopyFile(strPathSrc,strPathDest,FALSE);//强制覆盖文件

	}

	int r;
	r=4;
	int fieldnum;
	fieldnum=19;

	c[0]="A";
	c[1]="B";
	c[2]="C";
	c[3]="D";
	c[4]="E";
	c[5]="F";
	c[6]="G";
	c[7]="H";
	c[8]="I";
	c[9]="J";
	c[10]="K";
	c[11]="L";
	c[12]="M";
	c[13]="N";
	c[14]="O";
	c[15]="P";
	c[16]="Q";
	c[17]="R";
	c[18]="S";
	var[0]=long(1);
	var[1]=long(3);
	var[2]=long(24);
	var[3]=long(4);
	var[4]=long(13);
	var[5]=long(6);
	var[6]=long(17);
	var[7]=long(8);
	var[8]=long(10);
	var[9]=long(31);
	var[10]=long(32);
	var[11]=long(9);
	var[12]=long(12);
	var[13]=long(25);
	var[14]=long(38);
	var[15]=long(30);
	var[16]=long(49);
	var[17]=long(50);
	var[18]=long(26);
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(100);
	try
	{
		CString strTmpPath = gsTemplateDir;
		CString strPrj = gsProjectDBDir;
		CString strTmpName = gsShareMdbDir +"TableFormat.mdb";

		if( !DLFillExcelContent( strTmpPath.GetBuffer( 255 ), strPrj.GetBuffer( 255 ),
			     38, m_pRecordset, strTmpName.GetBuffer( 255 )) )
		{
      		return false;
		}

	//	OutoExcel(m_pRecordset,strPathDest,_T("bwsm"),r,fieldnum);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

//	AfxMessageBox(_T("数据已导出！"));
	return true;

}
//输出油漆说明表
bool ToExcel::Menu46()
{
      CWaitCursor wait;
	int pos=1;
	CString sql;
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
	try
	{
		theApp.m_pConnection->Execute("drop table temp_print",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionError(e);
			return false;
		}
	}

      ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
	sql.Format(_T("select * into temp_print from e_paiexp where EnginID='%s'"),EDIBgbl::SelPrjID);
	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
		theApp.m_pConnection->Execute("UPDATE temp_print SET pai_amou=0 where pai_size=0",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
      ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);

	_RecordsetPtr	m_pRecordset;
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	
	sql=_T("select * from temp_print ORDER BY 序号");
	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);

		if(m_pRecordset->adoEOF && m_pRecordset->BOF)
		{
			AfxMessageBox(_T("没有数据可以导出!"));
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);

			return false;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

     ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
	CString strPathSrc;
	CString strPathDest;

	strPathSrc.Format(_T("%s%s"),gsTemplateDir,"paint.xls");
	strPathDest.Format(_T("%s%s"),gsProjectDBDir,"paint.xls");
	if( !::FileExists(strPathDest) )
	{
		CopyFile(strPathSrc,strPathDest,FALSE);//强制覆盖文件 
	}
    ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);

	int r=3, fieldnum=15;

	c[0]="A";
	c[1]="B";
	c[2]="C";
	c[3]="D";
	c[4]="E";
	c[5]="F";
	c[6]="G";
	c[7]="H";
	c[8]="I";
	c[9]="J";
	c[10]="K";
	c[11]="L";
	c[12]="M";
	c[13]="N";
	c[14]="O";

	var[0]=long(2);
	var[1]=long(3);
	var[2]=long(5);
	var[3]=long(6);
	var[4]=long(7);
	var[5]=long(9);
	var[6]=long(10);
	var[7]=long(11);
	var[8]=long(12);
	var[9]=long(13);
	var[10]=long(14);
	var[11]=long(15);
	var[12]=long(16);
	var[13]=long(25);
	var[14]=long(8);

    while(pos<98)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		Sleep(10);
		pos++;
	}

      try
	{
		CString strTmpPath = gsTemplateDir;
		CString strPrj = gsProjectDBDir;
		CString strTmpName = gsShareMdbDir +"TableFormat.mdb";

		if( !DLFillExcelContent( strTmpPath.GetBuffer( 255 ), strPrj.GetBuffer( 255 ),
			     40, m_pRecordset, strTmpName.GetBuffer( 255 )) )
		{
    		return false;
		}
	//	OutoExcel(m_pRecordset,strPathDest,_T("yqsm"),r,fieldnum);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
//	AfxMessageBox(_T("数据已导出！"));
	return true;
}

//保温材料汇总表
bool ToExcel::Menu54()
{
/*
IF FILE("p.dbf")
    ERASE p.dbf
ENDIF
COPY FILE e_precol.dbf TO p.dbf
SELE 1
**USE e_precol
USE p
APPE FROM e_preast FOR .NOT. DELETED()
REPL ALL col_num WITH RECNO()
//*/
	CWaitCursor wait;
	short pos,pos1;
	CString sql;

	pos=1;
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	try
	{
		theApp.m_pConnection->Execute("drop table p",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionError(e);
			return false;
		}
	}

    ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	sql.Format(_T("select * into p from e_precol where EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

    sql.Format(_T("insert into p select * from e_preast where EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	_RecordsetPtr	m_pRecordset;

	m_pRecordset.CreateInstance(__uuidof(Recordset));
	
	sql=_T("select * from p");

	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);

		if(m_pRecordset->adoEOF && m_pRecordset->BOF)
		{
			AfxMessageBox(_T("没有数据可以导出!"));
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
			return false;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
	
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);

	pos++;
    pos1=0;

	try
	{
		if(m_pRecordset->GetRecordCount()>0)//9/17
		{
			m_pRecordset->MoveFirst();//9/17
		}
		while((m_pRecordset->GetRecordCount()>0)&&(!m_pRecordset->adoEOF))
		{
			pos1++;
			if(pos%20==0 && pos<90)
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;
			}

		//	ReplNext(m_pRecordset,_variant_t("col_num"),_variant_t(RecNo(m_pRecordset)),1);//9/17原代码
			//9/17
			m_pRecordset->PutCollect(_T("col_num"),_variant_t(pos1));
			m_pRecordset->Update();
			//9/17
			m_pRecordset->MoveNext();
		}

		if(m_pRecordset->GetRecordCount()>0)//9/17
		{
		m_pRecordset->MoveFirst();
	
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	CString strPathSrc;
	CString strPathDest;

	strPathSrc.Format(_T("%s%s"),gsTemplateDir,"insu.xls");
	strPathDest.Format(_T("%s%s"),gsProjectDBDir,"insu.xls");
	if( !::FileExists(strPathDest) )
	{
		CopyFile(strPathSrc,strPathDest,FALSE);//强制覆盖文件
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	int r;
	r=3;
	int fieldnum;
	fieldnum=6;

	c[0]="A";
	c[1]="B";
	c[2]="C";
	c[3]="D";
	c[4]="E";
	c[5]="F";

	var[0]=long(7);
	var[1]=long(1);
	var[2]=long(3);
	var[3]=long(4);
	var[4]=long(6);
	var[5]=long(5);

    while(pos<98)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		Sleep(20);
		pos++;
	}


	try
	{

		CString strTmpPath = gsTemplateDir;
		CString strPrj = gsProjectDBDir;
		CString strTmpName = gsShareMdbDir +"TableFormat.mdb";

		if( !DLFillExcelContent( strTmpPath.GetBuffer( 255 ), strPrj.GetBuffer( 255 ),
			     39, m_pRecordset, strTmpName.GetBuffer( 255 )) )
		{
			return false;
		}
	//	OutoExcel(m_pRecordset,strPathDest,"clhz",r,fieldnum);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

//	AfxMessageBox(_T("数据已导出！"));
	return true;

}

//油漆材料汇总表
bool ToExcel::Menu58()
{
    CWaitCursor wait;
	short pos,pos1;
	CString sql;

	pos=1;
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	try
	{
		theApp.m_pConnection->Execute("drop table p",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(	e.Error() != DB_E_NOTABLE )
		{
			ReportExceptionError(e);
			return false;
		}
	}

    ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	sql.Format(_T("select * into p from e_paicol where EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;
    
	sql.Format(_T("insert into p select * from e_paiast where EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false ;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;
	
	_RecordsetPtr	m_pRecordset;

	m_pRecordset.CreateInstance(__uuidof(Recordset));
	
	sql=_T("select * from p");

	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);

		if(m_pRecordset->adoEOF && m_pRecordset->BOF)
		{
			AfxMessageBox(_T("没有数据可以导出!"));
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
			return false;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;
    pos1=0;

	try
	{
		if(m_pRecordset->GetRecordCount()>0)//9/17
		{
			m_pRecordset->MoveFirst();//9/17
		}
		while(!m_pRecordset->adoEOF)
		{
			pos1++;
			if(pos1%20==0 && pos<90)
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;
			}

		//	ReplNext(m_pRecordset,_variant_t("col_num"),_variant_t(RecNo(m_pRecordset)),1);//9/17原代码
			m_pRecordset->PutCollect(_T("col_num"),_variant_t(pos1));//9/17
			m_pRecordset->Update();//9/17
			m_pRecordset->MoveNext();
		}
		if(m_pRecordset->GetRecordCount()>0)//9/17
		{
		m_pRecordset->MoveFirst();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	CString strPathSrc;
	CString strPathDest;

	strPathSrc.Format(_T("%s%s"),gsTemplateDir,"paint.xls");
	strPathDest.Format(_T("%s%s"),gsProjectDBDir,"paint.xls");

	if( !::FileExists(strPathDest) )
	{
		CopyFile(strPathSrc,strPathDest,FALSE);//强制覆盖文件
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	int r;
	r=2;
	int fieldnum;
	fieldnum=6;

	c[0]="A";
	c[1]="B";
	c[2]="C";
	c[3]="D";
	c[4]="E";
	c[5]="F";

	var[0]=long(6);
	var[1]=long(1);
	var[2]=long(3);
	var[3]=long(4);
	var[4]=long(2);
	var[5]=long(5);

    while(pos<98)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		Sleep(20);
		pos++;
	}


	try
	{
		CString strTmpPath = gsTemplateDir;
		CString strPrj = gsProjectDBDir;
		CString strTmpName = gsShareMdbDir +"TableFormat.mdb";

		if( !DLFillExcelContent( strTmpPath.GetBuffer( 255 ), strPrj.GetBuffer( 255 ),
			     41, m_pRecordset, strTmpName.GetBuffer( 255 )) )
		{
			return false;
		}
	//	OutoExcel(m_pRecordset,strPathDest,_T("yqhz"),r,fieldnum);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

//	AfxMessageBox(_T("数据已导出！"));
	return true;
}

//保温材料规格表
bool ToExcel::Menu511()
{
    CWaitCursor wait;
	short pos1,pos;
	CString sql;

	pos=1;
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	try
	{
		theApp.m_pConnection->Execute("drop table p",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		if(e.Error()!=DB_E_NOTABLE)
		{
			ReportExceptionError(e);
			return false;
		}
	}

    ((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	sql.Format(_T("select * into p from e_presiz where EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		theApp.m_pConnection->Execute(_bstr_t(sql),NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;
 
	_RecordsetPtr	m_pRecordset;
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	
	sql=_T("select * from p");

	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);

		if(m_pRecordset->adoEOF && m_pRecordset->BOF)
		{
			AfxMessageBox(_T("没有数据可以导出!"));
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
			return false;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;
    pos1=0;
	

	try
	{
		if((!m_pRecordset->BOF)&&(!m_pRecordset->adoEOF))//9/17
		{
			m_pRecordset->MoveFirst();//9/17
		}
		while((!m_pRecordset->BOF)&&(!m_pRecordset->adoEOF))
		{
			pos1++;

			if(pos1%20==0 && pos<90)
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
				pos++;
			}

		//	ReplNext(m_pRecordset,_variant_t("size_num"),_variant_t(RecNo(m_pRecordset)),1);//9/17原代码
			//9/17
			m_pRecordset->PutCollect(_T("size_num"),_variant_t(pos1));
			m_pRecordset->Update();
			//9/17
			m_pRecordset->MoveNext();
		}

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		pos++;

		m_pRecordset->Close();
		theApp.m_pConnection->Execute("UPDATE p SET 规格=0 where 规格>=2000",NULL,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	try
	{
		m_pRecordset->Open(_variant_t(sql),
						   theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenStatic,
								adLockOptimistic,
								adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}

	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	CString strPathSrc;
	CString strPathDest;

	strPathSrc.Format(_T("%s%s"),gsTemplateDir,"insu.xls");
	strPathDest.Format(_T("%s%s"),gsProjectDBDir,"insu.xls");

	if( !::FileExists(strPathDest) )
	{
		CopyFile(strPathSrc,strPathDest,FALSE);//强制覆盖文件
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
	pos++;

	int r;
	int fieldnum;

	r=4;
//	fieldnum=7;pfg2005/1/6
	fieldnum=8;//pfg2005/1/6

	c[0]="A";
	c[1]="B";
	c[2]="C";
	c[3]="D";
	c[4]="E";
	c[5]="F";
	c[6]="G";
	c[7]="H";//pfg2005/1/6

	var[0]=long(1);
	var[1]=long(3);
	var[2]=long(2);
	var[3]=long(4);
	var[4]=long(9);
	var[5]=long(11);
	var[6]=long(13);
	var[7]=long(8);//pfg2005/1/6

    while(pos<98)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos);
		Sleep(20);
		pos++;
	}

	try
	{
		CString strTmpPath = gsTemplateDir;
		CString strPrj = gsProjectDBDir;
		CString strTmpName = gsShareMdbDir +"TableFormat.mdb";
 
		if( !DLFillExcelContent( strTmpPath.GetBuffer( 255 ), strPrj.GetBuffer( 255 ),
			     36, m_pRecordset, strTmpName.GetBuffer( 255 )) )
		{
			return false;
		}
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
	//	OutoExcel(m_pRecordset,strPathDest,_T("gghz"),r,fieldnum);
	}
	catch(_com_error &e)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
		ReportExceptionError(e);
		return false;
	}

//	AfxMessageBox(_T("数据已导出！"));
	return true;
}
