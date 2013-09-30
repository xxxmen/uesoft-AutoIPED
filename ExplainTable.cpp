// ExplainTable.cpp: implementation of the CExplainTable class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "ExplainTable.h"
#include "EDIBgbl.h"
#include "vtot.h"
#include "Mainfrm.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CExplainTable::CExplainTable()
{

}

CExplainTable::~CExplainTable()
{

}
//功能：     生成保温说明表。
//引用表：   1. [pare_calc]
//			 2. [a_color]
//			 3. [pass]
//生成表：   [e_preexp]
bool CExplainTable::CreateBWTable()
{ 
	try
	{  
		::CoInitialize(NULL);
		_RecordsetPtr  pRs, pRsPre_calc;
    	pRs.CreateInstance(__uuidof(Recordset));
	//	pRs->CursorLocation = adUseClient;
    	pRsPre_calc.CreateInstance(__uuidof(Recordset));
	//	pRsPre_calc->CursorLocation = adUseClient;

		CString strSQL,        //SQL语句。
				str; 
		_variant_t var;
		pos = 1;
		((CMainFrame* )::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
/*  8/31  判断PASS表*/
		strSQL = "SELECT * FROM [pass] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";

		pRs->Open(_variant_t(strSQL),(IDispatch*)m_Pconnection,\
			adOpenStatic, adLockOptimistic, adCmdText);

		if( pRs->adoEOF && pRs->BOF )
		{
			AfxMessageBox("准则库中没有名为'"+EDIBgbl::SelPrjName+"'的工程， 不能生成保温说明表！");
			return false;
		}
		str = "PASS_MOD1='C_PREEXP'";
		pRs->MoveFirst();
		pRs->Find(_bstr_t(str), 0, adSearchForward);      //定位说明表的标记。
		if( pRs->adoEOF || pRs->BOF )
		{
			AfxMessageBox("数据库错误，不能生成保温说明表！");
			return false;
		}
		var = pRs->GetCollect("PASS_LAST1");
		CString	strLastMod = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""): (CString)var.bstrVal;

		str="PASS_MOD1='"+strLastMod+"'";
		pRs->MoveFirst();
		pRs->Find(_bstr_t(str), 0, adSearchForward);          //定位生成说明表的前一项操作，
	
		if( pRs->adoEOF || pRs->BOF )
		{
			AfxMessageBox("数据库错误，不能生成保温说明表！");
			return false;
		}
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		CString strPassMark;
		var = pRs->GetCollect("PASS_MARK1");
		strPassMark = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""): (CString)var.bstrVal;
		//////////////////////////////////////////////////////////////////////////
		strPassMark = _T("T");		// 不对说明表与计算的模块关联
		//////////////////////////////////////////////////////////////////////////
		
		strPassMark.MakeUpper();
		if( strPassMark != _T("T") )
		{
			strSQL = "SELECT DISTINCT LEFT(C_VOL,5) AS C_VOL FROM [pre_calc] WHERE EnginID='"+EDIBgbl::SelPrjID+"' AND (c_amount=0 OR c_amount IS NULL) ";	
			if( pRsPre_calc->State == adStateOpen )
			{
				pRsPre_calc->Close();
			}
			pRsPre_calc->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
			CString strMss;
			for( int i=1; !pRsPre_calc->adoEOF; pRsPre_calc->MoveNext(), i++ )
			{
				var = pRsPre_calc->GetCollect("C_VOL");
				if( var.vt == VT_NULL || var.vt == VT_EMPTY )
				{
					i--;
					continue;
				}
				if( i == 8 )
				{
					strMss += (CString)var.bstrVal+"\n";
					i = 0;
				}
				else
				{
					strMss += (CString)var.bstrVal+"\t";
				}
			}
			if( strMss.IsEmpty() )
			{
				AfxMessageBox("请先进行计算！");
				return false;
			}
			strMss += "\n\n以上保温数量未编辑完, 不能生成保温说明表!";
			AfxMessageBox(strMss);
			return false;
		}
		strSQL = "SELECT * FROM [pre_calc] WHERE EnginID = '"+EDIBgbl::SelPrjID+"' ORDER BY ID ";
		if( pRs->State == adStateOpen )
		{
			pRs->Close();
		}
		pRs->Open( _variant_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->RecordCount <= 0 )
		{
			AfxMessageBox("在[pre_calc]表中没有' "+EDIBgbl::SelPrjID+" '工程的记录，不能生成保温说明表！");
			return false;
		}
		//将字段C_NUM和ID设为相同的数. (两个字段的含义相同)
		int lastC_NUM, lastID;
		pRs->MoveLast();
		lastID = vtoi( pRs->GetCollect(_variant_t("ID")) );
		lastC_NUM = vtoi( pRs->GetCollect(_variant_t("c_num")) );

		if( lastID != lastC_NUM )
		{
			pRs->MoveFirst();
			while( !pRs->adoEOF )
			{
				pRs->PutCollect(_variant_t("C_NUM"), pRs->GetCollect(_variant_t("ID")) );
				pRs->MoveNext();
			}
		}
		//

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		/////将表[pre_calc]当前工程的记录复制到表[e_preexp]中
		strSQL = "DELETE  FROM [e_preexp] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);

		strSQL = "INSERT INTO [e_preexp] (ID,序号,c_vol,名称,规格,c_pi_thi,温度t";
		strSQL += ",c_wind,外保温材料,保护材料,外保温厚,cal_thi,保护厚,地点";
		strSQL += ",c_price,c_hour,c_rate,温度t0,c_pre_wei,c_pro_wei,c_pipe_wei,c_wat_wei";
		strSQL += ",c_with_wat,c_no_wat,数量,c_area,备注,c_srsb";
		strSQL += ",cal_srsb,c_lost,外表面温度,外单体积,外全体积,保护单体积";
		strSQL += ",保护全体积,c_pi_mat,c_distan,c_pg,外表面积";
		strSQL += ",c_pass,c_steam,c_p_s,内保温材料,内保温厚";
		strSQL += ",c_in_wei,内单体积,内全体积,c_ts,c_color,EnginID)";

		strSQL += "	SELECT ID,c_num AS 序号,c_vol,c_name1 AS 名称,c_size AS 规格,c_pi_thi,c_temp AS 温度t";
		strSQL += ",c_wind,c_name2 AS 外保温材料,c_name3 AS 保护材料,c_pre_thi AS 外保温厚,cal_thi,c_pro_thi AS 保护厚,c_place AS 地点";
		strSQL += ",c_price,c_hour,c_rate,c_con_temp AS 温度t0,c_pre_wei,c_pro_wei,c_pipe_wei,c_wat_wei";
		strSQL += ",c_with_wat,c_no_wat,c_amount AS 数量,c_area,c_mark AS 备注,c_srsb";
		strSQL += ",cal_srsb,c_lost,c_tb3 AS 外表面温度,c_v_pre AS 外单体积,c_tv_pre AS 外全体积,c_v_pro AS 保护单体积";
		strSQL += ",c_tv_pro AS 保护全体积,c_pi_mat,c_distan,c_pg,c_area*c_amount AS 外表面积";
		strSQL += ",c_pass,c_steam,c_p_s,c_name_in AS 内保温材料,c_in_thi AS 内保温厚";
		strSQL += ",c_in_wei,c_v_in AS 内单体积,c_tv_in AS 内全体积,c_ts,c_color,EnginID";
		
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);

		strSQL += "	FROM [pre_calc] WHERE EnginID='"+EDIBgbl::SelPrjID+"' AND C_PASS='Y' "; //2005.3.7
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		strSQL = "SELECT DISTINCT colr_code, colr_arrow,colr_ring FROM [a_color] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		if(pRs->State == adStateOpen)
		{
			pRs->Close();
		}
		pRs->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection,
			    adOpenStatic, adLockOptimistic, adCmdText);

		pRs->MoveFirst();
		int colrCode;
		CString colrRing, colrArrow,strColrCode;
		//对色环，箭头两个字段赋值，（该值从表a_color中取colr_ring,colr_arrow字段）
		while( !pRs->adoEOF )
		{
			var = pRs->GetCollect("colr_code");
			colrCode = vtoi(var);
			strColrCode.Format("%d",colrCode);			
			var = pRs->GetCollect("colr_ring");
			colrRing = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(" "): (CString)var.bstrVal;

			var = pRs->GetCollect("colr_arrow");
			colrArrow = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(" "): (CString)var.bstrVal;

			if( !colrRing.Compare("N") )
			{
				colrRing = "─";
			}
			if( !colrArrow.Compare("N") )
			{
				colrArrow = "—";
			}
			strSQL = "UPDATE [e_preexp] SET 色环='"+colrRing+"' , 箭头='"+colrArrow+"'  WHERE EnginID='"+EDIBgbl::SelPrjID+"' AND c_color="+strColrCode+" ";
			m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
			pRs->MoveNext();
			if( pos<100 )
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
			}
		}
/* 8/31 取消写pass表8/31	*/
		strSQL = "SELECT * FROM [pass] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		if ( pRs->State == adStateOpen )
		{
			pRs->Close();
		}
		pRs->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection,
				 adOpenStatic, adLockOptimistic, adCmdText);
		pRs->MoveFirst();
		////将生成保温说明表之前的操作状态置“T”
		bool bPass = true;
		while( !pRs->adoEOF )
		{
			var = pRs->GetCollect("PASS_MOD1");		
			str = (var.vt == VT_NULL||var.vt==VT_EMPTY)?_T("") : (CString)var.bstrVal;
			if( bPass )
			{
				pRs->Fields->GetItem("PASS_MARK1")->PutValue(_variant_t("T"));
			}else
			{
				pRs->Fields->GetItem("PASS_MARK1")->PutValue(_variant_t("F"));//之后的操作置“F”
			}
			if( !str.Compare("C_PREEXP") )
			{			
				bPass = false;
			}
		    pRs->MoveNext();
			if( pos< 100 )
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
			//	Sleep(20);
			}
		}
		while( pos< 100 )
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		//	Sleep(20);
		}

//		AfxMessageBox("生成保温说明成功！");
		return true;	
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
}

//生成油漆说明表。
bool CExplainTable::CreatePaintTable()
{
	try
	{
		::CoInitialize(NULL); 
		_RecordsetPtr pRsPass,     //pass表的记录集。
					  pRs;      //其他表。
		pRs.CreateInstance(__uuidof(Recordset));
	//	pRs->CursorLocation = adUseClient;
		pRsPass.CreateInstance(__uuidof(Recordset));
	//	pRsPass->CursorLocation = adUseClient;

		CString strSQL, //SQL语句。
				str;     
		_variant_t  var;
		pos = 1;
/* 8/31 PASS表   8/31 */
		strSQL = "SELECT * FROM [pass] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		if( pRsPass->State == adStateOpen )
		{
			pRsPass->Close();
		}
		pRsPass->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection,
				  adOpenStatic, adLockOptimistic, adCmdText);
		if( pRsPass->adoEOF && pRsPass->BOF )
		{
			AfxMessageBox("准则库中没有名为'"+EDIBgbl::SelPrjName+"'的工程, 不能生成油漆说明表!");
			return false;
		}
/*		// 判断油漆提资单中出现油漆量=0的记录，不能生成油漆说明表！");

		strSQL = "PASS_MOD2='C_PAIEXP'";
		//将游标定位第一条记录，然后才进行查找。
		pRsPass->MoveFirst();       
		pRsPass->Find(_bstr_t(strSQL), 0, adSearchForward);

		var = pRsPass->GetCollect("PASS_LAST2");
		str = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""): (CString)var.bstrVal;
		strSQL = "PASS_MOD2='"+str+"'";
		pRsPass->MoveFirst();
		pRsPass->Find(_bstr_t(strSQL), 0, adSearchForward);
		
		var = pRsPass->GetCollect("PASS_MARK2");
		str = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""): (CString)var.bstrVal;
		str.MakeUpper();
		if( str != "T" )   //
		{
			AfxMessageBox("油漆提资单中出现油漆量=0的记录。\n不能生成油漆说明表！");
			return false;
		}
*/

/*	
		//有受保温说明表的控制。
		strSQL = "PASS_MOD1='E_PREEXP'";
		pRsPass->MoveFirst();
		pRsPass->Find(_bstr_t(strSQL), 0, adSearchForward);

		var = pRsPass->GetCollect("PASS_MARK1");
		str = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""): (CString)var.bstrVal;
		str.MakeUpper();
	 	if( str != "T" )
		{
			AfxMessageBox("保温说明表未编辑，不能生成油漆说明表!");
			return false;
		}

*/	
		
		//////////////////////////////////////////////////////////
/*		if( pRs->State == adStateOpen )				by zsy    油漆原始数据为空的时候，可能还有保温数据需要油漆的.所以不能返回      [2005/09/08]
		{
			pRs->Close();
		}
		strSQL = "SELECT * FROM [paint] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
		pRs->Open(_bstr_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->RecordCount <= 0 )
		{
			AfxMessageBox("油漆原始数据为空,请先添加油漆原始记录然后再生成！");
//			return false;
		}
*/
		//清空T120表中的当前工程中的记录。
		EmptyT120Table();

		//调用函数moduleT120()生成t120表]
		if ( gbAutoPaint120 )
		{
			if( !moduleT120() )
			{
				return false;
			}
		}
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		////////////////
		strSQL = "DELETE FROM [e_paiexp] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

		//////////将paint表中的相应工程的内容加入到e_paiexp表中。{{  1
		strSQL = "INSERT INTO [e_paiexp] (ID, PAI_VOL, 序号";
		strSQL += ",名称, PAI_TYPE, PAI_SIZE, PAI_AMOU, 油漆面积, 备注";
		strSQL += ",底漆名, 底漆度数, 底每度用量, 底合计用量, 面漆名";
		strSQL += ",面漆度数, 面每度用量, 面合计用量, PAI_A1, PAI_A_T1, PAI_A2";
		strSQL += ",PAI_A_T2, PAI_AREA, PAI_A_C1, PAI_A_C2, PAI_CODE, PAI_C_FACE, EnginID)";
		/////////paint表中的字段	//ID AS 序号
		strSQL += " SELECT ID, PAI_VOL, ID AS 序号";
		strSQL += ",PAI_NAME AS 名称, PAI_TYPE, PAI_SIZE, PAI_AMOU, PAI_P_AREA AS 油漆面积, PAI_MARK AS 备注";
		strSQL += ",PAI_NAME1 AS 底漆名, PAI_TIMES1 AS 底漆度数, PAI_COST1 AS 底每度用量, PAI_TCOST1 AS 底合计用量, PAI_NAME2 AS 面漆名";
		strSQL += ",PAI_TIMES2 AS 面漆度数, PAI_COST2 AS 面每度用量, PAI_TCOST2 AS 面合计用量, PAI_A1, PAI_A_T1, PAI_A2";
		strSQL += ",PAI_A_T2, PAI_AREA, PAI_A_C1, PAI_A_C2, PAI_CODE, PAI_C_FACE, EnginID";

		strSQL += " FROM [paint] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";

		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText); /////  1 }}

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);

		//e_paiexp表重新接 卷册号,名称,规格 进行编号。  2005.3.12
		short NO = 1;
//		strSQL = "SELECT 序号  FROM [e_paiexp] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ORDER BY PAI_VOL,PAI_SIZE,名称 ";
		strSQL = "SELECT  序号 FROM [e_paiexp] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ORDER BY 序号";
		if(pRs->GetState() == adStateOpen )
		{
			pRs->Close();
		}
		pRs->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection,
				adOpenStatic, adLockOptimistic, adCmdText);
		if( !pRs->adoEOF && !pRs->BOF )
		{
		/*	for(pRs->MoveFirst(); !pRs->adoEOF ; pRs->MoveNext(), NO++)
			{
				pRs->PutCollect(_variant_t("序号"), _variant_t(NO));
				pRs->Update();
			}
		*/
			pRs->MoveLast();
			NO = vtoi(pRs->GetCollect(_variant_t("序号")));
			NO++;
		}
		//将t120表按名称排序，接上面的序号进行编号
		strSQL = "SELECT 序号  FROM [T120] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ORDER BY 名称 ";
		if(pRs->GetState() == adStateOpen )
		{
			pRs->Close();
		}
		pRs->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection,
				adOpenStatic, adLockOptimistic, adCmdText);
		if( !pRs->adoEOF && !pRs->BOF )
		{
			for( pRs->MoveFirst(); !pRs->adoEOF; NO++, pRs->MoveNext() )
			{
				pRs->PutCollect(_variant_t("序号"), _variant_t(NO));
				pRs->Update();
			}
		}
		if(pRs->GetState() == adStateOpen )
		{
			pRs->Close();
		}

		//

		//////////将t120表中的相应工程的内容加入到e_paiexp表中。{{  2
		CString strSQL1;
		strSQL1 = "INSERT INTO [e_paiexp] ( PAI_VOL, 序号";
		strSQL1 += ",名称, PAI_TYPE, PAI_SIZE, PAI_AMOU, 油漆面积, 备注";
		strSQL1 += ",底漆名, 底漆度数, 底每度用量, 底合计用量, 面漆名";
		strSQL1 += ",面漆度数, 面每度用量, 面合计用量, PAI_A1, PAI_A_T1, PAI_A2";
		strSQL1 += ",PAI_A_T2, PAI_AREA, PAI_A_C1, PAI_A_C2, EnginID)";
		//////////t120表中的字段。
		strSQL1 += " SELECT PAI_VOL, 序号";
		strSQL1 += ",名称, PAI_TYPE, PAI_SIZE, PAI_AMOU, 油漆面积, 备注";
		strSQL1 += ",底漆名, 底漆度数, 底每度用量, 底合计用量, 面漆名";
		strSQL1 += ",面漆度数, 面每度用量, 面合计用量, PAI_A1, PAI_A_T1, PAI_A2";
		strSQL1 += ",PAI_A_T2, PAI_AREA, PAI_A_C1, PAI_A_C2, EnginID";

		strSQL1 += " FROM [T120] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";

		m_Pconnection->Execute(_bstr_t(strSQL1), NULL, adCmdText); ////  2}}

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		strSQL = "SELECT DISTINCT colr_code, colr_face FROM [a_color] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		if(pRs->GetState() == adStateOpen )
		{
			pRs->Close();
		}
		pRs->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection,
				adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->adoEOF && pRs->BOF )
		{
			AfxMessageBox("不能生成油漆说明表，规则库被破坏，请重新选择工程！");
			return false;
		}
		int code;
		CString  strCode, colrFace;
		while( !pRs->adoEOF )
		{
			var = pRs->GetCollect("COLR_CODE");
			code = (var.vt==VT_NULL || var.vt==VT_EMPTY)?0: (int)var.dblVal;
			strCode.Format("%d",code);

			var = pRs->GetCollect("COLR_FACE");
			colrFace = (var.vt==VT_NULL || var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
			if( !colrFace.CompareNoCase("N") )
			{
				colrFace = _T(" ");
			}
			strSQL = "UPDATE [e_paiexp] SET pai_c_face='"+colrFace+"' WHERE EnginID='"+EDIBgbl::SelPrjID+"' AND pai_code="+strCode+"";
			m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
			
			pRs->MoveNext();
			if( pos < 100 )
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
			}
			else
				pos = 1;
		}
/*	8/31  PASS表8/31 */
		pRsPass->MoveFirst();
		bool bPass = true;
		CString strTab;
		while( !pRsPass->adoEOF )       //将本模块前的所有操作状态设为通过，之后的设为F.
		{
			var = pRsPass->GetCollect("PASS_MOD2");
			strTab = (var.vt==VT_NULL || var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
			if( bPass )
			{
				pRsPass->Fields->GetItem("PASS_MARK2")->PutValue(_variant_t("T"));
			}else
			{
				pRsPass->Fields->GetItem("PASS_MARK2")->PutValue(_variant_t("F"));
			}
			if( !strTab.Compare("C_PAIEXP") )
			{
				bPass=false;
			}
			pRsPass->MoveNext();
		}
		if( pRsPass->State == adStateOpen )
			pRsPass->Close();
		pRsPass.Release();

		if( pRs->State == adStateOpen )
			pRs->Close();
		while( pos < 100 )
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}
		//		AfxMessageBox("生成油漆说明表成功！");
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}

BOOL CExplainTable::EmptyT120Table() const
{
	CString strSQL;
	strSQL.Format( _T("DELETE FROM [T120] WHERE EnginID='%s' "), EDIBgbl::SelPrjID );
	m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
	return TRUE;
}

//生成T120表。
BOOL CExplainTable::moduleT120()
{
	try
	{
		_RecordsetPtr  pRs,					//临时表记录集        
			pRsT120ls,			//Table120ls表的记录集。
			pT120;				//T120表的记录集。
		pRs.CreateInstance(__uuidof(Recordset));
		pRs->CursorLocation = adUseClient;
		pRsT120ls.CreateInstance(__uuidof(Recordset));
		pRsT120ls->CursorLocation = adUseClient;
		pT120.CreateInstance(__uuidof(Recordset));
		pT120->CursorLocation = adUseClient;

		CString strSQL, strOldName, strName;
		double   areaSum=0, fArea=0;
		int      flg=1;
		_variant_t  var;
		//创建一个临时表‘Table120ls’
		if( !IsTableExists(m_Pconnection,"Table120ls") )
		{
			strSQL = "CREATE TABLE [Table120ls] (名称 char,外表面积  double, 规格 double, 数量 double, C_PI_THI double)";
			m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		}
		//清空临时表。
		strSQL = "DELETE FROM [Table120ls]";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		//
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		//将表[e_preexp]中的值存到临时表中。
		strSQL = "INSERT INTO [Table120ls] SELECT 名称,外表面积 ,规格 ,数量 ,C_PI_THI FROM [e_preexp] WHERE EnginID='"+EDIBgbl::SelPrjID+"' AND 温度t<=120 ";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

		//修改字段'外表面积'。
		strSQL = "UPDATE [Table120ls] SET 外表面积=规格*3.1415926*数量/1000 WHERE C_PI_THI<>0";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

		if( pRsT120ls->State == adStateOpen )
		{
			pRsT120ls->Close();
		}
		strSQL = "SELECT trim(名称) AS 名称, 外表面积 FROM [Table120ls] ORDER BY trim(名称)";
		pRsT120ls->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
		if( pRsT120ls->RecordCount <= 0 )
			return EmptyT120Table();

		//清空T120表中的当前工程中的记录，重新生成新的记录。
		EmptyT120Table();
		strSQL = "SELECT * FROM [T120] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		pT120->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
	
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		pRsT120ls->MoveFirst();
		while( !pRsT120ls->adoEOF )
		{
			var = pRsT120ls->GetCollect("名称");
			strName = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
			strName.TrimRight();
			var = pRsT120ls->GetCollect("外表面积");
			fArea = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
			if( flg==1 )
			{
				strOldName = strName;
				areaSum = fArea;
				flg=0;
				pRsT120ls->MoveNext();
				continue;
			}
			if( !strName.Compare(strOldName) ) //名称相同，对外表面积的值求和。
			{
				areaSum += fArea;
			}
			else                             //将新生成的记录增加到T120表中。
			{
				pT120->AddNew();
				pT120->PutCollect(_T("名称"), _variant_t(strOldName));
				pT120->PutCollect(_T("油漆面积"), _variant_t(areaSum));
				pT120->PutCollect(_T("PAI_TYPE"), _variant_t("≤120℃管道"));
				pT120->PutCollect(_T("EnginID"), _variant_t(EDIBgbl::SelPrjID));
				pT120->Update();
				strOldName = strName;
				areaSum = fArea;
			}
			pRsT120ls->MoveNext();
			if( pos < 100 )
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);			
			}
			else
			{
				pos = 1;
			}
		}
		if( !flg )      //对最后一条名称处理。
		{
			pT120->AddNew();
			pT120->PutCollect(_T("名称"), _variant_t(strOldName));
			pT120->PutCollect(_T("油漆面积"), _variant_t(areaSum));
			pT120->PutCollect(_T("PAI_TYPE"), _variant_t("≤120℃管道"));
			pT120->PutCollect(_T("EnginID"), _variant_t(EDIBgbl::SelPrjID));
			pT120->Update();
		}
		if( pRs->State == adStateOpen )
		{
			pRs->Close();
		}
		strSQL = "SELECT * FROM [a_pai] WHERE EnginID='"+EDIBgbl::SelPrjID+"' AND PAI_TYPE='≤120℃管道'";
		pRs->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->GetRecordCount() <=0 )
		{
			AfxMessageBox("规则库中没有类型为‘≤120℃管道’的记录。");
			return false;
		}
		CString str_N1, str_N2, str_A1, str_A2;
		CString strd_T1, strd_C1, strd_T2, strd_C2, strd_A_T1, strd_A_T2;
		//从表a_pai取出相应的值。
		var    = pRs->GetCollect("PAI_N1");       //底漆名        
		str_N1 = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(" "):(CString)var.bstrVal;

		var    = pRs->GetCollect("PAI_N2");			//面漆名
		str_N2 = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(" "):(CString)var.bstrVal;

		var    = pRs->GetCollect("PAI_A1");			//PAI_A1
		str_A1 = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(" "):(CString)var.bstrVal;

		var    = pRs->GetCollect("PAI_A2");			//PAI_A2
		str_A2 = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(" "):(CString)var.bstrVal;

		var  = pRs->GetCollect("PAI_T1");           //以下所有数字型转为字符串。 //底漆度数
		strd_T1.Format("%f", (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal);

		var  = pRs->GetCollect("PAI_C1");			//底每度用量
		strd_C1.Format("%f", (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal);

		var  = pRs->GetCollect("PAI_T2");			//面漆度数
		strd_T2.Format("%f", (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal);

		var  = pRs->GetCollect("PAI_C2");			//
		strd_C2.Format("%f", (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal);

		var  = pRs->GetCollect("PAI_A_T1");			//
		strd_A_T1.Format("%f", (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal);

		var  = pRs->GetCollect("PAI_A_T2");			//
		strd_A_T2.Format("%f", (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal);

		strSQL = "UPDATE [T120] SET 底漆名='"+str_N1+"', 底漆度数="+strd_T1+" ";
		strSQL += ", 底每度用量="+strd_C1+", 面漆名='"+str_N2+"', 面漆度数="+strd_T2+" ";
		strSQL += ", 面每度用量="+strd_C2+", PAI_A1='"+str_A1+"', PAI_A_T1="+strd_A_T1+" ";
		strSQL += ", PAI_A2='"+str_A2+"', PAI_A_T2="+strd_A_T2+" ";
		strSQL += " WHERE EnginID='"+EDIBgbl::SelPrjID+"' AND PAI_TYPE='≤120℃管道'";

		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		//
		strSQL = "UPDATE [T120] SET 底合计用量=底漆度数*底每度用量*油漆面积 ";
		strSQL += ", 面合计用量=面漆度数*面每度用量*油漆面积 ";
		strSQL += ", PAI_A_C1=PAI_A_T1*油漆面积, PAI_A_C2=PAI_A_T2*油漆面积 ";
		strSQL += "  WHERE EnginID='"+EDIBgbl::SelPrjID+"'";

		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}

//判断该表是否存在
BOOL CExplainTable::IsTableExists(_ConnectionPtr pCon, CString tb)
{
	if(pCon==NULL || tb=="")
		return false;
	_RecordsetPtr rs;
	if(tb.Left(1)!="[")
		tb="["+tb;
	if(tb.Right(1)!="]")
		tb+="]";
	try{
		rs=pCon->Execute(_bstr_t(tb), NULL, adCmdTable);
		rs->Close();
		return true;
	}
	catch(_com_error e)
	{
		return false;
	}

}
bool CExplainTable::EditBWTable()
{
	if( !IsPass("pass_mod1", "E_PREEXP", "pass_last1", "pass_mark1") )
	{
		AfxMessageBox("保温说明表未生成，不能编辑！");
		return false;
	}
	WritePass("pass_mod1","E_PREEXP","pass_mark1");
	return true;
}
//功能。判断PASS表中的，strMode模块是否通过。通过返回true
//strField:为不同的模块对应的字段（pass_mod1，pass_mod2...)
//strMode:当前模块。
//pass_Last: 前一个模块字段。
//pass_mark；为不同的模块对应的状态。

bool CExplainTable::IsPass(CString strField, CString strMode, CString pass_Last, CString pass_mark, bool flg, CString strCur)
{
	try{

		_RecordsetPtr pRsPass;
		pRsPass.CreateInstance(__uuidof(Recordset));
		_variant_t var;
		CString strSQL, strLastMode, strTmp;
		strSQL = "SELECT * FROM [PASS] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
		pRsPass->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic,adCmdText);
		
		if( pRsPass->adoEOF && pRsPass->BOF )
		{
			return false;	
		}
		strTmp = ""+strField+"='"+strMode+"'";
		pRsPass->Find(bstr_t(strTmp), 0, adSearchForward);
		if( pRsPass->adoEOF )
		{
			return false;
		}
		var = pRsPass->GetCollect(_variant_t(pass_Last));
		strLastMode = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
		//判断strMode模块的前一次操作。
		pRsPass->MoveFirst();
		strTmp = ""+strField+"='"+strLastMode+"'";
		pRsPass->Find(bstr_t(strTmp), 0, adSearchForward);
		if( pRsPass->adoEOF )
		{
			return false;
		}
		var = pRsPass->GetCollect(_variant_t(pass_mark));
		strLastMode = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;	
		strLastMode.TrimRight();
		strLastMode.MakeUpper();

		if( strLastMode != "T" )
		{
			return false;
		}
		//判断能否生成油漆时的特殊情况，多加了一个条件。
		if( flg && !strCur.IsEmpty() )
		{
			pRsPass->MoveFirst();
			strTmp = "pass_mod1='"+strCur+"'";
			pRsPass->Find(_bstr_t(strTmp), 0, adSearchForward);
			if( pRsPass->adoEOF )
			{
				return false;
			}
			var = pRsPass->GetCollect("pass_mark1");
			strLastMode = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
			strLastMode.TrimRight();
			strLastMode.MakeUpper();
			if( strLastMode != "T" )
			{
				return false;
			}
		}
		pRsPass->Close();
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}
//功能。写PASS表。
//pass_mode:为不同的模块对应的字段（pass_mod1，pass_mod2...)
//strMode:当前模块。
//pass_mark；为不同的模块对应的状态。
bool CExplainTable::WritePass(CString pass_mode, CString strMode, CString pass_mark)
{
	try
	{
		_RecordsetPtr pRsPass;
		pRsPass.CreateInstance(__uuidof(Recordset));
		CString strSQL, strTmp;
		_variant_t var;
		bool flg = true;
		strMode.MakeUpper();
		strSQL = "SELECT * FROM [PASS] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";

		pRsPass->Open(_variant_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
		if( pRsPass->adoEOF && pRsPass->BOF )
		{
			return false;
		}
		pRsPass->MoveFirst();
		while( !pRsPass->adoEOF )
		{
			var = pRsPass->GetCollect(_bstr_t(pass_mode));
			strTmp = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
			if( flg )
			{
				pRsPass->PutCollect(bstr_t(pass_mark), "T");
				pRsPass->Update();
				strTmp.MakeUpper();
				if( strTmp == strMode )
				{
					flg = false;
				}
			}
			else
			{
				pRsPass->PutCollect(bstr_t(pass_mark), "F");
				pRsPass->Update();
			}
			pRsPass->MoveNext();
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}

bool CExplainTable::EditPaint()
{
	if( !IsPass("pass_mod2", "E_PAIEXP", "pass_last2", "pass_mark2") )
	{
		AfxMessageBox("油漆说明表未生成,不能编辑!");
		return false;
	}
	WritePass("pass_mod2", "E_PAIEXP", "pass_mark2");
	return true;
}
