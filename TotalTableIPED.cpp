// TotalTableIPED.cpp: implementation of the CTotalTableIPED class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "TotalTableIPED.h"
#include "edibgbl.h"
#include "Mainfrm.h"
#include "vtot.h"
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
#define	 pi			3.1415927

CTotalTableIPED::CTotalTableIPED()
{

}

CTotalTableIPED::~CTotalTableIPED()
{

}

//************************************
//* 模 块 名: c_paicol()	         *
//* 功    能: 生成油漆材料汇总表     *
//* 上级模块: MAIN_MENU              *
//* 下级模块:  modC_RING()           *
//* 引 用 库: 1区paint_c,            *
//*           2区e_paicol,           *
//* 修改日期:                        *
//************************************
BOOL CTotalTableIPED::c_paicol()
{
	_RecordsetPtr paicol_set;
	paicol_set.CreateInstance(__uuidof(Recordset));
	_variant_t key;
	CString str;
	int pos=0;

	if(paicol_set->State==adStateOpen)
	{
		paicol_set->Close();
	}
	if(!modC_RING())//9/2
	{
		return false;
	}
	while(pos<10)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
	}

	CString strSQL; //*1

	strSQL="delete from e_paicol where enginid='"+EDIBgbl::SelPrjID+"'";
	_bstr_t sql;
	sql=strSQL;
	//清空表e_paicol中的所有记录,此表用来保存油漆材料表的结果
	try
	{
	m_Pconnection->Execute(sql,NULL,adCmdText);
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	} //*1

	//对表paint_c的预处理,程序就是对该表进行汇总处理
	try
	{
		
		strSQL="update paint_c set 面漆名=' ',面漆度数=0,面合计用量=0 where((pai_c_face='N') or (pai_c_face='n') and enginid='"+EDIBgbl::SelPrjID+"')"; //*2
		m_Pconnection->Execute((_bstr_t)strSQL,NULL,adCmdText);//*2
		
		//对表paint_c进行相应的修改
		if(InitPaint_c()==false)	//*3 //*3
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
			return false;
		}
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);

		return	false;
	}
	while(pos<20)
	{
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
	}

	//对表paint_c进行汇总
	try
	{
		if(paicol_set->State==adStateOpen)
		{
			paicol_set->Close();
		}
		_RecordsetPtr TargetSet;//目标记录集
		TargetSet.CreateInstance(__uuidof(Recordset));	//*4
		TargetSet->Open(_bstr_t(_T("select * from e_paicol")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);

		m_strPaint_c="col_name";
		m_valPaint_c="col_amount";

		//对底漆进行统计1
		strSQL="select * from paint_c where enginid='"+EDIBgbl::SelPrjID+"' order by 底漆名 desc";
		sql=strSQL;
		paicol_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		TotalTable(paicol_set,TargetSet,"底漆名","底合计用量"); //*4
		while(pos<30)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}

		//对面漆进行统计2 //*5
		strSQL="select * from paint_c where enginid='"+EDIBgbl::SelPrjID+"' order by 面漆名 desc";
		if(paicol_set->State==adStateOpen)
		{
			paicol_set->Close();
		}
		sql=strSQL;
		paicol_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		TotalTable(paicol_set,TargetSet,"面漆名","面合计用量"); //*5
		while(pos<40)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}

		//对pai_r_colr进行统计5  //*9
		strSQL="select * from paint_c where enginid='"+EDIBgbl::SelPrjID+"' order by pai_r_colr";
		sql=strSQL;
		if(paicol_set->State==adStateOpen)
		{
			paicol_set->Close();
		}
		paicol_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		TotalTable(paicol_set,TargetSet,"pai_r_colr","pai_r_t");
		while(pos<50)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}   //*9

		//对pai_a_colr进行统计6
		strSQL="select * from paint_c where enginid='"+EDIBgbl::SelPrjID+"' order by pai_a_colr";
		sql=strSQL;
		if(paicol_set->State==adStateOpen)  //*10
		{
			paicol_set->Close();
		}
		paicol_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		TotalTable(paicol_set,TargetSet,"pai_a_colr","pai_a_t");
		while(pos<60)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}   

		//8/26
		//对字段:col_unit进行赋值
		strSQL="update e_paicol set col_unit='kg' where enginid='"+EDIBgbl::SelPrjID+"'";
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);
		//8/26
		//对pai_a1进行统计3
		strSQL="select * from paint_c where enginid='"+EDIBgbl::SelPrjID+"' order by pai_a1 desc";
		sql=strSQL;
		if(paicol_set->State==adStateOpen)//*7
		{
			paicol_set->Close();
		}
		paicol_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		TotalTable(paicol_set,TargetSet,"pai_a1","pai_a_c1");
		while(pos<70)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}//*7


		//对pai_a2进行统计4
		strSQL="select * from paint_c where enginid='"+EDIBgbl::SelPrjID+"' order by pai_a2 desc";
		sql=strSQL;
		if(paicol_set->State==adStateOpen)
		{
			paicol_set->Close();
		}
		paicol_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		TotalTable(paicol_set,TargetSet,"pai_a2","pai_a_c2");

		while(pos<80)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}
		// 将上面汇总出的油漆再进行一次汇总（因为不同的类型可能会使用相同的油漆） zsy [2005/12/23]
		TotalSelfTable();
		
		//8/26对字段col_unit行赋值
		strSQL="update e_paicol set col_unit='m2' where enginid='"+EDIBgbl::SelPrjID+"' and col_unit is NULL";
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);

		//删除汇总值为零的记录
		strSQL="delete from e_paicol where enginid='"+EDIBgbl::SelPrjID+"' and col_amount=0";
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);
		while(pos<100)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}

		//8/26
		//8/28
//		AfxMessageBox("生成油漆材料汇总表成功!");//10/13

		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);		//8/28
		
	}catch(_com_error e)
	{
		AfxMessageBox("生成油漆材料汇总表失败!");
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
		return false;
	}

	//9/18
	//对表：pass 进行处理
	_RecordsetPtr pass_set;
	int flagPass=0;
	pass_set.CreateInstance(__uuidof(Recordset));
	pass_set->Open(_bstr_t(_T("select * from pass where enginid='"+EDIBgbl::SelPrjID+"'")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
	if(pass_set->GetRecordCount()>0)
	{
		pass_set->MoveFirst();
		for(;!pass_set->adoEOF;pass_set->MoveNext())
		{
			if(flagPass==0)
			{
				pass_set->PutCollect(_T("pass_mark2"),_T("T"));
			}
			else
			{
				pass_set->PutCollect(_T("pass_mark2"),_T("F"));
			}
			key=pass_set->GetCollect(_T("pass_mod2"));
			if((!(key.vt==VT_NULL||key.vt==VT_EMPTY))&&(flagPass==0))
			{
				str=key.bstrVal;
				flagPass=(str.Compare("C_PAICOL")==0)?1:0;
			}
		}
	}	

	//9/18

	return TRUE;
}

//删除字符串的前后空格符
CString CTotalTableIPED::Trim(LPCTSTR pcs)
{
	CString s = pcs;
	s.TrimLeft();
	s.TrimRight();
	return s;
	
}

//对表:paint_t进行相应的修改
BOOL CTotalTableIPED::InitPaint_c()
{
	_RecordsetPtr Paint_cSet;
	_RecordsetPtr pRsRing_T;
	pRsRing_T.CreateInstance(__uuidof(Recordset));
	Paint_cSet.CreateInstance(__uuidof(Recordset));
	_variant_t key;
	CString strRingPaintName;
	CString str;
	CString strTemp;
	try
	{
		//打开色环箭头油漆材料库，取出油漆名称。
		pRsRing_T->Open(_variant_t("SELECT * FROM [a_ring_paint]"), m_pConnectionCODE.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		if (pRsRing_T->adoEOF && pRsRing_T->BOF)
		{
			strRingPaintName = "调和漆";		//始果没有记录时默认为"调和漆".
		}
		else
		{
			pRsRing_T->MoveFirst();
			strRingPaintName = vtos(pRsRing_T->GetCollect(_variant_t("paint_name")));
			if (strRingPaintName.IsEmpty())
			{
				strRingPaintName = "调和漆";		//取出的是一个空字符串时，用默认值
			}
		}
		Paint_cSet->Open(_bstr_t(_T("select * from paint_c where enginid='"+EDIBgbl::SelPrjID+"'")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(Paint_cSet->GetRecordCount()==0)
		{
			AfxMessageBox(" 提示:没有油漆材料表的原始记录!");
			return false;
		}
		for(;!Paint_cSet->adoEOF;Paint_cSet->MoveNext())
		{//4
			key=Paint_cSet->GetCollect(_T("pai_c_face"));
			//修改面漆名
			if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
			{//3
				str=key.bstrVal;
				if(!((str=="n")||(str=="N")))
				{//2
					key=Paint_cSet->GetCollect(_T("面漆名"));
					if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
					{//1
					strTemp=key.bstrVal;
					 strTemp=Trim(strTemp);
					str=""+strTemp+"("+str+")";
					Paint_cSet->PutCollect(_T("面漆名"),_variant_t(str));
					}//1
				}//2
			}//3


			key=Paint_cSet->GetCollect(_T("pai_r_colr"));
			//修改字段pai_r_colr的值
			if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
			{//2
				str=key.bstrVal;
				if(!(str=="n"||str=="N"))
				{//1
					str=strRingPaintName+"("+str+")";
					Paint_cSet->PutCollect(_T("pai_r_colr"),_variant_t(str));
				}//1
				else if(str=="n"||str=="N")
				{
					str=" ";
					Paint_cSet->PutCollect(_T("pai_r_colr"),_variant_t(str));
				}
			}//2

			//修改字段pai_a_colr中的值
			key=Paint_cSet->GetCollect(_T("pai_a_colr"));
			if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
			{//2
				str=key.bstrVal;
				if(!(str=="n"||str=="N"))
				{//1
				str=strRingPaintName+"("+str+")";
				Paint_cSet->PutCollect(_T("pai_a_colr"),_variant_t(str));
				}//1
				else if(str=="n"||str=="N")
				{
					str=" ";
					Paint_cSet->PutCollect(_T("Pai_a_colr"),_variant_t(str));
				}
			}//2
			Paint_cSet->Update();
		}//4

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;

}


//对给定的字段值进行汇总,构建通用的汇总函数
//源记录集SourceSet必须按字段source_str进行排序
BOOL CTotalTableIPED::TotalTable(_RecordsetPtr SourceSet,_RecordsetPtr TargetSet,CString source_str,CString source_val)
{
	SourceSet->MoveFirst();
	double totalval=0.0;
	double temptotal=0.0;
	CString strTemp;
	CString last_str;
	last_str="";//赋初始值
	int flag=0;
	_variant_t key;
	try
	{
		for(;!SourceSet->adoEOF;SourceSet->MoveNext())
		{//4

			key=SourceSet->GetCollect(_variant_t(source_str));//在源记录集中取值
			if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
			{//3
				strTemp=key.bstrVal;
				strTemp=Trim(strTemp);

				key=SourceSet->GetCollect(_variant_t(source_val));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{//1
					temptotal=key.dblVal;
				}//1
				else 
				{//1
					temptotal=0.0;
				}//1

				//相邻两次SourceSet记录集中的字段(source_str)值是否相等
				if(!(strTemp.Compare(last_str)==0))//不相等
				{//2
					
					//当它不是第一次,在汇总表中增加一条记录
					if(flag!=0&&last_str!="")
					{//1
						TargetSet->AddNew();
						TargetSet->PutCollect(_T(_variant_t(m_strPaint_c)),_variant_t(last_str));
						TargetSet->PutCollect(_T(_variant_t(m_valPaint_c)),_variant_t(totalval));
						TargetSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));//8/25
						TargetSet->Update();
					}//1
					last_str=strTemp;
					totalval=0.0;
				}//2
				flag=1;
				totalval+=temptotal;//求和

			}//3没有取到值
			else 
			{
				if(flag!=0&&last_str!="")
				{
					TargetSet->AddNew();
					TargetSet->PutCollect(_T(_variant_t(m_strPaint_c)),_variant_t(last_str));
					TargetSet->PutCollect(_T(_variant_t(m_valPaint_c)),_variant_t(totalval));
					TargetSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));//8/25
					TargetSet->Update();
					totalval=0.0;
				}
				last_str="";
			}
		}//4


		//到了记录尾还要增加一条记录
		if(last_str!="")
		{
			TargetSet->AddNew();
			TargetSet->PutCollect(_T(_variant_t(m_strPaint_c)),_variant_t(last_str));
			TargetSet->PutCollect(_T(_variant_t(m_valPaint_c)),_variant_t(totalval));
			TargetSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));
			TargetSet->Update();
		
		}

	}catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}


//************************************
//* 模 块 名: c_presiz()			 *
//* 功    能: 生成保温材料规格汇总表 *
//* 上级模块: MAIN_MENU              *
//* 下级模块:                        *
//* 引 用 库: 1区e_presiz,           *
//*              pre_calc,col        *
//* 修改日期:                        *
//************************************

BOOL CTotalTableIPED::c_presiz()
{
	int pos=1;
	_RecordsetPtr presiz_set;
	presiz_set.CreateInstance(__uuidof(Recordset));
	_variant_t key;
	CString str;
	try
	{

		//清空表e_presiz中的所有记录,该表为保温材料汇总表(规格)的结果表
		CString strSQL;
		_bstr_t sql;
		strSQL="delete from e_presiz where enginid='"+EDIBgbl::SelPrjID+"'";//8/27
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);
		while(pos<30)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}
		//对表：e_presiz进行预处理
		if(InitPresiz()==false)
		{
			AfxMessageBox("生成保温材料汇总表没有成功!");
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
			return false;
		}
		while(pos<60)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}
		//对保温材料进行汇总8/28
		_RecordsetPtr sourceSet;
		sourceSet.CreateInstance(__uuidof(Recordset));
		_RecordsetPtr TargetSet;
		TargetSet.CreateInstance(__uuidof(Recordset));
		sourceSet->Open(_bstr_t(_T("select * from temp_presiz order by size_key")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		TargetSet->Open(_bstr_t(_T("select * from e_presiz")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		while(pos<80)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}

		if(totalPresiz(sourceSet,TargetSet)==false)
		{
			AfxMessageBox("生成保温材料汇总表没有成功!");
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
			return false;
		}
		else
		{
			while(pos<100)
			{
				((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
			}
//			AfxMessageBox("生成保温材料汇总表成功!");//10/13
		}
		//8/28
		//对表：pass 进行处理
		_RecordsetPtr pass_set;
		int flagPass=0;
		pass_set.CreateInstance(__uuidof(Recordset));
		pass_set->Open(_bstr_t(_T("select * from pass where enginid='"+EDIBgbl::SelPrjID+"'")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(pass_set->GetRecordCount()>0)
		{
			pass_set->MoveFirst();
			for(;!pass_set->adoEOF;pass_set->MoveNext())
			{
				if(flagPass==0)
				{
					pass_set->PutCollect(_T("pass_mark1"),_T("T"));
				}
				else
				{
					pass_set->PutCollect(_T("pass_mark1"),_T("F"));
				}
				key=pass_set->GetCollect(_T("pass_mod1"));
				if((!(key.vt==VT_NULL||key.vt==VT_EMPTY))&&(flagPass==0))
				{
					str=key.bstrVal;
					flagPass=(str.Compare("C_PRESIZ")==0)?1:0;
				}
			}
		}
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
		return false;
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);

	return true;
}


//初始化表:presiz 
BOOL CTotalTableIPED::InitPresiz()
{
	_RecordsetPtr presiz_set;
	CString strSQL;
	_bstr_t sql;
	CString str;
	_variant_t key;
	double val;

	presiz_set.CreateInstance(__uuidof(Recordset));
	try
	{
		//判断表temp_presiz是否存在
		if(tableExists(m_Pconnection,"temp_presiz")==false)
		{
			//创建表temp_presiz
			strSQL="create table temp_presiz (id long,size_num double,规格 double,外保温材料 char,外保温厚 double,c_pre_wei double,c_pipe_wei double,";
			strSQL=strSQL+"数量 double,size_amoun double,size_unit char,外单体积 double,外全体积 double,size_key char,size_mark char,内保温材料 char,内保温厚 double,EnginID char)";
			sql=strSQL;

			m_Pconnection->Execute(sql,NULL,adCmdText);
		}
		else
		{
			//清空表：temp_presiz
			strSQL="delete from temp_presiz";
			sql=strSQL;
			m_Pconnection->Execute(sql,NULL,adCmdText);

		}
		

		//将表:e_preexp中的内容导入到表:temp_presiz
		strSQL="insert into temp_presiz(id,规格,外保温材料,外保温厚,c_pre_wei,c_pipe_wei,";
		strSQL=strSQL+"数量,外单体积,外全体积,内保温材料,内保温厚,EnginID )";
		strSQL=strSQL+"  select id,规格,外保温材料,外保温厚,c_pre_wei,c_pipe_wei , ";
		strSQL=strSQL+"数量,外单体积,外全体积,内保温材料,内保温厚,EnginID ";
		strSQL=strSQL+"from e_preexp where 内保温材料 is not NULL and EnginID='"+EDIBgbl::SelPrjID+"'";
		strSQL=strSQL+" and 内保温材料<>''";
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);

		strSQL="select * from temp_presiz";
		sql=strSQL;
		presiz_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);

		if(presiz_set->GetRecordCount()>0)
		{//3
			presiz_set->MoveFirst();
			int count=0;
			for(;!presiz_set->adoEOF;presiz_set->MoveNext())
			{//2
				key=presiz_set->GetCollect(_T("内保温材料"));
				str=key.bstrVal;
				str=Trim(str);
				presiz_set->PutCollect(_T("外保温材料"),_variant_t(str));
				key=presiz_set->GetCollect(_T("内保温厚"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{//1
					val=key.dblVal;
				}//1
				else
				{//1
					val=0.0;
				}//1
				presiz_set->PutCollect(_T("外保温厚"),_variant_t(val));
				val=0.0;
				presiz_set->PutCollect(_T("内保温厚"),_variant_t(val));
				str="";
				presiz_set->PutCollect(_T("内保温材料"),_variant_t(str));
				str="内保温层";
				presiz_set->PutCollect(_T("size_mark"),_variant_t(str));
				presiz_set->Update();
				count++;//统计记录个数

			}//2
		}//3


		//把表:e_preexp中的所有记录都添加到表：e_presiz中
		strSQL="insert into e_presiz(id,规格,外保温材料,外保温厚,c_pre_wei,c_pipe_wei,";
		strSQL=strSQL+"数量,外单体积,外全体积,内保温材料,内保温厚,EnginID )";
		strSQL=strSQL+"  select id,规格,外保温材料,外保温厚,c_pre_wei,c_pipe_wei , ";
		strSQL=strSQL+"数量,外单体积,外全体积,内保温材料,内保温厚,EnginID ";
		strSQL=strSQL+"from e_preexp where  EnginID='"+EDIBgbl::SelPrjID+"' ";
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);

		if(presiz_set->State==adStateOpen)
		{
			presiz_set->Close();
		}
		
		//对表：e_presiz中的当前记录中的规格进行修改
		strSQL="select * from e_presiz	where enginid='"+EDIBgbl::SelPrjID+"'";
		sql=strSQL;

		presiz_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(presiz_set->GetRecordCount()>0)
		{//2
			double val1;

			presiz_set->MoveFirst();
			for(;!presiz_set->adoEOF;presiz_set->MoveNext())
			{//1
				//取规格
				key=presiz_set->GetCollect(_T("规格"));
				val=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
				key=presiz_set->GetCollect(_T("内保温厚"));
				val1=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
				val=val+2*val1;
				presiz_set->PutCollect(_T("规格"),_variant_t(val));
				presiz_set->Update();
				
			}//1
		}//2

		//将表：e_presiz 和表:temp_presiz连接在一起
		strSQL="insert into temp_presiz (id ,size_num ,规格 ,外保温材料 ,外保温厚 ,c_pre_wei ,c_pipe_wei ,";
		strSQL=strSQL+"数量 ,size_amoun ,size_unit ,外单体积 ,外全体积 ,size_key ,size_mark ,内保温材料 ,内保温厚 ,EnginID )";
		strSQL=strSQL+"select id ,size_num ,规格 ,外保温材料 ,外保温厚 ,c_pre_wei ,c_pipe_wei ,";
		strSQL=strSQL+"数量 ,size_amoun ,size_unit ,外单体积 ,外全体积 ,size_key ,size_mark ,内保温材料 ,内保温厚 ,EnginID ";
		strSQL=strSQL+"from e_presiz where enginid='"+EDIBgbl::SelPrjID+"'";

		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);
		//8/27
		strSQL="delete from e_presiz where enginid='"+EDIBgbl::SelPrjID+"'";
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);
		//8/27

		if(presiz_set->State==adStateOpen)
		{
			presiz_set->Close();
		}

		strSQL="select * from temp_presiz where enginid='"+EDIBgbl::SelPrjID+"'";
		sql=strSQL;
		presiz_set->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(presiz_set->GetRecordCount()>0)
		{//4
			presiz_set->MoveFirst();
			CString strTemp;
			double valTemp;

			for(;!presiz_set->adoEOF;presiz_set->MoveNext())
			{//3
				//修改size_key
				key=presiz_set->GetCollect(_T("外保温材料"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					str=key.bstrVal;
				}
				else
				{
					str="";
				}
				str=Trim(str);
				key=presiz_set->GetCollect(_T("规格"));
				val=(!(key.vt==VT_EMPTY||key.vt==VT_NULL))?key.dblVal:0.0;
				valTemp=val;
				strTemp.Format("%6.1f",val);
				str=str+strTemp;
				key=presiz_set->GetCollect(_T("外保温厚"));
				val=(!(key.vt==VT_EMPTY||key.vt==VT_NULL))?key.dblVal:0.0;
				strTemp.Format("%4.0f",val);
				str=str+strTemp;
				presiz_set->PutCollect(_T("size_key"),_variant_t(str));

				//修改size_amount
				key=presiz_set->GetCollect(_T("数量"));
				val=(!(key.vt==VT_EMPTY||key.vt==VT_NULL))?key.dblVal:0.0;
				presiz_set->PutCollect(_T("size_amoun"),_variant_t(val));

				//修改size_unit
				str="m3";
				presiz_set->PutCollect(_T("size_unit"),_variant_t(str));

				//修改size_mark
				if(valTemp>=2000)
				{//2
					key=presiz_set->GetCollect("size_mark");
					if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
					{//1
						str=key.bstrVal;
						str=Trim(str);
					}//1
					else
					{//1
						str="";
					}//1
					str=str+"设备";
					presiz_set->PutCollect(_T("size_mark"),_variant_t(str));
				}//2	
				presiz_set->Update();

			}//3
		}//4
		if(presiz_set->State==adStateOpen)
		{
			presiz_set->Close();
		}

}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	
	return true;

}


//判断该表是否存在
BOOL CTotalTableIPED::tableExists(_ConnectionPtr pCon, CString tbn)
{
	if(pCon==NULL || tbn=="")
		return false;
	_RecordsetPtr rs;
	if(tbn.Left(1)!="[")
		tbn="["+tbn;
	if(tbn.Right(1)!="]")
		tbn+="]";
	try{
		rs=pCon->Execute(_bstr_t(tbn),NULL,adCmdTable);
		rs->Close();
		return true;
	}
	catch(_com_error e)
	{
		return false;
	}

}


//对表：e_presiz 进行汇总
//记录集 sourceSet 必须按指定的字段进行排序
BOOL CTotalTableIPED::totalPresiz(_RecordsetPtr sourceSet, _RecordsetPtr TargetSet)
{
	_variant_t key;
	double totalSize_amoun=0.0;//保存字段size_amoun的和
	double tempSize_amoun=0.0;
	
	double totalVol=0.0;//保存外全体积的和
	double tempVol=0.0;

	CString strTempKey;
	CString last_strKey;
	last_strKey="";//赋初始值
	int flag=0;

	double TempSize_num,lastSize_num;
	double TempSpec,lastSpec;
	CString TempOut_heat,lastOut_heat;
	double TempOut_warm,lastOut_warm;
	CString TempSize_unit,lastSize_unit;
	CString TempSize_mark,lastSize_mark;

	
	if(sourceSet->GetRecordCount()>0)
	{

		sourceSet->MoveFirst();
	}
	else 
	{
		AfxMessageBox("提示:保温材料没有原始记录!");
		return false;
	}

	try
	{
		for(;!sourceSet->adoEOF;sourceSet->MoveNext())
		{//4

			key=sourceSet->GetCollect(_T("size_key"));//在源记录集中取键值
			if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
			{//3
				strTempKey=key.bstrVal;
				strTempKey=Trim(strTempKey);//9/2

				//取字段：size_amoun 的值
				key=sourceSet->GetCollect(_T("size_amoun"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{//1
					tempSize_amoun=key.dblVal;
				}//1
				else 
				{//1
					tempSize_amoun=0.0;
				}//1

				//取字段: 外全体积 的值
				key=sourceSet->GetCollect(_T("外全体积"));
				tempVol=(!(key.vt==VT_EMPTY||key.vt==VT_NULL))?key.dblVal:0.0;
					
				//取字段:size_num 中的值
				key=sourceSet->GetCollect(_T("size_num"));
				TempSize_num=(!(key.vt==VT_NULL||key.vt==VT_EMPTY))?key.dblVal:0.0;

				//取字段:规格
				key=sourceSet->GetCollect(_T("规格"));
				TempSpec=(!(key.vt==VT_NULL||key.vt==VT_EMPTY))?key.dblVal:0.0;

				//外保温材料
				key=sourceSet->GetCollect(_T("外保温材料"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					TempOut_heat=key.bstrVal;
					TempOut_heat=Trim(TempOut_heat);
				}
				else
				{
					TempOut_heat=" ";
				}

				//外保温厚
				key=sourceSet->GetCollect(_T("外保温厚"));
				TempOut_warm=(!(key.vt==VT_EMPTY||key.vt==VT_NULL))?key.dblVal:0.0;

				//字段：size_unit
				key=sourceSet->GetCollect(_T("size_unit"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					TempSize_unit=key.bstrVal;
					TempSize_unit=Trim(TempSize_unit);
				}
				else
				{
					TempSize_unit=" ";
				}

				//字段：size_mark
				key=sourceSet->GetCollect(_T("size_mark"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					TempSize_mark=key.bstrVal;
					TempSize_mark=Trim(TempSize_mark);
				}
				else
				{
					TempSize_mark=" ";
				}

				//相邻两次SourceSet记录集中的字段(source_str)值是否相等
				if(!(strTempKey.Compare(last_strKey)==0))//不相等
				{//2
					
					//当它不是第一次,在汇总表中增加一条记录
					if(flag!=0&&last_strKey!="")
					{//1
						CString strGet;
						TargetSet->AddNew();
						TargetSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));//9/2
						TargetSet->PutCollect(_T("size_key"),_variant_t(last_strKey));
						TargetSet->PutCollect(_T("size_amoun"),_variant_t(totalSize_amoun));
						TargetSet->PutCollect(_T("外全体积"),_variant_t(totalVol));
						
						TargetSet->PutCollect(_T("size_num"),_variant_t(lastSize_num));

						TargetSet->PutCollect(_T("规格"),_variant_t(lastSpec));

						TargetSet->PutCollect(_T("外保温材料"),_variant_t(lastOut_heat));

						TargetSet->PutCollect(_T("外保温厚"),_variant_t(lastOut_warm));

						TargetSet->PutCollect(_T("size_unit"),_variant_t(lastSize_unit));
						TargetSet->PutCollect(_T("size_mark"),_variant_t(lastSize_mark));

						TargetSet->Update();
					}//1
					last_strKey=strTempKey;
					totalSize_amoun=0.0;
					totalVol=0.0;//9/2
					lastSize_num=TempSize_num;
					lastSpec=TempSpec;
					lastOut_heat=TempOut_heat;
					lastOut_warm=TempOut_warm;
					lastSize_unit=TempSize_unit;
					lastSize_mark=TempSize_mark;
				}//2
				flag=1;
				totalSize_amoun+=tempSize_amoun;//求和
				totalVol+=tempVol;//9/2

			}//3没有取到值
			else 
			{
				if(flag!=0&&last_strKey!="")
				{
					CString strGet;
					TargetSet->AddNew();
					TargetSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));
					TargetSet->PutCollect(_T("size_key"),_variant_t(last_strKey));
					TargetSet->PutCollect(_T("size_amoun"),_variant_t(totalSize_amoun));
					TargetSet->PutCollect(_T("外全体积"),_variant_t(totalVol));
					
					TargetSet->PutCollect(_T("size_num"),_variant_t(lastSize_num));

					TargetSet->PutCollect(_T("规格"),_variant_t(lastSpec));

					TargetSet->PutCollect(_T("外保温材料"),_variant_t(lastOut_heat));

					TargetSet->PutCollect(_T("外保温厚"),_variant_t(lastOut_warm));

					TargetSet->PutCollect(_T("size_unit"),_variant_t(lastSize_unit));

					TargetSet->PutCollect(_T("size_mark"),_variant_t(lastSize_mark));

					TargetSet->Update();
					totalSize_amoun=0.0;
					totalVol=0.0;
				}
				last_strKey="";
			}
		}//4


		//到了记录尾还要增加一条记录
		if(last_strKey!=""&&sourceSet->GetRecordCount()>0)
		{
			CString strGet;
			TargetSet->AddNew();
			TargetSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));
			TargetSet->PutCollect(_T("size_key"),_variant_t(last_strKey));
			TargetSet->PutCollect(_T("size_amoun"),_variant_t(totalSize_amoun));
			TargetSet->PutCollect(_T("外全体积"),_variant_t(totalVol));
			
			TargetSet->PutCollect(_T("size_num"),_variant_t(lastSize_num));

			TargetSet->PutCollect(_T("规格"),_variant_t(lastSpec));

			TargetSet->PutCollect(_T("外保温材料"),_variant_t(lastOut_heat));

			TargetSet->PutCollect(_T("外保温厚"),_variant_t(lastOut_warm));

			TargetSet->PutCollect(_T("size_unit"),_variant_t(lastSize_unit));

			TargetSet->PutCollect(_T("size_mark"),_variant_t(lastSize_mark));

			TargetSet->Update();

		}

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}


//生成保温材料汇总表(标准+辅助)
void CTotalTableIPED::c_preast()
{
	CString strpi;
	strpi.Format("%9.7f",pi);
	_RecordsetPtr a_matastSet;
	a_matastSet.CreateInstance(__uuidof(Recordset));
	_RecordsetPtr a_wireSet;
	a_wireSet.CreateInstance(__uuidof(Recordset));
	_RecordsetPtr e_preexpSet;
	e_preexpSet.CreateInstance(__uuidof(Recordset));
	_RecordsetPtr e_preexpSet1;
	e_preexpSet1.CreateInstance(__uuidof(Recordset));
	_RecordsetPtr a_vertSet;
	a_vertSet.CreateInstance(__uuidof(Recordset));
	CString strSQL;
	_bstr_t sql;

	CString main_mat;
	_variant_t key;
	CString ast;
	CString sum_field;
	CString con;
	CString str_wire_base;
	CString sum_field1;
	CString vert_form;
	CString ver_amou;
	double wire_tot;
	int a_vertFieldsCount;
	CString strast;//9/3
	double wire_amoun;
	CString ast2;
	double preexp_tot;
	double vert_amouD;

	try
	{

		strSQL="select * from a_wire where enginid='"+EDIBgbl::SelPrjID+"'"; //9/22
		sql=strSQL;
		a_wireSet->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		
		strSQL="update a_wire set wire_amoun=0 where enginid='"+EDIBgbl::SelPrjID+"'";
		sql=strSQL;
		m_Pconnection->Execute(sql,NULL,adCmdText);//9/22

		strSQL="select * from a_matast where EnginID='"+EDIBgbl::SelPrjID+"' ";
		sql=strSQL;
		a_matastSet->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);//*01
		
		strSQL="select * from e_preexp where enginid='"+EDIBgbl::SelPrjID+"'";	//*2
		sql=strSQL;
		e_preexpSet->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		
		strSQL="select * from a_vert where enginid='"+EDIBgbl::SelPrjID+"'";//9/22
		sql=strSQL;
		a_vertSet->Open(sql,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		
		//获取表：a_vert中的字段个数
		FieldsPtr a_vert;
		a_vertSet->get_Fields(&a_vert);
		a_vertFieldsCount=a_vert->Count;
		a_vertFieldsCount=a_vertFieldsCount/5;	

		int i;									
		CString stri;
		CString str;
		for(i=1;i<=a_vertFieldsCount;i++)
		{
			stri.Format("%d",i);
			str="vert_amou"+stri;
			strSQL="update a_vert set "+str+"=0 where enginid='"+EDIBgbl::SelPrjID+"'";//9/22
			sql=strSQL;
			m_Pconnection->Execute(sql,NULL,adCmdText);
		}

		//获取表:a_matast中的字段个数
		FieldsPtr a_matast;
		a_matastSet->get_Fields(&a_matast);
		int a_matastFieldsCount;
		a_matastFieldsCount=a_matast->Count;
		a_matastFieldsCount--; //减去工程字段。zsy


		if(a_matastSet->GetRecordCount()<=0)//??是继续运行还是返回
		{
			return;
		}								//*2

		for(i=1;i<a_matastFieldsCount;i++)  //*01
		{ //10
			a_matastSet->MoveFirst();//*01
			for(;!a_matastSet->adoEOF;a_matastSet->MoveNext()) //*02
			{  //9
				key=a_matastSet->GetCollect(_T("mat_name"));///????
				if(!(key.vt==VT_NULL||key.vt==VT_EMPTY))
				{
					main_mat=key.bstrVal;
				}
				else
				{
					main_mat="";
				}
				

				key=a_matastSet->GetCollect(_variant_t((short)i));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					ast=key.bstrVal;
					ast=Trim(ast);
				}
				else if(key.vt==VT_EMPTY||key.vt==VT_NULL)
				{
					ast="";
				}
				if(ast=="")
				{
					continue;
				}

				//定位记录
				strast="wire_name='"+ast+"'";
				if(a_wireSet->GetRecordCount()>0)
				{
					a_wireSet->MoveFirst();//9/3
					a_wireSet->Find(_bstr_t(strast),0,adSearchForward);
				}									//*02


				for(;(a_wireSet->GetRecordCount()>0)&&(!a_wireSet->adoEOF);)  //*03
				{ //7
					key=a_wireSet->GetCollect(_T("wire_sum"));
					if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
					{
						sum_field=key.bstrVal;
					}
					else 
					{
						sum_field="";
					}
					key=a_wireSet->GetCollect(_T("wire_con"));
					if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
					{
						con=key.bstrVal;
						con=Trim(con);
						con.Replace(".and."," and ");
						con.Replace(".AND."," AND ");
						con.Replace("TRIM(main_mat)","'"+main_mat+"'");//9/3
						con.Replace("#","<>");
					}
					else
					{
						con="";
					}		//*03

					key=a_wireSet->GetCollect(_T("wire_base"));//*04 
					if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
					{ 
						str_wire_base=key.bstrVal;
						str_wire_base=Trim(str_wire_base);
					}
					else
					{
						str_wire_base="";
					}

					if(str_wire_base!="")  //*04
					{//6
						//9/8
						key=a_wireSet->GetCollect(_T("wire_size"));//*05
						if(!(key.vt==VT_NULL||key.vt==VT_EMPTY))
						{
							ast2=key.bstrVal;
							ast2=Trim(ast2);
						}
						else
						{
							ast2="";
						}
						strSQL="vert_name='"+ast2+"'";
						//9/8
						//9/2
						if(a_vertSet->GetRecordCount()>=0)
						{//5
							a_vertSet->MoveFirst();
							
							//9/2
							a_vertSet->Find(_bstr_t(strSQL),0,adSearchForward);
							short k;

							for(k=1;k<=a_vertFieldsCount&&!a_vertSet->adoEOF;k++)//9/2
							{//4
								strSQL.Format("%d",k);
								vert_form="vert_form"+strSQL;
								key=a_vertSet->GetCollect(_variant_t(vert_form));
								if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
								{//3
									sum_field1=key.bstrVal;
									sum_field1=Trim(sum_field1);
									if(sum_field1=="") break;
									//对表:e_preexp的当前字段进行求和
									CString strSQL1;
									if(e_preexpSet1->State==adStateOpen)
									{
										e_preexpSet1->Close();
									}
									sum_field1.Replace("pi",""+strpi+"");	//9/17
									strSQL1="select sum("+sum_field1+") as tot from e_preexp where "+con+" and enginid='"+EDIBgbl::SelPrjID+"'";
									e_preexpSet1->Open((_bstr_t)strSQL1,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
									if(e_preexpSet1->GetRecordCount()>0)
									{//2
										e_preexpSet1->MoveFirst();
										key=e_preexpSet1->GetCollect(_T("tot"));
										preexp_tot=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
										//对表：a_vert中的ver_amou字段进行修改
										ver_amou="vert_amou"+strSQL;
										key=a_vertSet->GetCollect(_variant_t(ver_amou));
										vert_amouD=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
										vert_amouD+=preexp_tot;
										a_vertSet->PutCollect(_variant_t(ver_amou),_variant_t(vert_amouD));
										a_vertSet->Update();

									}//2
									
								}//3
								else
								{
									break;
								}
							}//4
						}//9/2 
					}//不是组合键则 //6
					else  //*04
					{//2
						if(e_preexpSet1->State==adStateOpen)
						{
							e_preexpSet1->Close();
						}
						if(!(sum_field==""||con==""))
						{
							sum_field.Replace("pi",""+strpi+"");//9/17
							strSQL="select sum("+sum_field+") as aa from e_preexp where "+con+" and enginid='"+EDIBgbl::SelPrjID+"'";
						
							e_preexpSet1->Open((_bstr_t)strSQL,(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
							if(e_preexpSet1->GetRecordCount()>0)  
							{//1
								e_preexpSet1->MoveFirst();
								key=e_preexpSet1->GetCollect(_T("aa"));
								wire_tot=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
								key=a_wireSet->GetCollect(_T("wire_amoun"));
								wire_amoun=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
								wire_amoun=wire_amoun+wire_tot;
								a_wireSet->PutCollect(_T("wire_amoun"),_variant_t(wire_amoun));
								a_wireSet->Update(); //*5
							}//1
						}
					}//2 *04
				
					a_wireSet->MoveNext(); //*03
					if(!a_wireSet->adoEOF)
					{
						a_wireSet->Find(_bstr_t(strast),0,adSearchForward); //*03
					}
				}//7

		
			}  //9  *02

		}  //10		*01
		
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return;
	}
	
	try
	{	
		//修改表：a_wire.wire_amoun(取整)  //*3
		strSQL="update a_wire set wire_amoun=int(wire_amoun) ";//9/22
		strSQL=strSQL+" where (wire_unit='个' or wire_unit='只' or wire_unit='块') and enginid='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);  //9/22

		//删除表:e_ast中的所有记录
		strSQL="delete from e_ast where enginid='"+EDIBgbl::SelPrjID+"'";//9/22		
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22
		
		//将表:a_wire中的记录复制到表：e_ast
		strSQL="insert into e_ast (col_num,col_name,gen_amount,col_size,col_unit,EnginID)";//将新增加的记录设置为当前工程. by zsy
		strSQL=strSQL+" select wire_num,wire_name,wire_amoun,wire_size,wire_unit,EnginID ";
		strSQL=strSQL+" from a_wire where (trim(wire_base)='' and wire_amoun<>0 ";//9/22
		strSQL=strSQL+" or wire_base is NULL and wire_amoun<>0) and enginid='"+EDIBgbl::SelPrjID+"'";//9/22
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);	//9/22

		//将表：a_vert中的相应字段复制到表：e_ast	//*5

		if(!(a_vertCopyE_ast(a_vertFieldsCount)))
		{
			return;
		}			//*5

		//设置表：e_ast中的字段：col_mark的值
		strSQL="update e_ast set col_mark=trim(col_name)+trim(col_size) where enginid='"+EDIBgbl::SelPrjID+"'";  //9/22
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22
		strSQL="update e_ast set col_mark=col_name where col_size is NULL and enginid='"+EDIBgbl::SelPrjID+"'";//9/22
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22

		//汇总
		_RecordsetPtr e_astSet;
		e_astSet.CreateInstance(__uuidof(Recordset));
		_RecordsetPtr epreastSet;
		epreastSet.CreateInstance(__uuidof(Recordset));
		strSQL="select * from e_ast  where enginid='"+EDIBgbl::SelPrjID+"' order by col_mark ";//9/22
		e_astSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);//9/22
		strSQL="delete from epreast where enginid='"+EDIBgbl::SelPrjID+"'";//9/22
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22
		strSQL="select * from epreast where enginid='"+EDIBgbl::SelPrjID+"'";//9/22
		epreastSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);//9/22
		if(!TotalPreast(e_astSet,epreastSet))
		{
			return;
		}   //*10

		strSQL="update epreast set col_mark=' ' where enginid='"+EDIBgbl::SelPrjID+"'"; //9/22
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22

		strSQL="update epreast set gen_amount=int(gen_amount+0.999) where (col_unit='个' ";//9/22
		strSQL=strSQL+" or col_unit='只' or col_unit='块') and enginid='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22

		strSQL="update epreast set col_amount=gen_amount where enginid='"+EDIBgbl::SelPrjID+"'";//9/22
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22
	
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return;
	}

}


//对记录集中的str字段求和 并返回和
//注意必须确保记录集不为空
double CTotalTableIPED::sum(CString str, _RecordsetPtr set)
{
	double tot=0.0;
	double val;
	_variant_t key;
	set->MoveFirst();
	try
	{
		for(;!set->adoEOF;set->MoveNext())
		{
			key=set->GetCollect((_variant_t)str);
			val=(key.vt==VT_NULL||key.vt==VT_EMPTY)?0.0:key.dblVal;
			tot=tot+val;
			
		}
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return -1;//表示求和失败
	}
	return tot;
}

//将表：a_vert中的记录复制到表:e_ast中
bool CTotalTableIPED::a_vertCopyE_ast(int count)
{
	CString vert_num;
	CString vert_mat;
	CString vert_amou;
	CString vert_size;
	CString vert_unit;
	CString strSQL;
	CString stri;
	int i;
	try{
			for(i=1;i<=count;i++)
			{
				stri.Format("%d",i);
				vert_size="vert_size"+stri;
				vert_mat="vert_mat"+stri;
				vert_amou="vert_amou"+stri;
				vert_unit="vert_unit"+stri;
				strSQL="insert into e_ast (col_num,col_name,Gen_amount,col_size,col_unit,EnginID) "; //将新增加的记录设置为当前工程. by zsy
				strSQL=strSQL+" select vert_num,"+vert_mat+","+vert_amou+","+vert_size+","+vert_unit+",EnginID ";
				strSQL=strSQL+" from a_vert where "+vert_amou+" <> 0 and enginid='"+EDIBgbl::SelPrjID+"'";//9/22
				m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);//9/22

			}
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}

//对保温材料进行汇总函数
bool CTotalTableIPED::TotalPreast(_RecordsetPtr sourceSet, _RecordsetPtr TargetSet)
{
	_variant_t key;
	short iCol_num=0;
	
	double totalGet_amount=0.0;//保存get_amount的和
	double tempGet_amount=0.0;

	double totalCol_amount;//保存col_amount的和
	double tempCol_amount;

	CString strCol_name;
	CString lastCol_name;
	CString strCol_size;
	CString lastCol_size;
	CString strCol_unit;
	CString lastCol_unit;

	CString strTempKey;
	CString last_strKey;
	last_strKey="";//赋初始值
	int flag=0;
	
	if(sourceSet->GetRecordCount()>0)
	{
		sourceSet->MoveFirst();
	}

	try
	{
		for(;!sourceSet->adoEOF;sourceSet->MoveNext())
		{//4

			key=sourceSet->GetCollect(_T("col_mark"));//在源记录集中取键值
			if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
			{//3
				strTempKey=key.bstrVal;


				//取字段: get_amount 的值
				key=sourceSet->GetCollect(_T("gen_amount"));
				tempGet_amount=(!(key.vt==VT_EMPTY||key.vt==VT_NULL))?key.dblVal:0.0;


				//取字段:col_amount的值
				key=sourceSet->GetCollect(_T("col_amount"));
				tempCol_amount=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
				
				key=sourceSet->GetCollect(_T("col_name"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					strCol_name=key.bstrVal;
				}
				else
				{
					strCol_name="";
				}

				
				key=sourceSet->GetCollect(_T("col_size"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					strCol_size=key.bstrVal;
				}
				else
				{
					strCol_size="";
				}


				key=sourceSet->GetCollect(_T("col_unit"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					strCol_unit=key.bstrVal;
				}
				else
				{
					strCol_unit="";
				}

				//相邻两次SourceSet记录集中的字段(source_str)值是否相等
				if(!(strTempKey.Compare(last_strKey)==0))//不相等
				{//2
					
					//当它不是第一次,在汇总表中增加一条记录
					if(flag!=0&&last_strKey!="")
					{//1
						
						iCol_num++;
						TargetSet->AddNew();
						TargetSet->PutCollect(_T("col_mark"),_variant_t(last_strKey));
						TargetSet->PutCollect(_T("col_num"),_variant_t(iCol_num));
						TargetSet->PutCollect(_T("gen_amount"),_variant_t(totalGet_amount));
						TargetSet->PutCollect(_T("col_amount"),_variant_t(totalCol_amount));
						TargetSet->PutCollect(_T("col_name"),_variant_t(lastCol_name));
						TargetSet->PutCollect(_T("col_size"),_variant_t(lastCol_size));
						TargetSet->PutCollect(_T("col_unit"),_variant_t(lastCol_unit));
						TargetSet->PutCollect(_T("EnginID"),_variant_t(EDIBgbl::SelPrjID));  //设置为当前工程.  by zsy

						TargetSet->Update();
					}//1
					last_strKey=strTempKey;
					lastCol_name=strCol_name;
					lastCol_size=strCol_size;
					lastCol_unit=strCol_unit;
					totalGet_amount=0.0;
					totalCol_amount=0.0;
				}//2
				flag=1;
				totalGet_amount+=tempGet_amount;
				totalCol_amount+=tempCol_amount;

			}//3没有取到值
			else 
			{
				if(flag!=0&&last_strKey!="")
				{
					iCol_num++;
					TargetSet->AddNew();
					TargetSet->PutCollect(_T("col_mark"),_variant_t(last_strKey));
					TargetSet->PutCollect(_T("col_num"),_variant_t(iCol_num));
					TargetSet->PutCollect(_T("gen_amount"),_variant_t(totalGet_amount));
					TargetSet->PutCollect(_T("col_amount"),_variant_t(totalCol_amount));
					TargetSet->PutCollect(_T("col_name"),_variant_t(lastCol_name));
					TargetSet->PutCollect(_T("col_size"),_variant_t(lastCol_size));
					TargetSet->PutCollect(_T("col_unit"),_variant_t(lastCol_unit));
					TargetSet->PutCollect(_T("EnginID"),_variant_t(EDIBgbl::SelPrjID));  //设置为当前工程. by zsy

					TargetSet->Update();

					totalGet_amount=0.0;
					totalCol_amount=0.0;
					
				}
				last_strKey="";
			}
		}//4


		//到了记录尾还要增加一条记录
		if(last_strKey!=""&&sourceSet->GetRecordCount()>0)
		{
			TargetSet->AddNew();
			iCol_num++;
			TargetSet->PutCollect(_T("col_mark"),_variant_t(last_strKey));
			TargetSet->PutCollect(_T("col_num"),_variant_t(iCol_num));
			TargetSet->PutCollect(_T("gen_amount"),_variant_t(totalGet_amount));
			TargetSet->PutCollect(_T("col_amount"),_variant_t(totalCol_amount));
			TargetSet->PutCollect(_T("col_name"),_variant_t(lastCol_name));
			TargetSet->PutCollect(_T("col_size"),_variant_t(lastCol_size));
			TargetSet->PutCollect(_T("col_unit"),_variant_t(lastCol_unit));
			TargetSet->PutCollect(_T("EnginID"),_variant_t(EDIBgbl::SelPrjID));//设置为当前工程. by zsy

			TargetSet->Update();

		}

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}

//功能：产生色环箭头油漆汇总表
bool CTotalTableIPED::modC_RING()
{
	try
	{
		int pos=1;  //9/4
		_RecordsetPtr pRsPaint_c;
		pRsPaint_c.CreateInstance(__uuidof(Recordset));

		CString strSQL;
		_variant_t var;
		//清空表PAINT_C
		strSQL = "DELETE FROM [paint_c] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
		//	*****将油漆提资单库paint和保温对象油漆颜色表库paint_c2(改为直接从e_preexp表中提取)  ****/
		//	*****合并到油漆颜色设计库paint_c中						****/

		//将表E_PAIEXP中的内容复制到表PAINT_C中//*3
		strSQL = "INSERT INTO [paint_c] (PAI_VOL,序号,名称,PAI_TYPE,PAI_SIZE,PAI_AMOU";
		strSQL += ",油漆面积,备注,底漆名,底漆度数,底每度用量,底合计用量,面漆名,面漆度数";
		strSQL += ",面每度用量,面合计用量,PAI_A1,PAI_A_T1,PAI_A2,PAI_A_T2,PAI_AREA,PAI_A_C1";
		strSQL += ",PAI_A_C2,PAI_CODE,PAI_C_FACE,PAI_R_COLR,PAI_R_COST,PAI_R_AREA";
		strSQL += ",PAI_R_T,PAI_A_COLR,PAI_A_COST,PAI_A_AREA,PAI_A_T,EnginID)";

		strSQL += "	 SELECT PAI_VOL,序号,名称,PAI_TYPE,PAI_SIZE,PAI_AMOU";
		strSQL +=",油漆面积,备注,底漆名,底漆度数,底每度用量,底合计用量,面漆名,面漆度数";
		strSQL += ",面每度用量,面合计用量,PAI_A1,PAI_A_T1,PAI_A2,PAI_A_T2,PAI_AREA,PAI_A_C1";
		strSQL += ",PAI_A_C2,PAI_CODE,PAI_C_FACE,PAI_R_COLR,PAI_R_COST,PAI_R_AREA";
		strSQL += ",PAI_R_T,PAI_A_COLR,PAI_A_COST,PAI_A_AREA,PAI_A_T,EnginID";

		strSQL += " FROM [e_paiexp] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText); //*3
		//将e_preexp表中的内容复制到表paint_c中.     //*2
		strSQL = "INSERT INTO [paint_c] (PAI_VOL, 名称, PAI_SIZE ";
		strSQL += ", PAI_AMOU, 油漆面积, 备注, pai_code, EnginID ) ";

		strSQL += "SELECT C_VOL AS PAI_VOL, 名称, 规格 AS PAI_SIZE ";
		strSQL += ", 数量 AS PAI_AMOU, 外表面积 AS 油漆面积, 备注, c_color AS pai_code, EnginID";
		strSQL += " FROM [e_preexp] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";

		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);//*2
		/// ************************************* /
		CString strFace="", strRing="", strArrow="";
		double  dw, dCode, dLength, dFace_l,
				dRing_l, dS_ring,dArrow_l, dS_arrow,
				dS_word, dev_area;
		
		double	dRing_Dosage;		//色环箭头的度数
		
		//打开表 pRsColr_l	调和漆单位消耗量库
		//		pRsRing_Paint 色环箭头油漆材料库
		//		pRsColor	色环箭头油漆颜色准则库
		//		pRsArrow	色环箭头尺寸准则库		 将各表做为类成员,一次性全部打开.提高速度	 [2005/06/10]
		if ( !OpenCodeDB() )
		{
			return false;
		}
		m_strExecuteSQLTblName = "ExecuteSQL";		//执行SQL语句的临时表名
		if (!CreateTempTable(m_Pconnection, m_strExecuteSQLTblName))
		{
			return FALSE;
		}
		if (pRsExecSQL == NULL)
		{
			pRsExecSQL.CreateInstance(__uuidof(Recordset));
		}
		if (pRsExecSQL->State != adStateOpen)
		{
			pRsExecSQL->Open(_variant_t("select * from ["+m_strExecuteSQLTblName+"]"), m_Pconnection.GetInterfacePtr(), adOpenStatic, adLockOptimistic,adCmdText);
		}
		//色环箭头油漆材料的度数
		if (pRsRing_Paint->adoEOF && pRsRing_Paint->BOF)
		{
			dRing_Dosage = 1;		//如果没有记录时默认为1度.
		}else
		{
			pRsRing_Paint->MoveFirst();
			dRing_Dosage = vtof(pRsRing_Paint->GetCollect(_variant_t("paint_dosage")));
			if (dRing_Dosage < 0)
			{
				dRing_Dosage = 1;
			}
		}
		////油漆原始数据表
		if( pRsPaint_c->State == adStateOpen )
		{
			pRsPaint_c->Close();
		}
		strSQL = "SELECT * FROM [paint_c] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
		pRsPaint_c->Open(_bstr_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
		//9/28
		if ((!pRsPaint_c->adoEOF) && (!pRsPaint_c->BOF))
		{
			pRsPaint_c->MoveFirst();
			while (!pRsPaint_c->adoEOF)
			{
				//初始化。
				strFace = strRing =  strArrow = " "; //*4
				dFace_l = dRing_l = dS_ring = dArrow_l = dS_arrow = dS_word = 0.0;
				
				var = pRsPaint_c->GetCollect("pai_size");      //外径
				dw  = (var.vt==VT_NULL || var.vt==VT_EMPTY)?0.0:var.dblVal;
				
				var = pRsPaint_c->GetCollect("pai_code");       //油漆颜色类型
				dCode = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
				
				var = pRsPaint_c->GetCollect("pai_amou");         //管长
				dLength = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
				
				var = pRsPaint_c->GetCollect("油漆面积");         //设备面积
				dev_area = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
				
				
				//调用模块 
				if( !moduleCOLOR_S(dw/*外径*/, dCode/*油漆颜色类型*/, dLength/*管长*/, strFace/*面漆色*/,
					dFace_l/*面漆每度用量*/, strRing/*色环*/, dRing_l/*每度用量*/,
					dS_ring/*色环总面积*/, strArrow/*箭头/文字*/, dArrow_l/*每度用量*/,
					dS_arrow/*箭头总面积*/, dS_word/*文字总面积*/, dev_area/*设备面积*/) )																							
				{
					return false;
				}
				pRsPaint_c->Fields->GetItem("pai_c_face")->PutValue(_variant_t(strFace));
				pRsPaint_c->Fields->GetItem("面每度用量")->PutValue(_variant_t(dFace_l));
				pRsPaint_c->Fields->GetItem("pai_r_colr")->PutValue(_variant_t(strRing));
				pRsPaint_c->Fields->GetItem("pai_r_cost")->PutValue(_variant_t(dRing_l));
				
				pRsPaint_c->Fields->GetItem("pai_r_area")->PutValue(_variant_t(dS_ring));
				pRsPaint_c->Fields->GetItem("pai_a_colr")->PutValue(_variant_t(strArrow));
				pRsPaint_c->Fields->GetItem("pai_a_cost")->PutValue(_variant_t(dArrow_l));
				pRsPaint_c->Fields->GetItem("pai_a_area")->PutValue(_variant_t((dS_arrow+dS_word)));

				//色环的用量 
				pRsPaint_c->PutCollect("PAI_R_T",_variant_t(dRing_Dosage*dRing_l*dS_ring));
				//箭头文字的用量
				pRsPaint_c->PutCollect("pai_a_t",_variant_t(dRing_Dosage*dArrow_l*(dS_arrow+dS_word)));
				
				pRsPaint_c->Update();
				
				pRsPaint_c->MoveNext();
				
				if(pos<100)
				{
					((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
				}
				if(pos>=100) pos=1;
			}
		}

		while(pos<100)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);			
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);

		return false;
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
	return true;

}
//功    能: 求色环,箭头,文字的   颜色及总面积              
bool CTotalTableIPED::moduleCOLOR_S(double& dw/*外径*/, double& dCode/*油漆颜色类型*/,
		double& dLength/*管长*/, CString& strFace/*面漆色*/, double& dFace_l/*面漆每度用量*/, 
		CString& strRing/*色环*/, double& dRing_l/*每度用量*/, double& dS_ring/*色环总面积*/, 
		CString& strArrow/*箭头/文字*/, double& dArrow_l/*每度用量*/, double& dS_arrow/*箭头总面积*/, 
		double& dS_word/*文字总面积*/,   double& dev_area/*设备面积*/)
{
	try
	{
		CString strSQL, str, strMedia, strExpression;
		_variant_t var;
		double  dTmpSum, cTmpSum;
		//设置状态
		strFace="N";
		strRing="N";
		strArrow="N";

		if( pRsColor->RecordCount <= 0 )
		{
			return false;
		}
		else
		{
			pRsColor->MoveFirst();
		}
		str.Format("%f",dCode);
		strSQL = "colr_code="+str+" ";
		pRsColor->Find(_bstr_t(strSQL), NULL, adSearchForward);//*1
		if( !pRsColor->adoEOF )	//*2
		{
			var      = pRsColor->GetCollect("colr_media");
			strMedia = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
			var      = pRsColor->GetCollect("colr_face");
			strFace  = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;

			var      = pRsColor->GetCollect("colr_ring");
			strRing  = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
			var      = pRsColor->GetCollect("colr_arrow");
			strArrow = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
		}  //*2

		if( pRsExecSQL->State != adStateOpen )
		{
			strSQL="SELECT * FROM ["+m_strExecuteSQLTblName+"] ";// WHERE "+strExpression+" ";
			pRsExecSQL->Open(_bstr_t(strSQL), (IDispatch*)m_Pconnection, adOpenStatic, adLockOptimistic, adCmdText);
		}
		long rCount = pRsExecSQL->GetRecordCount();
		if ( rCount> 1)
		{
/*			strSQL = "DELETE FROM ["+m_strExecuteSQLTblName+"]";
			m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

			//增中一条记录.
			str.Format("%f",dw);
			strSQL = "INSERT INTO ["+m_strExecuteSQLTblName+"] (DW) VALUES ("+str+") ";
*/
			for (pRsExecSQL->MoveFirst(), pRsExecSQL->MoveNext(); !pRsExecSQL->adoEOF; pRsExecSQL->MoveNext())
			{
				pRsExecSQL->Delete(adAffectCurrent);
			}
			pRsExecSQL->MoveFirst();
		}else if (rCount == 1)
		{
			pRsExecSQL->MoveFirst();
		}else
		{
			pRsExecSQL->AddNew();
		}
		//将DW保存到临时表中，然后执行从表中取出的表达式。
		pRsExecSQL->PutCollect(_variant_t("DW"),_variant_t(dw));
		
		//9/28
		if(pRsArrow->GetRecordCount()>0)
		{
			pRsArrow->MoveFirst();
			//9/28
			while( !pRsArrow->adoEOF )//*3
			{
				var = pRsArrow->GetCollect("arrow_dw");
				strExpression = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
				//将字符串.AND.转换为 AND ，
				strExpression.Replace(".AND.", " AND ");
				strExpression.Replace(".OR.",  " OR ");

				pRsExecSQL->PutFilter((long)adFilterNone);
				pRsExecSQL->PutFilter(_variant_t(strExpression));
				if( pRsExecSQL->RecordCount > 0 )
				{
					break;
				}
				pRsArrow->MoveNext();
			}
		}
		if( pRsArrow->adoEOF || pRsArrow->BOF )
		{
			return false;
		}//*3
		////
		CString strSQLArrow_c,  //字段arrow_c的值（一个表达式）
				strSQLArrow_d;  //字段arrow_d的值（一个表达式）
		double  dTmpArrow_b,
				dTmpArrow_g;
		var = pRsArrow->GetCollect("arrow_c");         //{{取出字段值 //*4
		strSQLArrow_c = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
		strSQLArrow_c.Replace(".AND.", " AND ");
		strSQLArrow_c.Replace(".OR.",  " OR ");
		//取出ARROW_D的值,用SQL语言进行计算。
		var = pRsArrow->GetCollect("arrow_d");
		strSQLArrow_d = (var.vt==VT_NULL||var.vt==VT_EMPTY)?_T(""):(CString)var.bstrVal;
		strSQLArrow_d.Replace(".AND.", " AND ");
		strSQLArrow_d.Replace(".OR.",  " OR ");      
		//arrow_b字段的值
		var = pRsArrow->GetCollect("arrow_b");
		dTmpArrow_b = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
		//arrow_g字段的值
		var = pRsArrow->GetCollect("arrow_g");
		dTmpArrow_g = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;	//}}

		//执行取出的表达式。
		strSQL = "UPDATE ["+m_strExecuteSQLTblName+"] SET D_ARROW="+strSQLArrow_d+", C_ARROW="+strSQLArrow_c+" ";
		m_Pconnection->Execute(_bstr_t(strSQL), NULL, adCmdText);

		//从表ExecuteSQL中取出计算后的结果。
		var = pRsExecSQL->GetCollect("D_ARROW");
		dTmpSum = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
		var = pRsExecSQL->GetCollect("C_ARROW");
		cTmpSum = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;

		//***无环则有字,无字则有环,箭头必有
		//***计算色环,箭头,文字的总面积
		dS_ring = dS_arrow = dS_word = 0.0;	//*4
		if( strRing.CompareNoCase("N") ) //*5
		{
			dS_ring = dLength/5*pi*dw*dTmpArrow_b/1000000;	
		}
		else
		{	
			dS_word = dLength/5*dTmpSum*dTmpArrow_g/1000000; //*5
		}
		dS_arrow = dLength/5*dTmpSum*(cTmpSum+dTmpSum)/1000000;

		//***平台扶梯支吊架无箭头文字 
		if( !strArrow.CompareNoCase("N") )
		{
			dS_arrow = 0.0;
			dS_word = 0.0;
		}

		//
		if( pRsColr_l->RecordCount <= 0)
		{
			return false;
		}
		else
		{
			pRsColr_l->MoveFirst();
		}
		//定位COLR_NAME='"+strFace+"'之后取出COLR_LOST字段的值。 
		strSQL = "COLR_NAME='"+strFace+"' ";
		pRsColr_l->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( !pRsColr_l->adoEOF )
		{
			var = pRsColr_l->GetCollect("colr_lost");
			dFace_l = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
		}
		//重新定位，取出COLR_LOST字段的值。 
		pRsColr_l->MoveFirst();
		strSQL = "COLR_NAME='"+strRing+"' ";
		pRsColr_l->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( !pRsColr_l->adoEOF )
		{
			var = pRsColr_l->GetCollect("colr_lost");
			dRing_l = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
		}
		//重新定位，取出COLR_LOST字段的值。 
		pRsColr_l->MoveFirst();
		strSQL = "COLR_NAME='"+strArrow+"' ";
		pRsColr_l->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( !pRsColr_l->adoEOF )
		{
			var = pRsColr_l->GetCollect("colr_lost");
			dArrow_l = (var.vt==VT_NULL||var.vt==VT_EMPTY)?0.0:var.dblVal;
		}
		//
		if( dw >= 2000 || dw == 0 )
		{
			strRing = "N";
			dRing_l = dS_ring = dS_arrow = 0;
			dS_word = 0.08 * (int)((dev_area + 49) / 50);
		}
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return true;
}


//************************************
//* 模 块 名: c_precol()             *
//* 功    能: 生成保温材料汇总表     *
//*           和保温辅助材料汇总表	 *
//* 上级模块: MAIN_MENU              *
//* 下级模块: c_preast()             *
//* 修改日期:                        *
//************************************
BOOL CTotalTableIPED::c_precol()
{
	//删除表：e_precol中的记录
	_RecordsetPtr e_precolSet;
	_RecordsetPtr e_preexpSet;
	e_preexpSet.CreateInstance(__uuidof(Recordset));
	e_precolSet.CreateInstance(__uuidof(Recordset));
	_RecordsetPtr epreastSet;
	epreastSet.CreateInstance(__uuidof(Recordset));
	CString strSQL;
	m_strPaint_c="col_name";
	m_valPaint_c="col_amount";
	_variant_t key;
	double val;
	CString str;
	int pos=0;
	try
	{
		strSQL="delete from e_precol where enginid='"+EDIBgbl::SelPrjID+"'";	//*1
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText);    //*1
		
		strSQL="select * from e_precol where enginid='"+EDIBgbl::SelPrjID+"'";   //*2
		e_precolSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);

		//表：e_preexp中的 "外全体积"汇总
		strSQL="select trim(外保温材料) as 外保温材料,外全体积 from e_preexp  where enginid='"+EDIBgbl::SelPrjID+"' order by trim(外保温材料)";
		e_preexpSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(e_preexpSet->GetRecordCount()>0)
		{
			TotalTable(e_preexpSet,e_precolSet,"外保温材料","外全体积");		//汇总
		}

		for(;pos<10;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}		//*2

		//表:e_preexp中的 "内保温材料"  汇总
		if(e_preexpSet->State==adStateOpen)
		{
			e_preexpSet->Close();
		}
		//*3
		strSQL="select trim(内保温材料) as 内保温材料,内全体积	from e_preexp   where enginid='"+EDIBgbl::SelPrjID+"' order by trim(内保温材料)";
		e_preexpSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(e_preexpSet->GetRecordCount()>0)
		{
			TotalTable(e_preexpSet,e_precolSet,"内保温材料","内全体积");
		}
		for(;pos<20;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}//*3

		//表:e_preexp中的"保护材料 " 汇总
		if(e_preexpSet->State==adStateOpen)//*4
		{
			e_preexpSet->Close();
		}
		strSQL="select trim(保护材料) as 保护材料,外表面积 from e_preexp where enginid='"+EDIBgbl::SelPrjID+"'";
		strSQL=strSQL+" and 保护厚<=5 order by trim(保护材料)";
		e_preexpSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(e_preexpSet->GetRecordCount()>0)
		{
			TotalTable(e_preexpSet,e_precolSet,"保护材料","外表面积");
		}
		if(e_preexpSet->State==adStateOpen)
		{
			e_preexpSet->Close();//*4
		}
		for(;pos<30;)	//*6
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}
		
		//表：e_preexp中的 "保护材料"汇总
		strSQL="select trim(保护材料) as 保护材料,保护全体积 from e_preexp where enginid='"+EDIBgbl::SelPrjID+"'";
		strSQL=strSQL+" and 保护厚>5 order by trim(保护材料)";
		e_preexpSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if(e_preexpSet->GetRecordCount()>0)
		{
			TotalTable(e_preexpSet,e_precolSet,"保护材料","保护全体积");
		}
		for(;pos<40;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}		//*6
			

		//删除表：e_precol中字段(col_name)值为0 的记录 //*7
		strSQL="delete from e_precol where (col_name is NULL or col_name=' ' )and enginid='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute(_bstr_t(strSQL),NULL,adCmdText); //*7
		
		//根据表：a_yl的值对表：e_precol进行相应的修改
		//修改字段：gen_anount col_unit
		if(e_precolSet->GetRecordCount()>0)//*8
		{
			e_precolSet->MoveFirst();
			a_yl(e_precolSet);
		}
		for(;pos<50;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}//*8
		

		//修改字段:col_size
		if(e_precolSet->GetRecordCount()>0)//*9
		{
			e_precolSet->MoveFirst();
			CString col_name;
			int i,j;
			for(;!e_precolSet->adoEOF;e_precolSet->MoveNext())
			{
				key=e_precolSet->GetCollect(_T("col_name"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					col_name=key.bstrVal;
					col_name=Trim(col_name);
					i=col_name.Find("(",0);
					j=col_name.Find(")",0);
					if(!(i<0||j<0))
					{
						col_name=col_name.Mid(i+1,j-i-1);
						e_precolSet->PutCollect(_T("col_size"),_variant_t(col_name));
						e_precolSet->Update();
					}
				}
			}
		}
		for(;pos<60;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}//*9

		c_preast();

		for(;pos<80;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}
		
		//将表:epreast中的所有记录拷贝到表：e_precol中//*10
		if(e_precolSet->State==adStateOpen)
		{
			e_precolSet->Close();
		}
		strSQL="select * from e_precol where enginid='"+EDIBgbl::SelPrjID+"'";
		e_precolSet->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		epreastSet->Open(_bstr_t(_T("select * from epreast where enginid='"+EDIBgbl::SelPrjID+"'")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);//9/22
		if(epreastSet->GetRecordCount()>0)
		{
			epreastSet->MoveFirst();
			for(;!epreastSet->adoEOF;epreastSet->MoveNext())
			{
				e_precolSet->AddNew();
				key=epreastSet->GetCollect(_T("col_name"));
				if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
				{
					str=key.bstrVal;
				}
				else 
				{
					str="";
				}
				e_precolSet->PutCollect(_T("col_name"),_variant_t(str));
				 
				key=epreastSet->GetCollect(_T("gen_amount"));
				val=(key.vt==VT_NULL||key.vt==VT_EMPTY)?0.0:key.dblVal;
				e_precolSet->PutCollect(_T("gen_amount"),_variant_t(val));

				key=epreastSet->GetCollect(_T("col_size"));
				if(key.vt==VT_EMPTY||key.vt==VT_NULL)
				{
					str="";
				}
				else
				{
					str=key.bstrVal;
				}
				e_precolSet->PutCollect(_T("col_size"),_variant_t(str));

				key=epreastSet->GetCollect(_T("col_unit"));
				if(key.vt==VT_EMPTY||key.vt==VT_NULL)
				{
					str="";
				}
				else 
				{
					str=key.bstrVal;
				}
				e_precolSet->PutCollect(_T("col_unit"),_variant_t(str));

				key=epreastSet->GetCollect(_T("col_mark"));
				if(key.vt==VT_EMPTY||key.vt==VT_NULL)
				{
					str="";
				}
				else 
				{
					str=key.bstrVal;
				}
				e_precolSet->PutCollect(_T("col_mark"),_variant_t(str));

				key=epreastSet->GetCollect(_T("col_amount"));
				val=(key.vt==VT_NULL||key.vt==VT_EMPTY)?0.0:key.dblVal;
				e_precolSet->PutCollect(_T("col_amount"),_variant_t(val));
				e_precolSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));

				e_precolSet->Update();
			}
		
		}
		for(;pos<90;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}
		if(e_precolSet->GetRecordCount()>0)//9/13
		{
			a_config(e_precolSet);//*10
		}

		//对字段：col_num重新编号
		long n=1;
		if(e_precolSet->GetRecordCount()>0)
		{
			e_precolSet->MoveFirst();
		}
		for(;e_precolSet->GetRecordCount()>0&&(!e_precolSet->adoEOF);n++,e_precolSet->MoveNext())
		{
			e_precolSet->PutCollect(_T("col_num"),_variant_t((double)n));
			e_precolSet->Update();
		}
		for(;pos<100;)
		{
			((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(pos++);
		}

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
		return false;
	}

	//9/18
	//对表：pass 进行处理
	_RecordsetPtr pass_set;
	int flagPass=0;
	pass_set.CreateInstance(__uuidof(Recordset));
	pass_set->Open(_bstr_t(_T("select * from pass where enginid='"+EDIBgbl::SelPrjID+"'")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
	if(pass_set->GetRecordCount()>0)
	{
		pass_set->MoveFirst();
		for(;!pass_set->adoEOF;pass_set->MoveNext())
		{
			if(flagPass==0)
			{
				pass_set->PutCollect(_T("pass_mark1"),_T("T"));
			}
			else
			{
				pass_set->PutCollect(_T("pass_mark1"),_T("F"));
			}
			key=pass_set->GetCollect(_T("pass_mod1"));
			if((!(key.vt==VT_NULL||key.vt==VT_EMPTY))&&(flagPass==0))
			{
				str=key.bstrVal;
				flagPass=(str.Compare("C_PRECOL")==0)?1:0;
			}
		}
	}
	//9/18

	if(e_precolSet->GetRecordCount()>0)
	{
//		AfxMessageBox("生成保温材料汇总表成功!");//10/13
	}
	else
	{
		AfxMessageBox("保温材料原始记录为空!");
		((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
		return false;
	}
	((CMainFrame*)::AfxGetMainWnd())->m_wndStatusBar.OnProgress(0);
	return true;

}

//将表：a_config 中的信息加入到表达式:e_precol中
void CTotalTableIPED::a_config(_RecordsetPtr sourceSet)
{
	// test add valve  [2005/07/01]
/*	_RecordsetPtr	pRsPipe_Valve;		//管道与阀门的映射表
	CString			strSQL;				//SQL语句
	long	nResCount = 0;				//管道上增加阀门的个数.
	pRsPipe_Valve.CreateInstance(__uuidof(Recordset));
	strSQL = "SELECT * FROM [Pipe_Valve] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ORDER By valveID ";
	pRsPipe_Valve->Open(_bstr_t(strSQL), m_Pconnection.GetInterfacePtr(),
						adOpenStatic, adLockOptimistic, adCmdText);				
		
	nResCount = pRsPipe_Valve->GetRecordCount();
	
	if (nResCount <= 0)
	{	
		nResCount = 0;
	}
	//*/
	_variant_t key;
	CString val_mat;
	CString str;
	double val_amount;
	_RecordsetPtr a_configSet;
	a_configSet.CreateInstance(__uuidof(Recordset));
	a_configSet->Open(_bstr_t("select * from a_config where enginid='"+EDIBgbl::SelPrjID+"'"),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
	if(a_configSet->GetRecordCount()<=0)
	{
		return;
	}
	key=a_configSet->GetCollect(_T("阀保温材料"));
	if(key.vt==VT_EMPTY||key.vt==VT_NULL)
	{
		val_mat="";
	}
	else 
	{
		val_mat=key.bstrVal;
	}
	sourceSet->AddNew();
	sourceSet->PutCollect(_T("col_name"),_variant_t(val_mat));
	
	key=a_configSet->GetCollect(_T("阀门总数"));
	val_amount=(key.vt==VT_NULL||key.vt==VT_EMPTY)?0.0:key.dblVal;
	val_amount=val_amount*0.015;
//	val_amount+=nResCount;		// test add valve [2005/07/01]
	sourceSet->PutCollect(_T("gen_amount"),_variant_t(val_amount));
	sourceSet->PutCollect(_T("col_amount"),_variant_t(val_amount));
	str="m3";
	sourceSet->PutCollect(_T("col_unit"),_variant_t(str));
	str="阀门保温用";
	sourceSet->PutCollect(_T("col_mark"),_variant_t(str));
	sourceSet->PutCollect(_T("enginid"),_variant_t(EDIBgbl::SelPrjID));
	sourceSet->Update();

}


void CTotalTableIPED::a_yl(_RecordsetPtr sourceSet)
{
	_RecordsetPtr a_ylSet;
	_variant_t key;
	double col_yl;
	double col_amount;
	CString col_unit;

	CString col_name;
	CString strSQL;
	a_ylSet.CreateInstance(__uuidof(Recordset));
	a_ylSet->Open(_bstr_t("select * from a_yl where enginid='"+EDIBgbl::SelPrjID+"'"),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);//9/21
	if(sourceSet->GetRecordCount()>0&&a_ylSet->GetRecordCount()>0)
	{
		sourceSet->MoveFirst();
		for(;!sourceSet->adoEOF;sourceSet->MoveNext())
		{
			key=sourceSet->GetCollect(_T("col_name"));
			if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
			{
				col_name=key.bstrVal;
				strSQL="col_name='"+col_name+"'";
				a_ylSet->MoveFirst();
				a_ylSet->Find(_bstr_t(strSQL),0,adSearchForward);
				if(!a_ylSet->adoEOF)
				{
					key=a_ylSet->GetCollect(_T("col_yl"));
					col_yl=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
					key=sourceSet->GetCollect(_T("col_amount"));
					col_amount=(key.vt==VT_EMPTY||key.vt==VT_NULL)?0.0:key.dblVal;
					col_amount=col_amount*(1+col_yl);
					sourceSet->PutCollect(_T("gen_amount"),_variant_t(col_amount));
					key=a_ylSet->GetCollect(_T("col_unit"));
					if(!(key.vt==VT_EMPTY||key.vt==VT_NULL))
					{
						col_unit=key.bstrVal;
						sourceSet->PutCollect(_T("col_unit"),_variant_t(col_unit));
					}
					sourceSet->Update();					
				}
				else
				{
					//在施工余量库中没有该材料,则施工数量为设计数量.		//2005.4.2	ZSY
					sourceSet->PutCollect(_T("gen_amount"), sourceSet->GetCollect(_T("col_amount")));
					//没有默认的单位.

				}
			
			}
		}
	}

}


//修改pass 表
void CTotalTableIPED::passModiTatal(CString str1, CString str2, CString str3)
{
	_variant_t key;
	CString str;
	_RecordsetPtr pass_set;
	int flagPass=0;
	pass_set.CreateInstance(__uuidof(Recordset));
	pass_set->Open(_bstr_t(_T("select * from pass")),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
	if(pass_set->GetRecordCount()>0)
	{
		pass_set->MoveFirst();
		for(;!pass_set->adoEOF;pass_set->MoveNext())
		{
			if(flagPass==0)
			{
				pass_set->PutCollect(_variant_t(str3),_T("T"));
			}
			else
			{
				pass_set->PutCollect(_variant_t(str3),_T("F"));
			}
			key=pass_set->GetCollect(_variant_t(str1));
			if((!(key.vt==VT_NULL||key.vt==VT_EMPTY))&&(flagPass==0))
			{
				str=key.bstrVal;
				flagPass=(str.Compare(str2)==0)?1:0;
			}
		}
	}	
	if(pass_set->State==adStateOpen)
	{
		pass_set->Close();
	}

}


//------------------------------------------------------------------                
// DATE         :[2005/06/10]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :打开准则库
//				pRsColr_l;	调和漆单位消耗量库
//				pRsColor;	色环箭头油漆颜色准则库
//				pRsArrow;	色环箭头尺寸准则库
//------------------------------------------------------------------
short CTotalTableIPED::OpenCodeDB()
{
	try
	{
		CString	strSQL;				//SQL语句

		//调和漆单位消耗量库
		if ( pRsColr_l == NULL )
		{
			pRsColr_l.CreateInstance(__uuidof(Recordset));
		}
		if ( pRsColr_l->GetState() == adStateOpen )
		{
			pRsColr_l->Close();
		}
		strSQL = "SELECT * FROM [a_colr_l] ";
		pRsColr_l->Open(_bstr_t(strSQL), (IDispatch*)m_pConnectionCODE,
						adOpenStatic, adLockOptimistic, adCmdText);

		//色环箭头油漆材料库。
		if (pRsRing_Paint == NULL)
		{
			pRsRing_Paint.CreateInstance(__uuidof(Recordset));
		}
		if (pRsRing_Paint->State == adStateOpen)
		{
			pRsRing_Paint->Close();
		}
		strSQL = "SELECT * FROM [a_ring_paint]";
		pRsRing_Paint->Open(_variant_t(strSQL), m_pConnectionCODE.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);
		
		//色环箭头油漆颜色准则库
		if ( pRsColor == NULL )
		{
			pRsColor.CreateInstance(__uuidof(Recordset));
		}
		if ( pRsColor->GetState() == adStateOpen )
		{
			pRsColor->Close();
		}
		strSQL = "SELECT * FROM [a_color] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		pRsColor->Open(_bstr_t(strSQL), (IDispatch*)m_Pconnection,
						adOpenStatic, adLockOptimistic, adCmdText);

		//	色环箭头尺寸准则库
		if ( pRsArrow == NULL )
		{
			pRsArrow.CreateInstance(__uuidof(Recordset));
		}
		if ( pRsArrow->GetState() == adStateOpen )
		{
			pRsArrow->Close();
		}
		strSQL = "SELECT * FROM [a_arrow] WHERE EnginID='"+EDIBgbl::SelPrjID+"'";
		pRsArrow->Open(_bstr_t(strSQL), (IDispatch*)m_Pconnection,
						adOpenStatic, adLockOptimistic, adCmdText);
		
	}
	catch (_com_error& e) 
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	return 1;
}

//------------------------------------------------------------------
// DATE         :[2005/11/14]
// Parameter(s) :
// Return       :
// Remark       :创建一个临时表用于计算色环箭头的用量
//------------------------------------------------------------------
BOOL CTotalTableIPED::CreateTempTable(_ConnectionPtr pCon, CString strTblName)
{
	try
	{
		CString strSQL;
		if (pCon == NULL || strTblName.IsEmpty())
		{
			return FALSE;
		}
		if( !tableExists(m_Pconnection, strTblName) )//9/2
		{
			strSQL = "CREATE TABLE ["+strTblName+"] (DW double, D_ARROW double, C_ARROW double)";
			pCon->Execute(_bstr_t(strSQL), NULL, adCmdText);
		}
	}catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}

	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/12/23]
// Parameter(s) :
// Return       :
// Remark       :在一个记录集中，按照指定的字段进行汇总
//------------------------------------------------------------------
BOOL CTotalTableIPED::TotalSelfTable()
{
	try
	{
		_RecordsetPtr pRsSou;
		_RecordsetPtr pRsDes;
		pRsSou.CreateInstance( __uuidof(Recordset) );
		pRsDes.CreateInstance( __uuidof(Recordset) );
		CString strSQL;		// SQL语句
		CString strTempTblName = "p";		//临时表名-
		if ( tableExists( m_Pconnection, strTempTblName ) )
		{
			m_Pconnection->Execute( _bstr_t("DROP TABLE ["+strTempTblName+"] "), NULL, -1 );
		}
		//对分类的油漆记录汇总
		strSQL = " SELECT [col_name],[col_size],[col_unit],[col_mark],sum(col_num) AS [col_num],\
			sum(col_amount) AS [col_amount] INTO "+strTempTblName+" FROM [E_PAICOL] \
			WHERE EnginID='"+EDIBgbl::SelPrjID+"' GROUP BY [col_name],[col_size],[col_unit],[col_mark] \
			ORDER BY [col_name],[col_size],[col_unit],[col_mark]";
		m_Pconnection->Execute( _bstr_t(strSQL), NULL, -1 );
		
		strSQL = " delete from e_paicol where enginid='"+EDIBgbl::SelPrjID+"'";
		m_Pconnection->Execute( _bstr_t(strSQL), NULL, -1 );
		if ( pRsDes->State  == adStateOpen )
		{
			pRsDes->Close();
		}
		pRsDes->Open(_variant_t("select * from [e_paicol] "), m_Pconnection.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText );

		strSQL = " SELECT * FROM ["+strTempTblName+"] ";
		pRsSou->Open(_bstr_t(strSQL),(IDispatch*)m_Pconnection,adOpenStatic,adLockOptimistic,adCmdText);
		if ( pRsSou->adoEOF &&pRsSou->BOF )
		{
			return TRUE;
		}
		// 将统一汇总的数据写入到油漆表中
		for ( ; !pRsSou->adoEOF; pRsSou->MoveNext() )
		{
			pRsDes->AddNew();

			pRsDes->PutCollect( _variant_t("col_name"), pRsSou->GetCollect( _variant_t("col_name")));
			pRsDes->PutCollect( _variant_t("col_size"), pRsSou->GetCollect( _variant_t("col_size")));
			pRsDes->PutCollect( _variant_t("col_unit"), pRsSou->GetCollect( _variant_t("col_unit")));
			pRsDes->PutCollect( _variant_t("col_mark"), pRsSou->GetCollect( _variant_t("col_mark")));
			pRsDes->PutCollect( _variant_t("col_num"),  pRsSou->GetCollect( _variant_t("col_num")));
			pRsDes->PutCollect( _variant_t("col_amount"), pRsSou->GetCollect( _variant_t("col_amount")));

			pRsDes->PutCollect( _variant_t("EnginID"), _variant_t(EDIBgbl::SelPrjID));
			pRsDes->Update();
		}
		
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE;
}