// HeatPreCal.cpp: implementation of the CHeatPreCal class.
//
//////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "AutoIPED.h"
#include "HeatPreCal.h"
#include "ParseExpression.h"
#include "vtot.h"
#include <math.h>
#include <stdlib.h>
#include "AutoIPEDView.h"
#include "EDIBgbl.h"
#include "EditOriginalData.h"

#include "CalcThickBase.h"
#include "uematerial.h"

//打开材料库。输入许用应用，弹性模量。膨胀系数
//extern BOOL InputMaterialParameter(CString strDlgCaption, CString strMaterialName, int TableID);
extern CAutoIPEDApp theApp;
BOOL InputMaterialParameter(CString strDlgCaption, CString strMaterialName, int TableID)
{
	// 使用动态库进行显示
	CUeMaterial MatDlg;

	MatDlg.SetCodeID(1);	// 规范
	MatDlg.SetMaterialTableToOpen( TableID );
	MatDlg.SetDlgCaption(strDlgCaption);	// 对话框标题
	MatDlg.SetShowMaterial(TRUE);		// 显示新旧材料对照表
	MatDlg.SetMaterialConnection( theApp.m_pConMaterial );	// 材料数据库的连接
	MatDlg.SetCurMaterialName( strMaterialName );	// 当前显示的材料
	MatDlg.AddCustomMaterialToDB( NULL );

	return TRUE;
}

//#import"c:\program files\common files\system\ado\msado15.dll"	\
//		no_namespace rename("EOF","adoEOF")

#define GRAVITY_ACCELERATION 9.807
//#define D_MAXTEMP	50000.0		//

#define OPENATABLE(INTERFACE,TABLENAME)						\
		_RecordsetPtr INTERFACE;							\
		OpenAProjectTable(INTERFACE,#TABLENAME);			\


#define OPENASSISTANTPROJECTTABLE(INTERFACE,TABLENAME)						\
		_RecordsetPtr INTERFACE;											\
		OpenAssistantProjectTable(INTERFACE,#TABLENAME);					\

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CHeatPreCal::CHeatPreCal()
{
	m_AssistantConnection = NULL;

	m_ProjectName=_T("");
	ch_cal=2;
}

CHeatPreCal::~CHeatPreCal()
{

}

//------------------------------------------------------------------                
// DATE         :[2005/04/29]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :将当前要计算的原始数据赋给选择的计算方法的类成员变量.
//------------------------------------------------------------------
void CHeatPreCal::InitOriginalToMethod()
{
	
#define	INITDATA(CLASSN)		CLASSN::D0 = D0;\
	CLASSN::in_mat		= in_mat;\
	CLASSN::out_mat		= out_mat;\
	CLASSN::pro_mat		= pro_mat;\
	CLASSN::pro_name	= pro_name;\
	CLASSN::in_a0		= in_a0;\
	CLASSN::in_a1		= in_a1;\
	CLASSN::in_a2		= in_a2;\
	CLASSN::in_ratio	= in_ratio;\
	CLASSN::out_a0		= out_a0;\
	CLASSN::out_a1		= out_a1;\
	CLASSN::out_a2		= out_a2;\
	CLASSN::out_ratio	= out_ratio;\
	CLASSN::ts			= temp_ts;\
	CLASSN::tb			= temp_tb;\
	CLASSN::ta			= Temp_env;\
	CLASSN::t			= Temp_pip;\
	CLASSN::Mhours		= Mhours;\
	CLASSN::Yong		= Yong;\
	CLASSN::heat_price	= heat_price;\
	CLASSN::in_price	= in_price;\
	CLASSN::out_price	= out_price;\
	CLASSN::pro_price	= pro_price;\
	CLASSN::S			= S;\
	CLASSN::in_dens		= in_dens;\
	CLASSN::out_dens	= out_dens;\
	CLASSN::pro_dens	= pro_dens;\
	CLASSN::speed_wind	= speed_wind;\
	CLASSN::m_nID		= m_nID;\
	CLASSN::co_alf		= co_alf;\
	CLASSN::ALPHA		= co_alf;\
	CLASSN::strPlace	= strPlace;\
	CLASSN::m_HotRatio	= m_HotRatio;\
	CLASSN::m_CoolRatio	= m_CoolRatio;\
	CLASSN::MaxTs		= MaxTs;\
	CLASSN::hour_work	= hour_work;\
	CLASSN::hedu		= hedu;\
	CLASSN::m_dMaxD0	= m_dMaxD0;\
	CLASSN::distan		= distan;\
	CLASSN::in_tmax		= in_tmax;\
	CLASSN::m_MaxHeatDensityRatio=m_HotRatio;\
	CLASSN::m_thick1Start		= m_thick1Start;\
	CLASSN::m_thick1Max			= m_thick1Max;\
	CLASSN::m_thick1Increment	= m_thick1Increment;\
	CLASSN::m_thick2Start		= m_thick2Start;\
	CLASSN::m_thick2Max			= m_thick2Max;\
	CLASSN::m_thick2Increment	= m_thick2Increment;\
	CLASSN::m_pRsHeatLoss		= m_pRsHeatLoss;\
	CLASSN::m_pRsHeat_alfa		= m_pRsHeat_alfa;\
	CLASSN::m_pRsCongeal		= m_pRsCongeal;\
	CLASSN::m_pRsSubterranean	= m_pRsSubterranean;\
	CLASSN::m_pRsFaceResistance	= m_pRsFaceResistance;\
	CLASSN::m_pRsLSo			= m_pRsLSo;\
	CLASSN::thick				= thick_2;\
	CLASSN::thick1				= thick_1;\
	CLASSN::thick2				= thick_2;\
	CLASSN::pi_thi				= pi_thi;\
	CLASSN::dQ					= 0.0;\
	CLASSN::nMethodIndex		= nMethodIndex;\
	CLASSN::nAlphaIndex			= m_nIndexAlpha;\
	CLASSN::nTaIndex			= nTaIndex;\
	CLASSN::B					= B;\
	CLASSN::strExceptionInfo	= "";


/*
//防冻结保温计算时，新增加的一些参数.
#define INITCONGEALDATA(CLASSN)			\
		CLASSN::Kr		=	Kr;\
		CLASSN::taofr	=	taofr;\
		CLASSN::tfr		=	tfr;\
		CLASSN::V		=	V;\
		CLASSN::ro		=	ro;\
		CLASSN::C		=	C;\
		CLASSN::Vp		=	Vp;\
		CLASSN::rop		=	rop;\
		CLASSN::Cp		=	Cp;\
		CLASSN::Hfr		=	Hfr;\
		CLASSN::dFlux	=	dFlux;
*/

}

//根据厚度计算外表面温度.
void CHeatPreCal::temp_plain_one(double &thick_t, double &ts)
{
	double ts1,condu_r_t,condu_out_t;

	ts=(Temp_pip-Temp_env)/2.0+Temp_env;

	while(TRUE)
	{
		condu_r_t=in_a0+in_a1*(ts+Temp_pip)/2.0+in_a2*(ts+Temp_pip)*(ts+Temp_pip)/4.0;

		condu_out_t=co_alf;
		ts1=thick_t*Temp_env/1000.0/condu_r_t+Temp_pip/condu_out_t;
		ts1=ts1/(thick_t/1000.0/condu_r_t+1.0/condu_out_t);
		
		if(fabs(ts-ts1)<0.1)
			break;
		ts=ts1;
	}

	return;
}



void CHeatPreCal::temp_pip_one(double &thick_t, double &ts)
{
	double ts1,condu_r_t,condu_out_t,D1_t;

	ts=(Temp_pip-Temp_env)/2.0+Temp_env;
	D1_t=D0+thick_t*2.0;
	while(TRUE)
	{
//		* the following line is corrected in next step
		condu_out_t=co_alf;
		condu_r_t=in_a0+in_a1*(ts+Temp_pip)/2.0+in_a2*(ts+Temp_pip)*(ts+Temp_pip)/4.0;
		ts1=1.0/condu_r_t*log(D1_t/D0)*Temp_env+2000.0*Temp_pip/condu_out_t/D1_t;
		ts1=ts1/(log(D1_t/D0)/condu_r_t+2000.0/condu_out_t/D1_t);
		if (fabs(ts-ts1)<0.1)
			break;
		ts=ts1;
	}
	return;
}

void CHeatPreCal::temp_plain_com(double &thick1_t, double &thick2_t, double &ts, double &tb)
{
	double tb1,ts1,condu_out_t,condu_r1_t;
	double condu_r2_t;

	ts=(Temp_pip-Temp_env)/3.0+Temp_env;
	tb=(Temp_pip-Temp_env)/3.0*2.0+Temp_env;
	while(TRUE)
	{
	//	* the following line is corrected in next step
		condu_out_t=co_alf;
		condu_r1_t=in_a0+in_a1*(Temp_pip+tb)/2.0+in_a2*(Temp_pip+tb)*(Temp_pip+tb)/4.0;
		condu_r2_t=out_a0+out_a1*(tb+ts)/2.0+out_a2*(tb+ts)*(tb+ts)/4.0;
		
		ts1=thick1_t*Temp_env/1000.0/condu_r1_t+thick2_t*Temp_env/1000.0/condu_r2_t+Temp_pip/condu_out_t;
		ts1=ts1/(thick1_t/1000.0/condu_r1_t+thick2_t/1000.0/condu_r2_t+1.0/condu_out_t);
		tb1=thick1_t*Temp_env/1000.0/condu_r1_t+thick2_t*Temp_pip/1000.0/condu_r2_t+Temp_pip/condu_out_t;
		tb1=tb1/(thick1_t/1000.0/condu_r1_t+thick2_t/1000.0/condu_r2_t+1.0/condu_out_t);
		
		if ((fabs(ts-ts1)<0.1)&&(fabs(tb-tb1)<0.1))
			break;
		ts=ts1;
		tb=tb1;
	}
	return;
}

void CHeatPreCal::temp_pip_com(double &thick1_t, double &thick2_t, double &ts, double &tb)
{
	double tb1,ts1,condu_r1_t,condu_out_t,D1_t,D2_t;
	double condu_r2_t;

	ts=(Temp_pip-Temp_env)/3.0+Temp_env;
	tb=(Temp_pip-Temp_env)/3.0*2.0+Temp_env;
	D1_t=D0+2.0*thick1_t;
	D2_t=D1_t+2.0*thick2_t;

	while(TRUE)
	{
//		* the following line is corrected in next step
		condu_out_t=co_alf;
		condu_r1_t=in_a0+in_a1*(Temp_pip+tb)/2.0+in_a2*(Temp_pip+tb)*(Temp_pip+tb)/4.0;
		condu_r2_t=out_a0+out_a1*(tb+ts)/2.0+out_a2*(tb+ts)*(tb+ts)/4.0;
		ts1=Temp_env*log(D1_t/D0)/condu_r1_t+Temp_env*log(D2_t/D1_t)/condu_r2_t+2000.0*Temp_pip/condu_out_t/D2_t;
		ts1=ts1/(log(D1_t/D0)/condu_r1_t+log(D2_t/D1_t)/condu_r2_t+2000.0/condu_out_t/D2_t);
		tb1=Temp_env*log(D1_t/D0)/condu_r1_t+Temp_pip*log(D2_t/D1_t)/condu_r2_t+2000.0*Temp_pip/condu_out_t/D2_t;
		tb1=tb1/(log(D1_t/D0)/condu_r1_t+log(D2_t/D1_t)/condu_r2_t+2000.0/condu_out_t/D2_t);
		if ((fabs(ts-ts1)<0.1)&&(fabs(tb-tb1)<0.1))
			break;
		
		ts=ts1;
		tb=tb1;
	}
	return;


}


//////////////////////////////////////////
//模 块 名: support.prg            //
//功    能: 计算支吊架间距         //
//上级模块: C_ANALYS               //
//下级模块:                        //
//引 用 库: (6)  a_pipe,           //
//          (8)a_elas			   //
//修改日期:                        //
///////////////////////////////////////////
void CHeatPreCal::SUPPORT(double &dw, double &s, double &temp, double &wei1,double &wei2,double &lmax, 
						  CString &v, CString &cl,BOOL &pg,CString &st, double &t0)
{
	//管径,壁厚,温度, 无水单重,有水单重, 间距, 
	//卷册号,材质,汽水标识,冷态温度,环境温度
/*	double d=0,w=0,i=0;
	double elastic;
	double l1,l2;
	double stres_t0=0.0,stres_t1=0.0,stress=0.0,stress3=0.0;
	double circle=0,l3,maxl12;
	CString k,mk;
	CString TempStr;
	_variant_t TempValue;
	CString strDlgCaption;
	int		TableID;
	 
	cl.TrimLeft();
	cl.TrimRight();

	SetMapValue(CString(_T("S")),&S);
	//将记录集IRecPipe改为成员的。每次计算只打开一次。
	//目的：提高速度。		ZSY	2005/4/8
//	OPENASSISTANTPROJECTTABLE(IRecPipe,A_PIPE);

/*	try
	{
		IRecPipe->MoveFirst();
		while(!IRecPipe->adoEOF)
		{
			TempValue=IRecPipe->GetCollect(_variant_t("PIPE_DW"));
			tempVtf = vtof(TempValue);
			if(fabs(tempVtf - dw) > 1E-6)
			{
				IRecPipe->MoveNext();
				continue;
			}

			TempValue=IRecPipe->GetCollect(_variant_t("PIPE_S"));
			tempVtf = vtof(TempValue);
			if(fabs(tempVtf - s) > 1E-6)
			{
				IRecPipe->MoveNext();
				continue;
			}
			
			GetTbValue(IRecPipe,_variant_t("PIPE_I"),i);
			GetTbValue(IRecPipe,_variant_t("PIPE_W"),w);

			break;
		}

	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	if(IRecPipe->adoEOF)
	{
		d=(dw-2*s);
		w=pi*(dw*dw*dw*dw/10000-d*d*d*d/10000)/(dw/10*32);
		i=w*dw/20;

	}

*/
	//管道断面惯性矩(PIPE_I) 厘米^4
	//管道断面抗弯矩(PIPE_W) 厘米^3
	//从管道参数库中查找对应外径和壁厚的记录。
/*	if( 0 == GetPipeWeithAndWaterWeith(IRecPipe,dw,s,"PIPE_I",i,"PIPE_W",w) )
	{
		d=(dw-2*s);
		//管道断面抗弯矩(PIPE_W) 厘米^3
		w=pi*(dw*dw*dw*dw/10000-d*d*d*d/10000)/(dw/10*32);
		//管道断面惯性矩(PIPE_I) 厘米^4
		i=w*dw/20;
	}
*/
/*

	//AutoIPED7.0已经采用新的方法计算许用应力、弹性模量
	//即利用动态库函数从 material.mdb中查找许用应力/弹性模量，查出的单位是MPa,kN/mm2
	stress3=CCalcDll.GetMaterialSIGMAt(cl,(float)t0,(float)s)/GRAVITY_ACCELERATION;

	CString strDlgCaption, strMaterialName, *pStrArr;
	int		TableID, MatNum;
	if( stress3 < 0 )
	{
		if( FindMaterialNameOldOrNew(cl, pStrArr, MatNum) ) //查找新旧材料对照表.
		{
			strMaterialName = pStrArr[0];
			delete[] pStrArr;
		}
		else
		{
			strMaterialName = cl;
		}
		TableID = 0;
		strDlgCaption.Format(_T("请输入 %s 在环境温度下的许用应力(MPa):"),strMaterialName);
		//输入数据
		InputMaterialParameter(strDlgCaption,  strMaterialName,  TableID);
		//再找一次
		stress3=CCalcDll.GetMaterialSIGMAt(cl,(float)t0,(float)s)/GRAVITY_ACCELERATION;
	}

	stress=CCalcDll.GetMaterialSIGMAt(cl,(float)temp,(float)s)/GRAVITY_ACCELERATION;
	if( stress < 0 )
	{
		if( FindMaterialNameOldOrNew(cl, pStrArr, MatNum) ) //查找新旧材料对照表.
		{
			strMaterialName = pStrArr[0];
			delete[] pStrArr;
		}
		else
		{
			strMaterialName = cl;
		}

		TableID = 0;
		strDlgCaption.Format(_T("请输入 %s 在 %f  ℃下的许用应力(MPa):"),strMaterialName,temp);
//		strMaterialName = cl;
		//输入数据
		InputMaterialParameter(strDlgCaption,  strMaterialName,  TableID);
		stress=CCalcDll.GetMaterialSIGMAt(cl,(float)temp,(float)s)/GRAVITY_ACCELERATION;
	}
	//
	elastic=CCalcDll.GetMaterialEt(cl,(float)temp)/GRAVITY_ACCELERATION*100000;
	if( elastic < 0 )
	{
		if( FindMaterialNameOldOrNew(cl, pStrArr, MatNum) ) //查找新旧材料对照表.
		{
			strMaterialName = pStrArr[0];
			delete[] pStrArr;
		}
		else
		{
			strMaterialName = cl;
		}

		TableID = 1;
		strDlgCaption.Format(_T("请输入 %s 在 %f  ℃下的弹性模量(kN/mm2)"),strMaterialName,temp);
//		strMaterialName = cl;
		//输入数据
		InputMaterialParameter(strDlgCaption,  strMaterialName,  TableID);
		elastic=CCalcDll.GetMaterialEt(cl,(float)temp)/GRAVITY_ACCELERATION*100000;
	}

//	***应力   stress
//	***弹性模量   elastic
	BOOL IsFind=FALSE;
	//当没有数据，由用户输入时才循环一次
	while( 1 )
	{
		//将记录集IRec改为成员的。每次计算只打开一次。
		//目的：提高速度。		ZSY	2005/4/8
	//	OPENASSISTANTPROJECTTABLE(IRec,A_C09);
		m_pRs_a_c09->MoveFirst();
		if(LocateFor(m_pRs_a_c09,_variant_t("PIPE_MAT"),CFoxBase::EQUAL,_variant_t(cl)))
		{
			try
			{
				GetTbValue(m_pRs_a_c09,_variant_t("PIPE_C"),circle);
				break;
			}
			catch(_com_error &e)
			{
				TempStr.Format(_T("在\"环向焊缝应力系数库(A_C09)\"的PIPE_MAT=%s时读取数据有错"),cl);

				Exception::SetAdditiveInfo(TempStr);

				ReportExceptionErrorLV2(e);

				throw;
			}
		}
		else
		{	// 如果上面找不到可能是因为材料的新旧名称
			//
			CString *pMatNameArr;
			int Num;
			if(FindMaterialNameOldOrNew(cl,pMatNameArr,Num))
			{
				while(Num>0)
				{
					Num--;
					strMaterialName = pMatNameArr[Num]; //zsy
					if(LocateFor(m_pRs_a_c09,_variant_t("PIPE_MAT"),CFoxBase::EQUAL,_variant_t(pMatNameArr[Num])))
					{
						try
						{
							GetTbValue(m_pRs_a_c09,_variant_t("PIPE_C"),circle);
							IsFind=TRUE;
							break;
						}
						catch(_com_error &e)
						{
							TempStr.Format(_T("在\"环向焊缝应力系数库(A_C09)\"的PIPE_MAT=%s时读取数据有错"),cl);
							Exception::SetAdditiveInfo(TempStr);
							ReportExceptionErrorLV2(e);
							throw;
						}
					}
				}
				delete[] pMatNameArr;
			}
			if(IsFind==FALSE)
			{
				if( strMaterialName.IsEmpty() )
				{
					strMaterialName = cl;
				}
				TableID = 3;
				strDlgCaption.Format(_T("请输入 %s 的环向焊缝应力系数:"),strMaterialName);
				//输入数据
				InputMaterialParameter(strDlgCaption,  strMaterialName,  TableID);	
				IsFind = TRUE;
			}
			else
			{
				break;
			}
		}
	} //while(1)

	if(st==_T("水管"))
	{

		if(pg==_T("Y") || pg==_T("y"))
		{
			l1=sqrt(w*circle*stress/wei2)*2.0;
		}
		else
		{
			l1=sqrt(w*circle*stress/wei2)*2.24;
		}
		l2=0.0241*pow((elastic*i/wei2),(1.0/3.0));
		lmax=min(l1,l2);
	}
	else if(st==_T("汽管"))
	{
		if(pg==_T("Y") || pg==_T("y"))
		{
			l1=sqrt(w*circle*stress/wei1)*2.0;
			l3=sqrt(w*circle*stress3/wei2)*2.0;
		}
		else
		{
			l1=sqrt(w*circle*stress/wei1)*2.24;
			l3=sqrt(w*circle*stress3/wei2)*2.24; 
		}
		l2=0.0241*pow((elastic*i/wei1),(1.0/3.0));
		maxl12=min(l1,l2);
		lmax=min(maxl12,l3);
	}
	else
	{
	    lmax=0;
	}
*/
}	

//////////////////////////////////////
//economic Pip thickness calculation//
//////////////////////////////////////

void CHeatPreCal::pip_com(double &thick1_resu, double &thick2_resu, 
						  double &tb_resu, double &ts_resu)
{
	double condu1_r,condu2_r,condu_out,cost_q,cost_s,cost_tot,cost_min,thick1,thick2,tb,ts;
	double D1,D2;


//	double MaxTb = in_tmax*m_HotRatio;	//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段控制界面温度值.

	double nYearVal, nSeasonVal, Q;
	double QMax;		//允许最大散热密度
	bool   flg=true;	//进行散热密度的判断

	thick1_resu=0;
	thick2_resu=0;
	cost_min=0;
	
/*  //应该没有温度的限制
	if (Temp_pip<350.0)
	{
		return;
	}
*/
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		if( !GetHeatLoss(m_pRsHeatLoss, Temp_pip, nYearVal, nSeasonVal, QMax) )
		{		
		}
	}
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表thicknessRegular获得
	//
	for(thick1=m_thick1Start;thick1<=m_thick1Max;thick1+=m_thick1Increment)
	{
		for(thick2=m_thick2Start;thick2<=m_thick2Max;thick2+=m_thick2Increment)
		{
		//	do temp_pip_com with thick1,thick2,ts,tb 
			temp_pip_com(thick1,thick2,ts,tb);
			condu1_r=in_a0+in_a1*(Temp_pip+tb)/2.0+in_a2*(Temp_pip+tb)*(Temp_pip+tb)/4.0;
			condu2_r=out_a0+out_a1*(tb+ts)/2.0+out_a2*(tb+ts)*(tb+ts)/4.0;
			condu_out=co_alf;
			D1=2*thick1+D0;
			D2=2*thick2+D1;

			cost_q=7.2*pi*Mhours*heat_price*Yong*fabs(Temp_pip-Temp_env)*1e-6
					/(1.0/condu1_r*log(D1/D0)+1.0/condu2_r*log(D2/D1)+2000.0/condu_out/D2);

			cost_s=(pi/4.0*(D1*D1-D0*D0)*in_price*1e-6+pi/4.0*(D2*D2-D1*D1)*out_price*1e-6
					+pi*D2*pro_price*1e-3)*S;

			cost_tot=cost_q+cost_s;

			//tb<350)
			if((ts<MaxTs && tb<MaxTb) && (fabs(cost_min)<1E-6 || cost_min>cost_tot))
			{
				flg = true;
				//选项中选择的判断散热密度,就进行比较.
				if(bIsHeatLoss)
				{
					//求得本次厚度的散热密度
					//公式:
					//(介质温度-环境温度)/((外保温层外径/2000.0)*(1.0/内保温材料导热率*ln(内保温层外径/管道外径) + 1.0/外保温材料导热率*ln(外保温层外径/内保温层外径) ) + 1/传热系数)
					Q = (Temp_pip-Temp_env) / ( (D2/2000.0) * (1.0/condu1_r*log(D1/D0)+1.0/condu2_r*log(D2/D1)) + 1.0/co_alf );

					if(fabs(QMax) <= 1E-6 || Q <= QMax)
					{	//当前厚度满足允许最大散热规则.
						flg = true;
					}
					else
					{	
						flg = false;
					}
				}
				//满足条件.
				if( flg )
				{
					cost_min=cost_tot;
					thick1_resu=thick1;
					thick2_resu=thick2;
					ts_resu=ts;
					tb_resu=tb;//
				}
			}
		}
	}
	return;

}


////////////////////////////////////////////////////
// economic plain thickness for complex calculation//
/////////////////////////////////////////////////////

void CHeatPreCal::plain_com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double condu1_r,condu2_r,condu_out,cost_q,cost_s,cost_tot,cost_min,thick1,thick2,tb,ts;
	
//	double MaxTb = in_tmax*m_HotRatio;//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	double MaxTs;	//外表面允许最大温度 
	//对于防烫伤保温，保温结构外表面温度不应超过60摄氏度。 
	if(	Temp_env <= 27 )
	{
		//环境温度不高于27摄氏度时，设备和管道保温结构外表面温度不应超过50摄氏度；
		MaxTs = 50;
	}
	else
	{
		//环境温度高于27摄氏度时，保温结构外表面温度可比环境温度高25x摄氏度。
		MaxTs = Temp_env + 25;
	}

	double nYearVal, nSeasonVal,
		   Q;			//保温结构外表面散热密度
	double QMax;		//允许最大散热密度
	bool   flg=true;	//进行散热密度的判断

	thick1_resu=0;
	thick2_resu=0;
	cost_min=0;
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		if( !GetHeatLoss(m_pRsHeatLoss, Temp_pip, nYearVal, nSeasonVal, QMax) )
		{		
		}
	}
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表thicknessRegular获得
	for(thick1=m_thick1Start;thick1<=m_thick1Max;thick1+=m_thick1Increment)
	{
		for(thick2=m_thick2Start;thick2<=m_thick2Max;thick2+=m_thick2Increment)
		{
		//	do temp_plain_com with thick1,thick2,ts,tb 
			temp_plain_com(thick1,thick2,ts,tb);
			condu1_r=in_a0+in_a1*(Temp_pip+tb)/2.0+in_a2*(Temp_pip+tb)*(Temp_pip+tb)/4.0;
			condu2_r=out_a0+out_a1*(tb+ts)/2.0+out_a2*(tb+ts)*(tb+ts)/4.0;
			condu_out=co_alf;

			cost_q=3.6*Mhours*heat_price*Yong*fabs(Temp_pip-Temp_env)*1e-6
				/(thick1/1000.0/condu1_r+thick2/1000.0/condu2_r+1.0/condu_out);

			cost_s=(thick1/1000.0*in_price+thick2/1000.0*out_price+pro_price)*S;
			cost_tot=cost_q+cost_s;
			if((ts<MaxTs && tb<MaxTb))	//ts<50 && tb<350
			{
				if(fabs(cost_min)<1E-6 )
				{//第一次
					flg = true;
					//选项中选择的判断散热密度,就进行比较.
					if(bIsHeatLoss)
					{
						//求得本次厚度的散热密度
						//公式:
						//(介质温度-环境温度)/(内保温厚/1000/内保温材料导热率 + 外保温厚/1000/外保温材料导热率 + 1/传热系数)
						Q = (Temp_pip-Temp_env)/(thick1/1000.0/condu1_r+thick2/1000.0/condu2_r+1.0/co_alf);
						if(fabs(QMax) <= 1E-6 || Q <= QMax)
						{	//当前厚度满足允许最大散热规则.
							flg = true;
						}
						else
						{	
							flg = false;
						}
					}
					//满足条件.
					if( flg )
					{
						cost_min=cost_tot;
						thick1_resu=thick1;
						thick2_resu=thick2;
						ts_resu=ts;
						tb_resu=tb;
					}
			//		InsertToReckonCost(ts,tb,thick1,thick2,cost_tot,true);
				}
				else
				{//不是第一次
					if(cost_min>cost_tot)
					{
						flg = true;
						//选选项中选择的判断散热密度,就进行比较.
						if(bIsHeatLoss)
						{
							//求得本次厚度的散热密度
							//公式:
							//(介质温度-环境温度)/(内保温厚/1000/内保温材料导热率 + 外保温厚/1000/外保温材料导热率 + 1/传热系数)
							Q = (Temp_pip-Temp_env)/(thick1/1000.0/condu1_r+thick2/1000.0/condu2_r+1.0/co_alf);
							if(fabs(QMax) <= 1E-6 || Q <= QMax)
							{	//当前厚度满足允许最大散热规则.
								flg = true;
							}
							else
							{
								flg = false;
							}
						}
						//满足条件.
						if( flg )
						{
							//前次费用大于本次费用
							cost_min=cost_tot;
							thick1_resu=thick1;
							thick2_resu=thick2;
							ts_resu=ts;
							tb_resu=tb;
						}
					//	InsertToReckonCost(ts,tb,thick1,thick2,cost_tot,false);
					}
					else
					{//前次费用小于本次费用
					}
				}
			}
		}
	}
	return;
}


//////////////////////////////////////
//economic Pip thickness calculation//
//////////////////////////////////////

void CHeatPreCal::pip_one(double &thick_resu, double &ts_resu)
{
	double condu_r,condu_out,cost_q,cost_s,cost_tot,cost_min,thick,ts;
	double D1;
	double nYearVal, nSeasonVal, Q;
	double QMax;		//允许最大散热密度
	bool   flg=true;	//进行散热密度的判断

	thick_resu=0;
	ts_resu=0;
	cost_min=0;
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		if( !GetHeatLoss(m_pRsHeatLoss, Temp_pip, nYearVal, nSeasonVal, QMax) )
		{		
		}
	}
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表thicknessRegular获得
	//
	for(thick=m_thick2Start;thick<=m_thick2Max;thick+=m_thick2Increment)
	{
	//	do temp_pip_one with thick,ts
		temp_pip_one(thick,ts);
		condu_r=in_a0+in_a1*(ts+Temp_pip)/2.0+in_a2*(ts+Temp_pip)*(ts+Temp_pip)/4.0;
		condu_out=co_alf;
		D1=2.0*thick+D0;

		cost_q=7.2*pi*Mhours*heat_price*Yong*fabs(Temp_pip-Temp_env)*1e-6
				/(1.0/condu_r*log(D1/D0)+2000.0/condu_out/D1);
 
		cost_s=(pi/4.0*(D1*D1-D0*D0)*in_price*1e-6+pi*D1*pro_price*1e-3)*S;
		cost_tot=cost_q+cost_s;
		if(ts<MaxTs && (cost_tot<cost_min || fabs(cost_min)<1E-6))
		{
			flg = true;
			//选项中选择的判断散热密度,就进行比较.
			if(bIsHeatLoss)
			{
				//求得本次厚度的散热密度
				//公式:
				//(介质温度-环境温度)/((保温层外径/2000.0/材料导热率*ln(保温层外径/管道外径)) + 1/传热系数)
				Q = (Temp_pip-Temp_env)/(D1/2000.0/condu_r*log(D1/D0) + 1.0/co_alf);
				if(fabs(QMax) <= 1E-6 || Q <= QMax)
				{	//当前厚度满足允许最大散热规则.
					flg = true;
				}
				else
				{	
					flg = false;
				}
			}
			//满足条件.
			if( flg )
			{
				cost_min=cost_tot;
				thick_resu=thick;
				ts_resu=ts;
			}
		}
	}
	return;
}


/////////////////////////////////////////
//economic plain thickness calculation //
//main program                         //
////////////////////////////////////////
void CHeatPreCal::plain_one(double &thick_resu, double &ts_resu)
{
	double condu_r,condu_out,cost_q,cost_s,cost_tot,cost_min,thick,ts;
	double nYearVal, nSeasonVal, Q ;	
	double QMax;		//允许最大散热密度
	bool   flg=true;	//进行散热密度的判断

	thick_resu=0;
	ts_resu=0;
	cost_min=0;
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		if( !GetHeatLoss(m_pRsHeatLoss, Temp_pip, nYearVal, nSeasonVal, QMax) )
		{ 		
		}
	}
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表thicknessRegular获得
	//
	for(thick=m_thick2Start;thick<=m_thick2Max;thick+=m_thick2Increment)
	{
		//	do temp_plain_one with thick,ts
		temp_plain_one(thick,ts);
		condu_r=in_a0+in_a1*(ts+Temp_pip)/2.0+in_a2*(ts+Temp_pip)*(ts+Temp_pip)/4.0;
		condu_out=co_alf;

		cost_q=3.6*Mhours*heat_price*Yong*fabs(Temp_pip-Temp_env)*1e-6
				/((thick/1000.0/condu_r)+1.0/condu_out);

		cost_s=(thick/1000*in_price+pro_price)*S;
		cost_tot=cost_q+cost_s;
	
		if((cost_tot<cost_min || fabs(cost_min)<1E-6) && ts<MaxTs)
		{
			flg = true;
			//选项中选择的判断散热密度,就进行比较.
			if(bIsHeatLoss)
			{
				//求得本次厚度的散热密度
				//公式:
				//(介质温度-环境温度)/(当前保温厚/(1000*材料导热率) + 1/传热系数)
				Q = (Temp_pip-Temp_env) / (thick / (1000*condu_r) + 1.0 / co_alf); 
				if(fabs(QMax) <= 1E-6 || Q <= QMax)
				{	//当前厚度满足允许最大散热规则.
					flg = true;
				}
				else
				{	
					flg = false;
				}
			}
			//满足条件.
			if( flg )
			{
				cost_min=cost_tot;
				thick_resu=thick;
				ts_resu=ts;
			}
		}
	}
	return;
}

//------------------------------------保温计算------------------------------------------
//////////////////////////////////////////////////////////////
//                        C_ANALYS                         //
/////////////////////////////////////////////////////////////

/////////////////////////////////////
//模 块 名: c_analys.prg           //
//功    能: 保温结构经济厚度计算   //
//上级模块: MAIN_MENU              //
//下级模块:  *                     //
//          SUPPORT*			   //
//引 用 库: (1)pre_calc,           //
//          (3)a_mat,              //
//          (6)a_pipe,             //
//          (8)a_stress,a_elas     //
//修改日期:                        //
/////////////////////////////////////

void CHeatPreCal::C_ANALYS()
{
	_variant_t TempValue;
	double thick_1,thick_2,temp_ts,temp_ts1,temp_tb;
	double pre_v1,pre_v2,pro_v,pro_wei;
	double pro_conduct;
	double pro_thi,pi_thi;
	double pre_wei1,pre_wei2;
	double d1,d2,d3,p_wei,w_wei,no_w,with_w,face_area;
	double a_config_hour;
	double dAmount;	//数量.
	double tempVtf;
	long   rec;
	long   ERR;
	int    IsStop;
	HRESULT hr;
	BOOL	pi_pg=FALSE;	//压力标识，大于3.92MPa为TRUE，否则为FALSE
	CString tempVts;
	CString pi_mat,mark,vol,steam;
	CString pi_site;
	CString last_mod;
	CString unit;
	CString strPlace;		//安装地点。（室内或室外）
	CString TempStr;		//临时的，输出到屏幕的字符串
	CString cPipeType;		//管道或设备的类型,为卷册号最后一个字符.
	CString strMedium;		//介质名称
	int		nMethodIndex=0;	//计算方法的索引值
	int		nTaIndex=0;		//环境温度的取值索引
	bool	bNoCalInThi;	//是否计算内保温层厚度
	bool	bNoCalPreThi;	//是否计算外保温层厚度
	double  Temp_dew_point = 23;	// 露点温度
	double  Temp_coefficient = 1;	// 不同的规范增加的不同系数
	double  dUnitLoss;
	double  dAreaLoss;
	BOOL	bIsReCalc = FALSE;		// 如果是保冷计算，是否需要重新用表面温度法校验外表面温度的值，和厚度
	
	ch_cal=GetChCal();
	if(ch_cal==4)
	{
		return;
	}

//	*注册
	OPENATABLE(IRecordset,PASS);

	if(!LocateFor(IRecordset,_variant_t("PASS_MOD1"), 
			CFoxBase::EQUAL,_variant_t("C_ANALYS")))
	{
		TempStr.Format("无法在PASS表中的\"PASS_MOD1\"字段找到C_ANALYS \n有可能数据表为空或未添加数据");
		ExceptionInfo(TempStr);
		return;
	}
	try
	{
		GetTbValue(IRecordset,_variant_t("PASS_LAST1"),last_mod);
	}
	catch(_com_error &e)
	{
		TempStr.Format(_T("在\"PASS表\"的PASS_MOD1=C_ANALYS时读取PASS_LAST1数据有错"));
		Exception::SetAdditiveInfo(TempStr);
		ReportExceptionErrorLV2(e);
		throw;
	}

	if(!LocateFor(IRecordset,_variant_t("PASS_MOD1"),
			CFoxBase::EQUAL,_variant_t(last_mod)))
	{
		TempStr.Format("无法在PASS表中的\"PASS_MOD1\"字段找到%s \r\n有可能数据表为空或未添加数据",last_mod);
		ExceptionInfo(TempStr);

		return;
	}
	TempValue=IRecordset->GetCollect(_variant_t("PASS_MARK1"));
	if(TempValue!=_variant_t("T"))
	{
		this->MessageBox(_T("不能计算, 必须先进行原始数据编辑!"));
		IRecordset->Close();
		return;
	}
	
	IRecordset->Close();
	CString sql;
    if (ch_cal==1)
	{
		try
		{
			//将C_PASS字段全部赋'Y'
			sql = "UPDATE [PRE_CALC] SET C_PASS='Y' WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
			theApp.m_pConnection->Execute(_bstr_t(sql), NULL, adCmdText );

			OPENATABLE(IRecordset2,PRE_CALC);
			if(IRecordset2->adoEOF && IRecordset2->BOF)
			{
				ExceptionInfo(_T("原始数据表为空\n"));
				return;
			}
			start_num=1;
			IRecordset2->MoveLast();
			stop_num=RecNo(IRecordset2);     
			IRecordset2->Close();
		}
		catch(_com_error &e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}
	}
	else if(ch_cal==2)
	{
		try
		{
			//首先将C_PASS字段全部清空。
			sql = "UPDATE [PRE_CALC] SET C_PASS='' WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
			theApp.m_pConnection->Execute(_bstr_t(sql), NULL, adCmdText );

			sql = "SELECT * FROM [PRE_CALC] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ORDER BY ID ASC";
			_RecordsetPtr pRsCalc;
			pRsCalc.CreateInstance(__uuidof(Recordset));
			pRsCalc->Open(_variant_t(sql), (IDispatch*)GetConnect(),
					adOpenStatic, adLockOptimistic, adCmdText);
			if(pRsCalc->adoEOF && pRsCalc->BOF)
			{
				ExceptionInfo(_T("原始数据表为空\n"));
				return;
			}

			start_num=1;
			pRsCalc->MoveLast();
			stop_num=RecNo(pRsCalc);
			if( FALSE == RangeDlgshow(start_num,stop_num) )
			{
				return;
			}
			//然后在指定的范围内，将C_PASS字段赋'Y'
			ReplaceAreaValue(pRsCalc, "c_pass", "Y", start_num, stop_num);
			pRsCalc->Close();
		}
		catch(_com_error &e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}

	}
	else if(ch_cal==3)
	{
		try
		{
			_RecordsetPtr IRecordset2;
			IRecordset2.CreateInstance(__uuidof(Recordset));
			int iIsStop=0;
			//首先将C_PASS字段全部清空。
			sql = "UPDATE [PRE_CALC] SET C_PASS='' WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
			theApp.m_pConnection->Execute(_bstr_t(sql), NULL, adCmdText );
			//
			sql.Format(_T("SELECT c_pass,id,c_vol,c_name1,c_size,c_pi_thi,c_temp,c_name2,							\
				c_name3,c_pre_thi,c_place,c_mark ,EnginID FROM PRE_CALC where EnginID='%s' ORDER BY id ASC"),GetProjectName());
			
			IRecordset2->CursorLocation = adUseClient;
			
			IRecordset2->Open(_variant_t(sql),_variant_t((IDispatch*)GetConnect()),
				adOpenDynamic,adLockOptimistic,adCmdText);
			SelectToCal(IRecordset2,&iIsStop);
			//MessageBox("在欲计算的记录C_PASS字段处输入'Y'");
			IRecordset2->Close();
			if(iIsStop)
			{
				return;
			}
		}
		catch(_com_error &e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}
	}
	OPENATABLE(IRecConfig,A_CONFIG);
	
	if(IRecConfig->adoEOF && IRecConfig->BOF)
	{
		ExceptionInfo(_T("保温设计常数库(A_CONFIG)不能为空\n"));
		return;
	}
	try
	{
		IRecConfig->MoveFirst();
		
		GetTbValue(IRecConfig,_variant_t("单位制"),unit);
		unit.MakeUpper();
		GetTbValue(IRecConfig,_variant_t("主汽热价"),heat_price);
		GetTbValue( IRecConfig, _variant_t("年费用率"), S ); // by zsy 2007.01.11
		GetTbValue(IRecConfig,_variant_t("年运行小时"),a_config_hour);
		m_dMaxD0 = vtof(IRecConfig->GetCollect(_variant_t("平壁与圆管的分界外径")) );
		if( m_dMaxD0 <= 0 )
		{	//当分界外径不合理时,赋一个默认值.
			m_dMaxD0 = 2000;
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}
	if(unit != _T("SI"))
	{
		//内部公式全部使用国际单位制计算,西南院程序公式使用工程制计算.
		//而西南院的放热系数ALPHA采用的是国际单位制,两者不一致计算结果有错误.
		heat_price=heat_price/4.1868;
	}
	IRecConfig->Close();
	
	if( IRecCalc == NULL )
	{
		hr=IRecCalc.CreateInstance(__uuidof(Recordset));
	}
	TempStr.Format(_T("SELECT * FROM PRE_CALC WHERE EnginID='%s' ORDER BY ID ASC"),this->GetProjectName());
	try
	{
		if(IRecCalc->State == adStateOpen)
		{
			IRecCalc->Close();
		}
		IRecCalc->Open(_variant_t(TempStr),_variant_t((IDispatch*)GetConnect()),
			adOpenStatic,adLockOptimistic,adCmdText);
		
		EDIBgbl::pRsCalc = IRecCalc;				//将原始数据的记录集可以在其他方法中使用.
		IRecCalc->MoveFirst();
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}
	OPENATABLE(IRecMat,A_mat);
	if(IRecMat->adoEOF && IRecMat->BOF)
	{
	//	ExceptionInfo(_T("保温材料参数库(A_MAT)不能为空\n"));
	//	return;
	}
	//打开标准的材料库.(工程名为空的记录)
	OpenAProjectTable(m_pRs_CodeMat,"A_MAT","");
	//打开防冻计算时新增数据的表。
	OpenAProjectTable(m_pRsCongeal,"pre_calc_congeal",GetProjectName());
	// 打开埋地管道保温计算时新增加的数据表
	OpenAProjectTable(m_pRsSubterranean, "Pre_calc_Subterranean", GetProjectName());
	//
	
	
	//将记录集IRecPipe改为成员的。每次计算只打开一次。
	//目的：提高速度。		
	//打开管道参数库。
	if(!OpenAssistantProjectTable(IRecPipe, "A_PIPE") )
	{
		return;
	}
	//打开保温结构外表面允许最大散热密度
	//在IPEDcode.mdb库中。
	if( !OpenAssistantProjectTable(m_pRsHeatLoss, "HeatLoss") )
	{
		return ;
	}
	//打开管道/设置外表面传热系数.
	if( !OpenAssistantProjectTable(m_pRsCon_Out, "CON_OUT"))
	{
		return ;
	}
	//打开计算放热系数时的计算公式。
	if( !OpenAssistantProjectTable(m_pRsHeat_alfa, "放热系数取值表") )
	{
		return ;
	}
	//打开保护材料的黑度
	if( !OpenAssistantProjectTable(m_pRs_a_Hedu, "a_hedu") )
	{
		return ;
	}
	// 土壤的导热系数
	if( !OpenAssistantProjectTable(m_pRsLSo, "土壤的导热系数"))
		return;

/*
	// 保温层外表面向周围空气的放热阻
	if ( !OpenAssistantProjectTable(m_pRsFaceResistance, "Resistance") )
	{
		return;
	}
*/
	
	// 获得气象参数的露点温度
	GetWeatherProperty( "Td_DewPoint", Temp_dew_point );

	CCalcThickBase *lpMethodClass;		// 指向不同方法对应的类
	bool	bIsPlane;	 	// 保温对象类型的标志，true--平面,false--管道
	
	//获得
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表（thicknessRegular）中取出上面的值。
	if( !GetThicknessRegular() )
	{
		return ;
	}

	//保温界面温度不应超过外层保温材料最高使用温度的比例.	
	if(	!this->GetConstantDefine("ConstantDefine", "Ratio_MaxHotTemp", TempStr) )
	{
		m_HotRatio = 1;	//默认的值
	}else
	{
		m_HotRatio = _tcstod(TempStr, NULL);
	}
	//保冷界面温度不应超过外层保冷材料最高使用温度的比例.   
	if(	!this->GetConstantDefine("ConstantDefine", "Ratio_MaxCoolTemp", TempStr) )
	{
		m_CoolRatio = 1;	//默认的值
	}else
	{
		m_CoolRatio = _tcstod(TempStr, NULL);
	}
	// 外表面允许最大温度
	if (!GetConstantDefine("ConstantDefine","FaceMaxTemp",TempStr))
	{
		MaxTs = 50;
	}else
	{
		MaxTs = strtod(TempStr,NULL);
	}

	// 保冷时,规定为露点温度增加的一个常数
	if ( !GetConstantDefine( "ConstantDefine", "DeltaTsd", Temp_coefficient ) )
	{
		Temp_coefficient = 1;
	}

	// 如果是重新自动选择保温结构
	if( giInsulStruct == 1 )
	{
		if(IDYES==AfxMessageBox(IDS_REALLYAUTOSELALLMAT,MB_YESNO+MB_DEFBUTTON2+MB_ICONQUESTION))
		{
			CEditOriginalData  Original;
			Original.SetCurrentProjectConnect(theApp.m_pConnection);
			Original.SetPublicConnect(theApp.m_pConnectionCODE);
			Original.SetProjectID(EDIBgbl::SelPrjID);

			if( !Original.AutoSelAllMat() )
			{
				return;
			}
		}
	}
	MinimizeAllWindow();	// 最小化所有的浏览窗口

	Say(_T("开始保温计算 "));
	
	//	进入循环体
	while(!IRecCalc->adoEOF)
	{
		try
		{
			GetTbValue(IRecCalc,_variant_t("c_pass"),TempStr);
			if( TempStr.CompareNoCase("Y") )
			{
				IRecCalc->MoveNext();
				continue;
			}
			//原始数据的序号
			m_nID = vtoi(IRecCalc->GetCollect(_variant_t("ID")));
			
			// ***查有关材料价格、导热系数计算公式(或值)
			//	 	&&外层保温材料
			GetTbValue(IRecCalc,_variant_t("c_name2"),out_mat);
			out_mat.TrimRight();
			
			//		&&内层保温材料
			GetTbValue(IRecCalc,_variant_t("c_name_in"),in_mat);
			in_mat.TrimRight();
			
			//	 	&&保护层材料
			GetTbValue(IRecCalc,_variant_t("c_name3"),pro_mat);
			pro_mat.TrimRight();
			
			//在黑度表中查找保护材料的黑度的时候。
			//保护材料名称,是去掉后面括号的名称.如"铝皮(0.75)"则该值为"铝皮"
			pro_name = pro_mat;
			if( -1 != pro_name.Find("(",0) )
			{
				pro_name = pro_name.Mid( 0, pro_name.Find("(",0) );
			}
		}
		catch(_com_error &e)
		{
			TempStr.Format(_T("原始数据序号为%d的记录在读取\"原始数据表(PRE_CALC)\"中出错\r\n"),
				m_nID);
			if(Exception::GetAdditiveInfo())
			{
				TempStr+=Exception::GetAdditiveInfo();
			}
			ExceptionInfo(TempStr);
			Exception::SetAdditiveInfo(NULL);
			ReportExceptionErrorLV2(e);
			IRecCalc->MoveNext();
			continue;
		}   
		//		***单层保温
		if (out_mat.IsEmpty())
		{
			TempStr.Format(_T("原始数据序号为%d的记录,没有外保温材料!"),m_nID);
			ExceptionInfo(TempStr);
			IRecCalc->MoveNext();
			continue;
		}
		if (pro_mat.IsEmpty())
		{
			TempStr.Format(_T("原始数据序号为%d的记录,没有保护层材料!"),m_nID);
			ExceptionInfo(TempStr);			
			IRecCalc->MoveNext();
			continue;			
		}
		try
		{
			if (!(IRecMat->adoEOF && IRecMat->BOF))
			{
				IRecMat->MoveFirst();
			}
		}
		catch(_com_error &e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}
		
		if(!LocateFor(IRecMat,_variant_t("MAT_NAME"),CFoxBase::EQUAL,
			_variant_t(out_mat)))
		{
			//在当前工程的材料库中没找到,则去标准库中找(工程名为空).找到就将标准库中的材料增加到当前的工程中来			
			if (!FindStandardMat(IRecMat,out_mat))
			{
				TempStr.Format(_T("原始数据序号为%d的记录无法在\"保温材料参数库(A_MAT)\"中的\"材料名称\"字段")
					_T("找到%s\r\n有可能数据表为空或未添加数据"),
					m_nID,out_mat);
				ExceptionInfo(TempStr);
				IRecCalc->MoveNext();
				continue;
			}
		}
		
		try
		{
			//价格
			GetTbValue(IRecMat,_variant_t("MAT_PRIC"),in_price);
			
			//外层保温材料的最高使用温度
			//GetTbValue(IRecMat,_variant_t("MAT_TEMP"),in_tmax);
			
			//"MAT_TMAX"字段是"复合面允许最高温度"
			GetTbValue(IRecMat,_variant_t("MAT_TMAX"),in_tmax);
			if (in_tmax <= 0)
			{   //若复合面允许最高温度未设置
				double dTemp =  vtof(IRecMat->GetCollect(_variant_t("MAT_TEMP")));				
				in_tmax = dTemp / 2;//复合面允许最高温度设置外层保温材料的最高使用温度的1/2
				if (in_tmax <= 0)
				{
					//若外层保温材料的最高使用温度未设置，则复合面允许最高温度默认为350C。
					in_tmax = 350;
				}
			}
			//导热系数基数
			GetTbValue(IRecMat,_variant_t("MAT_A0"),in_a0);
			//导热系数一次项系数
			GetTbValue(IRecMat,_variant_t("MAT_A"),in_a1);
			//导热系数二次项系数
			GetTbValue(IRecMat,_variant_t("MAT_A1"),in_a2);
			//容重
			GetTbValue(IRecMat,_variant_t("MAT_DENS"),in_dens);
			// 导热系数增加的比率
			GetTbValue( IRecMat, _variant_t("MAT_RATIO"), in_ratio);

			//内部公式全部使用国际单位制计算,西南院程序公式使用工程制计算.
			//而西南院的放热系数ALPHA采用的是国际单位制,两者不一致计算结果有错误.
			if(unit != _T("SI"))
			{
				in_a0/=0.8598;
				in_a1/=0.8598;
				in_a2/=0.8598;
			}
		}
		catch(_com_error &e)
		{
			TempStr.Format(_T("原始数据序号为%d的记录在读取\"保温材料参数库(A_MAT)\"中")
				_T("材料名称=%s时出错\r\n"),m_nID,out_mat);
			if(Exception::GetAdditiveInfo())
			{
				TempStr+=Exception::GetAdditiveInfo();
			}
			ExceptionInfo(TempStr);
			Exception::SetAdditiveInfo(NULL);
			ReportExceptionErrorLV2(e);
			IRecCalc->MoveNext();
			continue;
		}
		IRecMat->MoveFirst();
		
		//查找保护材料的数据
		if(!LocateFor(IRecMat,_variant_t("MAT_NAME"),CFoxBase::EQUAL,
			_variant_t(pro_mat)))
		{
			//在当前工程的材料库中没找到,则去标准库中找(工程名为空).找到就将标准库中的材料增加到当前的工程中来
			if (!FindStandardMat(IRecMat, pro_mat))
			{
				TempStr.Format(_T("原始数据序号为%d的记录无法在\"保温材料参数库(A_MAT)\"中的\"材料名称\"字段")
					_T("找到%s\r\n有可能数据表为空或未添加数据"),
					m_nID,pro_mat);			
				ExceptionInfo(TempStr);
				IRecCalc->MoveNext();
				continue;
			}
		}
		try
		{
			//价格
			GetTbValue(IRecMat,_variant_t("MAT_PRIC"),pro_price);
			//导热系数基数
			GetTbValue(IRecMat,_variant_t("MAT_A0"),pro_conduct);
			//容重
			GetTbValue(IRecMat,_variant_t("MAT_DENS"),pro_dens);
		}
		catch(_com_error &e)
		{
			TempStr.Format(_T("原始数据序号为%d的记录在读取\"保温材料参数库(A_MAT)\"中")
				_T("材料名称=%s时出错\n"),m_nID,pro_mat);
			if(Exception::GetAdditiveInfo())
			{
				TempStr+=Exception::GetAdditiveInfo();
			}
			ExceptionInfo(TempStr);
			Exception::SetAdditiveInfo(NULL);
			ReportExceptionErrorLV2(e);
			IRecCalc->MoveNext();
			continue;
		}
		
		//复合保温
		in_mat.TrimRight();
		if (!in_mat.IsEmpty())
		{
			//将前面获得的外保温材料的参数赋给外保温材料的变量
			out_a0 = in_a0;
			out_a1 = in_a1;
			out_a2 = in_a2;
			out_price = in_price;
			out_dens  = in_dens;
			out_ratio = in_ratio;
			
			//查找内保温材料
			IRecMat->MoveFirst();
			if(!LocateFor(IRecMat,_variant_t("MAT_NAME"),CFoxBase::EQUAL,
							_variant_t(in_mat)))
			{
				//在当前工程的材料库中没找到,则去标准库中找(工程名为空).找到就将标准库中的材料增加到当前的工程中来
				if (!FindStandardMat(IRecMat, in_mat))
				{
					TempStr.Format(_T("原始数据序号为%d的记录无法在\"保温材料参数库(A_MAT)\"中的\"材料名称\"字段")
						_T("找到%s\r\n有可能数据表为空或未添加数据"),
						m_nID,in_mat);			
					ExceptionInfo(TempStr);
					IRecCalc->MoveNext();
					continue;
				}
			}
			try
			{
				//价格
				GetTbValue(IRecMat,_variant_t("MAT_PRIC"),in_price);
				//导热系数基数
				GetTbValue(IRecMat,_variant_t("MAT_A0"),in_a0);
				//导热系数一次项系数
				GetTbValue(IRecMat,_variant_t("MAT_A"),in_a1);
				//导热系数二次项系数
				GetTbValue(IRecMat,_variant_t("MAT_A1"),in_a2);
				//容重
				GetTbValue(IRecMat,_variant_t("MAT_DENS"),in_dens);
				// 导热系数增加的比率，如果为保冷(1.5 －　1.7)
				GetTbValue( IRecMat, _variant_t("MAT_RATIO"), in_ratio );

				//内部公式全部使用国际单位制计算,西南院程序公式使用工程制计算.
				//而西南院的放热系数ALPHA采用的是国际单位制,两者不一致计算结果有错误.
				if(unit != _T("SI"))
				{
					in_a0/=0.8598;
					in_a1/=0.8598;
					in_a2/=0.8598;
				}				
			}
			catch(_com_error &e)
			{
				TempStr.Format(_T("原始数据序号为%d的记录在读取\"保温材料参数库(A_MAT)\"中")
					_T("材料名称=%s时出错\r\n"),m_nID,in_mat);
				if(Exception::GetAdditiveInfo())
				{
					TempStr+=Exception::GetAdditiveInfo();
				}
				ExceptionInfo(TempStr);
				Exception::SetAdditiveInfo(NULL);
				ReportExceptionErrorLV2(e);
				IRecCalc->MoveNext();
				continue;
			}
		}
		try
		{
			//热价比主汽价
			GetTbValue(IRecCalc,_variant_t("c_price"),Yong);
			//设计温度
			GetTbValue(IRecCalc,_variant_t("c_temp"),Temp_pip);
			//外径
			GetTbValue(IRecCalc,_variant_t("c_size"),D0);
			// Modifier zsy 2007.01.11
			// 年费用率应该直接用常数据库中的数据.
//			//年费用率
//			GetTbValue(IRecCalc,_variant_t("c_rate"),S);
			//年运行小时
			GetTbValue(IRecCalc,_variant_t("c_hour"),Mhours);
			//判断运行工况
			if( Mhours >= a_config_hour )
			{
				hour_work = 1;	//年运行工况
			}
			else
			{
				hour_work = 0; //季节运行工况.
			}
			//环境温度
			GetTbValue(IRecCalc,_variant_t("c_con_temp"),Temp_env);
			//环境温度的取值索引
			nTaIndex = vtoi(IRecCalc->GetCollect(_variant_t("c_Env_Temp_Index")));
			//风速
			GetTbValue(IRecCalc,_variant_t("c_wind"),speed_wind);
			if (speed_wind < 0)
			{
				speed_wind = 0;
			}
			//保护层厚
			GetTbValue(IRecCalc,_variant_t("c_pro_thi"),pro_thi);
			//壁厚
			GetTbValue(IRecCalc,_variant_t("c_pi_thi"),pi_thi);
			//管道材质
			GetTbValue(IRecCalc,_variant_t("c_pi_mat"),pi_mat);
			//压力标示
			//GetTbValue(IRecCalc,_variant_t("c_pg"),pi_pg);
			//管内压力
			m_dPipePressure = vtof(IRecCalc->GetCollect(_variant_t("c_Pressure")));
			if (m_dPipePressure <= 3.92) // 用于管道上计算支吊架间距的参数
			{
				pi_pg = TRUE;
			}
			else{
				pi_pg = FALSE;
			}

			//备注
			GetTbValue(IRecCalc,_variant_t("c_mark"),mark);
			//卷册号
			GetTbValue(IRecCalc,_variant_t("c_vol"),vol);			
			//管道或设备的类型,为卷册号最后一个字符.
			cPipeType = vol.Right(1);
			//汽管或水管的标示．主要用于计算管道上支吊架的间距。
			GetTbValue(IRecCalc,_variant_t("c_steam"),steam);			
			//安装地点.
			strPlace = vtos(IRecCalc->GetCollect(_variant_t("c_place")));			
			//获得计算放热系数的公式,的索引.
			m_nIndexAlpha = vtoi(IRecCalc->GetCollect(_variant_t("c_Alfa_Index")));			
			//数量
			dAmount = vtof(IRecCalc->GetCollect(_variant_t("c_amount")));			
			//计算方法的索引值
			nMethodIndex = vtoi(IRecCalc->GetCollect(_variant_t("C_CalcMethod_Index")));
			// 
			switch( bIsReCalc )
			{
			case 1: nMethodIndex = nSurfaceTemperatureMethod;	// 设置重新使用外表面温度法校核。
					bIsReCalc = FALSE;		// 只有当前的一条记录才校核
					break;
			default: ;	// 不改变计算方法
			}
			//是否计算内保温层厚度
			bNoCalInThi = vtob(IRecCalc->GetCollect(_variant_t("c_CalInThi")));
			//是否计算外保温层厚度
			bNoCalPreThi = vtob(IRecCalc->GetCollect(_variant_t("c_CalPreThi")));
			//外表温度
			temp_ts = vtof(IRecCalc->GetCollect(_variant_t("c_tb3")));
			//内层保温厚
			thick_1 = vtof(IRecCalc->GetCollect(_variant_t("c_in_thi")));
			//外层保温厚
			thick_2 = vtof(IRecCalc->GetCollect(_variant_t("c_pre_thi")));
			//沿风速方向的平壁宽度
			B = vtof(IRecCalc->GetCollect(_variant_t("c_B")));
			if (B <= 0)
			{
				//默认值
				B = 1000;
			}
			//介质名称
			strMedium = vtos(IRecCalc->GetCollect(_variant_t("c_Medium")));
			m_dMediumDensity = GetMediumDensity(strMedium);	//根据介质名,查找其密度
		}
		catch(_com_error &e)
		{
			TempStr.Format(_T("原始数据序号为%d的记录在读取\"原始数据库(PRE_CALC)\"中时出错\r\n"),m_nID);
			if(Exception::GetAdditiveInfo())
			{
				TempStr+=Exception::GetAdditiveInfo();
			}
			ExceptionInfo(TempStr);
			Exception::SetAdditiveInfo(NULL);
			ReportExceptionErrorLV2(e);
			IRecCalc->MoveNext();
			continue;
		}
		ERR=0;
		//查找设备和管道保温结构外表面传热系数.
		try
		{
			if(pro_mat.Find(_T("抹面")) != -1)	//保护材料为抹面时.
			{
				GetAlfaValue(m_pRsCon_Out, "out_dia", "con2", D0, co_alf);
			}
			else 
			{
				GetAlfaValue(m_pRsCon_Out, "out_dia", "con1", D0, co_alf);
			}
		}
		catch(_com_error &e)
		{
			TempStr.Format(_T("原始数据序号为%d的记录在读取\"con_out表\"中时出错\r\n"),m_nID);
			if(Exception::GetAdditiveInfo())
			{
				TempStr+=Exception::GetAdditiveInfo();
			}
			ExceptionInfo(TempStr);
			Exception::SetAdditiveInfo(NULL);
			ReportExceptionErrorLV2(e);
			IRecCalc->MoveNext();
			continue;
		}
		//用一个成员的记录指针，在计算之前就打开表[a_hedu]
		m_pRs_a_Hedu->MoveFirst();
		//保护材料的黑度
		if(LocateFor(m_pRs_a_Hedu,_variant_t("c_name"),CFoxBase::EQUAL,
			_variant_t(pro_name))==TRUE)
		{
			TempValue=m_pRs_a_Hedu->GetCollect(_variant_t("n_hedu"));
			hedu=TempValue;
		}
		else
		{
			hedu=0.23;     //默认镀锌铁皮对应的黑度
		}
		
		if(Yong*Temp_pip<-0.001)
		{
			TempStr.Format(_T("出现错误在记录: %d"),RecNo(IRecCalc));
			AfxMessageBox(TempStr);
			ERR=1;
		}
		
		if(D0*in_price*in_a0*S*Mhours*speed_wind*pro_thi*pi_thi<0)
		{
			TempStr.Format(_T("序号为 %d 的记录,出现原则性错误! "),RecNo(IRecCalc));
			AfxMessageBox(TempStr);
			ERR=1;			
		}
		
		if(ERR==1)
		{
			if( IDYES == AfxMessageBox("是否停止计算？", MB_YESNO ) )
			{
				return;
			}
			else
			{
				IRecCalc->MoveNext();
				continue;
			}
		}
		// 确定保温对象类型
		if( D0 >= m_dMaxD0 )		
		{	//平面
			bIsPlane = true;
		}
		else
		{	//管道
			bIsPlane = false;
		}

		switch( nMethodIndex )
		{
		case nSurfaceTemperatureMethod:
			//外表面温度法计算保温厚度 
			//外表面温度为已知
			INITDATA(CCalcThick_FaceTemp);		//初始各成员值。将当前记录的原始数据传入到计算方法的类中
			lpMethodClass = (CCalcThick_FaceTemp *)this;
			break;
		case nHeatFlowrateMethod:
			//允许散热密度法计算保温厚度
			INITDATA(CCalcThick_HeatDensity);	//
			lpMethodClass = (CCalcThick_HeatDensity *)this;
			break;
		case nPreventCongealMethod:
			//防冻结(热平衡)法计算保温厚度
			INITDATA(CCalcThick_PreventCongeal)	//
			lpMethodClass = (CCalcThick_PreventCongeal *)this;
			break;
		case nSubterraneanMethod: 
			// 埋地管道保温计算方法
			short nType;
			short nPipeCount;
			GetSubterraneanTypeAndPipeCount(m_nID, nType, nPipeCount);
			switch(nType)
			{
			case 1:	// 有管沟通行
			case 2:	// 有管沟不通行
				INITDATA(CCalcThick_SubterraneanObstruct);
				lpMethodClass = (CCalcThick_SubterraneanObstruct*)this;
				break;
			case 0:	// 埋地的无管沟直埋
			default:
				INITDATA(CCalcThick_SubterraneanNoChannel);		//.
				lpMethodClass = (CCalcThick_SubterraneanNoChannel*)this;
			}
			if (nPipeCount <= 1)
			{
				in_mat = "";	// 内保温材料对应埋地的第二根管道的保温材料名称
			}
			bIsPlane = false;	// 埋地计算方法,只有管道的保温的计算公式
			break;
		default:	//if( nMethodIndex == nEconomicalThicknessMethod )
			//默认为经济法计算保温厚
			INITDATA(CCalcThick_Economy);		//初始各成员值。将当前记录的原始数据传入到计算方法的类中.
			lpMethodClass = (CCalcThick_Economy *)this;
		}


		if( nMethodIndex == nSurfaceTemperatureMethod )
		{
			//外表面温度法计算保温厚度
			//外表面温度为已知
			INITDATA(CCalcThick_FaceTemp);		//初始各成员值。将当前记录的原始数据传入到计算方法的类中
			lpMethodClass = (CCalcThick_FaceTemp *)this;
		}
		else if( nMethodIndex == nHeatFlowrateMethod )
		{	
			//允许散热密度法计算保温厚度
			INITDATA(CCalcThick_HeatDensity);	//初始各成员值。将当前记录的原始数据传入到计算方法的类中
			lpMethodClass = (CCalcThick_HeatDensity *)this;
		}
		else if ( nMethodIndex == nPreventCongealMethod )
		{
			//防冻结(热平衡)法计算保温厚度
			INITDATA(CCalcThick_PreventCongeal);	//初始各成员值。将当前记录的原始数据传入到计算方法的类中
			lpMethodClass = (CCalcThick_PreventCongeal *)this;
		}
		else	//if( nMethodIndex == nEconomicalThicknessMethod )
		{	
			//默认为经济法计算保温厚
			INITDATA(CCalcThick_Economy);		//初始各成员值。将当前记录的原始数据传入到计算方法的类中.
			lpMethodClass = (CCalcThick_Economy *)this;
		}

				
		//计算前，先将各变量初始化
		
		temp_tb = 0.0;							//保温内外层界面温度
		pre_wei1 = pre_wei2 = pro_wei = 0.0;	//保温内外层、保护层重量
		p_wei = w_wei = with_w = no_w = 0.0;
		
		pre_v1 = pre_v2 = pro_v = 0.0;			//保温内外层体积
		distan = 0.0;							//支吊架间距
		face_area = 1;							//如果是平面时，则为1, 否则为管道将重新计算.		
		
		thick_3 = pro_thi;						//thick_3=保护厚
		
		in_mat.TrimRight();
		if( !in_mat.IsEmpty() )//为复合保温
		{
			if( bIsPlane )
			{
				//复合保温平壁.
				//计算保温厚.		2005.05.24
				if ( bNoCalInThi || bNoCalPreThi )		//输入保温层厚度
				{
					// 根据已知的保温厚度,计算外表面温度,和界面温度
					lpMethodClass->CalcPlain_Com_InputThick(thick_1,thick_2,temp_tb,temp_ts);
				}
				else
				{
					// 保温厚度没有输入，需要计算
					lpMethodClass->CalcPlain_Com(thick_1,thick_2,temp_tb,temp_ts);
				//////////////////////////////////////////////////////////////////////////
				// 放热系数在计算之前就确定好是查表还是按公式计算 [11/17/2005]
					/*
					while(1)
					{
						temp_ts1 = temp_ts;
						//根据用户选择的公式重新计算放热系数. 成功返回1.
						if( CalcAlfaValue(temp_ts, D0+2*thick_1+2*thick_2) )
						{
							//根据新的放热系数,再计算保温厚.
							lpMethodClass->CalcPlain_Com(thick_1,thick_2,temp_tb,temp_ts);
							if( fabs(temp_ts1-temp_ts) <= TsDiff)
							{
								break;
							}
						}
						else
						{
							//计算放热系数时,出错.或者是用查表所得的值去计算厚度,不需要再循环.
							break;
						}
					}
				*/
				}
				//输出到屏幕的字符串
				TempStr.Format(_T("平壁 \t记录号 \t温度 \t内层保温厚 \t外层保温厚 \t结合面温度 \t外表面温度 \r\n")
					_T("\t%d\t%.2f\t%.2f\t%.2f\t%.2f\t%.2f"),
					m_nID,Temp_pip,thick_1,thick_2,temp_tb,temp_ts);
				
				pre_v1=thick_1/1000;		//内保温层的体积.
				pre_v2=thick_2/1000;		//外保温层的体积.
				pro_v=thick_3/1000;			//保护层的体积
				pre_wei1=pre_v1*in_dens;	//内保温层的重量
				pre_wei2=pre_v2*out_dens;
				pro_wei=pro_v*out_dens;
				//end if ( 复合保温平壁 )
			}
			else
			{
				//复合保温管道
				if ( bNoCalPreThi || bNoCalInThi )	//2005.05.24
				{
					//指定保温厚度,计算外表面温度和界面温度
					lpMethodClass->CalcPipe_Com_InputThick(thick_1,thick_2,temp_tb,temp_ts);
				}
				else
				{
					//复合保温管道
					lpMethodClass->CalcPipe_Com(thick_1,thick_2,temp_tb,temp_ts);
					
				/*	while(1)
					{
						temp_ts1 = temp_ts;
						if(	CalcAlfaValue(temp_ts, D0+2*thick_1+2*thick_2) )
						{	//根据新的放热系数,再计算保温厚.
							lpMethodClass->CalcPipe_Com(thick_1,thick_2,temp_tb,temp_ts);
							if( fabs(temp_ts1-temp_ts) <= TsDiff )
							{
								break;
							}
						}
						else
						{
							//计算放热系数时,出错.或者是用查表所得的值去计算厚度,不需要再循环.
							break;
						}
					}
				*/
				}
				//输出到屏幕的字符串
				TempStr.Format(_T("管道 \t记录号 \t温度 \t直径 \t内层保温厚 \t外层保温厚 \t结合面温度 \t外表面温度 \r\n")
					_T("\t%d\t%.2f\t%.2f\t%.2f\t%.2f\t%.2f\t%.2f"),
					m_nID,Temp_pip,D0,thick_1,thick_2,temp_tb,temp_ts);
				
				//求体积和重量
				d1=D0+2*thick_1;
				d2=d1+2*thick_2;
				d3=d2+2*pro_thi;
				pre_v1=pi/4*(d1*d1-D0*D0)/1E6;
				pre_wei1=pre_v1*in_dens;
				pre_v2=pi/4*(d2*d2-d1*d1)/1E6;
				pre_wei2=pre_v2*out_dens;
				pro_v=pi/4*(d3*d3-d2*d2)/1E6;
				pro_wei=pro_v*pro_dens;
				
				//在管道参数库中获得指定外径和壁厚的管重和水重
				if( !GetPipeWeithAndWaterWeith(IRecPipe,D0,pi_thi,"PIPE_WEI",p_wei,"WATER_WEI",w_wei) )
				{
					//在管道参数库中没有找到对应外径的记录.
					//重新计算
					p_wei=pi/4*7850*(D0*D0-(D0-2*pi_thi)*(D0-2*pi_thi))/1000000;
					w_wei=pi/4*(D0-2*pi_thi)*(D0-2*pi_thi)/1000;
				}
				//无水单重
				no_w=p_wei+pre_wei1+pre_wei2+pro_wei;
				//有水单重
				with_w=no_w+w_wei;
				//外表面积
				face_area=pi*d3/1000;
				
				//管道或设备类型不为S，M，O
				if( cPipeType.CompareNoCase(_T("S")) && cPipeType.CompareNoCase(_T("M")) && cPipeType.CompareNoCase(_T("O")) )
				{
					p_wei=w_wei=no_w=with_w=0;
				}
				
				//***仅汽水管道、油管道算间距
				distan=0;
				
				//if( !cPipeType.CompareNoCase(_T("S")) || !cPipeType.CompareNoCase(_T("M")) || !cPipeType.CompareNoCase(_T("O")) )
				//{
					try
					{
						distan = GetSupportSpan(D0,pi_thi,pi_mat,m_dPipePressure,m_dMediumDensity,
											Temp_pip,no_w-p_wei);
					}
					catch(_com_error &e)
					{
						TempStr.Format(_T("原始数据序号为%d的记录在计算时出错\r\n"),
							m_nID);
						if(Exception::GetAdditiveInfo())
						{
							TempStr+=Exception::GetAdditiveInfo();
						}
						ExceptionInfo(TempStr);
						Exception::SetAdditiveInfo(NULL);
						ReportExceptionErrorLV2(e);
						IRecCalc->MoveNext();
						continue;
					}
					TempStr.Format(_T("管道 \t记录号 \t间距Lmax\t温度 \t直径 \t内层保温厚 \t外层保温厚 \t结合面温度 \t外表面温度 \r\n")
						_T("\t%d\t%.2f\t%.2f\t%.2f\t%.2f\t%.2f\t%.2f\t%.2f"),
						m_nID,distan,Temp_pip,D0,thick_1,thick_2,temp_tb,temp_ts);
				//}	
					
			}// end   (复合保温管道)
		}
		else//单层保温
		{
			temp_tb = 0;	// 单层没有界面温度
			thick_1 = 0;	// 单层没有内保温厚度

			if( bIsPlane )//单层保温平壁
			{
				if ( bNoCalPreThi || bNoCalInThi )	//手工输入保温层厚度
				{
					//根据保温厚度,计算外表面温度
					lpMethodClass->CalcPlain_One_InputThick(thick_2,temp_ts);
				}
				else
				{
					lpMethodClass->CalcPlain_One(thick_2,temp_ts);
				/*
					while(1)
					{
						temp_ts1 = temp_ts;
						//根据新的放热系数,再计算保温厚.
						if( CalcAlfaValue(temp_ts, D0+2*thick_2) )
						{
							lpMethodClass->CalcPlain_One(thick_2,temp_ts);
							if( fabs(temp_ts1-temp_ts) <= TsDiff)
							{
								break;
							}
						}
						else
						{
							//计算放热系数时,出错.或者是用查表所得的值去计算厚度,不需要再循环.
							break;
						}
					}
				*/
				}
				//输出到屏幕的字符串
				TempStr.Format(_T("平壁 \t记录号 \t温度 \t保温厚 \t外表面温度 \r\n")
					_T("\t%d\t%.2f\t%.2f\t%.2f"),
					m_nID,Temp_pip,thick_2,temp_ts);
				
				pre_v2=thick_2/1000;	//计算保温层体积	
				pro_v=pro_thi/1000;		//保护层体积
				pre_wei2=pre_v2*in_dens;//保温层重量
				pro_wei=pro_v*pro_dens;
				
				//end if (单层保温平壁)
			}
			else
			{//单层保温管道
				if ( bNoCalPreThi || bNoCalInThi )
				{
					//根据保温厚度,计算外表面温度.
					lpMethodClass->CalcPipe_One_InputThick(thick_2,temp_ts);					
				}
				else
				{
					lpMethodClass->CalcPipe_One(thick_2,temp_ts);
				/*
					while(1)
					{
						//根据新的放热系数,再计算保温厚.
						temp_ts1 = temp_ts;
						if( CalcAlfaValue(temp_ts, D0+2*thick_2) )
						{
							lpMethodClass->CalcPipe_One(thick_2,temp_ts);
							if( fabs(temp_ts1-temp_ts) <= TsDiff )
							{
								break;
							}
						}
						else
						{	//计算放热系数时,出错.或者是用查表所得的值去计算厚度,不需要再循环.
							break;
						}
					}
				*/
				}
				TempStr.Format(_T("管道 \t记录号 \t温度 \t直径 \t保温厚 \t外表面温度 \r\n")
					_T("\t%d\t%.2f\t%.2f\t%.2f\t%.2f"), 
					m_nID,Temp_pip,D0,thick_2,temp_ts);
				//计算体积与重量
				d1=D0+2*thick_2;
				d3=D0+2*thick_2+2*pro_thi;
				pre_v2=pi/4*(d1*d1-D0*D0)/1000000;
				pre_wei2=pre_v2*in_dens;
				pro_v=pi/4*(d3*d3-d1*d1)/1000000;
				pro_wei=pro_v*pro_dens; 
				
				//在管道参数库中获得指定外径和壁厚的管重和水重
				if( !GetPipeWeithAndWaterWeith(IRecPipe,D0,pi_thi,"PIPE_WEI",p_wei,"WATER_WEI",w_wei) )
				{
					//在管道参数库中没有找到对应外径的记录.
					//重新计算
					p_wei=pi/4*7850*(D0*D0-(D0-2*pi_thi)*(D0-2*pi_thi))/1000000;
					//	w_wei=pi/4*(D0-2*D0)*(D0-2*D0)/1000; // FoxPro中的公式
					w_wei = pi / 4 * ( D0 - 2 * pi_thi ) * ( D0 - 2 * pi_thi ) / 1000;
				}
				//无水单重
				no_w=p_wei+pre_wei2+pro_wei;
				//有水单重
				with_w=no_w+w_wei;
				//外表面积
				face_area=pi*(D0+2*thick_2+2*pro_thi)/1000;			
				//如果管道类型不为汽水管道或油管道
				if( cPipeType.CompareNoCase(_T("S")) && cPipeType.CompareNoCase(_T("M")) && cPipeType.CompareNoCase(_T("O")) )
				{
					p_wei=w_wei=no_w=with_w=0;
				}
				//***仅汽水管道或油管道算间距
				distan=0;

				
				//if( !cPipeType.CompareNoCase(_T("S")) || !cPipeType.CompareNoCase(_T("M")) || !cPipeType.CompareNoCase(_T("O")) )
				//{
					try
					{
						distan = GetSupportSpan(D0,pi_thi,pi_mat,m_dPipePressure,m_dMediumDensity,
											Temp_pip,no_w-p_wei);
					}
					catch(_com_error &e)
					{
						TempStr.Format(_T("原始数据序号为 %d 的记录在计算时出错\r\n"),
							m_nID);
						if(Exception::GetAdditiveInfo())
						{
							TempStr+=Exception::GetAdditiveInfo();
						}
						ExceptionInfo(TempStr);
						Exception::SetAdditiveInfo(NULL);
						ReportExceptionErrorLV2(e);
						IRecCalc->MoveNext();
						continue;
					}
					TempStr.Format(_T("管道 \t记录号 \t间距Lmax\t温度 \t直径 \t保温厚 \t外表面温度 \r\n")
						_T("\t%d\t%.2f\t%.2f\t%.2f\t%.2f\t%.2f"),
						m_nID,distan,Temp_pip,D0,thick_2, temp_ts);
				//}
				
			}// end （单层保温管道）
			
		}//end 单层

		if ( !bNoCalPreThi &&  !bNoCalInThi )	// 指定的保温厚度
		{
			if ( Temp_pip < Temp_env )
			{
				// 这种在不同的行业，计算方法
				if ( nMethodIndex != nSurfaceTemperatureMethod )
				{
					if ( temp_ts < Temp_dew_point + Temp_coefficient )	// 露点温度加一个修正值(不同的规范可能会有所不同,1-3℃)
					{
						//IRecCalc->PutCollect( _variant_t("c_CalcMethod_Index"), _variant_t((long)nSurfaceTemperatureMethod) );
						IRecCalc->PutCollect( _variant_t("c_tb3"), _variant_t(Temp_dew_point + Temp_coefficient));
						bIsReCalc = 1;	// 使用表面温度法重新校核

						TempStr.Format("提示: 序号为 %d 的记录在计算时,使用了外表面温度法进行校核！\r\n", m_nID);
						ExceptionInfo( TempStr );
						// 保存前一次计算的中间值
						IRecCalc->Update();
 						continue;
					}
				}
			}

		}
		//将字符串显示到屏幕上。（写入到文档中）
		DisplayRes(TempStr);
		
		//显示错误信息
		if ( lpMethodClass->GetExceptionInfo(TempStr) )
		{
			if(Exception::GetAdditiveInfo())
			{
				TempStr+=Exception::GetAdditiveInfo();
			}
			ExceptionInfo(TempStr);	
			Exception::SetAdditiveInfo(NULL);
		}		
		//是否停止运算
		IsStop=1;
		Cancel(&IsStop);
		if(IsStop==0)
		{
			return;
		}
		dUnitLoss = lpMethodClass->dQ;
		if (dUnitLoss < 0)
		{
			dUnitLoss = 0;
		}
		dAreaLoss = dUnitLoss * face_area * dAmount;
		
		//----------------------------
		//将计算的结果写入到数据库中
		try
		{
			//内保温厚
			IRecCalc->PutCollect(_variant_t("c_in_thi"),_variant_t(thick_1));
			//外保温厚/单层的保温层厚度
			IRecCalc->PutCollect(_variant_t("c_pre_thi"),_variant_t(thick_2));
			//外表面温度
			IRecCalc->PutCollect(_variant_t("c_tb3"),_variant_t(temp_ts));
			//界面温度
			IRecCalc->PutCollect(_variant_t("c_ts"),_variant_t(temp_tb));
			//单位散热密度
			IRecCalc->PutCollect(_variant_t("c_HeatFlowrate"),_variant_t(dUnitLoss));
			//热损失=外表面积*数量*单位散热密度
			IRecCalc->PutCollect(_variant_t("c_lost"),_variant_t(dAreaLoss));
			//保温结构投资年总费用
			IRecCalc->PutCollect(_variant_t("c_srsb"),_variant_t((double)0));
			//精确值对应保温结构投资年总费用
			IRecCalc->PutCollect(_variant_t("cal_srsb"),_variant_t((double)0));
			//外表温层厚度精确值
			IRecCalc->PutCollect(_variant_t("cal_thi"),_variant_t((double)0));
			//内保温层单位重量
			IRecCalc->PutCollect(_variant_t("c_in_wei"),_variant_t(pre_wei1));
			//外保温层单位重量/单层的保温层单重
			IRecCalc->PutCollect(_variant_t("c_pre_wei"),_variant_t(pre_wei2));
			//保护层单位重量
			IRecCalc->PutCollect(_variant_t("c_pro_wei"),_variant_t(pro_wei));
			//管道单重
			IRecCalc->PutCollect(_variant_t("c_pipe_wei"),_variant_t(p_wei));
			//水单重
			IRecCalc->PutCollect(_variant_t("c_wat_wei"),_variant_t(w_wei));
			//有水单重
			IRecCalc->PutCollect(_variant_t("c_with_wat"),_variant_t(with_w));
			//无水单重
			IRecCalc->PutCollect(_variant_t("c_no_wat"),_variant_t(no_w));
			//外表面积
			IRecCalc->PutCollect(_variant_t("c_area"),_variant_t(face_area));
			//内保温层单位体积
			IRecCalc->PutCollect(_variant_t("c_v_in"),_variant_t(pre_v1));
			//外保温层单位体积/单层的保温层单位体积
			IRecCalc->PutCollect(_variant_t("c_v_pre"),_variant_t(pre_v2));
			//保护层单位体积
			IRecCalc->PutCollect(_variant_t("c_v_pro"),_variant_t(pro_v));
			//管道支吊架间距
			IRecCalc->PutCollect(_variant_t("c_distan"),_variant_t(distan));
			
			IRecCalc->Update();
		}
		catch(_com_error &e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}		
		
		IRecCalc->MoveNext();
		continue;
		
		//以下代码为老版本的,现在没有用到.
		//删除，没有影响 by ligb on 2010.06.15

	}//ENDDO
	//计算保温材料的全体积.
	this->E_Predat_Cubage();
	//

	Say(_T("计算完毕!"));
//	CLOSE DATA

	OPENATABLE(IRecPass,PASS);
	
//	LOCATE FOR pass_mod1="C_ANALYS"
	if(!LocateFor(IRecPass,_variant_t("PASS_MOD1"),CFoxBase::EQUAL,
			  _variant_t("C_ANALYS")))
	{
		TempStr.Format(_T("无法在PASS表中的\"PASS_MOD1\"字段找到C_ANALYS \r\n有可能数据表为空或未添加数据"));
		ExceptionInfo(TempStr);

		return;
	}

//	rec=RECNO()
	rec=RecNo(IRecPass);
	IRecPass->MoveFirst();

//	REPL NEXT rec pass_mark1 WITH "T"
	ReplNext(IRecPass,_variant_t("PASS_MARK1"),_variant_t("T"),rec);

//	SKIP
	IRecPass->MoveNext();

//	REPL NEXT 130 pass_mark1 WITH "F"
	ReplNext(IRecPass,_variant_t("PASS_MARK1"),_variant_t("T"),130);

//	USE
//	RETURN
	IRecPass->Close();
	IRecCalc->Close();
	IRecMat->Close();
	IRecPipe->Close();
}

void CHeatPreCal::SetChCal(int Ch)
{
	ch_cal=Ch;
}

int CHeatPreCal::GetChCal()
{
	return ch_cal;
}

////////////////////////////////////////////////////////////
//
// 在需要指定计算的‘开始’与‘结束’范围时调用
//
// Start	需要计算的起始位置
// End		需要计算的结束位置
//
BOOL CHeatPreCal::RangeDlgshow(long &Start,long &End)
{
	AfxMessageBox("范围对话框");
	return TRUE;
	
}

///////////////////////////////////////////////////////////
//
// 当需要选择需要计算的数据时调用
//
// IRecordset	传入一张表的记录集合
//
// 在这张表的C_PASS的字段输入Y，这条记录将会被计算
//
void CHeatPreCal::SelectToCal(_RecordsetPtr &IRecordset,int *pIsStop)
{

}

////////////////////////////////////////////////////
//
// 设置工程名(ID)
//
// Name[in]	工程的ID号
//
void CHeatPreCal::SetProjectName(CString &Name)
{
	m_ProjectName=Name;
}

/////////////////////////////////////////////////////
//
//  返回工程名(ID)
//
CString& CHeatPreCal::GetProjectName()
{
	return m_ProjectName;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 打开某个工程的表的记录
//
// IRecordsetPtr[out]	返回需要打开的表的记录集
// pTableName[in]		需要打开的表名
// pProjectName[in]		需打开的工程名（ID）
//
// 如果函数成功返回TRUE，否则返回FALSE
//
// 调用此函数前需 SetConnect 设置数据库的连接
//
BOOL CHeatPreCal::OpenAProjectTable(_RecordsetPtr &IRecordset, LPCTSTR pTableName, LPCTSTR pProjectName)
{
	CString SQL;


	if(GetConnect()==NULL)
	{
		ReportExceptionErrorLV1(_T("连接不可用"));
		return FALSE;
	}

	if(pTableName==NULL)
	{
		ReportExceptionErrorLV1(_T("表名不可用"));
		return FALSE;
	}

	if(pProjectName==NULL)
	{
		SQL.Format(_T("SELECT * FROM %s WHERE EnginID='%s'"),pTableName,m_ProjectName);
	}
	else if(pProjectName==_T(""))
	{
		SQL.Format(_T("SELECT * FROM %s where EnginID is NULL"),pTableName);
	}
	else		
	{
		SQL.Format(_T("SELECT * FROM %s WHERE EnginID='%s'"),pTableName,pProjectName);
	}

	try
	{
		IRecordset.CreateInstance(__uuidof(Recordset));
		IRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)GetConnect()),
						 adOpenStatic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return FALSE;
	}

	return TRUE;
}


//////////////////////////////////////////////
//
// 判断是否停止计算时调用此函数
//
// pState[out]当pState返回0时则停止计算,
// 否则继续运算。
//
void CHeatPreCal::Cancel(int *pState)
{
	if(pState)
	{
		*pState=1;
	}
}


//////////////////////////////////////////////////
//
// 设置辅助数据库的连接
//
// IConnection[in]	数据库的连接的智能指针的引用
//
void CHeatPreCal::SetAssistantDbConnect(_ConnectionPtr &IConnection)
{
	m_AssistantConnection=IConnection; 
}

///////////////////////////////////////////////////////
//
//返回辅助数据库的连接
//
_ConnectionPtr& CHeatPreCal::GetAssistantDbConnect()
{
	return m_AssistantConnection;
}

/////////////////////////////////////////////////////////////////////////////////////
//
// 打开辅助数据库的表
// 
// IRecordset[out]	返回以打开表的记录集的智能指针
// pTableName[in]	需打开的表名
//
// 函数成功返回TRUE，否则返回FALSE
//
BOOL CHeatPreCal::OpenAssistantProjectTable(_RecordsetPtr &IRecordset, LPCTSTR pTableName)
{
	
	CString SQL;


	if(GetAssistantDbConnect()==NULL)
	{
		ReportExceptionErrorLV1(_T("连接不可用"));
		return FALSE;
	}

	if(pTableName==NULL)
	{
		ReportExceptionErrorLV1(_T("表名不可用"));
		return FALSE;
	}

	SQL.Format(_T("SELECT * FROM %s"),pTableName);

	try
	{
		if ( IRecordset == NULL )
			IRecordset.CreateInstance(__uuidof(Recordset));

		if ( IRecordset->State == adStateOpen )
			IRecordset->Close();
		
		IRecordset->Open(_variant_t(SQL),_variant_t((IDispatch*)GetAssistantDbConnect()),
						 adOpenStatic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return FALSE;
	}

	return TRUE;
}

/////////////////////////////////////////////////////////////
//
// 设置Material数据库的路径
// 
// pMaterialPath[in]	数据库所在的文件夹
//
void CHeatPreCal::SetMaterialPath(LPCTSTR pMaterialPath)
{
	HRESULT hr;
	CString strSQL;
	BOOL IsPathChange=FALSE;	//数据库路径与以前的路径是否相同
	CString FullDbFilePath;

	CalSupportSpan.m_strMaterialDbPath = m_MaterialPath+"Material.mdb";		//计算支吊架间距
	CalSupportSpan.m_pConMaterial = theApp.m_pConMaterial;
	m_IMaterialCon = theApp.m_pConMaterial;

	return ;

	try
	{
		if(pMaterialPath==NULL)
		{
			if(m_IMaterialCon==NULL)
			{
				return;
			}
			else if(m_IMaterialCon->State & adStateOpen)
			{
				m_IMaterialCon->Close();
			}
			return;
		}
		
		if(m_MaterialPath==pMaterialPath)
			IsPathChange=TRUE;
		
		m_MaterialPath=pMaterialPath;
//		CCalcDll.m_MaterialPath=m_MaterialPath+"material.mdb";
		FullDbFilePath=m_MaterialPath+"material.mdb";
		
		if(m_IMaterialCon==NULL)
		{
			hr=m_IMaterialCon.CreateInstance(__uuidof(Connection));
			
			if(FAILED(hr))
				return;
		}
		
		if(m_IMaterialCon->State==adStateClosed)
		{
			strSQL = _T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='"+FullDbFilePath+"'");
			m_IMaterialCon->Open(_bstr_t(strSQL),"","",-1);
			
			return;
		}

		if(IsPathChange==FALSE)
			return;

		strSQL =_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='"+FullDbFilePath+"'");

		m_IMaterialCon->Close();
		m_IMaterialCon->Open(_bstr_t(strSQL),"","",-1);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return;
	}
}

////////////////////////////////////////////////////////
//
// 返回Material数据库所在的文件夹
//
LPCTSTR CHeatPreCal::GetMaterialPath()
{
	return m_MaterialPath;
}

///////////////////////////////////////////////////////////////////////////////////////////
//
// 从Material数据库中的CL_OLD_NEW表中寻找与MatName材料名相对应的材料。
// 先找与MatName对应的新材料，如果未找到再找与MatName对应的旧材料名
//
// MatName[in]		输入寻找与之对应的材料名
// pFindName[out]	返回找到的材料名的数组
// Num[out]			pFindName的数组大小
//
// 如果函数成功找到返回TRUE，否则返回FALSE
// 如果返回TRUE，pFindName需用delete释放内存。
// 如果返回FALSE，pFindName与Num将无意义
//
BOOL CHeatPreCal::FindMaterialNameOldOrNew(LPCTSTR MatName, CString *&pFindName, int &Num)
{
	int FindNum=0;
	CString MatStr;
	CString SQL;
	_RecordsetPtr IRecord;

	if(m_IMaterialCon==NULL || m_IMaterialCon->State==adStateClosed || MatName==NULL)
		return FALSE;

	MatStr=MatName;
	MatStr.TrimLeft();
	MatStr.TrimRight();
	if(MatStr.IsEmpty())
		return FALSE;

	try
	{
		IRecord.CreateInstance(__uuidof(Recordset));

		SQL.Format(_T("SELECT NEWCL FROM Material_OLD_NEW WHERE OLD='%s'"),MatStr);

		IRecord->Open(_variant_t(SQL),_variant_t((IDispatch*)m_IMaterialCon),
					  adOpenStatic,adLockOptimistic,adCmdText);

		if(IRecord->adoEOF && IRecord->BOF)
		{
			SQL.Format(_T("SELECT OLD FROM Material_OLD_NEW WHERE NEWCL='%s'"),MatStr);

			IRecord->Close();
			IRecord->Open(_variant_t(SQL),_variant_t((IDispatch*)m_IMaterialCon),
						  adOpenStatic,adLockOptimistic,adCmdText);

			if(IRecord->adoEOF && IRecord->BOF)
				return FALSE;
		}

		IRecord->MoveFirst();
		FindNum=0;
		while(!IRecord->adoEOF)
		{
			FindNum++;
			IRecord->MoveNext();
		}

		pFindName=new CString[FindNum];
		Num=FindNum;

		IRecord->MoveFirst();
		FindNum=0;
		while(!IRecord->adoEOF)
		{
			GetTbValue(IRecord,_variant_t((short)0),pFindName[FindNum]);
			FindNum++;
			IRecord->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV1(e);
		return FALSE;
	}

	return TRUE;
}


//功能:
//根据单体积,计算全体积.
//
bool CHeatPreCal::E_Predat_Cubage()
{
	CString sql;
	CString amou,v_pre,v_in,tv_pre,tv_in,tv_pro,v_pro;
	double f_v_pre,f_amou,f_v_in,f_v_pro,amount,pro_thi;
	_variant_t num,varTemp;
	_RecordsetPtr m_pRecordset;
	m_pRecordset.CreateInstance(__uuidof(Recordset));

	sql.Format(_T("SELECT * FROM pre_calc where EnginID='%s'"),EDIBgbl::SelPrjID);

	try
	{
		m_pRecordset->Open(_variant_t(sql),               
								theApp.m_pConnection.GetInterfacePtr(),	 // 获取库接库的IDispatch指针
								adOpenDynamic,
								adLockOptimistic,
								adCmdText);

		while(!m_pRecordset->adoEOF)
		{
	
	//      不须要叛断序号,	
	//		num=m_pRecordset->GetCollect("c_num");
	//		if (num.vt != VT_NULL && num.vt != VT_EMPTY)
			{
				GetTbValue(m_pRecordset,_variant_t("c_amount"),amount);

				m_pRecordset->PutCollect(_variant_t("c_amou"),_variant_t(amount));

				GetTbValue(m_pRecordset,_variant_t("c_amou"),f_amou);

				GetTbValue(m_pRecordset,_variant_t("c_v_pre"),f_v_pre);

				GetTbValue(m_pRecordset,_variant_t("c_v_in"),f_v_in);

	    		tv_pre.Format(_T("%f"),f_v_pre*f_amou);
	    		tv_in.Format(_T("%f"),f_v_in*f_amou);

	    		m_pRecordset->PutCollect("c_tv_pre",_variant_t(tv_pre));
	    		m_pRecordset->PutCollect("c_tv_in",_variant_t(tv_in));

				GetTbValue(m_pRecordset,_variant_t("c_pro_thi"),pro_thi);

				if(pro_thi > 5.0)
				{
					GetTbValue(m_pRecordset,_variant_t("c_v_pro"),f_v_pro);

					tv_pro.Format("%f",f_v_pro*f_amou);
					m_pRecordset->PutCollect("c_tv_pro",_variant_t(tv_pro));
				}
				else
				{
					m_pRecordset->PutCollect("c_tv_pro",_variant_t("0"));
				}

				m_pRecordset->Update();

			}// if(num.vt!=VT_NULL)

			m_pRecordset->MoveNext();

		}//while(!m_pRecordset->adoEOF)
	}
	catch(_com_error &e)
	{
		ReportExceptionError(e);
		return false;
	}
	return true;
}
//
short CHeatPreCal::GetThicknessRegular()
{
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	CString strSQL,strMessage;
	_variant_t var;
	_RecordsetPtr pRs;
	pRs.CreateInstance(__uuidof(Recordset));
	try
	{
		strSQL = "SELECT * FROM [thicknessRegular] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
		pRs->Open(_variant_t(strSQL), GetConnect().GetInterfacePtr(),
				adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->GetRecordCount() > 0 )
		{
			var = pRs->GetCollect(_variant_t("thick1Start"));		//内层保温允许最小厚度	
			m_thick1Start = vtoi( var );

			var = pRs->GetCollect(_variant_t("thick1Max"));			//内层保温允许最大厚度
			m_thick1Max = vtoi( var );

			var = pRs->GetCollect(_variant_t("thick1Increment"));	//内层保温厚度最小增量
			m_thick1Increment = vtoi( var );

			var = pRs->GetCollect(_variant_t("thick2Start"));	//外层保温允许最小厚度
			m_thick2Start = vtoi( var );

			var = pRs->GetCollect(_variant_t("thick2Max"));		//外层保温允许最大厚度
			m_thick2Max = vtoi( var );

			var = pRs->GetCollect(_variant_t("thick2Increment"));//外层保温厚度最小增量
			m_thick2Increment = vtoi( var );
			
			//检查数据的正确性。
			if( (m_thick1Max<m_thick1Start) )//|| () || || fabs(m_thick2Increment) <= 1e-6)
			{
				strMessage.Format("内层保温允许最大厚度%dmm小于内层保温允许最小厚度%dmm",m_thick1Max,m_thick1Start);
				AfxMessageBox(strMessage);
				return false;
			}
			if(m_thick2Max < m_thick2Start)
			{
				strMessage.Format("外层保温允许最大厚度%dmm小于外层保温允许最小厚度%dmm",m_thick2Max,m_thick2Start);
				AfxMessageBox(strMessage);
				return false;
			}
			if( m_thick1Increment <= 1e-6 )
			{
				strMessage.Format("内层保温厚度最小增量%dmm太小.",m_thick1Increment);
				AfxMessageBox(strMessage);
				return false;
			}
			if( m_thick2Increment <= 1e-6 )
			{
				strMessage.Format("外层保温厚度最小增量%dmm太小.",m_thick2Increment);
				AfxMessageBox(strMessage);
				return false;
			}
			if( m_thick1Start <= 1e-6 )
			{
				strMessage.Format("内层保温允许最小厚度%dmm太小.",m_thick1Start);
				AfxMessageBox(strMessage);
				return false;
			}
			if( m_thick2Start <= 1e-6 )
			{
				strMessage.Format("外层保温允许最小厚度%dmm太小.",m_thick2Start);
				AfxMessageBox(strMessage);
				return false;
			}
			if( m_thick1Max <= 1e-6 )
			{
				strMessage.Format("内层保温允许最大厚度%dmm太小.",m_thick1Max);
				AfxMessageBox(strMessage);
				return false;
			}
			if( m_thick2Max <= 1e-6 )
			{
				strMessage.Format("外层保温允许最大厚度%dmm太小.",m_thick2Max);
				AfxMessageBox(strMessage);
				return false;
			}
		}
		else
		{
			if( IDNO == AfxMessageBox("保温厚度规则库中没有记录，是否按默认的数据进行计算?",MB_YESNO) )
			{
				return false;
			}
			m_thick1Start = 30;		//内层保温允许最小厚度	
			m_thick1Max= 300;		//内层保温允许最大厚度
			m_thick1Increment=10;	//内层保温厚度最小增量

			m_thick2Start = 30;		//外层保温允许最小厚度
			m_thick2Max = 300;			//外层保温允许最大厚度
			m_thick2Increment=10;	//外层保温厚度最小增量

			//将默认的数据写入到数据库中。
			strSQL.Format("INSERT INTO [thicknessRegular] (thick1Start,thick1Max ,thick1Increment ,thick2Start ,thick2Max ,thick2Increment ,EnginID )\
				VALUES(%d,%d,%d,%d,%d,%d,'%s') ",m_thick1Start,m_thick1Max,m_thick1Increment,m_thick2Start,m_thick2Max,m_thick2Increment,EDIBgbl::SelPrjID);
			theApp.m_pConnection->Execute(_bstr_t(strSQL), NULL, adCmdText );

		}

	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return false;
	}
	return true;
}
//
//
bool CHeatPreCal::InsertToReckonCost(double &ts, double &tb, double &thick1, double &thick2, double &cost_min, bool bIsFirst)
{
	CString strSQL;
	_RecordsetPtr pRs;
	pRs.CreateInstance(__uuidof(Recordset));

	try
	{
		//当为第一次时,先清空所有的记录.
		if( bIsFirst )
		{
			try
			{
				strSQL = "DELETE * FROM [ReckonCost]";
				theApp.m_pConnection->Execute(_bstr_t(strSQL),NULL,adCmdText);
			}
			catch(_com_error& e)
			{	
				//没有该表时， 则创建.
				if( e.Error() == DB_E_NOTABLE )
				{
					strSQL = "CREATE TABLE [ReckonCost] (ts double, tb double,thick1 double,thick2 double,cost_min double)";
					theApp.m_pConnection->Execute(_bstr_t(strSQL), NULL, adCmdText);
				}			
			}
		}
		//插入一条记录。
		strSQL.Format("INSERT INTO [ReckonCost] VALUES(%f,%f,%f,%f,%f)",ts,tb,thick1,thick2,cost_min); 
		theApp.m_pConnection->Execute(_bstr_t(strSQL),NULL,adCmdText);
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return false;
	}
	return true;
}
////////////////////////
//根据介质温度,取保温结构外表面允许最大散热密度
bool CHeatPreCal::GetHeatLoss(_RecordsetPtr& pRs,double &t/*温度*/, double &YearRun/*年运行*/, double &SeasonRun/*季节运行*/, double &QMax/*最大散热密度,根据运行小时数确定工况,(年运行或季节运行)*/)
{
	try
	{		
		YearRun = SeasonRun = QMax= 0;
		if( pRs->adoEOF && pRs->BOF )
		{
			//没有记录.
			return false;
		}
		//
		double nTempCur, nTempPre, nYearCur, nYearPre, nSeasonCur, nSeasonPre;
		bool   bFirst = true;
		for(pRs->MoveFirst(); !pRs->adoEOF; pRs->MoveNext() )
		{
			nTempCur = vtof(pRs->GetCollect(_variant_t("t")));
			nYearCur = vtof(pRs->GetCollect(_variant_t("YearRun"))); 
			nSeasonCur = vtof(pRs->GetCollect(_variant_t("SeasonRun")));

			if( nTempCur >= t )		//当前记录的温度值大于或等于要查找的温度.
			{
				if( bFirst )  //当为第一条记录时 
				{
					YearRun = nYearCur;
					SeasonRun = nSeasonCur;
					if( hour_work )
					{	//年运行工况 
						QMax = nYearCur;
					}
					else
					{	//季节运行工况
						QMax = nSeasonCur;
					}
					if( fabs(QMax) < 1e-6 )		//如果允许最大散热密度为0时(即，没有限制), 返回一个常数.
					{
						QMax  = D_MAXTEMP;
					}
					return true;
				}
				else
				{
					YearRun = (t-nTempPre)/(nTempCur-nTempPre)*(nYearCur-nYearPre)+nYearPre;
					SeasonRun = (t-nTempPre)/(nTempCur-nTempPre)*(nSeasonCur-nSeasonPre)+nSeasonPre;
					if( hour_work )
					{	//年运行工况
						QMax = nYearCur;
					}
					else
					{	//季节运行工况
						QMax = nSeasonCur;
					}
					if( fabs(QMax) < 1e-6 )		//如果允许最大散热密度为0时(即，没有限制), 返回一个常数.
					{
						QMax  = D_MAXTEMP;
					}
					return true;
				}
			}
			else
			{	//记住当前的值
				nTempPre = nTempCur;
				nYearPre = nYearCur;
				nSeasonPre = nSeasonCur;
			}
			bFirst = false;
		}
		//取最后一条记录的值
		if( pRs->adoEOF )
		{
			YearRun = nYearCur;
			SeasonRun = nSeasonCur;
			if( hour_work )
			{	//年运行工况
				QMax = nYearCur;
			}
			else
			{	//季节运行工况
				QMax = nSeasonCur;
			}
			if( fabs(QMax) < 1e-6 )		//如果允许最大散热密度为0时(即，没有限制), 返回一个常数.
			{
				QMax  = D_MAXTEMP;
			}

			return true;
		}
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return false;
	}
	return true;
}
//
// 从常量,常数库里取出数据.
bool CHeatPreCal::GetConstantDefine(CString strTblName, CString ConstName, CString &ConstValue)
{
	_RecordsetPtr pRs;
	pRs.CreateInstance(__uuidof(Recordset));

	CString strSQL;
	strTblName.TrimRight();
	strSQL = "SELECT * FROM ["+strTblName+"] ";
	try
	{
		pRs->Open(_variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(), 
			adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->GetRecordCount() <= 0)
		{
			return false; 
		}
		strSQL = "ConstantDefine_Name='"+ConstName+"' " ;
		pRs->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( !pRs->adoEOF  )
		{
			ConstValue = vtos(pRs->GetCollect("ConstantDefine_VALUE"));
			return true;
		}
		else
			return false;
	}
	catch(_com_error& e)
	{
		ReportExceptionError( e);
		return false;
	}
	return true;
}
// 在常量,常数库里取出数据.
bool CHeatPreCal::GetConstantDefine(CString strTblName, CString ConstName, double &dConstValue)
{
	CString strConstValue;
	if( GetConstantDefine(strTblName, ConstName, strConstValue) )
	{
		dConstValue = strtod( strConstValue, NULL );
		return true;
	}
	return false;
}

//功能:
//计算保温结构外表面传热系数.
//参数:
//dTs,外表面温度.
//dD1,保温层外径.
double CHeatPreCal::CalcAlfaValue(double dTs, double dD1)
{
	CString	sSQL,		//存放SQL语句.
			sVariableName,	//表中的变量名。
			sTmp,		//临时变量
			sFormula,	//计算放热系数的公式
			sIndex;		//将整数索引转换为字符串,临时的
	double  dTmpVal;

	CString sCalTblName = "tmpCalcInsulThick";//临时计算中间结果的表名.

	try
	{
		//查找计算外表面传热系数的公式.
		//{{如果公式为空,则是查表获得放热系数的.直接返回.
		if(m_pRsHeat_alfa->GetRecordCount() <= 0 )
		{
			//不存在记录时.
			return 0;	
		}
		sIndex.Format("%d", m_nIndexAlpha);
		sSQL = "Index= "+sIndex+" ";
		m_pRsHeat_alfa->MoveFirst();
		m_pRsHeat_alfa->Find(_bstr_t(sSQL), NULL, adSearchForward);
		if( !m_pRsHeat_alfa->adoEOF )
		{
			//定位到了给定索引值的记录.
			//提出计算放热系数的公式
			sFormula = vtos(m_pRsHeat_alfa->GetCollect(_variant_t("Formula")) );
			if( sFormula.IsEmpty() )
			{
				//没有计算公式的情况，查表取得放热系数。
				//查表已在计算每条记录的开始（初始化时）进行。
				return 0;
			}
		}
		else
		{
			//表中不存在该索引值对应的记录.
			return 0;
		}
		//}}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}

	
	_RecordsetPtr pRsTmp;		//临时的记录集.
	_RecordsetPtr pRsTmpVal;	//保存结果的记录。

	pRsTmp.CreateInstance(__uuidof(Recordset));
	pRsTmpVal.CreateInstance(__uuidof(Recordset));
	try
	{
		//先清空表中的所有记录。
//		theApp.m_pConWorkPrj->Execute(_bstr_t("DELETE * FROM ["+sCalTblName+"] "), NULL, adCmdText);
		//初始计算公式中的其他变量值
 		sSQL = "SELECT * FROM ["+sCalTblName+"] ";
		pRsTmp->Open(_variant_t(sSQL), theApp.m_pConWorkPrj.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		if( pRsTmp->GetRecordCount() <= 0 )
		{	//如果第一次使用该表,则新增加一条记录.
			pRsTmp->AddNew();
		}
		else
		{	//以后只改变第一条记录的值
			pRsTmp->MoveFirst();
		}
		//外表面温度
		pRsTmp->PutCollect(_variant_t("ts"), _variant_t(dTs));
		//保温层外径,(当为复合保温时, 应代入D2)
		pRsTmp->PutCollect(_variant_t("D1"), _variant_t(dD1));
		//环境温度
		pRsTmp->PutCollect(_variant_t("ta"), _variant_t(this->Temp_env));
		//风速
		pRsTmp->PutCollect(_variant_t("W"), _variant_t(this->speed_wind));
		//沿风速方向的平壁宽度。
		pRsTmp->PutCollect(_variant_t("B"), _variant_t(this->D0));
		//保护层材料黑度
		pRsTmp->PutCollect(_variant_t("Epsilon"), _variant_t(this->hedu));
		//管道外径。
		pRsTmp->PutCollect(_variant_t("D0"), _variant_t(this->D0));
		//管道与平壁的分界值。
		pRsTmp->PutCollect(_variant_t("Dmax"), _variant_t(this->m_dMaxD0));


		pRsTmp->Update();
		pRsTmp->Close();
		 
		//
		for(m_pRsHeat_alfa->MoveFirst(); !m_pRsHeat_alfa->adoEOF; m_pRsHeat_alfa->MoveNext())
		{
			sVariableName = vtos(m_pRsHeat_alfa->GetCollect(_variant_t("Variable_Name_Desc")));
			if( !sVariableName.IsEmpty() && sVariableName.CompareNoCase("ALPHA")  )	
			{	
				//计算放热系数的中间公式。
				sTmp = vtos(m_pRsHeat_alfa->GetCollect(_variant_t("Formula")));
				sSQL = "UPDATE ["+sCalTblName+"] SET "+sVariableName+" = "+sTmp+" ";
				theApp.m_pConWorkPrj->Execute(_bstr_t(sSQL), NULL, adCmdText);
			}
		}
		sSQL =    "SELECT ("+sFormula+") AS val  FROM ["+sCalTblName+"] ";
		pRsTmpVal->Open(_variant_t(sSQL), theApp.m_pConWorkPrj.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		//放热系数值
		dTmpVal = vtof( pRsTmpVal->GetCollect(_variant_t("val"))); 
		if( dTmpVal > 0 )
		{
			this->co_alf = dTmpVal;
			return 1;				//用公式计算出了一个有效的放热系数.成功返回
		}
	}
	catch(_com_error)
	{
		return false;
	}
	return 0;
}
//功能：	
//在一个表中根据指定的数据，求差值。
//例如：根据保温层外径，与材料的类型。
//获得室内的设备或管道保温结构外表面传热系数。
//
bool CHeatPreCal::GetAlfaValue(_RecordsetPtr pRs, CString sSearchFieldName,  CString sValFieldName, double dCompareVal,double& dAlfaVal)
{
	dAlfaVal = 0.0;			//如果表中没有记录,或字段名为空时,返回值为0。
	if( pRs->adoEOF && pRs->BOF )
	{
		return false;
	}
	sValFieldName.TrimLeft();
	sValFieldName.TrimRight();
	if(	sValFieldName.IsEmpty() || sSearchFieldName.IsEmpty() )
	{
		//要查找的字段名,或关键字段名为空。
		return false;
	}
	double  dCurDia,	//当前的外保温厚。
			dPreDia,	//前一次的外保温厚
			dCurValue,	//本次查找到的值。
			dPreValue;	//前一次的值

	bool	bFirst = true;	//标记查找的次数.

	for(pRs->MoveFirst(); !pRs->adoEOF; pRs->MoveNext())
	{
		dCurDia   = vtof(pRs->GetCollect(_variant_t(sSearchFieldName)));
		dCurValue = vtof(pRs->GetCollect(_variant_t(sValFieldName)));
		if( dCurDia >= dCompareVal )
		{	
			if( bFirst )
			{
				//第一次时。
				dAlfaVal = dCurValue;
				return true;
			}
			else
			{	//求差值.
				dAlfaVal = (dCompareVal-dPreDia)/(dCurDia-dPreDia)*(dCurValue-dPreValue)+dPreValue;
				return true;
			}
		}
		else
		{
			//保存当前的值，继续往下查找。
			dPreDia = dCurDia;
			dPreValue = dCurValue;
		}
		bFirst = false;
	}
	//没有找到满足条件的记录。取最后一条记录的值。
	if(pRs->adoEOF)
	{
		dAlfaVal = dCurValue;
		return true;
	}
	return true;
}

/////////////
//功能:
//传入计算保温厚度的公式,初始各参数,并计算.
//计算成功返回1.否则为0.
//参数：
//strFormula    保温厚度计算公式.
//strTsFormula	外表面温度计算公式
//strQFormula	散热密度计算公式
//thick1		保温厚(内层。) 
//thick2		外层保温厚
short CHeatPreCal::CalcInsulThick(CString strThickFormula,CString strThickFormula2,CString strTbFormula, CString strTsFormula, CString strQFormula ,double &dThick1, double &dThick2)
{
	CString	strTmp,		//临时变量
			strSQL;		//存入SQL语句.


	double	thick1, //内保温厚
			thick2,//外保温厚
			thick, //单层保温厚.
			ts,		//外表面温度。
			tb;		//界面温度

	double  condu_r,
			condu1_r,condu2_r,condu_out,cost_q,cost_s,cost_tot,
			cost_min=0;//最少的价格.
	double  D1,
			D2;

	double nYearVal, nSeasonVal,
		   Q;			//保温结构外表面散热密度
	double QMax;		//允许最大散热密度
	bool   flg=true;	//进行散热密度的判断
	
//	double MaxTb = in_tmax*m_HotRatio;//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).

	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	if( strThickFormula.IsEmpty() )	//计算保温厚的公式为空
	{
		return 0;
	}
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		if( !GetHeatLoss(m_pRsHeatLoss, Temp_pip, nYearVal, nSeasonVal, QMax) )
		{		
		}
	}
	//

	
	_RecordsetPtr pRsTmp;		//初始当前的中间变量的记录集.
	_RecordsetPtr pRsTmpVal;	//保存结果的记录。

	CString strCalTblName = "tmpCalcInsulThick";//临时计算中间结果的表名.

	pRsTmp.CreateInstance(__uuidof(Recordset));
	pRsTmpVal.CreateInstance(__uuidof(Recordset));
	try
	{
		//先清空表中的所有记录。
//		theApp.m_pConWorkPrj->Execute(_bstr_t("DELETE * FROM ["+sCalTblName+"] "), NULL, adCmdText);
		//初始计算公式中的其他变量值
 		strSQL = "SELECT * FROM ["+strCalTblName+"] ";
		pRsTmp->Open(_variant_t(strSQL), theApp.m_pConWorkPrj.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		if( pRsTmp->GetRecordCount() <= 0 )
		{	//如果第一次使用该表,则新增加一条记录.
			pRsTmp->AddNew();
		}
		else
		{	//以后只改变第一条记录的值
			pRsTmp->MoveFirst();
		}
		//初始计算保温厚度时要用到的,各个中间变量.
		
		//保温结构外表面温度
//		pRsTmp->PutCollect(_variant_t("ts"), _variant_t(dTs));
		//复合保温内外层界面处温度
		pRsTmp->PutCollect(_variant_t("tb"), _variant_t());
		//保温层外径,(当为复合保温时, 应代入D2)
//		pRsTmp->PutCollect(_variant_t("D1"), _variant_t(dD1));
		//环境温度
		pRsTmp->PutCollect(_variant_t("ta"), _variant_t(this->Temp_env));
		//风速
		pRsTmp->PutCollect(_variant_t("W"), _variant_t(this->speed_wind));
		//沿风速方向的平壁宽度。
		pRsTmp->PutCollect(_variant_t("B"), _variant_t(this->D0));
		//保护层材料黑度
		pRsTmp->PutCollect(_variant_t("Epsilon"), _variant_t(this->hedu));
		//管道外径。
		pRsTmp->PutCollect(_variant_t("D0"), _variant_t(this->D0));
		//管道与平壁的分界值。
		pRsTmp->PutCollect(_variant_t("Dmax"), _variant_t(this->m_dMaxD0));

		//计算保温厚度时用到的其它参数。
		//
		//隔热材料制品导热系系数，W/(m*C)
		pRsTmp->PutCollect(_variant_t("lamda"), _variant_t());
		//热价
		pRsTmp->PutCollect(_variant_t("Ph"), _variant_t());
		//介质拥质系数.
		pRsTmp->PutCollect(_variant_t("Ae"), _variant_t());
		//设备和管道温度
		pRsTmp->PutCollect(_variant_t("t"), _variant_t(this->Temp_pip));
		//年运行时间
		pRsTmp->PutCollect(_variant_t("tao"), _variant_t());
		//保温层单位造价,复合保温内层单位造价
		pRsTmp->PutCollect(_variant_t("P1"), _variant_t());
		//复合保温外层单位造价
		pRsTmp->PutCollect(_variant_t("P2"), _variant_t());
		//保护层单位造价
		pRsTmp->PutCollect(_variant_t("P3"), _variant_t());
		//保温工程投资贷款年分摊率
		pRsTmp->PutCollect(_variant_t("S"), _variant_t());
		//复合保温外层外径
		pRsTmp->PutCollect(_variant_t("D2"), _variant_t());
		//保温结构外表面散热密度
		pRsTmp->PutCollect(_variant_t("Q"), _variant_t());
		//圆周率
		pRsTmp->PutCollect(_variant_t("pi"), _variant_t((double)3.1415927));
		//管道通过支吊架处的热（或冷）损失的附加系数，可取1.05 ~ 1.15
		pRsTmp->PutCollect(_variant_t("Kr"), _variant_t((double)1.1));
		//防冻结管道允许液停留时间
		pRsTmp->PutCollect(_variant_t("taofr"), _variant_t());
		//介质在管内冻结温度
		pRsTmp->PutCollect(_variant_t("tfr"), _variant_t());
		//每米管长介质体积
		pRsTmp->PutCollect(_variant_t("V"), _variant_t());
		//介质密度
		pRsTmp->PutCollect(_variant_t("ro"), _variant_t());
		//介质的比热
		pRsTmp->PutCollect(_variant_t("C"), _variant_t());
		//每米管长管壁体积
		pRsTmp->PutCollect(_variant_t("Vp"), _variant_t());
		//管材密度
		pRsTmp->PutCollect(_variant_t("rop"), _variant_t());
		//管材的比热
		pRsTmp->PutCollect(_variant_t("Cp"), _variant_t());
		//介质溶解热
		pRsTmp->PutCollect(_variant_t("Hfr"), _variant_t());
		//保温层厚度。
		pRsTmp->PutCollect(_variant_t("delta"), _variant_t());
		//复合保温内层厚度
		pRsTmp->PutCollect(_variant_t("delta1"), _variant_t());
		//复合保温外层厚度
		pRsTmp->PutCollect(_variant_t("delta2"), _variant_t());
		//
		pRsTmp->Update();
		
		//开始计算

		//计算内层保温厚的公式为空时.作为单层处理。
		if( strThickFormula2.IsEmpty() )
		{
			//计算内层保温厚的公式为空时.作为单层处理。

			//thick1内层保温厚度。thick2外层保温厚度
			//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
			//从保温厚度规则表thicknessRegular获得
			//
			for(thick=m_thick2Start;thick<=m_thick2Max;thick+=m_thick2Increment)
			{
			//	do temp_pip_one with thick,ts
				temp_pip_one(thick,ts);
				condu_r=in_a0+in_a1*(ts+Temp_pip)/2.0+in_a2*(ts+Temp_pip)*(ts+Temp_pip)/4.0;
				condu_out=co_alf;
				D1=2.0*thick+D0;

				cost_q=7.2*pi*Mhours*heat_price*Yong*fabs(Temp_pip-Temp_env)*1e-6
						/(1.0/condu_r*log(D1/D0)+2000.0/condu_out/D1);

				cost_s=(pi/4.0*(D1*D1-D0*D0)*in_price*1e-6+pi*D1*pro_price*1e-3)*S;
				cost_tot=cost_q+cost_s;



				pRsTmp->PutCollect(_variant_t("ALPHA"), _variant_t(condu_out));

				//隔热材料制品导热系系数，W/(m*C)
				pRsTmp->PutCollect(_variant_t("lamda"), _variant_t(condu_r));

				//保温层的外径
				pRsTmp->PutCollect(_variant_t("D1"), _variant_t(D1));
				
				try
				{
					//
					//保温层厚度。
					pRsTmp->PutCollect(_variant_t("delta"), _variant_t(thick));
					//保温结构外表面温度
					pRsTmp->PutCollect(_variant_t("ts"), _variant_t(ts));
					//复合保温内外层界面处温度
					tb = ts;
					pRsTmp->PutCollect(_variant_t("tb"), _variant_t(tb));

					pRsTmp->Update();//更新;

					//求外表面温度，如果是双层加上界面温度。
					//strSQL = "SELECT ("+strTsFormula+") AS TSVal FROM ["+strCalTblName+"] ";

					strSQL = "UPDATE ["+strCalTblName+"] SET ts= ("+strTsFormula+") ";
					theApp.m_pConWorkPrj->Execute(_bstr_t(strSQL), NULL, adCmdText);

					//求得本次厚度的散热密度 
					//公式:
					//从数据库中获得，根据保温结构取得相应的计算公式
					strSQL = "UPDATE ["+strCalTblName+"] SET Q=("+strQFormula+") ";
					theApp.m_pConWorkPrj->Execute(_bstr_t(strSQL), NULL, adCmdText);
				}
				catch(_com_error& e)
				{
					AfxMessageBox ( e.Description() );					
				}


				if(ts<MaxTs && (cost_tot<cost_min || fabs(cost_min)<1E-6))
				{
					flg = true;
					//选项中选择的判断散热密度,就进行比较.
					if(bIsHeatLoss)
					{
						//求得本次厚度的散热密度
						//公式:为了兼容保冷计算，温度差应取绝对值 by ligb on 2010.06.14
						//(介质温度-环境温度)/((保温层外径/2000.0/材料导热率*ln(保温层外径/管道外径)) + 1/传热系数)
						Q = fabs(Temp_pip-Temp_env)/(D1/2000.0/condu_r*log(D1/D0) + 1.0/co_alf);
						if(fabs(QMax) <= 1E-6 || Q <= QMax)
						{	//当前厚度满足允许最大散热规则.
							flg = true;
						}
						else
						{	
							flg = false;
						}
					}
					//满足条件.
					if( flg )
					{
						cost_min=cost_tot;
					}
				}
			}

		}
		else
		{
			
			//test
			//thick1内层保温厚度。thick2外层保温厚度
			//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
			//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
			//从保温厚度规则表thicknessRegular获得

//			theApp.m_pConWorkPrj->BeginTrans();	//开始事务.
			for(thick1=m_thick1Start;thick1<=m_thick1Max;thick1+=m_thick1Increment)
			{
				for(thick2=m_thick2Start;thick2<=m_thick2Max;thick2+=m_thick2Increment)
				{
					temp_pip_com(thick1,thick2,ts,tb);
					condu1_r=in_a0+in_a1*(Temp_pip+tb)/2.0+in_a2*(Temp_pip+tb)*(Temp_pip+tb)/4.0;
					condu2_r=out_a0+out_a1*(tb+ts)/2.0+out_a2*(tb+ts)*(tb+ts)/4.0;
					condu_out=co_alf;
					D1=2*thick1+D0;
					D2=2*thick2+D1;
					//test	{{
					//计算保温结构外表面传热系数.
					condu_out = CalcAlfaValue(ts, tb);
					if( fabs(condu_out) < 1e-1 )
					{	//为零时用查表所得的传热系数.
						condu_out = co_alf;
					}
					pRsTmp->PutCollect(_variant_t("ALPHA"), _variant_t(condu_out));

					//隔热材料制品导热系系数，W/(m*C)
					pRsTmp->PutCollect(_variant_t("lamda"), _variant_t(condu1_r));
					//隔热材料制品导热系系数，W/(m*C)
					pRsTmp->PutCollect(_variant_t("lamda1"), _variant_t(condu1_r));
					//隔热材料制品导热系系数，W/(m*C)
					pRsTmp->PutCollect(_variant_t("lamda2"), _variant_t(condu2_r));

					//test }}

					//复合保温结构的内层保温外径
					pRsTmp->PutCollect(_variant_t("D1"), _variant_t(D1));
					//复合保温结构的外层保温外径
					pRsTmp->PutCollect(_variant_t("D2"), _variant_t(D2));

					cost_q=7.2*pi*Mhours*heat_price*Yong*fabs(Temp_pip-Temp_env)*1e-6
							/(1.0/condu1_r*log(D1/D0)+1.0/condu2_r*log(D2/D1)+2000.0/condu_out/D2);

					cost_s=(pi/4.0*(D1*D1-D0*D0)*in_price*1e-6+pi/4.0*(D2*D2-D1*D1)*out_price*1e-6
							+pi*D2*pro_price*1e-3)*S;

					cost_tot=cost_q+cost_s;

					
					try
					{
						//
						//保温层厚度。
						pRsTmp->PutCollect(_variant_t("delta"), _variant_t(thick1));
						//复合保温内层厚度
						pRsTmp->PutCollect(_variant_t("delta1"), _variant_t(thick1));
						//复合保温外层厚度
						pRsTmp->PutCollect(_variant_t("delta2"), _variant_t(thick2));
						//保温结构外表面温度
						pRsTmp->PutCollect(_variant_t("ts"), _variant_t(ts));
						//复合保温内外层界面处温度
						pRsTmp->PutCollect(_variant_t("tb"), _variant_t(tb));

						pRsTmp->Update();//更新;

						//求外表面温度，如果是双层加上界面温度。
						//strSQL = "SELECT ("+strTsFormula+") AS TSVal FROM ["+strCalTblName+"] ";

						strSQL = "UPDATE ["+strCalTblName+"] SET ts= ("+strTsFormula+") ";
						theApp.m_pConWorkPrj->Execute(_bstr_t(strSQL), NULL, adCmdText);

						//求得本次厚度的散热密度 
						//公式:
						//从数据库中获得，根据保温结构取得相应的计算公式
						strSQL = "UPDATE ["+strCalTblName+"] SET Q=("+strQFormula+") ";
						theApp.m_pConWorkPrj->Execute(_bstr_t(strSQL), NULL, adCmdText);

						//求保温厚度
				/**/	//	strSQL = "UPDATE ["+strCalTblName+"] SET delta=("+strThickFormula+") ";
					//	theApp.m_pConWorkPrj->Execute(_bstr_t(strSQL), NULL, adCmdText);

				/*		//求外表面温度，如果是双层加上界面温度。
 						strSQL = "SELECT ("+strTsFormula+") AS TSVal FROM ["+strCalTblName+"] ";
						if(pRsTmpVal->State == adStateOpen)
						{
							pRsTmpVal->Close();
						}
						pRsTmpVal->Open(_bstr_t(strSQL), theApp.m_pConWorkPrj.GetInterfacePtr(),
										adOpenStatic, adLockOptimistic, adCmdText);
						if(pRsTmpVal->GetRecordCount() <= 0 )
						{
							//计算外表面温度失败。
							pRsTmp->PutCollect(_variant_t("ts"), _variant_t((double)35));
						}
						else
						{
							ts = vtof(pRsTmpVal->GetCollect(_T("TSVal")));
							pRsTmp->PutCollect(_variant_t("ts"), _variant_t(ts));
						}

						//求界面温度。tb
						if( !strTbFormula.IsEmpty() )
						{
							strSQL = "SELECT ("+strTbFormula+") AS TBVal FROM ["+strCalTblName+"] ";
							if(pRsTmpVal->State == adStateOpen)
							{
								pRsTmpVal->Close();
							}
							pRsTmpVal->Open(_bstr_t(strSQL), theApp.m_pConWorkPrj.GetInterfacePtr(),
											adOpenStatic, adLockOptimistic, adCmdText);
							if(pRsTmpVal->GetRecordCount() <= 0 )
							{
								//计算界面温度失败
								pRsTmp->PutCollect(_variant_t("tb"), _variant_t((double)35));
							}
							else
							{
								tb = vtof(pRsTmpVal->GetCollect(_T("TBVal")));
								pRsTmp->PutCollect(_variant_t("tb"),_variant_t(tb));
							}

						}

						//求得本次厚度的散热密度
						//公式:
						//从数据库中获得，根据保温结构取得相应的计算公式
						strSQL = "SELECT ("+strQFormula+") AS QVal FROM ["+strCalTblName+"] ";
						if(pRsTmpVal->State == adStateOpen)
						{
							pRsTmpVal->Close();
						}
						pRsTmpVal->Open(_bstr_t(strSQL), theApp.m_pConWorkPrj.GetInterfacePtr(),
										adOpenStatic, adLockOptimistic, adCmdText);
						if(pRsTmpVal->GetRecordCount() <= 0 )
						{
							//计算散热密度失败。
							//赋一个默认的值,查找数据库所得。
							pRsTmp->PutCollect(_variant_t("Q"), _variant_t(QMax));
						}
						else
						{
							Q = vtof(pRsTmpVal->GetCollect(_T("QVal")));
							pRsTmp->PutCollect(_variant_t("Q"), _variant_t(Q));
						}
						
						//求保温厚度
						strSQL = "SELECT ("+strThickFormula+") AS Thick FROM ["+strCalTblName+"] ";
						if( pRsTmpVal->State == adStateOpen)
						{
							pRsTmpVal->Close();
						}

						pRsTmpVal->Open(_bstr_t(strSQL), theApp.m_pConWorkPrj.GetInterfacePtr(),
									adOpenStatic, adLockOptimistic, adCmdText);
						if( pRsTmpVal->GetRecordCount() <= 0 )
						{
							return 0; 
						}
						dThick1 = vtof(pRsTmpVal->GetCollect(_T("Thick")));
			/*	*/	
					}
					catch(_com_error& e)
					{
						AfxMessageBox(e.Description());
						return 0;
					}
					//tb<350)
					if((ts<MaxTs && tb<MaxTb) && (fabs(cost_min)<1E-6 || cost_min>cost_tot))
					{
						flg = true;
						//选项中选择的判断散热密度,就进行比较.
						if(bIsHeatLoss)
						{
							//求得本次厚度的散热密度
							//公式:
							//(介质温度-环境温度)/((外保温层外径/2000.0)*(1.0/内保温材料导热率*ln(内保温层外径/管道外径) + 1.0/外保温材料导热率*ln(外保温层外径/内保温层外径) ) + 1/传热系数)
							Q = (Temp_pip-Temp_env) / ( (D2/2000.0) * (1.0/condu1_r*log(D1/D0)+1.0/condu2_r*log(D2/D1)) + 1.0/co_alf );

							if(fabs(QMax) <= 1E-6 || Q <= QMax)
							{	//当前厚度满足允许最大散热规则.
								flg = true;
							}
							else
							{	
								flg = false;
							}
						}
						//满足条件.
						if( flg )
						{
							cost_min=cost_tot;
							dThick1=thick1;
							dThick2=thick2;
						//	ts_resu=ts;
						//	tb_resu=tb;//
						}
					}
				}
			}
	//		theApp.m_pConWorkPrj->CommitTrans();	//提交事务.
	//		theApp.m_pConWorkPrj->RollbackTrans();
		}//双层保温

	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}
	return 1;
}
//功能：
//获得计算当前记录的计算保温厚的公式.计算外表面温度的公式,计算散热密度的公式
//	nFormulaIndex	//结构对应的公式索引.
//	strFormula,		//返回计算公式.
int CHeatPreCal::GetCalcFormula(int nFormulaIndex, CString &strFormula, CString& strTsFormula, CString& strQFormula)
{
	CString	strTmp,		//临时变量
			strSQL;		//存入SQL语句.

	int		nMethod=0;	//计算方法索引。
	try
	{
		//在原始数据库中获得计算方法的索引。
		nMethod = vtoi(IRecCalc->GetCollect(_T("c_Method_Index")));

		if(m_pRsCalcThick->GetRecordCount() <= 0 )
		{
			return 0;
		}
		m_pRsCalcThick->Filter = (long)adFilterNone;
		strSQL.Format("InsulationCalcMethodIndex=%d AND Index=%d",nMethod,nFormulaIndex);
		m_pRsCalcThick->Filter = variant_t(strSQL);
		if( !m_pRsCalcThick->adoEOF )		 
		{
			//找到保温计算方法对应的公式.
			strFormula = vtos(m_pRsCalcThick->GetCollect(_T("Formula")));

			if( strFormula.IsEmpty() )
			{  
				//计算的公式为空
				return 0;
			}
		}
		else
		{
			return 0;		//找不到保温计算方法所对应的公式.
		}
		//查找求外表面ts值的计算公式
		m_pRsCalcThick->Filter = (long)adFilterNone;
		//设为一个常数,规定20为计算外表面温度的方法的索引.
		nMethod = 20;			
		strSQL.Format("InsulationCalcMethodIndex=%d AND Index=%d",nMethod,nFormulaIndex);
		m_pRsCalcThick->Filter = _variant_t(strSQL);
		if(m_pRsCalcThick->GetRecordCount() <= 0 )
		{
			return 0;
		}
		strTsFormula = vtos(m_pRsCalcThick->GetCollect(_T("Formula")));
		//
		//查找计算散热密度(Q)的计算公式
		m_pRsCalcThick->Filter = (long)adFilterNone;
		//设为一个常数,规定21为计算散热密度的方法的索引.
		nMethod = 21;	
		strSQL.Format("InsulationCalcMethodIndex=%d AND Index=%d",nMethod, nFormulaIndex);
		m_pRsCalcThick->Filter = _variant_t(strSQL);
		if(m_pRsCalcThick->GetRecordCount() <= 0 )
		{
			return 0;
		}
		strQFormula = vtos(m_pRsCalcThick->GetCollect(_variant_t("Formula")));

		//取消所有的过滤
		m_pRsCalcThick->Filter = (long)adFilterNone;
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	
	return 1;
}

//-----------------------------------------------------
//功能：
//在管道参数库中根据管道外径和壁厚，获得相应的管重，和水重.
short CHeatPreCal::GetPipeWeithAndWaterWeith(_RecordsetPtr &pRs, const double dPipe_Dw, const double dPipe_S, CString strField1, double &dValue1,CString strField2,  double &dValue2)
{
	try
	{
		if(pRs == NULL || pRs->State == adStateClosed || pRs->GetRecordCount() <= 0 )
		{	//记录集为空
			return 0;
		}

		CString strFilter;
		strFilter.Format("PIPE_DW=%f AND PIPE_S=%f",dPipe_Dw, dPipe_S);
		pRs->Filter = (long)adFilterNone;		//删除当前筛选条件并恢复查看的所有记录。

		pRs->Filter = _variant_t(strFilter);	//筛选出满足外径和壁厚的记录.

		while(!pRs->adoEOF)
		{
			try
			{
				//定位到外径和壁厚对应的记录
				//
				dValue1 = vtof(pRs->GetCollect(_variant_t(strField1) ));
				//取出给定字段名对应的值
				dValue2 = vtof(pRs->GetCollect(_variant_t(strField2) ));

				pRs->Filter = (long)adFilterNone;		//删除当前筛选条件并恢复查看的所有记录。			
				return 1;
			}
			catch(_com_error &e)
			{
				CString TempStr;
				TempStr.Format(_T("原始数据序号为%d的记录在读取\"管道参数库(A_PIPE)\"中时出错\r\n"),
							   m_nID );
				if(Exception::GetAdditiveInfo())
				{
					TempStr+=Exception::GetAdditiveInfo();
				}
				ExceptionInfo(TempStr);

				Exception::SetAdditiveInfo(NULL);
				ReportExceptionErrorLV2(e);

				IRecCalc->MoveNext();
				continue;
			}
		}
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		pRs->Filter = (long)adFilterNone;		//删除当前筛选条件并恢复查看的所有记录。			
		return 0;
	}
	//没有长到符合外径和壁厚的记录
	pRs->Filter = (long)adFilterNone;		//删除当前筛选条件并恢复查看的所有记录。			
	return 0;
}

//------------------------------------------------------------------
// DATE         :[2005/08/02]
// Parameter(s) :
// Return       :
// Remark       :在规范库中设置常数表中指定项的值.
//------------------------------------------------------------------
BOOL CHeatPreCal::SetConstantDefine(CString strTblName, CString ConstName, _variant_t ConstValue)
{
	_RecordsetPtr pRs;
	pRs.CreateInstance(__uuidof(Recordset));

	CString strSQL;
	strTblName.TrimRight();
	strSQL = "SELECT * FROM ["+strTblName+"] ";
	try
	{
		pRs->Open(_variant_t(strSQL), theApp.m_pConnectionCODE.GetInterfacePtr(), 
			adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->GetRecordCount() <= 0)
		{
			return false; 
		}
		strSQL = "ConstantDefine_Name='"+ConstName+"' " ;
		pRs->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if( !pRs->adoEOF  )
		{
			pRs->PutCollect(_variant_t("ConstantDefine_VALUE"),ConstValue);
			pRs->Update();
			return true;
		}
		else
		{
			return false;
		}
	}
	catch(_com_error& e)
	{
		ReportExceptionError( e);
		return false;
	}
	return TRUE;
}

//在标准库中查找保温材料
BOOL CHeatPreCal::FindStandardMat(_RecordsetPtr& pRsCurMat,CString strMaterialName)
{
	try
	{
		
		if (NULL == m_pRs_CodeMat || m_pRs_CodeMat->State != adStateOpen || (m_pRs_CodeMat->adoEOF && m_pRs_CodeMat->BOF))
		{
			return FALSE;
		}
		m_pRs_CodeMat->MoveFirst();
		if (!LocateFor(m_pRs_CodeMat,_variant_t("MAT_NAME"),CFoxBase::EQUAL, _variant_t(strMaterialName)))
		{
			return FALSE;
		}
		pRsCurMat->AddNew();
		pRsCurMat->Update();
		for (long i=0, nFieldCount=m_pRs_CodeMat->Fields->GetCount(); i < nFieldCount; i++)
		{
			pRsCurMat->PutCollect(_variant_t(i),m_pRs_CodeMat->GetCollect(_variant_t(i)));
		}
		pRsCurMat->PutCollect(_variant_t("EnginID"),_variant_t(m_ProjectName));
		pRsCurMat->Update();
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;	
	}
	return TRUE;
}

//------------------------------------------------------------------
// DATE         :[2005/09/13]
// Parameter(s) :
// Return       :
// Remark       :根据介质名称获得其介质的密度
//------------------------------------------------------------------
double CHeatPreCal::GetMediumDensity(const CString strMediumName, CString* pMediumName)
{
	double	dMediumDensity = 0.0;   //介质密度
	CString strSQL;					//SQL语句		
	
	if (strMediumName.IsEmpty())
	{
		return 0.0;
	}
	try
	{
		_RecordsetPtr pRs;
		pRs.CreateInstance(__uuidof(Recordset));
		strSQL = "SELECT * FROM [MediumSpec] WHERE MediumID IN (SELECT MediumID FROM [MediumAlias] WHERE MediumAlias='"+strMediumName+"') ";
		pRs->Open(_variant_t(strSQL), theApp.m_pConMedium.GetInterfacePtr(),
			adOpenStatic, adLockOptimistic, adCmdText);
		if (pRs->adoEOF && pRs->BOF)
		{
			//不用别名,直接在介质密度表中查找.
			pRs->Close();
			strSQL = "SELECT * FROM [MediumSpec] WHERE MediumName='"+strMediumName+"' ";
			pRs->Open(_variant_t(strSQL), theApp.m_pConMedium.GetInterfacePtr(),
				adOpenStatic, adLockOptimistic, adCmdText);
			if (pRs->adoEOF && pRs->BOF)
			{
				return 0.0;
			}
		}//
		dMediumDensity  = vtof( pRs->GetCollect(_variant_t("Density")) );
		if ( pMediumName != NULL )
		{
			*pMediumName = vtos( pRs->GetCollect(_variant_t("MediumName")));
		}
	}
	catch (_com_error& e) {
		AfxMessageBox(e.Description());
		return 0.0;
	}
	return dMediumDensity;
}

//------------------------------------------------------------------
// DATE         :[2005/09/15]
// Parameter(s) :dw:			管道的外径
//				:s:				壁厚
//				:dMediumDensity:管道内介质的密度 
//				:temp:			介质的温度
//				:dInsuWei:		保温材料的重量
// Return       :根据标准计算出的间距
// Remark       :使用动态库计算支吊架的间距 
//------------------------------------------------------------------
double CHeatPreCal::GetSupportSpan(const double dw, const double s, const CString strMaterialName,const double dPressure,const double dMediumDensity, const double temp, const double dInsuWei)
{
	double dLmax = 0.0;		//支吊架的间距

	CString strDlgCaption;	//对话框的标题
	double	LmaxR=0.0;
	double	LmaxS=0.0;
	double	fQ=0.0;
	double	fQp=0.0;
	int		TableID = -1;

	BOOL	bCalFlag;			//动态库计算的返回标志
	BOOL	bPreviousFlag=-120;	//前一次的计算标志
	BOOL	pg;
	if (3.92 < dPressure)	//压力
		pg = true;
	else
		pg = false;

	CalSupportSpan.SetMaterialCODE( EDIBgbl::sCur_CodeNO );	// 本工程设置的材料规范号

	for (int nCont=1; nCont <= 4; nCont++)		//如果计算没出错,则直接返回,否则增加了材料的属性,则重新计算一次
	{
		//计算支吊架间距
		if (0) 
		{
		}else
		{	//默认电力行业
			bCalFlag = CalSupportSpan.CalcSupportSpan(dw,s,dMediumDensity,temp,dInsuWei, 
							strMaterialName,pg,2,LmaxR,LmaxS,fQ,fQp); 
		}

		dLmax = min(LmaxR,LmaxS);		//刚性条件与强度条件中最小的一个值

		if (-1 == bCalFlag)
		{
			//指定的材料没有应力系数
			TableID = 0; 
			strDlgCaption.Format(_T("请输入 %s 在 %f  ℃下的许用应力(MPa):"),strMaterialName,temp);
			InputMaterialParameter(strDlgCaption,strMaterialName,TableID); 
		}else if (-2 == bCalFlag)
		{ 
			//指定的材料没有弹性模量
			TableID = 1;
			strDlgCaption.Format(_T("请输入 %s 在 %f  ℃下的弹性模量(KN/mm2)"),strMaterialName,temp);
			InputMaterialParameter(strDlgCaption,strMaterialName,TableID); 
		}else if (-3 == bCalFlag) 
		{
			//指定的材料没有环向焊缝应力系数
			TableID = 4;
			strDlgCaption.Format(_T("请输入 %s 的环向焊缝应力系数:"),strMaterialName);
			//输入数据 
			InputMaterialParameter(strDlgCaption,strMaterialName,TableID);
		}else
		{ 
			//计算成功
			break; 
		}
		//

		if(bPreviousFlag != bCalFlag)		//如果没有相应材料的属性，一种属性最多输入两次。
			bPreviousFlag = bCalFlag;
		else
			break;
	}
	return (dLmax<0)?0:dLmax;
}

//------------------------------------------------------------------
// DATE         :[2006/01/09]
// Parameter(s) :strVariableName[in] 气象参数的属性名
//				:dRValue[out] 对应的气象值
// Return       :
// Remark       : 获得气象属性值
//------------------------------------------------------------------
void CHeatPreCal::GetWeatherProperty(CString strVariableName, double &dRValue)
{
	if ( strVariableName.IsEmpty() )
		return;
	try
	{
		if ( !OpenAssistantProjectTable( m_pRsWeather, "Ta_Variable" ) )
		{
			return;
		}
		if ( m_pRsWeather->adoEOF && m_pRsWeather->BOF )
		{
			return;
		}

		CString strSQL;
		strSQL = " Ta_Variable_Name='"+strVariableName+"' ";
		
		m_pRsWeather->MoveFirst();
		m_pRsWeather->Find( _bstr_t( strSQL ), NULL, adSearchForward );
		if ( !m_pRsWeather->adoEOF && !m_pRsWeather->BOF )
		{
			dRValue = vtof( m_pRsWeather->GetCollect( _variant_t( "Ta_Variable_VALUE" ) ) );
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox( e.Description() );
		return ;
	}
}
//------------------------------------------------------------------
// Parameter(s) : m_nID[in] 记录号
// Return       : 埋地保温类型
// Remark       : 当前保温记录的埋地类型
//------------------------------------------------------------------
//  
BOOL CHeatPreCal::GetSubterraneanTypeAndPipeCount(const long m_nID, short& nType, short& nPipeCount) const
{
	nType = 0;
	nPipeCount = 1;
	try
	{
		if (NULL == m_pRsSubterranean || (m_pRsSubterranean->GetState() == adStateClosed) || (m_pRsSubterranean->GetadoEOF() && m_pRsSubterranean->GetBOF()))
		{
			return FALSE;
		}
		CString strSQL;
		m_pRsSubterranean->MoveFirst();
		strSQL.Format("ID=%d", m_nID);
		m_pRsSubterranean->Find(_bstr_t(strSQL), 0, adSearchForward);
		if (!m_pRsSubterranean->GetadoEOF())
		{
			nType			= vtoi(m_pRsSubterranean->GetCollect(_variant_t(_T("c_lay_mark"))));
			nPipeCount      = vtoi(m_pRsSubterranean->GetCollect(_variant_t(_T("c_Pipe_Count"))));
		}
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE;
}
