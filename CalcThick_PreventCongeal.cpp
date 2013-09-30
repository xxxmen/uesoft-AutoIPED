// CalcThick_PreventCongeal.cpp: implementation of the CCalcThick_PreventCongeal class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcThick_PreventCongeal.h"
#include "vtot.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThick_PreventCongeal::CCalcThick_PreventCongeal()
{

}

CCalcThick_PreventCongeal::~CCalcThick_PreventCongeal()
{

}


//------------------------------------------------------------------                
// DATE         :[2005/06/02]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :防冻结保温(热平衡)法计算保温层厚
//------------------------------------------------------------------
short CCalcThick_PreventCongeal::CalcPipe_One(double &thick_resu, double &ts_resu)
{
	CalcThickFormula(thick_resu,ts_resu);
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/02]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :防冻结保温(热平衡)法计算保温层厚
//------------------------------------------------------------------
short CCalcThick_PreventCongeal::CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	thick1_resu = thick2_resu = tb_resu = ts_resu = 0;
	//只计算外层的保温厚.
	CalcThickFormula(thick2_resu, ts_resu);
	
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/02]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :防冻结保温(热平衡)法计算保温层厚
//------------------------------------------------------------------
short CCalcThick_PreventCongeal::CalcPlain_One(double &thick_resu, double &ts_resu)
{
	CalcThickFormula(thick_resu,ts_resu, FALSE);
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/06/02]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :防冻结保温(热平衡)法计算保温层厚
//------------------------------------------------------------------
short CCalcThick_PreventCongeal::CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	thick1_resu = thick2_resu = tb_resu = ts_resu = 0 ;

	//只计算外层的保温厚．
	CalcThickFormula(thick2_resu,ts_resu, FALSE);

	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/02]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :防冻计算的公式，通用这一个公式。
//------------------------------------------------------------------
short CCalcThick_PreventCongeal::CalcThickFormula( double& thick_resu, double& ts_resu, BOOL bIsPipe)
{
	double ts;	//外表面温度
	double dCurQ = 0;
	bool   flg = true;
	thick_resu = ts_resu = 0;

	double QMax;
	//查表获得介质温度允许的最大散热密度
	GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax);

	//获得防冻计算的原始数据
	GetCongealOriginalData(m_nID);

	double Diff;					//差值
	double dMaxDiff = DZero;		//最小的差值
	//
	for (thick2=m_thick2Start; thick2 <= m_thick2Max; thick2 += m_thick2Increment)
	{
		//
		if (bIsPipe)
		{
			CalcPipe_One_TsTemp(thick2,ts);
		}
		else
		{
			CalcPlain_One_TsTemp(thick2,ts);
		}
		D1 = D0 + 2 * thick2;			//保温层外径
		lamda = GetLamda(t,ts);			//计算导热系数
		//根据不同的行业规范计算放热系数
		ALPHA = GetPipeAlpha(ts,D1);
		//
		Diff = log(D1/D0)+2*lamda/D1/ALPHA-2*lamda*(Kr*taofr/(2*(t-tfr)*(D0*D0/4*ro*C+(D0+pi_thi)*pi_thi*rop*Cp)/(t+tfr-2*ta)+0.25*D0*D0/4*ro*Hfr/(tfr-ta)));
		
		Calc_Q_PipeOne(dCurQ);
		if ( ts < MaxTs && (dMaxDiff <= DZero || fabs(Diff) < fabs(dMaxDiff)) )
		{
			flg = true;
			if (bIsHeatLoss)
			{
				Calc_Q_PipeOne(dCurQ);
			}
			if ( flg )
			{
				dMaxDiff	= Diff;		//记下当前最小的差值
				thick_resu	= thick2;	//保温厚
				ts_resu		= ts;
				Calc_Q_PipeOne(dQ);
			}
		}
	}
	return 1;
}

//------------------------------------------------------------------
// DATE         :[2005/06/06]
// Author       :ZSY
// Parameter(s) :当前记录的序号nID
// Return       :
// Remark       :获得当前计算记录的防冻保温的参数.
//------------------------------------------------------------------
BOOL CCalcThick_PreventCongeal::GetCongealOriginalData(long nID)
{
	try
	{	
		if (m_pRsCongeal==NULL || m_pRsCongeal->State == adStateClosed || m_pRsCongeal->GetRecordCount() <= 0 )
		{
			return FALSE;
		}
		CString	strSQL;		//SQL语句
		strSQL.Format("ID=%d",nID);

		m_pRsCongeal->MoveFirst();
		m_pRsCongeal->Find(_bstr_t(strSQL), 0, adSearchForward);

		//存在指定序号的防冻保温计算的参数
		if ( !m_pRsCongeal->adoEOF )
		{
			//管道通过支吊架处的热(或冷)损失的附加系数
			Kr		= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Heat_Loss_Data")) );
			//防冻结管道允许液体停留时间(小时)
			taofr	= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Stop_Time")) );
			taofr *= 3600;
			//介质在管内冻结温度
			tfr		= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Medium_Congeal_Temp")) );
			//介质密度
			ro		= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Medium_Density")) );
			//介质比热
			C		= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Medium_Heat")) );
			//管材密度
			rop		= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Pipe_Density")) );
			//管材比热
			Cp		= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Pipe_Heat")) );
			//介质融解热
			Hfr		= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Medium_Melt_Heat")) );
			//流量
			dFlux	= vtof( m_pRsCongeal->GetCollect(_variant_t("c_Flux")) );

			//每米管长介质体积，m3/m;	
			V	=  pi / 4 * pow(D0 - 2 * pi_thi, 2) / 1000000;
			//每米管长管壁体积,m3/m;
//			Vp	= pi / 4 * (D0 * D0 - pow(D0 - 2 * pi_thi, 2)) / 1000000;			
			Vp	= pi * pow(D0/2.0,2) / 1000000 - V;
		}
	}
	catch (_com_error)
	{
		return FALSE;
	}

	return TRUE;
}
