// CalcThickBase.cpp: implementation of the CCalcThickBase class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcThickBase.h"
#include <math.h>

#include "edibgbl.h"


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#define		PREVENT_SCALD    3			//防烫伤

extern CAutoIPEDApp	theApp;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThickBase::CCalcThickBase()
{
}

CCalcThickBase::~CCalcThickBase()
{

}
//初始计算放热系数时用到的变量
#define INIT_ALPHADATA(CLASSN) \
	CLASSN::dA_ta	= ta;\
	CLASSN::dA_ts	= ts;\
	CLASSN::dA_W	= speed_wind;\
	CLASSN::dA_B	= B;\
	CLASSN::dA_hedu	= hedu;\
	CLASSN::dA_AlphaIndex = nAlphaIndex;

//------------------------------------------------------------------                
// DATE         :[2005/04/25]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :计算平壁单层 ,外表面温度的值
//------------------------------------------------------------------	
short CCalcThickBase::CalcPlain_One_TsTemp(const double delta, double &dResTs)
{	
	double ts;	// 临时的外表面温度 [2005/06/07]
	int nBreak = 0;
	dResTs = ts = ( t - ta ) / 2 + ta;		//外表面温度赋初值。
	while (TRUE)
	{
		// 根据不同的行业规范计算放热系数
		ALPHA = GetPlainAlpha(ts);

		//根据改变的外表面温度计算导热系数
		lamda = GetLamda( t, ts );

		//保温层温差
		float tta=(t-ta)/(1+lamda/ALPHA*1000/delta);
		if(t>=ta)//保温
			ts = t-tta;
		else //保冷
			ts = t+tta;
		if (fabs(ts-dResTs) < TsDiff || nBreak >= MaxCycCount)
		{
			break;
		}
		dResTs = ts;	//
		nBreak++;
	}
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/04/25]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :计算管道单层,外表面温度的值
//------------------------------------------------------------------
short CCalcThickBase::CalcPipe_One_TsTemp(const double delta, double &dResTs)
{	
	double ts;	// 临时的外表面温度 [2005/06/07]
	int    nBreak = 0;
	D1 = D0 + 2 * delta;			//保温层外径。
	
	dResTs = ts = ( t - ta ) / 2 + ta;		//外表面温度赋初值。

	while (TRUE)
	{
		//根据不同的行业规范计算放热系数
		ALPHA = GetPipeAlpha(ts, D1);

		//根据改变的外表面温度计算导热系数
		lamda = GetLamda( t, ts );

		//保温层温差
		float tta=(t-ta)/(1+lamda/ALPHA*2000/log(D1/D0)/D1);
		ts = t-tta;
		
		if (fabs(ts-dResTs) < TsDiff || nBreak >= MaxCycCount)
		{
			break;
		}
		dResTs = ts;	//
		nBreak++;
	}
	return 1;
}

// 计算埋地管道的表面温度
void CCalcThickBase::CalcSubPipe_TsTemp(const double dPipeTemp, const double dDW, const double delta, double &dResTs, short nMark)
{
	double ts;//保温层表面温度
	double dD1;//保温层外径
	int	   nBreak = 0;

	dD1 = dDW + 2 * delta;
	dResTs = (dPipeTemp + ta) / 2.0 + ta;
	while ( TRUE )
	{
		ALPHA = GetPipeAlpha(ts, dD1);
		lamda = GetLamdaA(dPipeTemp, ts, nMark);

		//保温层温差
		float tta=(t-ta)/(1+lamda/ALPHA*2000/log(dD1 / dDW)/D1);
		if(t>=ta)//保温
			ts = t-tta;
		else //保冷
			ts = t+tta;
		if (fabs(ts - dResTs) < TsDiff || nBreak >= MaxCycCount)
		{
			break;
		}
		dResTs = ts;
		
	}
}
//------------------------------------------------------------------                
// DATE         :[2005/04/25]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :计算平壁双层,外表面和界面处的温度的值
//------------------------------------------------------------------
short CCalcThickBase::CalcPlain_Com_TsTb(const double delta1, const double delta2, double &dResTs, double &dResTb, int flg)
{	
	double ts;	// 临时的外表面温度 [2005/06/07]
	double tb;	// 临时的界面温度	[2005/06/07]
	int    nBreak = 0;
	if( flg == 1 )		//不求ts(表面温度)
	{
		//根据保温厚度和外表面温度,计算界面温度;	
		ts		= dResTs;						//外表面温度为已知
		dResTb	= tb = (t-ta)/3.0*2.0+ta;		//界面处的温度赋初值。
	}
	else
	{
		dResTs = ts = (t-ta)/3.0+ta;		//外表面温度赋初值。
		dResTb = tb = (t-ta)/3.0*2.0+ta;	//界面处的温度赋初值。
	}
	while (TRUE)
	{
		//根据不同的行业规范计算放热系数
		ALPHA = GetPlainAlpha(ts);
		
		//根据改变的外表面温度和界面温度计算导热系数
		lamda1 = GetLamda1( t, tb );
		lamda2 = GetLamda2( tb, ts);
		
		if( flg == 1 )
		{
			//根据保温厚度和外表面温度,计算界面温度;
			//外表面温度为已知的,不再计算.
		}
		else
		{
			//计算外表面温度的值。
			ts = (delta1/1000/lamda1*ta+delta2/1000/lamda2*ta+1/ALPHA*t)/(delta1/1000/lamda1+delta2/1000/lamda2+1/ALPHA);
		}

		//界面温度
		tb = (delta1/1000/lamda1*ta+delta2/1000/lamda2*t+1/ALPHA*t)/(delta1/1000/lamda1+delta2/1000/lamda2+1/ALPHA);
				
		//当相邻两次的结果之差不大于一个常数（0.1）时，返回。
		if ((fabs(dResTs-ts)<TsDiff) && (fabs(dResTb-tb)<TsDiff) || nBreak >= MaxCycCount)
		{
			break;
		}
		dResTs = ts;
		dResTb = tb;
		nBreak++;
	}
	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/25]
// Author       :ZSY
// Parameter(s) :flg = 1 时,根据保温厚度和外表面温度,计算界面温度;
//				:否则,根据保温厚度,计算界面温度和外表面温度.
// Return       :
// Remark       :计算管道双层,外表面和界面处的温度的值
//------------------------------------------------------------------
short CCalcThickBase::CalcPipe_Com_TsTb(const double delta1, const double delta2, double &dResTs, double &dResTb,int flg)
{
	double ts;	// 临时的外表面温度 [2005/06/07]
	double tb;	// 临时的界面温度	[2005/06/07]
	int    nBreak = 0;

	if( flg == 1 )		//不求ts(表面温度)
	{
		//根据保温厚度和外表面温度,计算界面温度;	
		ts		= dResTs;						//外表面温度为已知
		dResTb	= tb = (t-ta)/3.0*2.0+ta;		//界面处的温度赋初值.
	}
	else
	{
		dResTs = ts = (t-ta)/3.0+ta;		//外表面温度赋初值。
		dResTb = tb = (t-ta)/3.0*2.0+ta;	//界面处的温度赋初值。
	}
	D1 = D0 + 2.0 * delta1;
	D2 = D1 + 2.0 * delta2;
	while (TRUE)
	{
		//根据不同的行业规范计算放热系数
		ALPHA = GetPipeAlpha(ts,D2);

		//根据改变的外表面温度和界面温度计算导热系数
		lamda1 = GetLamda1( t, tb );
		lamda2 = GetLamda2( tb, ts);
	
		//计算外表面温度的值。
		if( flg == 1 )
		{
			//根据保温厚度和外表面温度,计算界面温度
			//外表面温度为已知的,不再计算
		}
		else
		{
			ts = (log(D1/D0)/lamda1*ta+log(D2/D1)/lamda2*ta+2000/ALPHA/D2*t)/(log(D1/D0)/lamda1+log(D2/D1)/lamda2+2000/ALPHA/D2);
		}
		//界面温度
		tb = (log(D1/D0)/lamda1*ta+log(D2/D1)/lamda2*t+2000/ALPHA/D2*t)/(log(D1/D0)/lamda1+log(D2/D1)/lamda2+2000/ALPHA/D2);
				
		//当相邻两次的结果之差不大于一个常数（0.1）时，返回。
		if ((fabs(dResTs-ts)<TsDiff) && (fabs(dResTb-tb)<TsDiff) || nBreak >= MaxCycCount)
		{
			break;
		}
		dResTs = ts;
		dResTb = tb;
		nBreak++;
	}
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :计算保温厚度，平面双层。
//------------------------------------------------------------------
short CCalcThickBase::CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :计算保温厚度，平面单层。
//------------------------------------------------------------------
short CCalcThickBase::CalcPlain_One(double &thick_resu, double &ts_resu)
{
	return 1;
}	

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :计算保温厚度，保温对象为管道，双层
//------------------------------------------------------------------
short CCalcThickBase::CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :计算保温厚度，保温对象为管道，单层
//------------------------------------------------------------------
short CCalcThickBase::CalcPipe_One(double &thick_resu, double &ts_resu)
{
	return 1;	
}

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :无
// Return       :
// Remark       :从数据库中取出常数,存入类成员变量中.
//
				//thick1内层保温厚度。thick2外层保温厚度
				//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
				//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
				//从保温厚度规则表（thicknessRegular）中取出上面的值。
//------------------------------------------------------------------
short CCalcThickBase::GetThicknessRegular()
{
	CString strSQL,strMessage;
	_variant_t var;
	_RecordsetPtr pRs;
	pRs.CreateInstance(__uuidof(Recordset));
	try
	{		
		strSQL = "SELECT * FROM [thicknessRegular] WHERE EnginID='"+EDIBgbl::SelPrjID+"' ";
		pRs->Open(_bstr_t(strSQL), m_pConAutoIPED.GetInterfacePtr(),
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
			if( (m_thick1Max<m_thick1Start) )
			{
				strMessage.Format("内层保温允许最大厚度%d mm小于内层保温允许最小厚度%d mm　。",m_thick1Max,m_thick1Start);
				AfxMessageBox(strMessage);
				return 0;
			}
			if(m_thick2Max < m_thick2Start)
			{
				strMessage.Format("外层保温允许最大厚度%d mm小于外层保温允许最小厚度%d mm　。",m_thick2Max,m_thick2Start);
				AfxMessageBox(strMessage);
				return 0;
			}
			if( m_thick1Increment <= DZero )
			{
				strMessage.Format("内层保温厚度最小增量%d mm太小。",m_thick1Increment);
				AfxMessageBox(strMessage);
				return 0;
			}
			if( m_thick2Increment <= DZero )
			{
				strMessage.Format("外层保温厚度最小增量%d mm太小.",m_thick2Increment);
				AfxMessageBox(strMessage);
				return 0;
			}
			if( m_thick1Start <= DZero )
			{
				strMessage.Format("内层保温允许最小厚度%d mm太小。",m_thick1Start);
				AfxMessageBox(strMessage);
				return 0;
			}
			if( m_thick2Start <= DZero )
			{
				strMessage.Format("外层保温允许最小厚度%d mm太小。",m_thick2Start);
				AfxMessageBox(strMessage);
				return 0;
			}
			if( m_thick1Max <= DZero )
			{
				strMessage.Format("内层保温允许最大厚度%d mm太小。",m_thick1Max);
				AfxMessageBox(strMessage);
				return 0;
			}
			if( m_thick2Max <= DZero )
			{
				strMessage.Format("外层保温允许最大厚度%d mm太小。",m_thick2Max);
				AfxMessageBox(strMessage);
				return 0;
			}
		}
		else
		{
			if( IDNO == AfxMessageBox("保温厚度规则库中没有记录，是否按默认的数据进行计算？",MB_YESNO) )
			{
				return 0;
			}
			m_thick1Start = 30;		//内层保温允许最小厚度	
			m_thick1Max= 300;		//内层保温允许最大厚度
			m_thick1Increment=10;	//内层保温厚度最小增量

			m_thick2Start = 30;		//外层保温允许最小厚度
			m_thick2Max = 300;		//外层保温允许最大厚度
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
		return 0;
	}
	return	1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/25]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :根据介质温度,取保温结构外表面允许最大散热密度
//------------------------------------------------------------------
bool CCalcThickBase::GetHeatLoss(_RecordsetPtr &pRs, double &t, double &YearRun, double &SeasonRun, double &QMax)
{
	try
	{		
		YearRun = SeasonRun = QMax= 0;
		if( pRs->adoEOF && pRs->BOF )
		{
			//没有记录.
			return false;
		}
			
		YearRun=GetMaxHeatLoss(pRs,t,"YearRun");
		SeasonRun=GetMaxHeatLoss(pRs,t,"SeasonRun");
		if( hour_work )
		{	//年运行工况
			QMax = YearRun;
		}
		else
		{	//季节运行工况
			QMax = SeasonRun;
		}
		return true;
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return false;
	}	
}


//------------------------------------------------------------------                
// DATE         :[2005/04/28]
// Author       :ZSY
// Parameter(s) :用于存放计算所得的散热密度.(其他成员变量在调用之前先初始化)
// Return       :成功返回1,否则返回0
// Remark       :计算保温结构外表面散热密度.(管道.单层)
//------------------------------------------------------------------
short CCalcThickBase::Calc_Q_PipeOne(double &Q)
{
	//求得本次厚度的散热密度
	//公式:
	//(介质温度-环境温度)/((保温层外径/2000.0/材料导热率*ln(保温层外径/管道外径)) + 1/传热系数)
	Q = fabs( t - ta ) / ( D1 / 2000.0 / lamda * log( D1 / D0 ) + 1.0 / ALPHA );

	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/28]
// Author       :ZSY
// Parameter(s) :用于存放计算所得的散热密度.(其他成员变量在调用之前先初始化)
// Return       :成功返回1,否则返回0
// Remark       :计算保温结构外表面散热密度.(管道.双层)
//------------------------------------------------------------------
short CCalcThickBase::Calc_Q_PipeCom(double &Q)
{
	//求得本次厚度的散热密度
	//公式:
	//(介质温度-环境温度)/((外保温层外径/2000.0)*(1.0/内保温材料导热率*ln(内保温层外径/管道外径) + 1.0/外保温材料导热率*ln(外保温层外径/内保温层外径) ) + 1/传热系数)
	Q = fabs( t - ta ) / ( (D2 / 2000.0) * (1.0 / lamda1 * log(D1 / D0) + 1.0 / lamda2 * log(D2 / D1)) + 1.0 / ALPHA );

	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/28]
// Author       :ZSY
// Parameter(s) :用于存放计算所得的散热密度.(其他成员变量在调用之前先初始化)
// Return       :成功返回1,否则返回0
// Remark       :计算保温结构外表面散热密度.(平壁.单层)
//------------------------------------------------------------------
short CCalcThickBase::Calc_Q_PlainOne(double &Q)
{
	//求得本次厚度的散热密度
	//公式:
	//(介质温度-环境温度)/(当前保温厚/(1000*材料导热率) + 1/传热系数)
	Q = fabs(t - ta) / (thick / (1000 * lamda) + 1.0 / ALPHA); 

	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/28]
// Author       :ZSY
// Parameter(s) :用于存放计算所得的散热密度.(其他成员变量在调用之前先初始化)
// Return       :成功返回1,否则返回0
// Remark       :计算保温结构外表面散热密度.(平壁.双层)
//------------------------------------------------------------------
short CCalcThickBase::Calc_Q_PlainCom(double &Q)
{
	//求得本次厚度的散热密度
	//公式:
	//(介质温度-环境温度)/(内保温厚/1000/内保温材料导热率 + 外保温厚/1000/外保温材料导热率 + 1/传热系数)
	Q = fabs(t - ta) / (thick1/1000.0/lamda1+thick2/1000.0/lamda2+1.0/ALPHA);

	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/28]
// Author       :ZSY
// Parameter(s) :用于存放计算所得的散热线密度.(其他成员变量在调用之前先初始化)
// Return       :成功返回1,否则返回0
// Remark       :计算保温结构散热线密度.(管道.单层)
//------------------------------------------------------------------
short CCalcThickBase::Calc_QL_PipeOne(double &QL)
{
	//求得本次厚度的散热线密度
	//公式:
	//(2.0*pi*(介质温度-环境温度)) / ((1.0/材料导热率*ln(保温层外径/管道外径)) + 2000/传热系数/保温层外径)
	QL = (2.0 * pi * fabs(t - ta)) / (1.0 / lamda * log(D1 / D0) + 2000.0 / ALPHA / D1);
	
	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/28]
// Author       :ZSY
// Parameter(s) :用于存放计算所得的散热线密度.(其他成员变量在调用之前先初始化)
// Return       :成功返回1,否则返回0
// Remark       :计算保温结构散热线密度.(管道.双层)
//------------------------------------------------------------------
short CCalcThickBase::Calc_QL_PipeCom(double &QL)
{
	//求得本次厚度的散热线密度
	//公式:
	//
	QL = 2 * pi * fabs(t - ta) / (log(D1 / D0) / lamda1 + log(D2 / D1) / lamda2 + 2000.0 / ALPHA / D2);
		
	return 1;
}



//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       ::给定厚度计算出界面温度(和外表面温度)，平面单层。
//------------------------------------------------------------------
short CCalcThickBase::CalcPlain_One_InputThick(double thick_resu, double &ts_resu)
{
	thick = thick_resu;
	//计算外表面温度.
	CalcPlain_One_TsTemp(thick,ts_resu);
	//散热密度
	Calc_Q_PlainOne(dQ);

	return 1;
}	

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :thick1_resu--内层保温厚度
// Parameter(s) :thick2_resu--外层保温厚度
// Return       :
// Remark       ::给定厚度计算出界面温度(和外表面温度)，保温对象为管道，双层
//------------------------------------------------------------------
short CCalcThickBase::CalcPipe_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu)
{
	thick1 = thick1_resu;
	thick2 = thick2_resu;
	//计算外表面温度和界面温度.
	CalcPipe_Com_TsTb(thick1,thick2,ts_resu,tb_resu);
	//D2是外保温层外径,在求温度时已赋值。求放热系数
	ALPHA = GetPipeAlpha(ts_resu, D2);
	//散热密度
	Calc_Q_PipeCom(dQ);

	return 1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       ::给定厚度计算出界面温度(和外表面温度)，保温对象为管道，单层
//------------------------------------------------------------------
short CCalcThickBase::CalcPipe_One_InputThick(double thick_resu, double &ts_resu)
{
	thick = thick_resu;
	//计算外表面温度.
	CalcPipe_One_TsTemp(thick, ts_resu);

	// 
	ALPHA = GetPipeAlpha(ts_resu, D1);
	//计算散热密度
	Calc_Q_PipeOne(dQ);

	return 1;
}
//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :thick1_resu--内层保温厚度
// Parameter(s) :thick2_resu--外层保温厚度
// Return       :
// Remark       :给定厚度计算出界面温度(和外表面温度)，平面双层。
//------------------------------------------------------------------
short CCalcThickBase::CalcPlain_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu)
{
	CString	strTmp;

	thick1 = thick1_resu;
	thick2 = thick2_resu;		//传入指定的厚度.
	
	//根据指定的厚度计算外表面温度和界面温度
	CalcPlain_Com_TsTb(thick1,thick2,ts_resu,tb_resu);
	
	//计算散热密度
	Calc_Q_PlainCom(dQ);
	
	//查表获得介质温度允许的最大散热密度
	if( GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax) )
	{
		if ( dQ > QMax )
		{
			//当前介质温度下的实际散热密度,大于允许的最大散热密度.
			strTmp.Format("原始数据序号为%d的记录,介质温度(%.1f℃)下的实际散热密度大于允许的最大散热密度\r\n", m_nID,t);
			strExceptionInfo += strTmp;
		}
	}


	return 1; 
}

//------------------------------------------------------------------                
// DATE         :[2005/05/26]
// Author       :ZSY
// Parameter(s) :
// Return       :如果没有提示字符，则返回0，否则返回1
// Remark       :获得计算当前记录时出现在错误提示或警告
//------------------------------------------------------------------
short CCalcThickBase::GetExceptionInfo(CString &strInfo)
{
	if ( strExceptionInfo.IsEmpty() )
	{
		return 0 ;
	}
	else
	{
		strInfo = strExceptionInfo;
		return	1;
	}

}

//------------------------------------------------------------------
// DATE         :[2005/12/21]
// Parameter(s) :t  介质温度
//				:ts	外表面温度
// Return       :保温材料的导热系数
// Remark       : 单层保温材料的导热系数
//------------------------------------------------------------------
double CCalcThickBase::GetLamda(double t, double ts)
{
	double dLamda;//导热系数
	dLamda = in_a0 + in_a1 * ( t + ts ) / 2.0 + in_a2 * ( t + ts ) * ( t + ts ) / 4.0;
	if ( t < ta )
	{	// 如果为保冷计算时，导热系数乘以一个比值
		in_ratio = ( in_ratio <= 0 ) ? 1 : in_ratio;
		dLamda *= in_ratio;
	}

	return dLamda;
}


//------------------------------------------------------------------                
// DATE         :[2005/06/06]
// Author       :ZSY
// Parameter(s) :t--介质温度
// Parameter(s) :tb--双层时界面温度,单层为外表面温度
// Return       :
// Remark       :计算导热系数
//------------------------------------------------------------------
double CCalcThickBase::GetLamda1(double t, double tb)
{
	double dLamda;
	dLamda = in_a0 + in_a1 * ( t + tb ) / 2.0 + in_a2 * ( t + tb ) * ( t + tb ) / 4.0;
	if ( t < ta )
	{	// 如果为保冷计算时，导热系数乘以一个比值
		in_ratio = ( in_ratio <= 0 ) ? 1 : in_ratio;
		dLamda *= in_ratio;
	}
	return dLamda;
}

// 获得导热系数，nMark=0是内保温材料或单层的， 1为外保温材料
double CCalcThickBase::GetLamdaA(double t1, double t2, short nMark)
{
	double dLamda;
	double dRatio;
	switch(nMark)
	{
	case 1:
		dRatio = ( out_ratio <= 0 ) ? 1 : out_ratio;
		dLamda = out_a0 + out_a1 * (t1 + t2) / 2.0  + out_a2 * (t1 + t2) * (t1 + t2) / 4.0;
		break;
	default:	// nMark == 0
		dRatio = ( in_ratio <= 0 ) ? 1 : in_ratio;
		dLamda = in_a0 + in_a1 * (t1 + t2) / 2.0 + in_a2 * (t1 + t2) * (t1 + t2) / 4.0;
	}
	if (t < ta)
	{
		dLamda *= dRatio;
	}
	
	return dLamda;

}
//------------------------------------------------------------------                
// DATE         :[2005/06/06]
// Author       :ZSY
// Parameter(s) :ts为外表面温度,tb界面温度
// Return       :
// Remark       :计算导热系数
//------------------------------------------------------------------
double CCalcThickBase::GetLamda2(double tb, double ts)
{
	double dLamda;
	dLamda = out_a0 + out_a1 * ( tb + ts ) / 2.0 + out_a2 * ( tb + ts ) * ( tb + ts ) / 4.0;
	if ( t < ta )
	{	// 如果为保冷计算时，导热系数乘以一个比值
		out_ratio = ( out_ratio <= 0 ) ? 1 : out_ratio;
		dLamda *= out_ratio;
	}
	return dLamda;
}

//------------------------------------------------------------------
// DATE         :[2005/08/15]
// Parameter(s) :ts:外表面温度.
// Parameter(s) :D1:当为双层管道时为保温外层的外径,单层为内层的外径. 平面的放热系数时D1为0
// Return       :
// Remark       :根据不同的行业规范,获得平面双层的放热系数
//				:城市热力网的计算放热系数使用国家标准。
//------------------------------------------------------------------
double CCalcThickBase::GetPlainAlpha(double ts, double D1)
{
	double dAlpha=co_alf;
	InitCalcAlphaData(ts);
	if (EDIBgbl::iCur_CodeKey == nKey_CODE_GB8175_2008 || EDIBgbl::iCur_CodeKey == nKey_CODE_CJJ34_2002) //国家标准和城市热力网标准
	{
		if (nMethodIndex == nEconomicalThicknessMethod || nMethodIndex == nHeatFlowrateMethod)
		{
			//在进行经济厚度法，最大允许散热损失法。
			CCalcAlpha_CodeGB::CalcEconomy_Alpha(dAlpha);
			
		}else if (nMethodIndex == nSurfaceTemperatureMethod)
		{

			if (PREVENT_SCALD == nTaIndex)
			{	//防烫伤
				CCalcAlpha_CodeGB::CalcPreventScald_Alpha(dAlpha);
				
			}else
			{	//表面温度法
				CCalcAlpha_CodeGB::CalcEconomy_Alpha(dAlpha);
			}			

		}else if (nMethodIndex == nPreventCongealMethod)
		{
			//防冻计算法，风速取冬季最多风向平均风速.
			CCalcAlpha_CodeGB::CalcEconomy_Alpha(dAlpha);
			
		}else
		{
			//默认的公式
			CCalcAlpha_CodeGB::CalcEconomy_Alpha(dAlpha);
		}
//		
	}else if (EDIBgbl::iCur_CodeKey == nKey_CODE_SH3010_2000)//石油化工标准
	{
		//在保温结构外表面温度计算中。

		if (nMethodIndex == nSurfaceTemperatureMethod)
		{
			CCalcAlpha_CodePC::CalcFaceTemp_Alpha(dAlpha);
		}else //if (nMethodIndex == nEconomicalThicknessMethod || nMethodIndex == nHeatFlowrateMethod)
		{	//在经济厚度计算及散热损失计算时
			CCalcAlpha_CodePC::CalcEconomy_Alpha(dAlpha);

		}
	}else											//默认为电力标准
	{	
		if (!strPlace.CompareNoCase("室内"))
		{
			dAlpha = co_alf;			//安装地点为室内时,放热系数用表中的值
		}
		else
		{
			CCalcAlpha_Code::CalcPlain_Alpha(dAlpha);
		}
	}
	return dAlpha;
}

//------------------------------------------------------------------
// DATE         :[2005/08/16]
// Parameter(s) :ts:外表面温度.
// Parameter(s) :D1:当为双层管道时为保温外层的外径,单层为内层的外径. 平面的放热系数时D1为0
// Return       :
// Remark       :根据不同的行业规范,获得平面单层的放热系数
//				:城市热力网的计算放热系数使用国家标准。
//------------------------------------------------------------------
double CCalcThickBase::GetPipeAlpha(double ts, double D1)
{
	double dresAlpha = co_alf;
	InitCalcAlphaData(ts);
	if (EDIBgbl::iCur_CodeKey == nKey_CODE_GB8175_2008 || EDIBgbl::iCur_CodeKey == nKey_CODE_CJJ34_2002)//国家标准标准或城市热力网标准
	{
		//平面与管道的放热系数的规定是一样的。
		dresAlpha = GetPlainAlpha(ts,D1);

	}else if (EDIBgbl::iCur_CodeKey == nKey_CODE_SH3010_2000)//石油化工标准
	{
		//平面与管道的放热系数的规定是一样的。
		dresAlpha = GetPlainAlpha(ts,D1);
	}else										
	{	//默认为电力标准		
		if (-1 != strPlace.Find("室内",0))
		{
			dresAlpha = co_alf;		//安装地点为室内时，用表中的放热系数
		}
		else
			
		{
			CCalcAlpha_Code::CalcPipe_Alpha(D1, dresAlpha);
		}
	}

	return dresAlpha;
}

//------------------------------------------------------------------
// DATE         :[2005/08/17]
// Parameter(s) :ts:外表面温度.
// Parameter(s) :D1:当为双层管道时为保温外层的外径,单层为内层的外径. 平面的放热系数时D1为0
// Return       :
// Remark       :初始计算放热系数的变量
//				:城市热力网的计算放热系数使用国家标准。
//------------------------------------------------------------------
BOOL CCalcThickBase::InitCalcAlphaData(double ts) 
{
	if (EDIBgbl::iCur_CodeKey == nKey_CODE_GB8175_2008 || EDIBgbl::iCur_CodeKey == nKey_CODE_CJJ34_2002)
	{ 
		//国家标准和城市热力网标准 
		INIT_ALPHADATA(CCalcAlpha_CodeGB);
	}else if (EDIBgbl::iCur_CodeKey == nKey_CODE_SH3010_2000) 
	{ 
		//石油化工 
		INIT_ALPHADATA(CCalcAlpha_CodePC); 
	}else 
	{
		//默认为电力
		INIT_ALPHADATA(CCalcAlpha_Code);
	}
	return TRUE;
}

//------------------------------------------------------------------                
// DATE         :[2010/07/09]
// Author       :LIGB
// Parameter(s) :
// Return       :
// Remark       :根据介质温度,查表取保温结构外表面允许最大散热密度
//------------------------------------------------------------------
float CCalcThickBase::GetMaxHeatLoss(_RecordsetPtr &pRs, double &t, CString strFieldName)
{
	try
	{
		float fTempCur, fTempPre, fValCur, fValPre, fVal;
		bool bFirst=true;
		for(pRs->MoveFirst(); !pRs->adoEOF; pRs->MoveNext() )
		{
			fTempCur = vtof(pRs->GetCollect(_variant_t("t")));
			fValCur = vtof(pRs->GetCollect(_variant_t(strFieldName))); 
			
			if( DZero > fValCur )
			{//若字段值为0或空，表明前一条记录是最后有值的记录
				fVal = fValPre;
				break;
			}
			else
			{//用插值法计算热流密度
				if( fTempCur >= t )
				{
					if(!bFirst)
					{//当前记录的温度值大于或等于要查找的温度.
						fVal = (t-fTempPre)/(fTempCur-fTempPre)*(fValCur-fValPre)+fValPre;					
					}
					else
						fVal=fValCur;
					break;
				}
				else
				{	//记住当前的值
					fTempPre = fTempCur;
					fValPre = fValCur;
				}
			}
			bFirst=false;
		}
			
		if( pRs->adoEOF )
		{//若是最后一条记录
			fVal = fValCur;
		}
		
		return fVal;
	}
	catch(_com_error& e)
	{
		ReportExceptionError(e);
		return 0;
	}	

}
