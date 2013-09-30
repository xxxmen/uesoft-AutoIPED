// CalcThick_HeatDensity.cpp: implementation of the CCalcThick_HeatDensity class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CalcThick_HeatDensity.h"
#include "EDIBgbl.h"
#include "AutoIPED.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThick_HeatDensity::CCalcThick_HeatDensity()
{

}

CCalcThick_HeatDensity::~CCalcThick_HeatDensity()
{

}


//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :使用允许散热密度法,计算保温层的厚度，平面双层。
//------------------------------------------------------------------
short CCalcThick_HeatDensity::CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	tb=0.0;			//界面处温度
	double	nYearVal;		//年运行工况的最大散热密度
	double	nSeasonVal;		//季节运行工况最大散热密度
	double	Q;				//当前厚度下的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度	
//	double	MaxTb = in_tmax*m_HotRatio;	//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	thick1_resu=0;
	thick2_resu=0;
	ts_resu=0;
	tb_resu=0;

	//查表获得介质温度允许的最大散热密度
	if( !GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax) )
	{ 		
	}
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表thicknessRegular获得
	for(thick1=m_thick1Start;thick1<=m_thick1Max;thick1+=m_thick1Increment)
	{
		for(thick2=m_thick2Start;thick2<=m_thick2Max;thick2+=m_thick2Increment)
		{
			CalcPlain_Com_TsTb(thick1,thick2,ts,tb);

			//根据不同的行业规范计算放热系数
			ALPHA = GetPlainAlpha(ts);

			//根据改变的外表面温度和界面温度计算导热系数
			lamda1 = GetLamda1( t, tb );
			lamda2 = GetLamda2( tb, ts);

			if( ts<MaxTs && tb<MaxTb )	//ts<50 && tb<350
			{
				//求得本次厚度的散热密度
				//公式:
				//(介质温度-环境温度)/(内保温厚/1000/内保温材料导热率 + 外保温厚/1000/外保温材料导热率 + 1/传热系数)
				Calc_Q_PlainCom(Q);
				//	Q = (t-ta)/(thick1/1000.0/lamda1+thick2/1000.0/lamda2+1.0/ALPHA);
				
				if(fabs(QMax) <= DZero || Q <= QMax)
				{	//当前厚度满足允许最大散热规则.
					thick1_resu=thick1;
					thick2_resu=thick2;
					ts_resu=ts;
					tb_resu=tb;
					//计算当前厚度下的散热密度
					this->Calc_Q_PlainCom(dQ);

					return 1;
				}
			}
		}
	}
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :使用允许散热密度法,计算保温层的厚度，平面单层。
//------------------------------------------------------------------
short CCalcThick_HeatDensity::CalcPlain_One(double &thick_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	nYearVal;		//年运行工况的最大散热密度
	double	nSeasonVal;		//季节运行工况最大散热密度
	double	Q;				//当前厚度下的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度	

	thick_resu=0;
	ts_resu=0;
	//查表获得介质温度允许的最大散热密度
	if( !GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax) )
	{ 		
	}

	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	for (thick=m_thick2Start; thick<=m_thick2Max; thick+=m_thick2Increment)
	{
		//根据指定的保温厚，计算外表面温度。
		CalcPlain_One_TsTemp(thick,ts);
		//导热系数			
		lamda = GetLamda( t, ts );

		//根据不同的行业规范计算放热系数
		ALPHA = GetPlainAlpha(ts);
		
		if(ts<MaxTs)
		{
			//求得本次厚度的散热密度
			//公式:	
			//(介质温度-环境温度)/(当前保温厚/(1000*材料导热率) + 1/传热系数)
			Calc_Q_PlainOne(Q);
			//	Q = (t-ta) / (thick / (1000*lamda) + 1.0 / ALPHA); 
			if(fabs(QMax) <= DZero || Q <= QMax)
			{
				//当前厚度满足允许最大散热规则. 
				thick_resu=thick;
				ts_resu=ts;
				//计算当前厚度下的散热密度
				this->Calc_Q_PlainOne(dQ);

				return 1;
			}
		}
	}
	return 1;
}	

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :使用允许散热密度法,计算保温层的厚度，保温对象为管道，双层
//------------------------------------------------------------------
short CCalcThick_HeatDensity::CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	tb=0.0;			//界面处温度
	double	nYearVal;		//年运行工况的最大散热密度
	double	nSeasonVal;		//季节运行工况最大散热密度

	double  dCur_Q;			//当前厚度下的散热密度.
	double	Q;				//允许最大散热密度与一个比值的乘积，作为计算保温厚度时的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度

	bool	flg=true;		//进行散热密度的判断

	double  dDiff1;			//确定内保温厚的差值
	double  dDiff2;
	double  dMinDiff1=0;	//最小的差值
	double  dMinDiff2=0;

//	double	MaxTb = in_tmax*m_HotRatio;	//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	thick1_resu=0;
	thick2_resu=0;
	ts_resu=0;
	tb_resu=0;

	//查表获得介质温度允许的最大散热密度
	if( !GetHeatLoss( m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax ) )
	{
	}

	Q = QMax * m_MaxHeatDensityRatio;			//保温结构外表面允许散热密度，按表中查出的允许最大散热密度的乘以一个系数取值.(规程规定为90%).

	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表thicknessRegular获得
	//
	for(thick1=m_thick1Start;thick1<=m_thick1Max;thick1+=m_thick1Increment)
	{
		for(thick2=m_thick2Start;thick2<=m_thick2Max;thick2+=m_thick2Increment)
		{
			CalcPipe_Com_TsTb(thick1,thick2,ts,tb);

			//内保温层外径
			D1=2*thick1+D0;
			//外保温层外径
			D2=2*thick2+D1;
			
			//根据不同的行业规范计算放热系数
			ALPHA = GetPipeAlpha(ts,D2);

			//根据改变的外表面温度和界面温度计算导热系数
			lamda1 = GetLamda1( t, tb );
			lamda2 = GetLamda2( tb, ts);
			
			//确定内保温厚的条件.
			dDiff1 = log(D1 / D0) - 2000 * lamda1 * fabs(t-tb) / Q / D2;
			//外保温厚
			dDiff2 = D2 * log(D2 / D0) - 2000 * ((lamda1 * fabs(t-tb) + lamda2 * fabs(tb-ta)) / Q - lamda2 / ALPHA);

			//  [2005/11/25]
			dDiff1 = fabs( dDiff1 );	
			dDiff2 = fabs( dDiff2 );
			//tb<350)
			if(	(ts<MaxTs && tb<MaxTb) && 
				((fabs(dMinDiff1) < DZero && fabs(dMinDiff2) < DZero) || (dDiff1 < dMinDiff1 && dDiff2 < dMinDiff2)))
			{
				flg = true;
				if ( bIsHeatLoss )
				{
					//计算当前厚度下的散热密度
					Calc_Q_PipeCom(dCur_Q);
					if ( fabs( QMax ) <= DZero || dCur_Q <= QMax )
						flg = true;
					else
						flg = false;
				}
				if ( flg )
				{
					dMinDiff1 = dDiff1;			//保存最小的差值
					dMinDiff2 = dDiff2;
		
 					thick1_resu=thick1;			//内保温厚
					thick2_resu=thick2;
					ts_resu=ts;
					tb_resu=tb;					//界面温度 

					//计算当前厚度下的散热密度
					this->Calc_Q_PipeCom(dQ);
				}
			}
		}
	}                     
	return 1;
}

 
//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :使用允许散热密度法,计算保温层的厚度，保温对象为管道，单层
//------------------------------------------------------------------
short CCalcThick_HeatDensity::CalcPipe_One(double &thick_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	nYearVal;		//年运行工况的最大散热密度
	double	nSeasonVal;		//季节运行工况最大散热密度
	double  dCur_Q;			//当前厚度下的散热密度.
	double	Q;				//允许最大散热密度的与一个比值的乘积，作为计算保温厚度时的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度

	double	dMinDiff=0;
	double	dDiff;			//差值,

	bool	flg=true;		//进行散热密度的判断

	thick_resu=0;
	ts_resu=0;

	//查表获得介质温度允许的最大散热密度
	if( !GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax) )
	{
	}
	Q = QMax * m_MaxHeatDensityRatio;			//保温结构外表面允许散热密度，按表中查出的允许最大散热密度的乘以一个系数取值.(规程规定为90%).

	//thick1内层保温厚度。thick2外层保温厚度
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	for (thick=m_thick2Start; thick<=m_thick2Max; thick+=m_thick2Increment)
	{
		//根据指定的保温厚，计算外表面温度。
		CalcPipe_One_TsTemp(thick,ts);
		//导热系数
		lamda = GetLamda(t, ts);

		//保温层外径
		D1=2.0*thick+D0;
		//根据不同的行业规范计算放热系数
		ALPHA = GetPipeAlpha(ts, D1);

		//确定保温厚的条件。求出差值
		dDiff = D1 * log(D1 / D0) - 2000 * lamda * (fabs(t - ta) / Q - 1 / ALPHA);
		dDiff = fabs( dDiff );
		
		if(ts<MaxTs && (fabs(dMinDiff) < DZero || fabs(dDiff) < fabs(dMinDiff)) )
		{
			flg = true;
			if ( bIsHeatLoss )
			{
				//计算当前厚度下的散热密度
				this->Calc_Q_PipeOne( dCur_Q );
				if ( fabs(QMax) <= DZero || dCur_Q <= QMax )
					flg = true;
				else
					flg = false;
			}
			if ( flg )
			{
 				dMinDiff = dDiff;
				thick_resu=thick;
				ts_resu=ts;
				//计算当前厚度下的散热密度
				this->Calc_Q_PipeOne(dQ);
			}
		}
	}

	return 1;	
}
