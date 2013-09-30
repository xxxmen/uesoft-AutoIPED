// CalcThick_Economy.cpp: implementation of the CCalcThick_Economy class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcThick_Economy.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThick_Economy::CCalcThick_Economy()
{

}

CCalcThick_Economy::~CCalcThick_Economy()
{

}
//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :使用经济厚度法,计算保温层的厚度，平面双层。
//------------------------------------------------------------------
short CCalcThick_Economy::CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double	ts=0.0,			//外表面温度值
			tb=0.0,			//界面处温度
			nYearVal,		//年运行工况的最大散热密度
			nSeasonVal,		//季节运行工况最大散热密度
			Q,				//当前厚度下的散热密度.
			QMax;			//根据当前记录的工况,允许最大散热密度	
//			MaxTb = in_tmax*m_HotRatio;	//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	double	cost_q,			//??费用
			cost_s,
			cost_tot,
			cost_min;		//最小的费用.

	bool	flg=true;		//进行散热密度的判断

	thick1_resu=0;
	thick2_resu=0;
	ts_resu=0;
	tb_resu=0;

	cost_min=0;
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax);
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
			//放热系数
			//根据不同的行业规范计算放热系数
			ALPHA = GetPlainAlpha(ts);

			//根据改变的外表面温度和界面温度计算导热系数
			lamda1 = GetLamda1(t, tb);
			lamda2 = GetLamda2( tb, ts);

			//保温结构的散热损失年费用  			公式见DL/T 5072—1997《火力发电厂保温油漆设计规程》Ｐ93
			cost_q=3.6*Mhours*heat_price*Yong*fabs(t-ta)*1e-6
				/(thick1/1000.0/lamda1+thick2/1000.0/lamda2+1.0/ALPHA);
			//保温结构投资年分摊费用　
			cost_s=(thick1/1000.0*in_price+thick2/1000.0*out_price+(gbWithoutProtectionCost?0:pro_price))*S;
			//保温对象年总费用
			cost_tot=cost_q+cost_s;

			if((ts<MaxTs && tb<MaxTb))	//ts<50 && tb<350
			{
				if(fabs(cost_min)<DZero/*第一次*/ || cost_min>cost_tot)
				{
					flg = true;
					//选项中选择的判断散热密度,就进行比较.
					if(bIsHeatLoss)
					{
						//求得本次厚度的散热密度
						//公式:
						//(介质温度-环境温度)/(内保温厚/1000/内保温材料导热率 + 外保温厚/1000/外保温材料导热率 + 1/传热系数)
						Calc_Q_PlainCom(Q);
						//	Q = (t-ta)/(thick1/1000.0/lamda1+thick2/1000.0/lamda2+1.0/ALPHA);
						
						if(fabs(QMax) <= DZero || Q <= QMax)
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

						//计算当前厚度下的散热密度
						this->Calc_Q_PlainCom(dQ);
					}
				}
			}
		}//end thick2
	}//end thick1
	return 1;
}

//------------------------------------------------------------------                
// DATE         :[2005/04/22]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :使用经济厚度法,计算保温层的厚度，平面单层。
//------------------------------------------------------------------
short CCalcThick_Economy::CalcPlain_One(double &thick_resu, double &ts_resu)
{
	double	ts=0.0,			//外表面温度值
			nYearVal,		//年运行工况的最大散热密度
			nSeasonVal,		//季节运行工况最大散热密度
			Q,				//当前厚度下的散热密度.
			QMax;			//根据当前记录的工况,允许最大散热密度

	double	cost_q,			//??费用
			cost_s,
			cost_tot,
			cost_min;		//最小的费用

	bool	flg=true;		//进行散热密度的判断

	thick_resu=0;
	ts_resu=0;
	cost_min=0;
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax);
	} 
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	for (thick=m_thick2Start; thick<=m_thick2Max; thick+=m_thick2Increment)
	{
		//根据指定的保温厚，计算外表面温度。
		CalcPlain_One_TsTemp(thick,ts);
		//导热系数
		lamda = GetLamda(t, ts);

		// 根据不同的行业规范计算放热系数
		ALPHA = GetPlainAlpha(ts);
	
		//保温结构的散热损失年费用  			公式见DL/T 5072—1997《火力发电厂保温油漆设计规程》Ｐ93
		cost_q=3.6*Mhours*heat_price*Yong*fabs(t-ta)*1e-6
				/((thick/1000.0/lamda)+1.0/ALPHA);
		//保温结构投资年分摊费用
		cost_s=(thick/1000*in_price+(gbWithoutProtectionCost?0:pro_price))*S;
		//保温对象年总费用
		cost_tot=cost_q+cost_s;	

		if( ts<MaxTs && (cost_tot<cost_min || fabs(cost_min)<DZero) )
		{
			flg = true;
			//选项中选择的判断散热密度,就进行比较.
			if(bIsHeatLoss)
			{
				//求得本次厚度的散热密度
				//公式:	
				//(介质温度-环境温度)/(当前保温厚/(1000*材料导热率) + 1/传热系数)
				Calc_Q_PlainOne(Q);
				//	Q = (t-ta) / (thick / (1000*lamda) + 1.0 / ALPHA); 
				if(fabs(QMax) <= DZero || Q <= QMax)
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
				//计算当前厚度下的散热密度
				Calc_Q_PlainOne(dQ);
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
// Remark       :使用经济厚度法,计算保温层的厚度，保温对象为管道，双层
//------------------------------------------------------------------
short CCalcThick_Economy::CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	tb=0.0;			//界面处温度
	double	nYearVal;		//年运行工况的最大散热密度
	double	nSeasonVal;		//季节运行工况最大散热密度
	double	Q;				//当前厚度下的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度	
//	double	MaxTb = in_tmax*m_HotRatio;	//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	double	cost_q;			//??费用
	double	cost_s;
	double	cost_tot;
	double	cost_min;		//最小的费用.

	bool	flg=true;		//进行散热密度的判断

	thick1_resu=0;
	thick2_resu=0;
	ts_resu=0;
	tb_resu=0;

	cost_min=0;
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		if( !GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax) )
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
			//计算当前厚度界面温度和外表面温度
			CalcPipe_Com_TsTb(thick1,thick2,ts,tb);

			//内保温层外径和外保温层外径
			D1=2*thick1+D0;
			D2=2*thick2+D1;	

			//根据不同的行业规范计算放热系数
			ALPHA =	GetPipeAlpha(ts,D2);

			//根据改变的外表面温度和界面温度计算导热系数
			lamda1 = GetLamda1(t, tb);
			lamda2 = GetLamda2( tb, ts);
			//保温结构的散热损失年费用  			公式见DL/T 5072—1997《火力发电厂保温油漆设计规程》Ｐ93
			cost_q=7.2*pi*Mhours*heat_price*Yong*fabs(t-ta)*1e-6
					/(1.0/lamda1*log(D1/D0)+1.0/lamda2*log(D2/D1)+2000.0/ALPHA/D2);
			//保温结构投资年分摊费用
			cost_s=(pi/4.0*(D1*D1-D0*D0)*in_price*1e-6+pi/4.0*(D2*D2-D1*D1)*out_price*1e-6
					+(gbWithoutProtectionCost?0:pi*D2*pro_price*1e-3))*S;
			//保温对象年总费用
			cost_tot=cost_q+cost_s;
			//tb<350)
			if((ts<MaxTs && tb<MaxTb) && (fabs(cost_min)<DZero || cost_min>cost_tot))
			{
				flg = true;
				//选项中选择的判断散热密度,就进行比较.
				if(bIsHeatLoss)
				{
					//求得本次厚度的散热密度
					//公式:
					//(介质温度-环境温度)/((外保温层外径/2000.0)*(1.0/内保温材料导热率*ln(内保温层外径/管道外径) + 1.0/外保温材料导热率*ln(外保温层外径/内保温层外径) ) + 1/传热系数)
					Calc_Q_PipeCom(Q);
					//	Q = (t-ta) / ( (D2/2000.0) * (1.0/lamda1*log(D1/D0)+1.0/lamda2*log(D2/D1)) + 1.0/ALPHA );

					if(fabs(QMax) <= DZero || Q <= QMax)
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

					//计算当前厚度下的散热密度
					Calc_Q_PipeCom(dQ);
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
// Remark       :使用经济厚度法,计算保温层的厚度，保温对象为管道，单层
//------------------------------------------------------------------
short CCalcThick_Economy::CalcPipe_One(double &thick_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	nYearVal;		//年运行工况的最大散热密度
	double	nSeasonVal;		//季节运行工况最大散热密度
	double	Q;				//当前厚度下的散热密度.
	double	QMax;			//根据当前记录的工况;允许最大散热密度

	double	cost_q;			//散热费用
	double	cost_s;			//材料费用
	double	cost_tot;	    //总费用
	double	cost_min;		//最小的费用.

	bool	flg=false;		//计算的保温厚度满足要求，初始为不满足

	thick_resu=0;
	ts_resu=0;
	cost_min=0;
	//选项中选择了判断散热密度,则查表。
	if(bIsHeatLoss)
	{
		//查表获得介质温度允许的最大散热密度
		if( !GetHeatLoss(m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax) )
		{ 		
		}
	}
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	for (thick=m_thick2Start; thick<=m_thick2Max; thick+=m_thick2Increment)
	{
		//根据指定的保温厚，计算外表面温度。
		CalcPipe_One_TsTemp(thick,ts);

		D1=2.0*thick+D0;
		//导热系数
		lamda = GetLamda(t, ts);
		//根据不同的行业规范计算放热系数
		ALPHA = GetPipeAlpha(ts,D1);

		//保温结构的散热损失年费用  			公式见DL/T 5072—1997《火力发电厂保温油漆设计规程》P93
		cost_q=7.2*pi*Mhours*heat_price*Yong*fabs(t-ta)*1e-6
				/(1.0/lamda*log(D1/D0)+2000.0/ALPHA/D1);
		//保温结构投资年分摊费用
		cost_s=(pi/4.0*(D1*D1-D0*D0)*in_price*1e-6+(gbWithoutProtectionCost?0:pi*D1*pro_price*1e-3))*S;
		//保温对象年总费用
		cost_tot=cost_q+cost_s;
		//求得本次厚度的散热密度
		Calc_Q_PipeOne(Q);
		
		if(ts<MaxTs && (cost_tot<cost_min || fabs(cost_min)<DZero))
		{
			flg = true;
			//选项中选择的判断散热密度,就进行比较.


			if(bIsHeatLoss)
			{
				//求得本次厚度的散热密度
				//公式:
				//(介质温度-环境温度)/((保温层外径/2000.0/材料导热率*ln(保温层外径/管道外径)) + 1/传热系数)
				Calc_Q_PipeOne(Q);
				//	Q = (t-ta)/(D1/2000.0/lamda*log(D1/D0) + 1.0/ALPHA);

				if(fabs(QMax) <= DZero || Q <= QMax)

					flg = true; // 当前厚度满足允许最大散热规则.

				else
					flg = false;
			}
			//满足条件
			if( flg )
			{

				cost_min = cost_tot;
				thick_resu = thick;
				ts_resu = ts;
				//计算当前厚度下的散热密度
				Calc_Q_PipeOne(dQ);
			}
		}
	}

	return 1;	
}
