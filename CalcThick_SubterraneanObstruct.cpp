// CalcThick_SubterraneanObstruct.cpp: implementation of the CCalcThick_SubterraneanObstruct class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcThick_SubterraneanObstruct.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThick_SubterraneanObstruct::CCalcThick_SubterraneanObstruct()
{

}

CCalcThick_SubterraneanObstruct::~CCalcThick_SubterraneanObstruct()
{

}

//------------------------------------------------------------------                
// Author       : ZSY
// Parameter(s) : 
// Return       : 
// Remark       : 计算埋地管道（双层不通行)保温层厚
//------------------------------------------------------------------
short CCalcThick_SubterraneanObstruct::CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double MaxTb = 350;
	double dDiff1;
	double dDiff2;
	double dMinDiff1 = 0;
	double dMinDiff2 = 0;
	double dCur_Q;
	
	double Q;
	
	bool   flg;

	thick1_resu=0;
	thick2_resu=0;
	ts_resu=0;
	tb_resu=0;

	//查表获得介质温度允许的最大散热密
	GetHeatLoss( m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax );

	// 获得原始数据
	GetSubterraneanOriginalData( m_nID );
	
	//thick1内层保温厚度。thick2外层保温厚度
	//thick1Start内层保温允许最小厚度,thick1Max内层保温允许最大厚度,thick1Increment内层保温厚度最小增量
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表thicknessRegular获得
	for(thick1=m_thick1Start;thick1<=m_thick1Max;thick1+=m_thick1Increment)
	{
		for(thick2=m_thick2Start;thick2<=m_thick2Max;thick2+=m_thick2Increment)
		{
			CalcPipe_Com_TsTb(thick1, thick2, ts, tb);
			
			//内保温层外径
			D1=2*thick1+D0;  
			//外保温层外径
			D2=2*thick2+D1;
			//根据不同的行业规范计算放热系数
			ALPHA = GetPipeAlpha(ts,D2);

			//根据改变的外表面温度和界面温度计算导热系数  
			lamda1 = GetLamda1( t, tb ); 
			lamda2 = GetLamda2( tb, ts); 

			Calc_Q_PipeCom(Q); 
			//确定内保温厚的条件 
			dDiff1 = log(D1 / D0) - 2000 * lamda1 * fabs(t-tb) / Q / D2; 
			//外保温厚  
			//lamda1 = GetLamda1()

			dDiff2 = D2 * log(D2 / D0) - 2000 * ( (lamda1 * fabs(t-tb) + lamda2 * fabs(tb-ta)) / Q - lamda2 / ALPHA );

			dDiff1 = fabs( dDiff1 );
			dDiff2 = fabs( dDiff2 );
			ALPHA  = GetPipeAlpha( ts, D2 );
			
			if ( ts<MaxTs && tb<MaxTb )
			{
				if ((fabs(dMinDiff1) < DZero && fabs(dMinDiff2) < DZero) || (dDiff1 < dMinDiff1 && dDiff2 < dMinDiff2))
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
						dMinDiff2 = fabs( dDiff2 );	// 
						
						thick1_resu = thick1;			//内保温厚
						thick2_resu = thick2;
						ts_resu = ts;					// 外表面温度
						tb_resu = tb;					//界面温度
						//计算当前厚度下的散热密度
						this->Calc_Q_PipeCom(dQ);
					}
				}
			}
		}
	}       

	return 1;
}

//------------------------------------------------------------------
// Author       : ZSY
// Parameter(s) : thick_resu[out] 保温层厚度
//				: ts_resu[out] 外表面温度
// Return       : 计算成功返加TRUE，否则返回FALSE
// Remark       : 埋地管道(不通行) 单层
//------------------------------------------------------------------
short CCalcThick_SubterraneanObstruct::CalcPipe_One(double &thick_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	dYearVal;		//年运行工况的最大散热密度
	double	dSeasonVal;		//季节运行工况最大散热密度
	double	Q = 1;				//当前厚度下的散热密度.
	double	QMax = 0;			//根据当前记录的工况;允许最大散热密度
	double	dLSo;	// 土壤的导热系数
	double	dASo;	//
	double  dDiff;
	double  dMinDiff = 0;
	bool	flg=true;		//进行散热密度的判断

	thick_resu=0;
	ts_resu=0;
	GetSubterraneanOriginalData(m_nID);
	GetHeatLoss(m_pRsHeatLoss, Tf1, dYearVal, dSeasonVal,QMax);
	if (fabs(QMax) >= DZero)
	{
		Q = QMax * m_MaxHeatDensityRatio;			//保温结构外表面允许散热密度，按表中查出的允许最大散热密度的乘以一个系数取值.(规程规定为90%).
	}

	//thick单层保温厚度。
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表ThicknessRegular获得
	for (thick=m_thick2Start; thick <= m_thick2Max; thick += m_thick2Increment)
	{
		//根据指定的保温厚，计算外表面温度
		CalcPipe_One_TsTemp(thick,ts);
		D1    = D0 + 2 * thick;
		
		dLSo  = GetLamdaSo(ts);		// 土壤的导热系数
		lamda = GetLamda(t, ts);	// 保温层的导热系数
		CalcPipe_Alpha(D0, dASo);
		dDiff = log(D1 / D0) - 2.0 * pi * lamda * ((Tf1 - Tso) / Q - (1 / (A1pre * pi * D1) + 1 / (Aaw * pi * Dag) + (log(4 * sub_h / Dag) / (2 * pi * dLSo))));
		if (ts<MaxTs && (dMinDiff <= DZero || dDiff < dMinDiff))
		{
			flg = true;
			//选项中选择的判断散热密度，就进行比较
			if(bIsHeatLoss)
			{	
				//求得本次厚度的散热密度
				Calc_Q_PipeOne(Q);
				
				if(fabs(QMax) <= DZero || Q <= QMax)
				{
					flg = true; // 当前厚度满足允许最大散热规则
				}else
				{
					flg = false;
				}
			}
			//满足条件 
			if( flg ) 
			{
				dMinDiff = dDiff;
				thick_resu = thick;
				ts_resu = ts;	
				//计算当前厚度下的散热密度 
				Calc_Q_PipeOne(dQ);
			}
		}
	}
	return TRUE;
}

//------------------------------------------------------------------
// Author       : ZSY
// Parameter(s) : dQ1[out]:第一根的热损失，单位[Kcal/(m*h)]
//				: dQ2[out]:第二根的热损失，单位[Kcal/(m*h)]
// Return       : void
// Remark       : 计算不通行地沟中敷设两根(多根)时的热损失
//------------------------------------------------------------------
void CCalcThick_SubterraneanObstruct::Calc_Q_PipeTwo(double &dQ1, double &dQ2)
{
	double dTaw;
	Calc_Taw_Two(dTaw);
	dQ1 = (Tf1 - dTaw) / (R11 + Rs1);
	dQ2 = (Tf2 - dTaw) / (R12 + Rs2);
	
}
//------------------------------------------------------------------
// Author       : ZSY
// Parameter(s) : dQ[out]: 单根的热损失，单位 [Kcal/(m*h)]
// Return       : short 1
// Remark       : 计算不通行地沟中敷设单根的热损失
//------------------------------------------------------------------
short CCalcThick_SubterraneanObstruct::Calc_Q_PipeOne(double &dOneQ)
{
	double dRso; 
	
//	double dQ1;
//	double dQ2;

	Calc_Rso(dRso);
	dOneQ = (Tf1 - Tso) / (R1 + Rs1 + Raw + dRso);
	return 1; 
}
