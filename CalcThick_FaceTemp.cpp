// CalcThick_FaceTemp.cpp: implementation of the CCalcThick_FaceTemp class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcThick_FaceTemp.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThick_FaceTemp::CCalcThick_FaceTemp()
{
}

CCalcThick_FaceTemp::~CCalcThick_FaceTemp()
{

}


//------------------------------------------------------------------                
// DATE         :[2005/04/25]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :外表面温度法计算保温厚度，平面单层。
//------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPlain_One(double &thick_resu, double &ts_resu)
{
	double	ts=ts_resu;		//外表面温度值


	//根据改变的外表面温度计算导热系数
	lamda = GetLamda( t, ts );
	//根据不同的行业规范计算放热系数
	ALPHA =	GetPlainAlpha(ts);

	//如果为表面温度法计算保温层厚度时,外表面温度和界面温度为已知的.函数参数传过来.
	thick_resu = 1000*lamda*(t-ts)/ALPHA/(ts-ta);
	
	// zsy 保温厚度的值为一个整数 [10/26/2005]
	thick_resu = 10 * (int)((thick_resu + INT_ADD) / 10);
	
	if ( thick_resu < 0 )
	{
		CString strMsg;
		strMsg.Format("原始数据序号为%d的记录, 计算出的保温层厚度为负数。请确认原始数据是否正确?", m_nID);
		if ( -1 == strExceptionInfo.FindOneOf(strMsg) )
		{
			strExceptionInfo += strMsg;
		}
	}
	//计算当前厚度下的散热密度
	thick = thick_resu;
	this->Calc_Q_PlainOne(dQ);

	return	1;

}	
//------------------------------------------------------------------                
// DATE         :[2005/04/25]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :用表面温度法计算保温厚度，平面双层。
//------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double	ts=ts_resu;		//函数参数传过来的外表面温度。
	double	tb;				//界面温度。

//	double	MaxTb = in_tmax*m_HotRatio;//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	double  dCur_Q;			//当前厚度下的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度
	bool	flg = true;		//进行散热密度的判断

	//查表获得介质温度允许的最大散热密度
	if( !GetHeatLoss( m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax ) )
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
			this->CalcPlain_Com_TsTb(thick1,thick2,ts,tb,1);

			//根据不同的行业规范计算放热系数
			ALPHA =	GetPlainAlpha(ts);

			//根据改变的外表面温度和界面温度计算导热系数
			lamda1 = GetLamda1(t, tb);
			lamda2 = GetLamda2( tb, ts);
			if(tb<MaxTb )
			{
				flg = true;
				if ( bIsHeatLoss )		//是否判断允许最大散热密度
				{
					Calc_Q_PlainCom(dCur_Q);
					if ( fabs(QMax) <= 1e-6 || dCur_Q <= QMax )
						flg = true;
					else
						flg = false;
				}
				if ( flg )
				{
					thick1_resu = thick1;
					thick2_resu = thick2;
					CString		strTmp;
					strTmp.Format("%.0f",tb);
					tb_resu		= _tcstod(strTmp,NULL);

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
// DATE         :[2005/04/26]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :表面温度法,计算管道双层保温厚度
//------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu)
{
	double	ts = ts_resu;			//表面温度法计算时，外表面温度为已知的。由参数传递过来。
	double	tb;						//界面温度。
	double	dDiff1;					//计算内层厚度的差值.
	double	dDiff2;					//计算外层厚度的差值.	
	double	dDiffMin1=0;			//内层厚度的最小差值.
	double	dDiffMin2=0;			//外层厚度的最小差值.

//	double	dMaxTb = in_tmax*m_HotRatio;//结合面允许最大温度=外保温材料的最大温度 * 一个系数(规程规定为0.9).
	double  MaxTb = in_tmax;//结合面允许最大温度,从保温材料参数库(a_mat)中取MAT_TMAX字段,可修改该字段进行控制界面温度值.

	thick1_resu = 0.0;
	thick2_resu = 0.0;

	double  dCur_Q;			//当前厚度下的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度
	bool	flg = true;		//进行散热密度的判断

	//查表获得介质温度允许的最大散热密度
	if( !GetHeatLoss( m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax ) )
	{ 		
	}

	for (thick1=m_thick1Start; thick1<=m_thick1Max; thick1+=m_thick1Increment)
	{
		D1 = D0 + 2*thick1;
		for (thick2=m_thick2Start; thick2<=m_thick2Max; thick2+=m_thick2Increment)
		{
			//根据当前的保温厚度和外表面温度,计算出界面温度.
			CalcPipe_Com_TsTb(thick1,thick2,ts,tb,1);

			D2 = D1 + 2*thick2;

			//根据不同的行业规范计算放热系数
			ALPHA =	GetPipeAlpha(ts,D2);
			
			//根据改变的外表面温度和界面温度计算导热系数
			lamda1 = GetLamda1(t, tb);
			lamda2 = GetLamda2( tb, ts);
			
			//计算外层的保温厚的公式,
			dDiff2 = D2*log(D2/D0)-2000/ALPHA/(ts-ta)*(lamda1*(t-tb)+lamda2*(tb-ts));	
//			dDiff2 = D2 * log( D2 / D0 ) - 2000 / ALPHA / fabs( ts - ta ) * ( lamda1 * fabs( t - tb ) + lamd2 * fabs( tb - ts ) );

			//计算内层保温厚的公式
			dDiff1 = log(D1/D0)-2000*lamda1*(t-tb)/(ALPHA*D2*(ts-ta));
//			dDiff1 = log( D1 / D0 ) - 2000 * lamda1 * fabs( t - tb ) / ( ALPHA * D2 * fabs( ts - ta ) );

			if ((tb < MaxTb) &&							//界面温度不应超过外层保温材料最高使用温度
					( fabs(dDiffMin1) < 1e-6 && fabs(dDiffMin2) < 1e-6  ||					//为第一次.
					  fabs(dDiff1) < fabs(dDiffMin1) && fabs(dDiff2) < fabs(dDiffMin2)) )	//最小差值时的厚度
			{
				flg = true;
				if ( bIsHeatLoss )
				{					
					Calc_Q_PipeCom( dCur_Q );		//计算当前厚度下的散热密度
					if ( fabs(QMax) <= 1e-6 || dCur_Q <= QMax )
						flg = true;
					else
						flg = false;
				}
				if ( flg )
				{
					thick1_resu = thick1;		//厚度
					thick2_resu = thick2;
					tb_resu		= tb;			//界面温度

					//计算当前厚度下的散热密度
					this->Calc_Q_PipeCom(dQ);

					dDiffMin1	= dDiff1;		//最小的差值
					dDiffMin2	= dDiff2;
				}
			}
		}
	}
	return	1;
}


//------------------------------------------------------------------                
// DATE         :[2005/04/26]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :表面温度法，计算管道单层
//------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPipe_One(double &thick_resu, double &ts_resu)
{
	double	ts = ts_resu;			//表面温度法计算时,外表面温度为已知的。
	double	dDiff;					//临时变量	
	double  dDiffMin=0;				//最小的差值. 
	
	thick_resu = 0.0;				//保温厚为0.
	
	double  dCur_Q;			//当前厚度下的散热密度.
	double	QMax;			//根据当前记录的工况,允许最大散热密度
	bool	flg = true;		//进行散热密度的判断

	//查表获得介质温度允许的最大散热密度
	if( !GetHeatLoss( m_pRsHeatLoss, t, nYearVal, nSeasonVal, QMax ) )
	{ 		
	}
 
	for(thick=m_thick2Start;thick<=m_thick2Max;thick+=m_thick2Increment)
	{
		D1 = D0 + 2*thick;
		//根据不同的行业规范计算放热系数
		ALPHA =	GetPipeAlpha(ts, D1);

		//根据改变的外表面温度计算导热系数
		lamda = GetLamda( t, ts );
		//表面温度法，计算管道单层保温厚的公式
		dDiff = D1*log(D1/D0)-2000*lamda*(t-ts)/ALPHA/(ts-ta);
//		dDiff = D1 * log( D1 / D0 ) - 2000 * lamda * fabs( t - ts ) / ALPHA / fabs( ts - ta );
		if (fabs(dDiffMin) < 1e-6 || fabs(dDiffMin) > fabs(dDiff) )
		{
			flg = true;
			if ( bIsHeatLoss )
			{
				this->Calc_Q_PipeOne(dCur_Q);
				if ( fabs(QMax) <= 1e-6 || dCur_Q <= QMax )
					flg = true;
				else
					flg = false;
			}
			if ( flg )
			{
				thick_resu	= thick;		//传回计算出的保温厚度。
				dDiffMin	= fabs( dDiff );		//保存最小的差值

				//计算当前厚度下的散热密度
				this->Calc_Q_PipeOne(dQ);
			}
		}
	}
	return	1;
}


//--------------------------------------------------------------------------
// DATE         :[2005/05/26]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :给定厚度和外表面温度计算出界面温度()，平面双层。
//--------------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPlain_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu)
{
	thick1 = thick1_resu;
	thick2 = thick2_resu;		//传入指定的厚度.
	ts	   = ts_resu;

	//根据指定的厚度和外表面温度,计算界面温度
	CalcPlain_Com_TsTb(thick1,thick2,ts_resu,tb_resu,1);
	//计算散热密度
	Calc_Q_PlainCom(dQ);

	return 1;

}


//--------------------------------------------------------------------------
// DATE         :[2005/05/26]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :给定厚度和外表面温度计算出界面温度()，保温对象为管道，双层.
//--------------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPipe_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu)
{
	thick1 = thick1_resu;
	thick2 = thick2_resu;
	//计算界面温度.
	CalcPipe_Com_TsTb(thick1,thick2,ts_resu,tb_resu,1);
	//计算散热密度时用到的导热系数和放热系数在计算温度时已经得到了值.
	//散热密度
	Calc_Q_PipeCom(dQ);


	return 1;
}


//--------------------------------------------------------------------------
// DATE         :[2005/05/26]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :指定厚度,和外表面温度.只须要计算散热密度. 管道单层
//--------------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPipe_One_InputThick(double thick_resu, double &ts_resu)
{
	thick = thick_resu;
	ts	  = ts_resu;
	D1	  = D0 + 2 * thick;

	//导热系数
	lamda = GetLamda( t, ts );
	//根据不同的行业规范计算放热系数
	ALPHA = GetPipeAlpha(ts, D1);
	//根据以上初始的变量,重新计算散热密度.
	Calc_Q_PipeOne(dQ);

	return 1;
}
//--------------------------------------------------------------------------
// DATE         :[2005/05/26]
// Author       :ZSY
// Parameter(s) :
// Return       :
// Remark       :指定厚度,和外表面温度.只须要计算散热密度.		平壁单层
//--------------------------------------------------------------------------
short CCalcThick_FaceTemp::CalcPlain_One_InputThick(double thick_resu, double &ts_resu)
{
	thick = thick_resu;
	ts	  = ts_resu;

	//导热系数
	lamda = GetLamda( t, ts );
	//根据不同的行业规范计算放热系数
	ALPHA = GetPlainAlpha(ts);
	//根据以上初始的变量,重新计算散热密度.
	Calc_Q_PlainOne(dQ);

	return 1;
}
