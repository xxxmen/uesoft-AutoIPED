// CalcThick_SubterraneanNoChannel.cpp: implementation of the CCalcThick_SubterraneanNoChannel class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcThick_SubterraneanNoChannel.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThick_SubterraneanNoChannel::CCalcThick_SubterraneanNoChannel()
{

}

CCalcThick_SubterraneanNoChannel::~CCalcThick_SubterraneanNoChannel()
{

}

//------------------------------------------------------------------
// Parameter(s) : dQ1 [out] 第一根管道的热损失
//				: dQ2 [out] 第二根管道的热损失
// Return       : void http://www.sailordata.com
// Remark       : 无管沟两根管道的热损失
//------------------------------------------------------------------
void CCalcThick_SubterraneanNoChannel::Calc_Q_PipeTwo(double &dReQ1, double &dReQ2)
{

	double dRso;	// 土壤的热阻
//	double dR1;		// 第一根管道保温层热阻
//	double dR2;		// 第二根
	double dQ1;
	double dQ2;
	double dRo;
	
	Calc_Rso(dRso);
	Calc_Ro(dRo);

//	Tf1	= 20;
	Tf2 = Tf1;	// test

//	dRo	= 0.1;
//	dRso = 2.2;

	double Tp1 = (Tf1 - Tso);	// 管一根管道，介质与地壤的温差
	double Tp2 = (Tf2 - Tso);	// 管二根管道，介质与地壤的温差
	
	double dTmpA;	// 中间的常数
	double dTmpB;
	double dTmpC;
	double dTmpDt1;
	double dTmpDt2;
	double dTmpMr;
	
	double dTmpDelta1;
	double dTmpDelta2;
//		double B1 = (dR1 + dRso);
//		double B2 = (dR2 + dRso);

	double dCalcQ1A;	// 方程的已知量，（ax^2 + bx + c = 0）
	double dCalcQ1B;
	double dCalcQ1C;
	double dCalcQ2A;
	double dCalcQ2B;
	double dCalcQ2C;
	
	double dQi	= QMax;
	double dQii = QMax;
//		double M;
//		double C1 = 0;
	
	// 设置一个初始值
	dReQ1 = dQ1 = QMax / 2;
	dReQ2 = dQ2 = QMax / 2;
//	do                                     
//	{
		dReQ1 = dQ1;
		dReQ2 = dQ2;
//		Calc_Rpipe_Two(dQ1, dQ2, dR1, dR2);
		
		Calc_Rso(dRso);
		Calc_Ro(dRo);

		// 中间的已知量用一些变量来代替
		dTmpA = Tp2 * dRo * dRso;
		dTmpB = Tp1 * dRo * dRso;
		dTmpC = Tp1 * Tp2 * pow(dRo, 2) + pow(dRso, 2) - pow(dRo, 2);
		dTmpMr = Tp1 * Tp2 * dRo;
		dTmpDt1 = Tp1 * dRso - dRo * Tp2;
		dTmpDt2 = Tp2 * dRso - dRo * Tp1;
		
		// 第一根管
		dCalcQ1A = dTmpA + dTmpB * pow((dTmpDt2 - dTmpMr) / (dTmpDt1 - dTmpMr), 2) + 
					(dTmpDt2 - dTmpMr) / (dTmpDt1 - dTmpMr) * dTmpC;

		dCalcQ1B = -(dTmpDt2 + (dTmpDt2 - dTmpMr) / (dTmpDt1 - dTmpMr) * dTmpMr);
		
		dCalcQ1C = 0;

		// 第二根管
		dCalcQ2A = dTmpA * pow((dTmpDt1 - dTmpMr) / (dTmpDt2 - dTmpMr), 2) + dTmpB + 
					(dTmpDt1 - dTmpMr) / (dTmpDt2 - dTmpMr) * dTmpC;
	
		dCalcQ2B = -(dTmpDt1 + (dTmpDt1 - dTmpMr) / (dTmpDt2 - dTmpMr) * dTmpMr);

		dCalcQ2C = 0;
 
		// b^2 - 4*a*c 
		dTmpDelta1 = pow(dCalcQ1B, 2) - 4 * dCalcQ1A * dCalcQ1C; 
		dTmpDelta2 = pow(dCalcQ2B, 2) - 4 * dCalcQ2A * dCalcQ2C;
		if (dTmpDelta1 >= 0)
		{
			// 管道的热损失值
			dReQ1 = (-dCalcQ1B - dCalcQ1B) / (2 * dCalcQ1A);		// (-b - sqrt(b^2 - 4*a*c) ) / (2*a)

			dReQ2 = (-dCalcQ2B - dCalcQ2B) / (2 * dCalcQ2A);	// (-b - sqrt(b^2 - 4*a*c) ) / (2*a)
		}
		else
		{
			dReQ1 = dReQ2 = 0;
		}

//		dQ1 = ((Tf1 - Tso) * (dR2 + dRso) - (Tf2 - Tso) * dRo) / ((dR1 + dRso) * (dR2 + dRso) - pow(dRo, 2));
//		dQ2 = ((Tf1 - Tso) * (dR1 + dRso) - (Tf1 - Tso) * dRo) / ((dR1 + dRso) * (dR2 + dRso) - pow(dRo, 2)); 

//	} while(fabs(dReQ1 - dQ1) > TsDiff || fabs(dReQ2 - dQ2) > TsDiff); 

}

//------------------------------------------------------------------
// Parameter(s) : dOneQ [out] 單根管道热损失
// Return       : 1
// Remark       : 无管沟單根管道热损失
//------------------------------------------------------------------
short CCalcThick_SubterraneanNoChannel::Calc_Q_PipeOne(double &dOneQ)
{
	double dRso;	// 土壤的热阻
	double dR1;		// 保温层热阻
	double dLamda;	// 保温层的导热系数
	
	dLamda = GetLamda(Tf1, Tso); //  
	D1 = D0 + 2 * thick;	// 保温层外径

	Calc_Rso(dRso);
	Calc_Rpipe_One(D1, dLamda, dR1);

	dOneQ = (Tf1 - Tso) / (dR1 + dRso);

	if (dOneQ < DZero)
	{ 		
		dOneQ = fabs(dOneQ);
	}

	return 1;
}

//------------------------------------------------------------------
// DATE         : [2006/03/28]
// Author       : ZSY
// Parameter(s) : dRso[out] 土壤热阻
// Return       : void
// Remark       : 无管沟敷设土壤热阻
//------------------------------------------------------------------
void CCalcThick_SubterraneanNoChannel::Calc_Rso(double &dRso)
{
	if (sub_h < DZero) // 管中心距地面的距离的变量小于零
	{
		dRso = Rso;
		return;
	}
	float dLSo;
	dLSo = GetLamdaSo(ts);

	D1 = D0 + 2 * thick;

	dRso = log(4 * sub_h / D1) / (2 * pi * dLSo);
 
	return;

	try
	{
		CString strSQL;
		CString strPipeLay;
		_RecordsetPtr pRs;
		pRs.CreateInstance(__uuidof(Recordset));
		strSQL = "select * from [埋地管道敷设状态]";
		pRs->Open(_variant_t(strSQL), _variant_t(m_pConAutoIPED.GetInterfacePtr()), \
			adOpenStatic, adLockOptimistic, adCmdText);
		if (pRs->GetRecordCount() < 0)
		{
			return;
		}
		strSQL.Format("index=%d", m_nID);
		pRs->Find(_bstr_t(strSQL), 1, adSearchForward);
		if (!pRs->adoEOF)
		{
			strPipeLay = vtos(pRs->GetCollect(_variant_t("pipelay")));
			m_dMaxD0 = vtof(pRs->GetCollect(_variant_t("Formula")));
		}
		if (strPipeLay.CompareNoCase(_T("有管沟不通行")))
		{
			for (pRs->MoveFirst(); !pRs->adoEOF; pRs->MoveNext())
			{
				strPlace = vtos(pRs->GetCollect(_T("Formula")));
				out_mat = vtos(pRs->GetCollect(_T("Index")));
			}
		}
	
	}
	catch (_com_error& e)
	{
		AfxMessageBox(e.Description());
		return;
	}
}

//------------------------------------------------------------------
// DATE         : [2006/03/28]
// Author       : ZSY
// Parameter(s) : dQ1[in] 第一根管道的散热密度
//				: dQ2[in] 第二根管道的散热密度
//				: dR1[out] 第一根管道的热阻
//				: dR2[out]	第二根管道的热阻
// Return       : void
// Remark       : 无管沟管道热阻 (两根管道敷设)
//------------------------------------------------------------------
void CCalcThick_SubterraneanNoChannel::Calc_Rpipe_Two(double dQ1, double dQ2, double &dR1, double &dR2)
{
	double dRo;
	Calc_Ro(dRo);

	dR1 = (Tf1 - Tso) * dQ2 * dRo / dQ1;
	dR2 = (Tf2 - Tso) * dQ1 * dRo / dQ2; 
}

//------------------------------------------------------------------
// Parameter(s) : D1[in] 保温层外径
//				: dLamda[in] 保温层的导热系数
//				: dR1[out] 保温层热阻
// Return       : void
// Remark       : 无管沟的单根保温层热阻
//------------------------------------------------------------------
void CCalcThick_SubterraneanNoChannel::Calc_Rpipe_One(double D1, double dLamda, double &dR1)
{
	D1 = D0 + 2 * thick;

	dR1 = log(D1 / D0) / (2 * pi * dLamda);
}

//------------------------------------------------------------------
// Parameter(s) : dRo [out] 当量热阻
// Return       : void
// Remark       : 无管沟敷设的两根管道之间的当量热阻
//------------------------------------------------------------------
void CCalcThick_SubterraneanNoChannel::Calc_Ro(double &dRo)
{
	float dLSo = GetLamdaSo(ts); 

	dRo = log(sqrt(1 + pow((2 * sub_h) / sub_b, 2))) / (2 * pi * dLSo); 
}

//------------------------------------------------------------------                
// Author       : ZSY
// Parameter(s) : thick_resu[out] 保温厚度， ts_resu[out] 保温层外表面温度
// Return       : 
// Remark       : 计算，无管沟直埋敷设保温层厚度, 和保温层外表面温度
//------------------------------------------------------------------
short CCalcThick_SubterraneanNoChannel::CalcPipe_One(double &thick_resu, double &ts_resu)
{
	double	ts=0.0;			//外表面温度值
	double	dYearVal;		//年运行工况的最大散热密度 
	double	dSeasonVal;		//季节运行工况最大散热密度 
	double	Q = 1;			//当前厚度下的散热密度.
	double	QMax = 0;		//根据当前记录的工况;允许最大散热密度

	double	dLSo;			// 土壤的导热系数
	double	dASo;			//
	double  dDiff;
	double  dMinDiff = 0;
	bool	flg		 = true;//进行散热密度的判断

	thick_resu  = 0;
	ts_resu		= 0;
	// 获得原始数据 	
	GetSubterraneanOriginalData(m_nID);
	GetHeatLoss(m_pRsHeatLoss, Tf1, dYearVal, dSeasonVal, QMax);
	dQA = QMax;
	Calc_Ro(Ro);
	if (Ro <= DZero)
	{
		Ro = 0.6;
	}
	if (fabs(QMax) >= DZero)
	{
		Q	= QMax * m_MaxHeatDensityRatio;			//保温结构外表面允许散热密度，按表中查出的允许最大散热密度的乘以一个系数取值.(电力行业规程规定为90%).
		dQA = QMax;
	}
	else
	{
		dQA = fabs(QMax);	// 当热损失为一个负数时
	}
	//thick单层保温厚度。
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表ThicknessRegular获得

	for (thick=m_thick2Start; thick <= m_thick2Max; thick += m_thick2Increment)
	{
		D1 = D0 + 2 * thick;

		//根据指定的保温厚，计算外表面温度    
		CalcPipe_One_TsTemp(thick, ts);
		
		dLSo  = GetLamdaSo(ts);		// 土壤的导热系数
		lamda = GetLamda(t, ts);	// 保温层的导热系数
		dASo  = GetPipeAlpha(ts, D1); 
		dDiff = log(D1 / D0) - 2.0 * pi * lamda * dLSo / (dLSo - lamda) * ((Tf1 - Tso) / dQA - (log(4 * sub_h / D0) / (2 * pi * dLSo)));

		if (ts<MaxTs)
		{
			if (fabs(dMinDiff) <= DZero || fabs(dDiff) < fabs(dMinDiff))
			{
				//选项中选择的判断散热密度,就进行比较
				if(bIsHeatLoss)
				{
					//求得本次厚度的散热密度   
					Calc_Q_PipeOne(Q);				
					if(fabs(QMax) <= DZero || Q <= QMax)
					{
						flg = true; // 当前厚度满足允许最大散热规则
					} 
					else
					{
						flg = false;
					}
				}
				if( flg )
				{
					dMinDiff   = fabs(dDiff);
					thick_resu = thick;
					ts_resu	   = ts;
					//计算当前厚度下的散热密度
					Calc_Q_PipeOne(dQ);
				}
			}
		}
	}
	return 1;
}

 
//------------------------------------------------------------------
// Parameter(s) : thick1_resu 第一根管道的保温厚度
//				：thcik2_resu 第二根管道的保温厚度
//				: ts1_resu 表面温度
//				: ts2_resu 
// Return       : 
// Remark       : 地下直埋敷设的双根管道保温厚度计算
//------------------------------------------------------------------
short CCalcThick_SubterraneanNoChannel::CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &ts1_resu, double &ts2_resu)
{

	double	ts = 0.0;			//外表面温度值
	double  ts2 = 0.0;
	double	dYearVal;		//年运行工况的最大散热密度 
	double	dSeasonVal;		//季节运行工况最大散热密度 
	double	Q = 1;				//当前厚度下的散热密度.
	double	QMax = 0;			//根据当前记录的工况;允许最大散热密度

	double	dLSo;	// 土壤的导热系数
	double	dASo;	//
	double  dMinDiff1 = DZero;
	double  dMinDiff2 = DZero;
	double  dDiff1;
	double	dDiff2;
	double  dQ1;
	double  dQ2;
// testzsy
	Tso = 12;
	double	Dii0;
	double  Dii1;
	CString strTmp;
	bool	flg=true;		//进行散热密度的判断
	
	thick1_resu = thick2_resu = 0;
	ts1_resu = ts2_resu = 0;
	if( !GetSubterraneanOriginalData(m_nID))
	{
		strTmp.Format("序号为%d的记录，在读取原始数据时出错！\r\n", m_nID);
		strExceptionInfo += strTmp;
		return 0;
	}
	GetHeatLoss(m_pRsHeatLoss, t, dYearVal, dSeasonVal, QMax);

	if (fabs(QMax) >= DZero)
	{
		Q = QMax * m_MaxHeatDensityRatio;			//保温结构外表面允许散热密度，按表中查出的允许最大散热密度的乘以一个系数取值.(规程规定为90%).
	}
	else
	{
		dQA = QMax; 
	}
	Calc_Ro(Ro);
	if (Ro <= DZero)
	{
		Ro = 0.6;
	}
	//thick单层保温厚度。
	//thick2Start外层保温允许最小厚度,thick2Max外层保温允许最大厚度,thick2Increment外层保温厚度最小增量
	//从保温厚度规则表ThicknessRegular获得 

	for (thick1=m_thick1Start; thick1 <= m_thick1Max; thick1 += m_thick1Increment)
	{
		for (thick2 = m_thick2Start; thick2 <= m_thick2Max; thick2 += m_thick2Increment)
		{
			D1 = D0 + 2 * thick1;
			D2 = D1 + 2 * thick2;
			Dii1 = D2;
			Dii0 = D0;
			
			//根据指定的保温厚，计算外表面温度
//			CalcPipe_Com_TsTb(thick1, thick2, ts, tb);
			
//			CalcSubPipe_TsTemp(D0, thick1, ts, 1);
//			CalcSubPipe_TsTemp(Dii1, thick2, ts2, 2);

			Calc_Q_PipeTwo(dQ1, dQ2);
			
			dLSo   = GetLamdaSo(ts);	// 土壤的导热系数
			lamda  = GetLamda(t, ts);	// 保温层的导热系数			
			dASo   = GetPipeAlpha(ts, D1);
			dDiff1 = log(D1 / D0) - 2.0 * pi * lamda * dLSo / (dLSo - lamda) *
				     ((Tf1 - Tso - dQ2 * Ro) / dQ1 - (log(4 * sub_h / D0) / (2 * pi * dLSo)));
			dDiff2 = log(Dii1 / Dii0) - 2.0 * pi * dLSo * lamda2 / (dLSo - lamda2) *
					 ((Tf2 - Tso - dQ1 * Ro) / dQ2 - log(4 * sub_h / Dii0) / (2 * pi * dLSo));
			dDiff1 = fabs(dDiff1);
			dDiff2 = fabs(dDiff2);
			
			if (ts < MaxTs && ts2 < MaxTs)
			{
				if ((dMinDiff1 <= DZero && dMinDiff2 <= DZero) || (dDiff1 < dMinDiff1 && dDiff2 < dMinDiff2))
				{
					flg = true;
					//选项中选择的判断散热密度,就进行比较
					if(bIsHeatLoss)
					{
						//求得本次厚度的散热密度
						Calc_Q_PipeTwo(dQ1, dQ2);
						
						if(fabs(QMax) <= DZero || (dQ1 <= QMax && dQ2 <= QMax))
						{
							flg = true; // 当前厚度满足允许最大散热规则 
						}
						else 
						{
							flg = false; 
						}
					}
					//满足条件
					if( flg )
					{
						dMinDiff1   = dDiff1;
						dMinDiff2   = dDiff2;
						thick1_resu = thick1;
						thick2_resu = thick2;
						ts1_resu	= tb;
						ts2_resu    = ts;
						
						//计算当前厚度下的散热密度
						Calc_Q_PipeOne(dQ);
						//	Calc_Q_PipeTwo(dQ1, dQ2);
					}
				}
			}
		}
	}
	return 1;
}/*

//------------------------------------------------------------------	
short CCalcThick_SubterraneanNoChanneal::CalcPlain_One_TsTemp(const double delta, double &dResTs)
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

		ts = (delta/1000/lamda*ta+1/ALPHA*t)/(delta/1000/lamda+1/ALPHA);
		if (fabs(ts-dResTs) < TsDiff || nBreak >= MaxCycCount)
		{
			break;
		}
		dResTs = ts;	//
		nBreak++;
	}
	return 1;
}
*/





