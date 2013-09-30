// CalcThickSubterranean.cpp: implementation of the CCalcThickSubterranean class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "CalcThickSubterranean.h"
#include "exceptioninfohandle.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCalcThickSubterranean::CCalcThickSubterranean()
{

}

CCalcThickSubterranean::~CCalcThickSubterranean()
{

}

//------------------------------------------------------------------
// DATE         : [2006/03/06]
// Author       : ZSY
// Parameter(s) : dQ[out]: 单根的热损失，单位 [Kcal/(m*h)]
// Return       : short 1
// Remark       : 计算不通行地沟中敷设单根的热损失
//------------------------------------------------------------------
short CCalcThickSubterranean::Calc_Q_PipeOne(double &dOneQ)
{
	double dRso;
	Calc_R_One(R1);
	Calc_Rso(dRso);
	
	dOneQ = (Tf1 - Tso) / (R1 + Rs1 + Raw + dRso);
	
	return 1;
}

//------------------------------------------------------------------
// Parameter(s) : dQ1[out]:第一根的热损失，单位[Kcal/(m*h)]
//				: dQ2[out]:第二根的热损失，单位[Kcal/(m*h)]
// Return       : void
// Remark       : 计算不通行地沟中敷设两根(多根)时的热损失
//------------------------------------------------------------------
void CCalcThickSubterranean::Calc_Q_PipeTwo(double &dQ1, double & dQ2)
{
	double dTaw;
	Calc_Taw_Two(dTaw);

	double dR1 = QMax / 2;
	double dR2;
	double dRo;
	double dTmpA1;
	double dTmpA2;
	
/* 
	( a - q2 * (A1 * a2 * Ro^2 + Rso^2 - Ro^2) )		
		= ( Ro * Rso * A1 * q2^2 - q2^2 * (B + C) ) / q1 + Ro * Rso * a2 * q1;

	q1 * ( a - q2 * (A1 * a2 * Ro^2 + Rso^2 - Ro^2) ) - Ro * Rso * a2 * q1^2		
		= ( Ro * Rso * A1 * q2^2 - q2^2 * (B - C) );
*/

	do 
	{
		R11 = dR1;
		R12 = dR2;

		Calc_R_One(dR1);
		Calc_R_Two(dR2);
		Calc_Ro(dRo);

		dQ = QMax / 2;
		
		dQ1 = (Tf1 - dTaw) / (R11 + Rs1);
		dQ2 = (Tf2 - dTaw) / (R12 + Rs2);
 
	} while(fabs(dQ1 - dR1) > DZero && fabs(dQ2 - dR2) > DZero);

}

//------------------------------------------------------------------
// DATE         : [2006/03/06]
// Author       : ZSY
// Parameter(s) : dR[out]:第一根管道的热阻
// Return       : void
// Remark       : 无管沟的管道热阻
//------------------------------------------------------------------
void CCalcThickSubterranean::Calc_R_One(double &dR)
{
	double q1;		// 第一根的热损失
	double q2;		// 第二根的热损失

	Calc_Q_PipeTwo( q1, q2 );
	Calc_Q_noOne(dQ); 

	dR = (Tf1 - Tso) * q2 * Ro / q1;
}

//------------------------------------------------------------------
// DATE         : [2006/03/06]
// Author       : ZSY
// Parameter(s) : dR[out]:第二根管道的热阻
// Return       : void
// Remark       : 无管沟的管道热阻
//------------------------------------------------------------------
void CCalcThickSubterranean::Calc_R_Two(double &dR)
{
	double q1;		// 第一根的热损失
	double q2;		// 第二根的热损失
	double dRo;
 
	Calc_Q_PipeTwo( q1, q2 );
	Calc_Ro(dRo);

	dR = (Tf1 - Tso) * q1 * dRo / q2; 
}

//------------------------------------------------------------------
// DATE         : [2006/03/06]
// Author       : ZSY
// Parameter(s) : dQ[out]:单根管道的热损失
// Return       : void
// Remark       : 无管沟单根热损失
//------------------------------------------------------------------
void CCalcThickSubterranean::Calc_Q_noOne(double &dQ)
{
	double dRso;	// 土壤的热阻
	Calc_R_One(R1);
	Calc_Rso(dRso);
	
	double dQTmp;
	dQ = QMax / 2;
	double dA;
	double dB;
	dA = pow((Tf1 - Tso) / dRso, 2);
	
	do
	{
		dQTmp = dQ;
		Calc_R_Two(R1);
		dQ = (Tf1 - Tso) / (R1 + dRso);
		
	} while(fabs(dQ - dQTmp) < DZero);
	
}

//------------------------------------------------------------------
// DATE         : [2006/03/06]
// Author       : ZSY
// Parameter(s) : dQ1[out] : 第一根管道的热损失
//				: dQ2[out] : 第二根管道的热损失
// Return       : void
// Remark       : 无管沟多根热损失
//------------------------------------------------------------------
void CCalcThickSubterranean::Calc_Q_noTwo(double &dQ1, double & dQ2)
{
	Calc_R_One(R1);
	Calc_R_Two(R2);
	Calc_Ro(Ro);
	double dPreQ1;
	double dPreQ2;

	do 
	{
		dPreQ1 = dQ1;
		dPreQ2 = dQ2;
		
		dQ1 = ((Tf1 - Tso) * (R2 + Rso) - (Tf2 - Tso) * Ro) / ((R1 + Rso) * (R2 + Rso) - (Ro * Ro));
		dQ2 = ((Tf1 - Tso) * (R1 + Rso) - (Tf1 - Tso) * Ro) / ((R1 + Rso) * (R2 + Rso) - (Ro * Ro));

		Calc_R_One(R1);
		Calc_R_Two(R2);
		
	} while(fabs(dPreQ1 - dQ1) < DZero && fabs(dPreQ2 - dQ2) < DZero);

}

//------------------------------------------------------------------
// DATE         : [2006/03/07]
// Author       : ZSY
// Parameter(s) : dRso[out] 土壤的热阻
// Return       : void
// Remark       : 计算土壤的热阻
//------------------------------------------------------------------
void CCalcThickSubterranean::Calc_Rso(double &dRso)
{
	double Hm;	// 埋设深度计算值量
	if (sub_h <= DZero)
	{
		return;
	}
	Hm = sub_h + LamdaSo / AlphaSo;
	dRso = 1 / (2 * pi * LamdaSo) * log(2 * Hm * sqrt(4 * Hm * Hm + Dag * Dag) / Dag);
}

//------------------------------------------------------------------
// DATE         : [2006/03/07]
// Author       : ZSY
// Parameter(s) : ts[in] 外表面温度
// Return       : 导热系数
// Remark       : 土壤的导热系数
//-----------------------------------------------------------------
double CCalcThickSubterranean::GetLamdaSo(double ts)
{
	double dLSo;
	double dHm = 20+1E-22; //testzsy
	// 土壤的导热系数 
	// 化工设计手册上根据土壤的湿度来选择导热系数
	dLSo = 0.58;	// 干土壤
	//	dLSo = 1.163;	// 不太湿的土壤//	dLSo = 1.74;	// 较湿的土壤//	dLSo = 2.33;	// 很湿的土壤

//testzsy 
	m_nID++;
	if (fabs(ts-dHm) <= DZero)
	{
		return dLSo;
	}
	
	try
	{
		CString strSQL;
		if (m_pRsLSo == NULL || m_pRsLSo->GetState() == adStateClosed || (m_pRsLSo->GetadoEOF() && m_pRsLSo->GetBOF()))
		{
			return  1.163;	// 不太湿的土壤的导热系数
		}

		strSQL.Format("H=%f", dHm);
		m_pRsLSo->MoveFirst();
		m_pRsLSo->Find(_bstr_t(strSQL), NULL, adSearchForward);
		if (!m_pRsLSo->GetadoEOF())
		{
			dLSo = vtof(m_pRsLSo->GetCollect(_variant_t("Lso")));
			return dLSo;
		}
	}
	catch (_com_error& e)
	{
		ReportExceptionErrorLV2(e);
		return 0;
	}
	return dLSo;
}

//------------------------------------------------------------------
// DATE         : [2006/03/07]
// Author       : ZSY
// Parameter(s) : dRo[out] 当量热阻
// Return       : void
// Remark       : 两根管道之间的当量热阻
//------------------------------------------------------------------
void CCalcThickSubterranean::Calc_Ro(double &dRo)
{
//	LamdaSo = GetLamdaSo(ts);

	dRo = log(sqrt(1 + pow((2 * sub_h) / sub_b, 2))) / (2 * pi * LamdaSo);
}

//------------------------------------------------------------------
// DATE         : [2006/03/06]
// Author       : ZSY
// Parameter(s) : nID[in]  记录的ID号
// Return       : 成功返回TRUE，否则返回FALSE
// Remark       : 获得当前计算记录的埋地管道的保温计算参数
//------------------------------------------------------------------
BOOL CCalcThickSubterranean::GetSubterraneanOriginalData(long nID)
{
	try
	{
		if (m_pRsSubterranean == NULL || m_pRsSubterranean->GetState() == adStateClosed || (m_pRsSubterranean->GetadoEOF() && m_pRsSubterranean->GetBOF()) )
		{
			return FALSE;
		}
		Tf1 = t;	// 用不同的变量名来代替
		Tso = ta;

		CString	strSQL;		// SQL语句
		double	dTwoD0;		// 第二根管道的外径

		strSQL.Format("ID=%d",nID);
		m_pRsSubterranean->MoveFirst();
		m_pRsSubterranean->Find(_bstr_t(strSQL), 0, adSearchForward);		
		//存在指定序号的埋地管道保温计算的参数
		if ( !m_pRsSubterranean->adoEOF )
		{
			// 埋地管道敷设状态
			subLayMark		= vtoi( m_pRsSubterranean->GetCollect( _variant_t("c_lay_mark") ) );
			// 管道的埋设深度（地表面到管道中心的距离）
			sub_h			= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Sub_Depth") ) );
			// 土壤的温度℃
			Tso				= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Edaphic_Temp") ) );
			// 两根管道之间的当量间距
			sub_b			= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Span") ) );
			// 土壤的导热系数
			LamdaSo			= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Edaphic_Lamda") ) );
			// 第二根管道的外径
			dTwoD0			= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Two_D0") ) );
			
			Dag				= vtof( m_pRsSubterranean->GetCollect( _variant_t("c_Pipe_Dag") ) );  // 管沟当量直径
		}
		Rs1 = GetSubterraneanRs(D0, Tf1);
		// 获得保温层表面放热阻
		if (D1 > 0)
		{
			Rs2 = GetSubterraneanRs(D1, Tf1);
		}
		// 单位转换（-> mm)
		sub_h *= 1000;
		sub_b *= 1000;
		
		Aaw = A1pre = 9;	// 常量
	}
	catch (_com_error)
	{
		return FALSE;
	}
	
	return TRUE;
}

//------------------------------------------------------------------
// DATE         : [2006/03/08]
// Author       : ZSY
// Parameter(s) : D0[in] 管道或平壁的外径
//				: dTemp[in] 介质温度
//				: strPlace[in] 安装地点
// Return       : 保温层表面放热阻
// Remark       : 取出管道或平壁保温结构的外表面至周围空气的放热阻
//------------------------------------------------------------------
double CCalcThickSubterranean::GetSubterraneanRs(const double &D0, const double &dTemp, CString strPlace /*=室内*/)
{
	double dRsVal = 0;
	if (D0 < DZero)
	{
		return dRsVal;
	}
	try
	{
		if (m_pRsFaceResistance == NULL || m_pRsFaceResistance->GetState() == adStateClosed || m_pRsFaceResistance->GetRecordCount() <= 0)
		{
			return dRsVal; //
		}

		CString strSQL;
		strSQL.Format("dw=%f", D0);
		m_pRsFaceResistance->MoveFirst();
		
		if (strPlace.Compare("室内"))
		{
			strSQL += " AND PLACE<>'室內' ";
		}
		m_pRsFaceResistance->Find(_bstr_t(strSQL), 0, adSearchForward);

		BOOL bFirst = TRUE;	// 没有进行循环。
		double dPreVal = 0;	// 前一次的值
		double dPreKey = 0;	// 前一次的关键字
		double dCurVal;		// 当前的值
		double dCurKey;		// 当前的关键字
		for (; !m_pRsFaceResistance->adoEOF; m_pRsFaceResistance->MoveNext())
		{
			dCurVal = vtof(m_pRsFaceResistance->GetCollect(_variant_t("Val")));
			dCurKey = vtof(m_pRsFaceResistance->GetCollect(_variant_t("Key")));
			
			if (dCurKey >= D0)
			{
				if (bFirst)
					dRsVal = dCurVal;	// 须要查找的关键字小于或等于表中最小的关键值
				else
					dRsVal = (dCurVal - dPreVal) / (dCurKey - dPreKey) * (D0 - dPreKey) + dPreVal;

				return dRsVal;
			}
			bFirst = FALSE;
			dPreKey = dCurKey;
			dPreVal = dCurVal;
		}
		
		if ( bFirst == FALSE && m_pRsFaceResistance->adoEOF)	// 查找的关键字超过了记录集中最大的
		{
			dRsVal = dPreVal;	// 使用最大的值
		}
	}
	catch (_com_error& e)
	{
		ReportExceptionError(e);
		return dRsVal;
	}

	return dRsVal;
}

// 多根管道的不通行管沟内空气的温度
BOOL CCalcThickSubterranean::Calc_Taw_One(double& dTaw)
{
	if (D0 < DZero)
	{
		return FALSE;
	}

	double dRsVal = 0;
	double dQ1;
	double dQ2;
	double dRso;
	double dRo;
	if (K < DZero)
	{
		K = 1;
	}
	
	Calc_Q_PipeTwo(dQ1, dQ2);
	Calc_Rso(dRso);
	Calc_Ro(dRo);
	
	if(bTawCalc)	// 计算管沟内空气的温度有两个公式，根据变量（bTawCalc）来选择计算
		dTaw = Tso + K * (dQ1 + dQ2) * (Raw + dRso);
	else
		dTaw = (Tf1 / (R11 + Rs1) + Tf2 / (R12 + Rs2) + Tso / (K * (Raw + dRso))) / (1 / (R11 + Rs1) + 1 / (R12 + Rs2) + 1 / ( K * (Raw + dRso)));

	return TRUE;
}

// 单根管道的不通行管沟内空气温度
BOOL CCalcThickSubterranean::Calc_Taw_Two(double& dTaw)
{
	if (D0 <= DZero)
	{
		return FALSE;
	}

	if (K <= DZero)//
	{
		K = 1;
	}

	double dRso;
	double dQ;

	Calc_Rso(dRso);
	Calc_Q_PipeOne(dQ); // 放热系数
	
	if (bTawCalc)
		dTaw = Tso + K * dQ * (Raw + dRso);
	else
		dTaw = Tso + (K * (Tf1 - Tso) * (Raw + Rso)) / (R1 + Rs1 + Raw + Rso);

	return TRUE;
}
