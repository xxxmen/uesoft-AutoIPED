// CalcThickSubterranean.h: interface for the CCalcThickSubterranean class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICKSUBTERRANEAN_H__7E20FEB9_1A1A_4D8D_B10A_345D92112F79__INCLUDED_)
#define AFX_CALCTHICKSUBTERRANEAN_H__7E20FEB9_1A1A_4D8D_B10A_345D92112F79__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
//#include "afx.h"
#include "CalcThickBase.h"
//#include <afxtempl.h>

class CCalcThickSubterranean : public CCalcThickBase
{
public:
	BOOL Calc_Taw_One(double& dTaw);
	BOOL Calc_Taw_Two(double& dTaw);
	double GetSubterraneanRs(const double& D0,const double& dTemp, CString strPlace="室内");
	virtual BOOL GetSubterraneanOriginalData(long nID);
	void Calc_Ro(double& dRo);
	double GetLamdaSo(double ts);	// 土壤的导热系数
	virtual void Calc_Rso(double& dRso);
	virtual void Calc_Q_noTwo(double &dQ1, double & dQ2);
	inline	virtual void Calc_Q_noOne(double& dQ);
	inline	virtual void Calc_R_Two(double& dR);
	inline	virtual void Calc_R_One(double& dR);
	inline	virtual void Calc_Q_PipeTwo(double& dQ1, double & dQ2);
	inline	virtual short Calc_Q_PipeOne(double& dQ);

	CCalcThickSubterranean();
	virtual ~CCalcThickSubterranean();
	
protected:		// 埋地管道保温新增加的变量
	int		subLayMark;		// 埋地的敷设状态
	BOOL	bTawCalc;		// 

//	double  subEdaphicTemp;	// 土壤的温度
	double  Tso;			// 土壤温度	

//	double  subEdaphicLamda;// 土壤的导热系数
	double  LamdaSo;		// 土壤的导热系数
	
//	double	subPipeDeep;	// 管道的埋设深度
	double  sub_h;			// 管道埋设深度( 管中心到地面距离 ) m 
//	double  subPipeSpan;	// 两管道之间的距离
	double  sub_b;			// 两根管道的中心距离 m 

//	double  Taw;
	double  R11;
	double  R12;
	double  Rs1;
	double  Rs2;
	double  Tf1;	// 第一根管道的介质温度
	double  Tf2;	// 第二根管道的介质温度
	double	Ro;		// 两根管道相互间影响的当量热阻
	double  Rso;	// 地壤的热阻 
	double  R1;		// 第一根管道热阻
	double  R2;		// 第二根管道热阻
	double  Raw;	// 管沟内空气致管沟内壁的热阻
	double  AlphaSo;// 土壤的放热系数
	double  Dag;	// 管沟的当量直径
	double  dQA;	// 直埋单根的允许热损失
	double	Aaw;	// 管沟内空气至管沟壁的给热系数
	double  A1pre;	// Aaw = A1pre = 9
	double  K;

};

#endif // !defined(AFX_CALCTHICKSUBTERRANEAN_H__7E20FEB9_1A1A_4D8D_B10A_345D92112F79__INCLUDED_)
