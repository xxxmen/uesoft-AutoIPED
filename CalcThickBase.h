// CalcThickBase.h: interface for the CCalcThickBase class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCTHICKBASE_H__1BF4CE0E_9FFC_41C7_B7AF_6E72380B8356__INCLUDED_)
#define AFX_CALCTHICKBASE_H__1BF4CE0E_9FFC_41C7_B7AF_6E72380B8356__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "CalcOriginal.h"
#include "CalcAlpha_Code.h"
#include "CalcAlpha_CodePC.h"
#include "CalcAlpha_CodeGB.h"
#include "CalcAlpha_CodeCJJ.h"
#include "vtot.h"

#define  INT_ADD	9			//保温厚度圆整时加的一个数
#define  DZero		(1e-6)		// 比较时作为零
//////////////////////////////
//表面温度法.			CCalcThick_FaceTemp
//散热(热流)密度法		CCalcThick_HeatDensity
//经济厚度法			CCalcThick_Economy
//防冻保温法			CCalcThick_Antifreezing
class CCalcThickBase  : public CCalcOriginalData ,public CCalcAlpha_Code,public CCalcAlpha_CodePC,public CCalcAlpha_CodeCJJ
{
public:
	//
	virtual short CalcPipe_One(double &thick_resu, double &ts_resu);
	virtual short CalcPipe_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPlain_One(double &thick_resu, double &ts_resu);
	virtual short CalcPlain_Com(double &thick1_resu, double &thick2_resu, double &tb_resu, double &ts_resu);

	//指定厚度
	virtual short CalcPipe_One_InputThick(double thick_resu, double &ts_resu);
	virtual short CalcPlain_One_InputThick(double thick_resu, double &ts_resu);
	virtual short CalcPipe_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu);
	virtual short CalcPlain_Com_InputThick(double thick1_resu, double thick2_resu, double &tb_resu, double &ts_resu);

	
	virtual short CalcPlain_One_TsTemp(const double delta, double &dResTs);		//平壁单层，外表面温度。
	virtual short CalcPipe_One_TsTemp(const double delta, double &dResTs);
	
	virtual short CalcPlain_Com_TsTb(const double delta1, const double delta2, double &dResTs, double &dResTb, int flg=0);
	virtual short CalcPipe_Com_TsTb(const double delta1, const double delta2, double &dResTs, double &dResTb, int flg=0);
	
	void CalcSubPipe_TsTemp(const double dPipeTemp, const double dDW, const double delta, double &dResTs, short nMark = 0);	// 埋地管的表面温度

	CCalcThickBase();
	virtual ~CCalcThickBase(); 

public:
	BOOL InitCalcAlphaData(double ts);
	double GetPipeAlpha(double ts, double D1);
	double GetPlainAlpha(double ts, double D1=0.0);
	short GetExceptionInfo(CString& strInfo);
//	short Calc_Q_PipeOneNoInit(double& Q);
//	short Calc_Q_PlainOneNoInit(double& Q);
	inline 	virtual short Calc_QL_PipeCom(double &QL);
	inline 	virtual short Calc_QL_PipeOne(double &QL);
	inline 	virtual short Calc_Q_PlainCom(double& Q);
	inline	virtual short Calc_Q_PlainOne(double& Q);
	inline	virtual short Calc_Q_PipeCom(double& Q);
	inline	virtual short Calc_Q_PipeOne(double& Q);
	
//	double	dQ;				//确定保温厚度下的散热密度

protected:
	inline	double GetLamda2(double tb, double ts);	// 外保温材料的导热系数
	inline	double GetLamda1(double t, double tb);	// 内保温材料的导热系数
	inline	double GetLamda(double t, double ts);	// 单层保温材料的导热系数

	inline  double GetLamdaA(double t1, double t2, short nMark=0);	// 获得导热系数，nMark=0是内保温材料或单层的， 1为外保温材料

	bool GetHeatLoss(_RecordsetPtr& pRs,double& t,double& YearRun, double& SeasonRun, double &QMax);

	virtual short GetThicknessRegular();

/*	_ConnectionPtr m_pConAutoIPED;		//连接项目数据库(AutoIPED.mdb)的指针.
	_ConnectionPtr m_pConIPEDCode;		//规范数据库(IPEDCode.mdb)
	_ConnectionPtr m_pConMaterial;		//材料数据库(Material.mdb)

	_RecordsetPtr m_pRsHeatLoss;		//保温材料放热系数表.()
	CString m_strThickFormula;		//计算保温厚的公式.
	CString m_strThick1Formula;		//计算复合保温时内层的公式.
	CString m_strThick2Formula;		//计算复合保温时外层的公式.

	CString m_strQFormula;			//计算放热系数的公式
	CString m_strTbFormula;			//计算界面温度的公式.
	CString m_strTsFormula;			//计算外表面温度的公式

	double	m_dThick;
	*/

private:
	float GetMaxHeatLoss(_RecordsetPtr &pRs, double &t, CString strFieldName);

};
#endif // !defined(AFX_CALCTHICKBASE_H__1BF4CE0E_9FFC_41C7_B7AF_6E72380B8356__INCLUDED_)
