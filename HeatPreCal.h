// HeatPreCal.h: interface for the CHeatPreCal class.
//
// 完成保温的计算
//
// 在调用C_ANALYS进行计算前
//
// 需要先调用		
//		SetProjectName;			设置工程名
//		SetConnect;				设置工程数据库连接
//		SetAssistantDbConnect(IConnectionCODE);	设置辅助（公共）数据库
//		SetMaterialPath(gsProjectDBDir);		设置材料数据库的路径
//
//////////////////////////////////////////////////////////////////////



#if !defined(AFX_HEATPRECAL_H__DE887BF7_4B9B_40F3_92E5_FAD4E263A3DF__INCLUDED_)
#define AFX_HEATPRECAL_H__DE887BF7_4B9B_40F3_92E5_FAD4E263A3DF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "FoxBase.h"
#include "GetPropertyofMaterial2.h"	// Added by ClassView
#include "CalcThick_FaceTemp.h"
#include "CalcThick_Economy.h"
#include "CalcThick_HeatDensity.h"
#include "CalcThick_PreventCongeal.h"
#include "CalcSupportSpanDll.h"
#include "CalcThick_SubterraneanObstruct.h"
#include "CalcThick_SubterraneanNoChannel.h"

class CHeatPreCal:public CFoxBase,public CCalcThick_FaceTemp,public CCalcThick_Economy,public CCalcThick_HeatDensity,public CCalcThick_PreventCongeal,public CCalcThick_SubterraneanObstruct, public CCalcThick_SubterraneanNoChannel
{
public:
	CHeatPreCal();
	virtual ~CHeatPreCal();

//method
public:

	//保温计算
	void C_ANALYS();

// property
public:
	virtual	void MinimizeAllWindow(){}
	void GetWeatherProperty(CString strVariableName, double& dRValue);
	double GetSupportSpan(const double dw,const double s,const CString strMaterialName,const double dPressure,const double dMediumDensity,const double temp,const double dInsuWei);
	static double GetMediumDensity(const CString strMediumName, CString* pMediumName = NULL);
	static BOOL SetConstantDefine(CString strTblName, CString ConstName, _variant_t ConstValue);
	//在规范数据库中访问常数表.
	static bool GetConstantDefine(CString strTblName, CString ConstName, CString& ConstValue);
	static bool GetConstantDefine(CString strTblName, CString ConstName, double& dConstValue);
	//返回Material数据库所在的文件夹
	LPCTSTR GetMaterialPath();

	//设置Material数据库的路径 
	void SetMaterialPath(LPCTSTR pMaterialPath); 
 
	//返回工程名(ID)
	CString& GetProjectName();

	//设置工程名(ID)
	void SetProjectName(CString &Name);

	//设置辅助数据库的连接
	void SetAssistantDbConnect(_ConnectionPtr &IConnection);

	//返回辅助数据库的连接
	_ConnectionPtr& GetAssistantDbConnect();

	//返回计算选择
	int GetChCal();

	//设置计算选择
	void SetChCal(int Ch);

protected:
	//打开辅助数据库的表(在规范数据库中)
	BOOL OpenAssistantProjectTable(_RecordsetPtr &IRecordset, LPCTSTR pTableName);

	//打开某个工程的表的记录
	BOOL OpenAProjectTable(_RecordsetPtr &IRecordset,LPCTSTR pTableName,LPCTSTR  pProjectName=NULL);


protected:

	int GetCalcFormula(int nFormulaIndex, CString& strFormula, CString& strTsFormula, CString& strQFormula);

	short CalcInsulThick(CString strThickFormula,CString strThickFormula2,CString strTbFormula, CString strTsFormula, CString strQFormula ,double &dThick1, double &dThick2);	//传入计算保温厚度的公式,初始各参数,并计算.

	//获得室内的设备或管道保温结构外表面传热系数。
	bool GetAlfaValue(_RecordsetPtr pRs, CString sSearchFieldName,  CString sValFieldName, double dCompareVal,double& dAlfaVal);
	double CalcAlfaValue(double dTs, double dD1);
	//根据介质温度,取保温结构外表面允许最大散热密度
	bool GetHeatLoss(_RecordsetPtr& pRs,double& t,double& YearRun, double& SeasonRun, double &QMax);

	bool InsertToReckonCost(double& ts,double&tb,double& thick1,double& thick2,double& cost_min, bool bIsFirst=false);

	virtual short GetThicknessRegular();
	bool E_Predat_Cubage();

	//当需要选择需要计算的数据时调用
	virtual void SelectToCal(_RecordsetPtr &IRecordset,int *pIsStop);

	//在需要指定计算的‘开始’与‘结束’范围时调用
	virtual	BOOL RangeDlgshow(long &Start,long &End);

	//判断是否停止计算时调用此函数
	virtual void Cancel(int *pState);

protected:
	// 从Material数据库中的CL_OLD_NEW表中寻找与MatName材料名相对应的材料。
	BOOL FindMaterialNameOldOrNew(LPCTSTR MatName,CString* &pFindName,int &Num);

	void temp_plain_one(double &thick_t,double &ts);
	void plain_one(double &thick_resu,double &ts_resu);
	void pip_one(double &thick_resu,double &ts_resu);
	void plain_com(double &thick1_resu,double &thick2_resu,double &tb_resu,double &ts_resu);
	void pip_com(double &thick1_resu,double &thick2_resu,double &tb_resu,double &ts_resu);
	void SUPPORT(double &dw, double &s, double &temp, double &wei1,double &wei2,double &lmax, CString &v, CString &cl,BOOL &pg,CString &st, double &t0);
	void temp_pip_com(double &thick1_t,double &thick2_t,double &ts,double &tb);
	void temp_plain_com(double &thick1_t,double &thick2_t,double &ts,double &tb);
	void temp_pip_one(double &thick_t,double &ts);

protected:
	double Temp_env;		// 环境温度
	double heat_price;		// 主汽热价
	double S;				// 年费用率

	double out_a0,out_a1,out_a2;	// 外保温材料的导热系数
	double in_a0,in_a1,in_a2;		// 内保温材料的导热系数
	double out_price,in_price,pro_price;	// 材料价格
	double in_dens,out_dens,pro_dens;		// 材料密度
	double in_ratio;	// 内保温材料的导热系数增加的比率
	double out_ratio;	// 外保温材料的导热系数增加的比率
	
	double	Mhours;		// 年运行小时数
	double  Yong;		// 热价比主汽价(介质用值系数)
	double	Temp_pip;	//介质温度
	double	D0;			//外径
	double	hedu;		//黑度

	CString out_mat;		//外保温材料名称
	CString	in_mat;			//内保温材料名称
	CString	pro_mat;		//保护材料名称
	CString	pro_name;		//保护材料名称,是去掉后面括号的名称.如"铝皮(0.75)"则该值为"铝皮"

	double	co_alf;
	double	distan;
	double	in_tmax;
	double	speed_wind;		//风速.
	double	B;				//沿风速方向的平壁宽度
	double	m_dPipePressure;	//管内压力
	double	m_dMediumDensity;	//介质密度
	//传热系数的索引.
	long	m_nIndexAlpha;

	double	m_HotRatio;		//保温界面温度不应超过外层保温材料最高使用温度的比例.
	double	m_CoolRatio;	//保冷界面温度不应超过外层保冷材料最高使用温度的比例.   
	double	MaxTs;			//外表面允许最大温度 

	//GLOBAL ?

	long	start_num;
	long	stop_num;

	//计算方式选择
	int ch_cal;
	//输入信息
	CString ch;
	//运行工况.
	//0:季节运行工况.
	//1:年运行工况.
	int hour_work;

	double	m_dMaxD0;	//平壁与圆管的分界外径.  以前为2000
	long	m_nID;			//原始数据当前记录的序号

	//
	double	thick_3;
	
	int m_thick1Start;			//内层保温允许最小厚度	
	int m_thick1Max;			//内层保温允许最大厚度
	int m_thick1Increment;		//内层保温厚度最小增量

	int m_thick2Start;			//外层保温允许最小厚度
	int m_thick2Max;			//外层保温允许最大厚度
	int m_thick2Increment;		//外层保温厚度最小增量

	_RecordsetPtr m_pRsCon_Out;		//对应设备的外表面放热系数表.
	_RecordsetPtr m_pRs_a_Hedu;		//外保温材料的黑度
	_RecordsetPtr IRecPipe;			//管道参数库
	_RecordsetPtr m_pRsHeat_alfa;	//计算放热系数公式取值索引表.
	_RecordsetPtr m_pRs_CodeMat;	//标准的保温材料表.

	_RecordsetPtr m_pRsCalcThick;	//计算保温厚度的公式.

	_RecordsetPtr IRecCalc;			//保温原始数据表。

	_RecordsetPtr m_pRsHeatLoss;	//保温结构外表面允许最大散热密度	

	_RecordsetPtr m_pRsCongeal;		//防冻结计算中新增加的一些参数
	
	_RecordsetPtr m_pRsSubterranean;// 埋地管道保温计算参数
	
	_RecordsetPtr m_pRsFaceResistance;	// 保温层外表面向周围空气的散热阻的记录集

	_RecordsetPtr m_pRsWeather;		//气象参数

	// 土壤的导热系数
	_RecordsetPtr m_pRsLSo;

protected:
	short GetPipeWeithAndWaterWeith(_RecordsetPtr &pRs, const double dPipe_Dw, const double dPipe_S, CString strField1, double &dValue1,CString strField2,  double &dValue2);
//	GetPropertyofMaterial CCalcDll;
	CCalcSupportSpanDll	  CalSupportSpan;

	CString m_ProjectName;	//工程名
	CString m_MaterialPath;
	_ConnectionPtr m_IMaterialCon;

	_ConnectionPtr m_AssistantConnection; //辅助数据库的连接
private:
	BOOL GetSubterraneanTypeAndPipeCount(const long m_nID, short& nType, short& nPipeCount) const;
	BOOL FindStandardMat(_RecordsetPtr& pRsCurMat,CString strMaterialName);
	void InitOriginalToMethod();

};

#endif // !defined(AFX_HEATPRECAL_H__DE887BF7_4B9B_40F3_92E5_FAD4E263A3DF__INCLUDED_)

BOOL InputMaterialParameter(CString strDlgCaption, CString strMaterialName, int TableID);
