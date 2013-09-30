// CalcOriginal.h: interface for the CCalcOriginalData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CALCORIGINAL_H__0EDA007A_D6E1_4AF5_A248_0CE0FA182794__INCLUDED_)
#define AFX_CALCORIGINAL_H__0EDA007A_D6E1_4AF5_A248_0CE0FA182794__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "GetPropertyofMaterial2.h"	// Added by ClassView


//计算放热系数时的相邻两个外表面温度的差值.
#define  TsDiff		0.1
#define  MaxCycCount 2000
//
#define	 pi			3.1415927
//如果允许最大温度为0时(即，没有温度限制), 返回一个常数.
#define	 D_MAXTEMP 	50000.0	

class CCalcOriginalData
{
public:
	CCalcOriginalData();
	virtual ~CCalcOriginalData();

	_RecordsetPtr m_pRsCon_Out;		//对应设备的外表面放热系数表.
	_RecordsetPtr m_pRs_a_Hedu;		//外保温材料的黑度
	_RecordsetPtr IRecPipe;			//管道参数库
	_RecordsetPtr m_pRsHeat_alfa;	//计算放热系数公式取值索引表.
	//  [2005/06/02] 
	_ConnectionPtr m_pConAutoIPED;	//连接项目数据库(AutoIPED.mdb)的指针.

	_RecordsetPtr m_pRsCalcThick;	//计算保温厚度的公式.

	_RecordsetPtr IRecCalc;			//保温原始数据表。

	_RecordsetPtr m_pRsHeatLoss;	// 允许最大散热密度

	_RecordsetPtr m_pRsCongeal;		//防冻结计算中新增加的一些参数
	
	_RecordsetPtr m_pRsSubterranean;// 埋地管道保温计算时的参数记录集
	
	_RecordsetPtr m_pRsFaceResistance;// 管道或平壁保温结构的外表面至周围空气的放热阻记录集
	
	// 土壤的导热系数表
	_RecordsetPtr m_pRsLSo;

	int m_thick1Start;		//内层保温允许最小厚度
	int m_thick1Max;		//内层保温允许最大厚度
	int m_thick1Increment;	//内层保温厚度最小增量

	int m_thick2Start;		//外层保温允许最小厚度
	int m_thick2Max;		//外层保温允许最大厚度
	int m_thick2Increment;	//外层保温厚度最小增量

	double	dQ;				//确定了保温厚度时，在此厚度下的散热密度值

protected:
	//{{//		计算保温厚度时所用到的变量		[2005/04/26]

	long		m_nID;			//原始数据当前记录的序号
	int			nMethodIndex;	//计算方法的索引 

	double		out_price;		//外保温材料的价格.
	double		in_price;		//内保温材料的价格.
	double		pro_price;		//保护材料的价格.

	double		out_dens;		//外保温材料的容重.
	double		in_dens;		//内保温材料的容重
	double		pro_dens;		//保护材料的容重

	double		co_alf; 		//根据外径在表中获得的传热系数
	double		Mhours;			//年运行小时数
	double		Yong;			//热价比主汽价
	double		heat_price;		//主汽热价
//	double		Temp_pip;		//介质温度
//	double		Temp_env;		//环境温度

	double		nYearVal;		//年运行工况散热密度
	double		nSeasonVal;		//季节运行工况散热密度.
	double		QMax;			//根据当前的工况，确定的散热密度。
	double		in_tmax;		//复合保温时的外保温材料允许的界面温度。

	int			hour_work;		//运行工况.0:季节运行工况.1:年运行工况.

	double		D0;	 			//设备或管道外径
	double		D1;				//保温层外径，复合保温内层外径
	double		D2;				//复合保温外层外径
	double		S;				//保温结构投资年费用率

	double		thick;			//单层保温层厚度
	double		thick1;			//复合保温内层厚度
	double		thick2;			//复合保温外层厚度
	double		thick3;			//保护层厚度
			
	double		lamda;			//单层保温层材料的导热系数
	double		lamda1;			//内保温层材料的导热系数.
	double		lamda2;			//外保温层材料的导热系数.
	double		ALPHA;			//计算中的传热系数

	double		t;				//设备或管道外表面温度（介质温度）
	double		ta;				//环境温度
	double		ts;				//保温结构外表面温度
	double		tb;				//复合保温内外界面处温度

	double		in_a0;			//内保温材料导热系数基数
	double		in_a1;			//内保温材料导热系数一次项系数
	double		in_a2;			//内保温材料导热系数二次项系数
	double		in_ratio;		//内保温材料的导热系数增加的比率

	double		out_a0;			//外保温材料导热系数基数
	double		out_a1;			//外保温材料导热系数一次项系数
	double		out_a2;			//外保温材料导热系数二次项系数
	double		out_ratio;		//外保温材料的导热系数增加的比率

	double		m_HotRatio;		//保温界面温度不应超过外层保温材料最高使用温度的比例.
	double		m_CoolRatio;	//保冷界面温度不应超过外层保冷材料最高使用温度的比例.
	
	double		m_MaxHeatDensityRatio; //保温结构外表面允许散热密度，按查表的允许最大散热密度的比值.
	double		m_dMaxD0;		//平壁与圆管的分界外径.  以前为2000
	double		MaxTs;			//外表面允许最大温度值
	double		speed_wind;		//风速.
	double		pi_thi;			//管道壁厚
	double		distan;			//管道支吊架间距
	
	int			nAlphaIndex;	//传热系数的取值索引.
	int			nTaIndex;		//环境温度的取值索引

	CString		out_mat;		//外保温材料名称
	CString		in_mat;			//内保温材料名称
	CString		pro_mat;		//保护材料名称
	CString		pro_name;		//保护材料名称,是去掉后面括号的名称.如"铝皮(0.75)"则该值为"铝皮"
	CString		strPlace;		//管道或设备的安装地点

	CString		strExceptionInfo;//计算时出现在错误信息。
 
	//}}
/*
	//防冻计算时新增加的变量
	double	Kr;					//管道通过支吊架处的热(或冷)损失的附加系数
	double	taofr;				//防冻结管道允许液体停留时间(小时)
	double	tfr;				//介质在管内冻结温度
	double	V;					//每米管长介质体积，m3/m;
	double	ro;					//介质密度
	double	C;					//介质比热
	double	Vp;					//每米管长管壁体积,m3/m;
	double	rop;				//管材密度
	double	Cp;					//管材比热
	double	Hfr;				//介质融解热
	double	dFlux;				//流量

*/
	//计算放热系数的变量
	double	hedu;				//黑度
	double	B;					//沿风速方向的平壁宽度，m
//	double	W;					//年平均风速
	
protected:
//	GetPropertyofMaterial CCalcDll;
//	CString m_ProjectName;	//工程名
//	CString m_MaterialPath;
//	_ConnectionPtr m_IMaterialCon;

//	_ConnectionPtr m_AssistantConnection; //辅助数据库的连接

};
#endif // !defined(AFX_CALCORIGINAL_H__0EDA007A_D6E1_4AF5_A248_0CE0FA182794__INCLUDED_)





