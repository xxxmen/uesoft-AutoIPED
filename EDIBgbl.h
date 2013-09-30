// EDIBgbl.h: interface for the EDIBgbl class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_EDIBGBL_H__264E7A48_66B4_440E_B867_3490AE90662E__INCLUDED_)
#define AFX_EDIBGBL_H__264E7A48_66B4_440E_B867_3490AE90662E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

enum IPEDCodeCallingKey{
   //保温设计规定标准关键字
   nKey_CODE_DL5072_1997 = 1,
   nKey_CODE_SH3010_2000, //2,
   nKey_CODE_GB8175_2008 , //3
   nKey_CODE_CJJ34_2002 ,//4
   nKey_CODE_DL5072_2007 ,//5
   nKey_CODE_GB4272_2008,  //6
   nKey_CODE_GB50264_1997,  //7
   nKey_CODE_GB15586_1995 ,//8
   nKey_CODE_GB4272_1992 ,//9
   nKey_CODE_GB8175_1987, //10
   nKey_CODE_SH3126_2001  //11
};

class EDIBgbl  
{
public:
	typedef struct CAPTION2FIELD		//EXCEL表中的标题对应的ACCESS库中的字段名
	{
		CString strCaption;			// 中文件标题，（EXCEL中的字段名）
		CString	strExcelNO;			// 对应Excel的列号
		CString strField;			// ACCESS中的字段名称
		CString strRemark;			// 字段的备注信息
		int		nFieldType;			// 字段的类型,0为字符型,1为数字型
	};

public:	// Method
	
	// 新建一个临时表，用于计算放热系数的值。
	static int NewCalcAlfaTable();

	//在数据库pCon中,对指定的表进行操作.将默认的记录插入到指定的工程中
	static bool InsertDefTOEng(_ConnectionPtr& pCon,CString strTblName,CString strPrjID,CString strDefKey="");

	// 将重新设定的数据库路径保存到文件中。
	static bool GetCurDBName();
	
	// 从文件中读取数据库的路径
	static bool SetCurDBName();

	// 替换数据库中一些字段的值
	static BOOL ReplaceFieldValue(CString strEnginID=_T(""));
	
	//获得共享数据库目录
	static CString GetShareDbPath();

public:		// Property
	
	// 当前选择的卷册ID
	static long SelVlmID;
	
	// 当前选择的行业ID
	static long SelHyID;	
	
	// 设计阶段ID
	static long SelDsgnID;
	
	// 工程代号
	static CString SelPrjID;

	//工程名称
	static CString SelPrjName;
	
	//卷册代号
    static CString SelVlmCODE;

	// Kgf到N的转换系数( = 9.80665 )
	const  static  double kgf2N;

	// 规范标准(行业)名称 如：电力行业
	static CString sCur_CodeName;

	// 规范标准代号 如：DL/T5054-1996
	static CString sCur_CodeNO;	

	// 规范标准代号数字标志，整数，节省空间，每个保温数据记录都应该设置这个字段，
	//这样每根管道都可以按不同的保温标准计算,如：1
	static long iCur_CodeKey;	

	// 专业ID
	static long SelSpecID;
	
// PATH
	//材料数据库所在的路径
	static CString sMaterialPath;

	//当前材料数据库的名称。
	static CString sCur_MaterialCodeName;

	//标准数据库所在的路径
	static CString sCritPath;

	//当前标准数据库的名称
	static CString sCur_CritDbName;
	
	//介质数据库的路径
//	static CString strCur_MediumDBPath;

	static _RecordsetPtr pRsCalc;

public:
	EDIBgbl();
	virtual ~EDIBgbl();

protected:
};

#endif // !defined(AFX_EDIBGBL_H__264E7A48_66B4_440E_B867_3490AE90662E__INCLUDED_)
