//
// ParseExpression.h: interface for the CParseExpression class.
//
// 对FOXPRO中的简单关系表达式进行解析
// 如"S>=46.AND.S<65"
// 
// 有两种方式对表达式进行解析：利用数据库连接、普通方式
// 利用数据库对表达式进行解析需首先调用SetConnectionForParse
// 设置数据库的连接
//
// Write by 黄伟卫
//
// Time:2004,7,25
//	

#if !defined(AFX_PARSEEXPRESSION_H__821D2B2D_B8D8_4177_912E_79157020E63A__INCLUDED_)
#define AFX_PARSEEXPRESSION_H__821D2B2D_B8D8_4177_912E_79157020E63A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <afxtempl.h>
class CParseExpression  
{
public:
	CParseExpression();
	virtual ~CParseExpression();

	//用于对表达式中的变量进行初始化
	struct tagExpressionVar
	{
		CString VarName;		//表达式中的变量名
		_variant_t VarValue;	//变量的值
	};

public:
	//对FOXPRO中的算术表达式进行解析
	double ParseUseConn_Evaluation(LPCTSTR pExpression,tagExpressionVar *pVarStruct,int VarStructNum,int *pState);

	//对FOXPRO中的逻辑表达式进行解析
	BOOL ParseUseConn_Bool(LPCTSTR pExpression,tagExpressionVar *pVarStruct,int VarStructNum,int* pState);

	//返回临时表的表名
	LPCTSTR GetTemporarilyTableName()const;

	//设置临时表的表名
	void SetTemporarilyTableName(LPCTSTR TableName);

	//返回用于解析的数据库连接
	_ConnectionPtr& GetConnectionForParse();

	//设置用于解析的数据库连接
	void SetConnectionForParse(_ConnectionPtr IConnect);

	//对FOXPRO中的简单关系表达式进行解析
	BOOL Parse(LPCTSTR pExpression,int *pStatue);

	//返回指定键值的指针
	BOOL GetMapValue(CString &Key,void*& pDouble);

	//设置指定键值的指针
	void SetMapValue(CString &Key,double *pDouble);

private:
	//在字符串中寻找非在浮点数中"."的位置
	int FindPoint(LPCTSTR szString, int nStart);

	//对长表达式中的小单元进行解析
	BOOL ElementParse(LPCTSTR pExpression,int *pStatue);

	//对输入的FOXPRO表达式进行初始化处理(用于有数据库连接方式)
	CString InitExpressionUseConnect(LPCTSTR pExpression);

	//建立用于表达式解析的临时表
	_Recordset* CreateTemporarilyTable(tagExpressionVar *pVarStruct,int VarStructNum);

	// 删除临时表
	void DropTemporarilyTable();

private:
//	CString m_ResultFieldName;
	CString m_TemporarilyTableName;			//用于数据库解析的临时表名
	CMapStringToPtr m_DoubleMap;			//键值与DOUBLE现变量指针对照表
	_ConnectionPtr m_ConnectionForParse;	//用于数据库解析的数据库连接
};

#endif // !defined(AFX_PARSEEXPRESSION_H__821D2B2D_B8D8_4177_912E_79157020E63A__INCLUDED_)
