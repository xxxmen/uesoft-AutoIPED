//
// FoxBase.h: interface for the CFoxBase class.
//
// 提供在FOXPRO中的一些基本的函数
//
// Write by 黄伟卫
// Time:	2004,7,29
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_FOXBASE_H__6C072190_9140_4AC7_8100_09DA2694738E__INCLUDED_)
#define AFX_FOXBASE_H__6C072190_9140_4AC7_8100_09DA2694738E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#import"c:\program files\common files\system\ado\msado15.dll"	\
		no_namespace rename("EOF","adoEOF")
#include "ParseExpression.h"

typedef struct _Txtup{
	int dID;
	CString txt;
} Txtup;

class CFoxBase:public CParseExpression  
{
public:
	CFoxBase();
	virtual ~CFoxBase();

public:
	enum{
		EQUAL,				//等于
		UNEQUAL,			//不等于
		GREATER,			//大于
		GREATER_OR_EQUAL,	//大于等于
		LESS,				//小于
		LESS_OR_EQUAL		//小于等于
	};

public:
	//在记录集中当前游标位置设置指定的字段的值
	void PutTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,double Value);
	void PutTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,float Value);
	void PutTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,int Value);
	void PutTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,CString Value);

	//从记录集中返回指定字段的值
	void GetTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,double &RetValue);
	void GetTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,float &RetValue);
	void GetTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,int &RetValue);
	void GetTbValue(_RecordsetPtr IRecordset,_variant_t FieldsName,CString &RetValue);

	//在记录集的当前位置插入一条新记录
	BOOL InsertNew(_RecordsetPtr &IRecord,int after);
	//如果after==1，则是insert blank
	//如果after为其他值就是insert blank befo



	//替换记录集中选中字段指定条记录的值(相当于FOXPRO中 Replace Next)
	BOOL ReplNext(_RecordsetPtr IRecordset, _variant_t FieldsName, _variant_t Value,int num);

	//判断文件是否存在
	BOOL FILE(LPCTSTR pFilePath);

	//替换记录集中选中字段所有的值
	BOOL ReplAll(_RecordsetPtr IRecordset,_variant_t FieldsName,_variant_t Value);

	//返回当前记录集的游标位置
	long RecNo(_RecordsetPtr IRecordset);

	//
	//从表的第一条记录开始寻找指定的逻辑关系
	//
	BOOL LocateFor(_RecordsetPtr IRecordset,_variant_t FieldsName,int Relations,_variant_t Value);

	//
	//从表的当前记录开始寻找指定的逻辑关系
	//
	BOOL LocateForCurrent(_RecordsetPtr IRecordset, _variant_t FieldsName, int Relations,_variant_t Value);


	//返回数据库的连接
	_ConnectionPtr& GetConnect();

	//设置数据库的连接
	void SetConnect(_ConnectionPtr &IConnection);

	BOOL CopyTo(LPCTSTR pSourceTable,LPCTSTR pDestinationTable);

	//替换指定范围的某字段的值.
	bool ReplaceAreaValue(_RecordsetPtr pRs, CString strField,CString strValue,int startRow=0, int endRow=0);

protected:
	//当有异常时调用此函数
	virtual void ExceptionInfo(LPCTSTR pErrorInfo);
	virtual void ExceptionInfo(_com_error &e);

	//当需显视信息时调用此函数
	virtual void DisplayRes(LPCTSTR pStr);

	//当需输入数值时调用此函数
	virtual void InputD(LPCTSTR pstr,double &val);

	//当等待输入字符串时调用此函数
	virtual void Wait(LPCTSTR pInfoStr,CString &strRet);

	//当需显视信息时调用此函数
	virtual void Say(LPCTSTR pStr);

	//当需退出应用程序时调用此函数
	virtual void Quit();

	//当需以对话框形式显示信息时调用此函数
	virtual int MessageBox(LPCTSTR pStr);

	BOOL CopyData(_RecordsetPtr &IRecordS,_RecordsetPtr &IRecordD);
	CString CreateTableSQL(_RecordsetPtr &IRecord,LPCTSTR pTableName);
private:
	HANDLE m_hStdout;				//标准输出句柄
	_ConnectionPtr m_Connection;	//数据库链接智能指针
};

#endif // !defined(AFX_FOXBASE_H__6C072190_9140_4AC7_8100_09DA2694738E__INCLUDED_)
