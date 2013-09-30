//
// MaterialName.h: interface for the CMaterialName class.
//
// 材料名称管理类
//
// 实现材料新，旧名称的替换
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_MATERIALNAME_H__3A9723FC_4F29_4C56_8144_D5C7B2018BF5__INCLUDED_)
#define AFX_MATERIALNAME_H__3A9723FC_4F29_4C56_8144_D5C7B2018BF5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "FoxBase.h"

class CMaterialName:public CFoxBase 
{
public:

	CMaterialName();
	virtual ~CMaterialName();

public:
	struct tagReplaceStruct			// 替换信息
	{
		CString strTableName;		// 需要替换的表名
		CString pstrFieldsName;		// 需要替换的字段名，多个字段用"|"隔开
	};

public:
	// 设置新，旧材料对照表的记录集连接
	void SetNameRefRecordset(_RecordsetPtr& IRecordset);
	void SetNameRefRecordset(_ConnectionPtr IConnection,LPCTSTR szTableName);	//throw(_com_error)
	// 返回新，旧材料对照表的记录集连接
	_RecordsetPtr& GetNameRefRecordset();


	// 用材料的新名称替换材料的旧名称 throw(_com_error);
	void ReplaceNameOldToNew(_ConnectionPtr &IConnection,tagReplaceStruct &ReplaceStruct);
	void ReplaceNameOldToNew(_RecordsetPtr &IRecordset, tagReplaceStruct &ReplaceStruct);


private:
	// 从包含字段的字符串中获得字段名
	CString GetFieldsNameFromString(LPCTSTR szFields,int Pos);

private:
	_RecordsetPtr m_IRecordsetNameRef;	//新，旧材料对照表的记录集连接
};

#endif // !defined(AFX_MATERIALNAME_H__3A9723FC_4F29_4C56_8144_D5C7B2018BF5__INCLUDED_)
