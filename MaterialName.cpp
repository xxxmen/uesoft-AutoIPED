// MaterialName.cpp: implementation of the CMaterialName class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "MaterialName.h"
#include "vtot.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMaterialName::CMaterialName()
{

}

CMaterialName::~CMaterialName()
{

}

///////////////////////////////////////////////////////////////////
//
// 设置新，旧材料对照表的记录集连接
//
// IRecordset[in]	新，旧材料对照表的记录集连接
//
void CMaterialName::SetNameRefRecordset(_RecordsetPtr& IRecordset)
{
	m_IRecordsetNameRef=IRecordset;
}

///////////////////////////////////////////////////////////////////
//
// 设置新，旧材料对照表的记录集连接
//
// IRecordset[in]	新，旧材料对照表的所在数据库的连接
// szTableName[in]	新，旧材料对照表的表名
//
// throw(_com_error)
//
void CMaterialName::SetNameRefRecordset(_ConnectionPtr IConnection, LPCTSTR szTableName)
{
	HRESULT hr;

	if(IConnection==NULL || szTableName==NULL)
	{
		_com_error e(E_POINTER);

		ReportExceptionErrorLV2(e);

		throw(e);
	}

	if(m_IRecordsetNameRef!=NULL)
	{
		m_IRecordsetNameRef.Release();
	}
	else
	{
		hr=m_IRecordsetNameRef.CreateInstance(__uuidof(Recordset));

		if(FAILED(hr))
		{
			_com_error e(hr);
			ReportExceptionErrorLV2(e);
			throw e;			
		}
	}

	try
	{
		m_IRecordsetNameRef->CursorLocation=adUseClient;

		m_IRecordsetNameRef->Open(_variant_t(szTableName),_variant_t((IDispatch*)IConnection),
								  adOpenStatic,adLockOptimistic,adCmdTable);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}
}

///////////////////////////////////////////////////////////////////
//
// 返回新，旧材料对照表的记录集连接
//
_RecordsetPtr& CMaterialName::GetNameRefRecordset()
{
	return m_IRecordsetNameRef;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 用材料的新名称替换材料的旧名称
//
// IConnection[in]		需要替换的表的数据库连接
// ReplaceStruct[in]	tagReplaceStruct结构
//
// throw(_com_error)
//
// 在调用此函数之前应先SetNameRefRecordset
//
void CMaterialName::ReplaceNameOldToNew(_ConnectionPtr &IConnection, tagReplaceStruct &ReplaceStruct)
{
	_RecordsetPtr IRecordset;
	HRESULT hr;

	if(IConnection==NULL || m_IRecordsetNameRef==NULL)
	{
		_com_error e(E_INVALIDARG);

		ReportExceptionErrorLV2(e);

		throw(e);
	}

	hr=IRecordset.CreateInstance(__uuidof(Recordset));

	if(FAILED(hr))
	{
		_com_error e(hr);

		ReportExceptionErrorLV2(e);

		throw e;
	}

	try
	{
		IRecordset->CursorLocation=adUseClient;

		IRecordset->Open(_variant_t(ReplaceStruct.strTableName),_variant_t((IDispatch*)IConnection),
						 adOpenStatic,adLockOptimistic,adCmdTable);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}

	try
	{
		ReplaceNameOldToNew(IRecordset,ReplaceStruct);
	}
	catch(_com_error &e)
	{
		ReportExceptionErrorLV2(e);
		throw;
	}
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 用材料的新名称替换材料的旧名称
//
// IRecordset[in]		需要替换的表的连接
// ReplaceStruct[in]	tagReplaceStruct结构,此结构中的strTableName不用填写
//
// throw(_com_error)
//
// 在调用此函数之前应先SetNameRefRecordset
//
void CMaterialName::ReplaceNameOldToNew(_RecordsetPtr &IRecordset, tagReplaceStruct &ReplaceStruct)
{
	CString strField,strTemp,SQL;
	_variant_t varTemp;
	int Pos;

	if(IRecordset==NULL || m_IRecordsetNameRef==NULL)
	{
		_com_error e(E_INVALIDARG);

		ReportExceptionErrorLV2(e);

		throw(e);
	}

	if(IRecordset->adoEOF && IRecordset->BOF)
	{
		return;
	}

	Pos=0;

	while(TRUE)
	{
		Pos++;
		strField=GetFieldsNameFromString(ReplaceStruct.pstrFieldsName,Pos);

		if(strField.IsEmpty())
			break;

		try
		{
			IRecordset->MoveFirst();

			while(!IRecordset->adoEOF)
			{
				GetTbValue(IRecordset,_variant_t(strField),strTemp);

				if(strTemp.IsEmpty())
				{
					IRecordset->MoveNext();
					continue;
				}

				SQL.Format(_T("OLD='%s'"),strTemp);

				m_IRecordsetNameRef->Filter=_variant_t(SQL);

				if(m_IRecordsetNameRef->adoEOF && m_IRecordsetNameRef->BOF)
				{
					IRecordset->MoveNext();
					continue;
				}

				varTemp=m_IRecordsetNameRef->GetCollect(_variant_t("NEWCL"));
				if (!strTemp.Compare(vtos(varTemp)))
				{
					//替换的名称相同则不修改数据库
					IRecordset->MoveNext();
					continue;
				}

				IRecordset->PutCollect(_variant_t(strField),varTemp);

				IRecordset->MoveNext();

			}//while(!IRecordset->adoEOF)

			IRecordset->MoveFirst();
			IRecordset->Update();

		}//try
		catch(_com_error &e)
		{
			ReportExceptionErrorLV2(e);
			throw;
		}
	}
}

///////////////////////////////////////////////////////////////////////
//
// 从包含字段的字符串中获得字段名
//
// szFields[in]		包含字段的字符串,多个字段名用“|”隔开
// Pos[in]			要返回的第几个字段名,Pos从1开始
//
// 返回指定位置的字段名
//
// 如果Pos<1或者大于最大所包含的字段名，将返回空字符串，可用IsEmpty判断
//
CString CMaterialName::GetFieldsNameFromString(LPCTSTR szFields,int Pos)
{
	CString retFieldName;
	CString strFields;
	int PrevPos,CurPos;

	retFieldName=_T("");

	if(szFields==NULL || Pos<=0)
	{
		return retFieldName;
	}


	strFields=szFields;
	strFields.TrimLeft();
	strFields.TrimRight();

	if(strFields.IsEmpty())
		return retFieldName;

	PrevPos=0;
	CurPos=strFields.Find(_T("|"));

	if(CurPos!=-1)
	{
		while(TRUE)
		{
			Pos--;

			if(Pos==0)
			{
				if(CurPos!=-1)
					retFieldName=strFields.Mid(PrevPos,CurPos-PrevPos);
				else
					retFieldName=strFields.Mid(PrevPos);

				break;
			}

			if(CurPos==-1)
				break;

			PrevPos=CurPos+1;

			CurPos=strFields.Find(_T("|"),CurPos+1);
		}
	}
	else
	{
		if(Pos==1)
		{
			retFieldName=strFields;
		}
		
	}

	return retFieldName;

}


