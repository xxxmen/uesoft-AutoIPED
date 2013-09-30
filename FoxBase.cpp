// FoxBase.cpp: implementation of the CFoxBase class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "FoxBase.h"
#include "vtot.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CFoxBase::CFoxBase()
{
	m_hStdout=NULL;
}

CFoxBase::~CFoxBase()
{

}

////////////////////////////////////////////////////
//
//从表的当前记录开始寻找指定的逻辑关系
//
//IRecordset[in] 记录集智能指针
//FieldsName[in] 字段名
//Relations[in]  逻辑关系，等于、不等于……
//Value[in]		 输入与FieldsName比较的值
//
//函数成功返回TRUE,否则返回FALSE
//如果需比较的字段与关系不匹配会抛出_com_error异常
//指针指向找到的记录上
//
BOOL CFoxBase::LocateForCurrent(_RecordsetPtr IRecordset, _variant_t FieldsName, int Relations, _variant_t Value)
{
	_variant_t tempvalue;
	CString tempstr;

	if(IRecordset==NULL || Value.vt==VT_NULL)
		return FALSE;
	

	while(!IRecordset->adoEOF)
	{
		tempvalue=IRecordset->GetCollect(FieldsName);

		if(tempvalue.vt==VT_NULL)
		{
			IRecordset->MoveNext();
			continue;
		}

		switch(Relations)
		{
		case CFoxBase::EQUAL:			// 等于
			if(Value.vt==VT_BSTR)
			{
			//	tempstr=(TCHAR*)(_bstr_t)tempvalue;
			//	if(tempstr.Find((TCHAR*)(_bstr_t)Value) != -1)
			//		return TRUE;
			//	else
			//		break;
				tempstr = vtos(tempvalue);
				if (!tempstr.Compare(vtos(Value)))
				{
					return TRUE;
				}
				else
				{
					break;
				}

			}

			if(tempvalue==Value)
				return TRUE;
			else
				break;

		case CFoxBase::UNEQUAL:				//不等于
			if(tempvalue!=Value)
				return TRUE;
			else
				break;

		case CFoxBase::GREATER:				//大于
			if((double)tempvalue>(double)Value)
				return TRUE;
			else
				break;

		case CFoxBase::GREATER_OR_EQUAL:	//大于等于
			if((double)tempvalue>=(double)Value)
				return TRUE;
			else
				break;

		case CFoxBase::LESS:				//小于
			if((double)tempvalue<(double)Value)
				return TRUE;
			else
				break;

		case CFoxBase::LESS_OR_EQUAL:		//小于等于
			if((double)tempvalue<=(double)Value)
				return TRUE;
			else
				break;

		default:
			return FALSE;

		}

		IRecordset->MoveNext();
	}

	if(IRecordset->adoEOF)
		return FALSE;

	return TRUE;
}

////////////////////////////////////////////////////
//
//从表的第一条记录开始寻找指定的逻辑关系
//
//IRecordset[in] 记录集智能指针
//FieldsName[in] 字段名
//Relations[in]  逻辑关系，等于、不等于……
//Value[in]		 输入与FieldsName比较的值
//
//函数成功返回TRUE,否则返回FALSE
//如果需比较的字段与关系不匹配会抛出_com_error异常
//指针指向找到的记录上
//
BOOL CFoxBase::LocateFor(_RecordsetPtr IRecordset, _variant_t FieldsName, 
						 int Relations, _variant_t Value)
{
	_variant_t tempvalue;
	CString tempstr;

	if(IRecordset==NULL || Value.vt==VT_NULL)
		return FALSE;
	
	try
	{
		IRecordset->MoveFirst();
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}

	return LocateForCurrent(IRecordset,FieldsName,Relations,Value);

}

///////////////////////////////////////////////
//
//返回当前记录集的游标位置
//
//IRecordset[in] 记录集智能指针
//
//函数成功返回当前记录集的游标位置，否则返回-1
//
long CFoxBase::RecNo(_RecordsetPtr IRecordset)
{
	long pos=0;
	long i;

	if(IRecordset==NULL)
		return -1;

	if(IRecordset->adoEOF && IRecordset->BOF)
		return -1;

	while(!IRecordset->GetBOF())
	{
		pos++;
		IRecordset->MovePrevious();
	}

	i=pos;

	i--;
	IRecordset->MoveFirst();

	while(i>0)
	{
		i--;
		IRecordset->MoveNext();
	}

	return pos;
}


///////////////////////////////////////////////////////////
//
//替换记录集中选中字段所有的值
//
//IRecordset[in] 记录集智能指针
//FieldsName[in] 字段名
//Value[in]		 输入替换FieldsName字段的值
//
//函数成功返回TRUE，否则返回FALSE
//当函数成功返回时，当前游标指针指向记录集最后一个记录的后面
//
BOOL CFoxBase::ReplAll(_RecordsetPtr IRecordset, _variant_t FieldsName, _variant_t Value)
{
	if(IRecordset==NULL)
		return FALSE;

	IRecordset->MoveFirst();

	while(!IRecordset->adoEOF)
	{
		IRecordset->PutCollect(FieldsName,Value);
		IRecordset->Update();
		IRecordset->MoveNext();
	}
	return TRUE;
}

//////////////////////////////////////////
//
//判断文件是否存在
//
//pFilePath[in]文件的绝对路径
//
//如果文件存在则返回TRUE,否则返回FALSE
//
BOOL CFoxBase::FILE(LPCTSTR pFilePath)
{
	HANDLE hFile;
	if(pFilePath==NULL)
		return FALSE;

	hFile=CreateFile(pFilePath,
					 GENERIC_READ,
					 FILE_SHARE_READ,
					 NULL,
					 OPEN_EXISTING,
					 FILE_ATTRIBUTE_NORMAL,
					 NULL);
	if(hFile==INVALID_HANDLE_VALUE)
	{
		//文件不存在  
		if(GetLastError()==ERROR_FILE_NOT_FOUND)
			return FALSE;
	}

	CloseHandle(hFile);
	return TRUE;
}

///////////////////////////////////////////////////////////
//
//替换从记录集中选中字段指定条记录的值(相当于FOXPRO中 Replace Next)
//
//IRecordset[in]	记录集智能指针
//FieldsName[in]	字段名
//Value[in]			输入替换FieldsName字段的值
//num[in]			需替换的记录数
//
//函数成功返回TRUE，否则返回FALSE
//函数从当前游标位置开始替换
//当函数成功返回时，当前游标指针指向最后修改的记录集上
//当需替换的记录数超过最后一条记录时,游标指向最后一条记录的后面
//
BOOL CFoxBase::ReplNext(_RecordsetPtr IRecordset, _variant_t FieldsName, _variant_t Value, int num)
{
	if(IRecordset==NULL || num<=0)
		return FALSE;


	while(!IRecordset->adoEOF)
	{
		IRecordset->PutCollect(FieldsName,Value);
		IRecordset->Update();

		num--;
		if(num<=0)
			break;
		
		IRecordset->MoveNext();
	}
	return TRUE;
}

int CFoxBase::MessageBox(LPCTSTR pStr)
{
	return ::MessageBox(NULL,"",pStr,MB_OK);
}

////////////////////////////////////////////
//
//退出应用程序调用
//
void CFoxBase::Quit()
{

}

///////////////////////////////////////////////
//
//	当需显视信息时调用此函数
//
//	pStr[in]	用于显示的信息字符串
//
void CFoxBase::Say(LPCTSTR pStr)
{
	::MessageBox(NULL,pStr,"",MB_OK);
}

////////////////////////////////////////////////////////
//
// 当等待输入字符串时调用此函数
//
// pInfoStr[in]		用于显示的信息字符串
// strRet[out]		需返回的字符串
//
void CFoxBase::Wait(LPCTSTR pInfoStr,CString &strRet)
{
	::MessageBox(NULL,"",pInfoStr,MB_OK);
	strRet=_T("F");
}

////////////////////////////////////////////////////////
//
// 当需输入数值时调用此函数
// 
// pstr[in]			用于显示的信息字符串
// val[out]			需返回的变量
//
void CFoxBase::InputD(LPCTSTR pstr, double &val)
{
	::MessageBox(NULL,pstr,"",MB_OK);
}

///////////////////////////////////////////////////////
//
// 当需显视信息时调用此函数
//
// pStr[in]			用于显示的信息字符串
//
void CFoxBase::DisplayRes(LPCTSTR pStr)
{
//	AfxMessageBox(pStr);
	DWORD cWritten;
    
	if(AllocConsole())
	{
		m_hStdout=GetStdHandle(STD_OUTPUT_HANDLE);
	}
	
	m_hStdout=GetStdHandle(STD_OUTPUT_HANDLE);
	WriteFile(m_hStdout,pStr,lstrlen(pStr),&cWritten,NULL);
	WriteFile(m_hStdout,"\n",lstrlen("\n"),&cWritten,NULL);	
}

//////////////////////////////////////////////////
//
//当有异常时调用此函数
//
//pErrorInfo[in]异常信息
//
void CFoxBase::ExceptionInfo(LPCTSTR pErrorInfo)
{
	AfxMessageBox(pErrorInfo);
}

void CFoxBase::ExceptionInfo(_com_error &e)
{
	if(e.Description().length()>0)
		ExceptionInfo(e.Description());
	else
		ExceptionInfo(e.ErrorMessage());
}
//////////////////////////////////////////////////////////////////////////////////////////////
//
// 从记录集中返回指定字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// RetValue[out]	返回的变量
//
// 如果有异常将调用ExceptionInfo函数
//
void CFoxBase::GetTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, double &RetValue)
{
	_variant_t TempValue;
	CString strTemp;

	if(IRecordset==NULL)
		return;
	
	try
	{
		TempValue=IRecordset->GetCollect(FieldsName);
		
		if(TempValue.vt!=VT_NULL)
		{
			if(TempValue.vt==VT_BSTR)
			{
				CString Temp;
				Temp=(TCHAR*)(_bstr_t)TempValue;
				Temp.TrimLeft();

				if(Temp.GetLength()==0)
				{
					RetValue=0.0;
					//break;
					return;
				}

			}
			RetValue=TempValue;
		}
		else
		{
			RetValue=0.0;
		}
	}
	catch(_com_error &e)
	{
		RetValue=0.0;
//		ExceptionInfo(e);

		strTemp.Format(_T("错误在从%s字段读取数据时"),(LPCTSTR)(_bstr_t)FieldsName);
		Exception::SetAdditiveInfo(strTemp);
		ReportExceptionErrorLV2(e);

		throw;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////
//
// 从记录集中返回指定字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// RetValue[out]	返回的变量
//
// 如果有异常将调用ExceptionInfo函数
//
void CFoxBase::GetTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, float &RetValue)
{
	_variant_t TempValue;
	CString strTemp;

	if(IRecordset==NULL)
		return;
	
	try
	{
		TempValue=IRecordset->GetCollect(FieldsName);
		
		if(TempValue.vt!=VT_NULL)
		{
			if(TempValue.vt==VT_BSTR)
			{
				CString Temp;
				Temp=(TCHAR*)(_bstr_t)TempValue;
				Temp.TrimLeft();

				if(Temp.GetLength()==0)
				{
					RetValue=0.0;
					//break;
					return;
				}

			}
			RetValue=TempValue;
		}
		else
		{
			RetValue=0.0;
		}
	}
	catch(_com_error &e)
	{
		RetValue=0.0;
//		ExceptionInfo(e);

		strTemp.Format(_T("错误在从%s字段读取数据时"),(LPCTSTR)(_bstr_t)FieldsName);
		Exception::SetAdditiveInfo(strTemp);

		ReportExceptionErrorLV2(e);
		throw;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////
//
// 从记录集中返回指定字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// RetValue[out]	返回的变量
//
// 如果有异常将调用ExceptionInfo函数
//
void CFoxBase::GetTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, int &RetValue)
{
	_variant_t TempValue;
	CString strTemp;

	if(IRecordset==NULL)
		return;
	
	try
	{
		TempValue=IRecordset->GetCollect(FieldsName);
		
		if(TempValue.vt!=VT_NULL)
		{
			if(TempValue.vt==VT_BSTR)
			{
				CString Temp;
				Temp=(TCHAR*)(_bstr_t)TempValue;
				Temp.TrimLeft();

				if(Temp.GetLength()==0)
				{
					RetValue=0;
					//break;
					return;
				}

			}
			RetValue=(long)TempValue;
		}
		else
		{
			RetValue=0;
		}
	}
	catch(_com_error &e)
	{
		RetValue=0;
//		ExceptionInfo(e);

		strTemp.Format(_T("错误在从%s字段读取数据时"),(LPCTSTR)(_bstr_t)FieldsName);
		Exception::SetAdditiveInfo(strTemp);

		ReportExceptionErrorLV2(e);
		throw;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////
//
// 从记录集中返回指定字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// RetValue[out]	返回的变量
//
// 如果有异常将调用ExceptionInfo函数
//
void CFoxBase::GetTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, CString &RetValue)
{
	_variant_t TempValue;
	CString strTemp;

	if(IRecordset==NULL)
		return;
	
	try
	{
		TempValue=IRecordset->GetCollect(FieldsName);
		
		if(TempValue.vt!=VT_NULL)
		{
			RetValue=(TCHAR*)(_bstr_t)TempValue;

			RetValue.TrimLeft();
			RetValue.TrimRight();
		}
		else
		{
			RetValue=_T("");
		}
	}
	catch(_com_error &e)
	{
		RetValue=_T("");
//		ExceptionInfo(e);

		strTemp.Format(_T("错误在从%s字段读取数据时"),(LPCTSTR)(_bstr_t)FieldsName);
		Exception::SetAdditiveInfo(strTemp);

		ReportExceptionErrorLV2(e);
		throw;
	}
}

/////////////////////////////////////////////////////////////////////////////////////////
//
// 在记录集中当前游标位置设置指定的字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// Value[in]		新的值
//
// 如果有异常将调用ExceptionInfo函数
//
void CFoxBase::PutTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, double Value)
{
	if(IRecordset==NULL)
		return;

	try
	{
		IRecordset->PutCollect(FieldsName,_variant_t(Value));
		IRecordset->Update();
	}
	catch(_com_error /*&e*/)
	{
//		ExceptionInfo(e);
		throw;
	}
}

/////////////////////////////////////////////////////////////////////////////////////////
//
// 在记录集中当前游标位置设置指定的字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// Value[in]		新的值
//
// 如果有异常将调用ExceptionInfo函数
//
void CFoxBase::PutTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, float Value)
{
	PutTbValue(IRecordset,FieldsName,(double)Value);
}

/////////////////////////////////////////////////////////////////////////////////////////
//
// 在记录集中当前游标位置设置指定的字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// Value[in]		新的值
//
// 如果有异常将调用ExceptionInfo函数
//
void CFoxBase::PutTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, int Value)
{
	PutTbValue(IRecordset,FieldsName,(double)Value);
}

/////////////////////////////////////////////////////////////////////////////////////////
//
// 在记录集中当前游标位置设置指定的字段的值
//
// IRecordset[in]	记录集智能指针	
// FieldsName[in]	字段名
// Value[in]		新的值
//
// 如果有异常将调用ExceptionInfo函数
//
// 如果Value为空或全为空格字符，当指定的字段的类型为字符串型时，字段的值为空
// 如果字段的类型为数字型，字段的值为0
//
void CFoxBase::PutTbValue(_RecordsetPtr IRecordset, _variant_t FieldsName, CString Value)
{
	adoDataTypeEnum DataType;

	if(IRecordset==NULL)
		return;

	Value.TrimLeft();
	Value.TrimRight();

	try
	{
		if(Value.IsEmpty())
		{
			IRecordset->GetFields()->GetItem(FieldsName)->get_Type(&DataType);
			
			if(DataType==adBSTR || DataType==adChar || DataType==adVarChar || 
			   DataType==adVarWChar ||DataType==adWChar)
			{
				Value=_T("");
			}
			else
			{
				Value=_T("0");
			}
		}

		IRecordset->PutCollect(FieldsName,_variant_t(Value));
		IRecordset->Update();
	}
	catch(_com_error /*&e*/)
	{
//		ExceptionInfo(e);
		throw;
	}

}

//////////////////////////////////////////////////
//
// 设置数据库的连接
//
// IConnection[in]	数据库的连接的智能指针的引用
//
void CFoxBase::SetConnect(_ConnectionPtr &IConnection)
{
	m_Connection=IConnection;
}

//////////////////////////////////////////////////
//
// 返回数据库的连接
//
// 函数返回数据库的连接的智能指针的引用
//
_ConnectionPtr& CFoxBase::GetConnect()
{
	return m_Connection;
}

CString CFoxBase::CreateTableSQL(_RecordsetPtr &IRecord,LPCTSTR pTableName)
{
	short Item;
	CString CreateSQL;
	CString DefList;
	FieldsPtr IFields;
	FieldPtr IField;

	CreateSQL=_T("");

	if(IRecord==NULL)
	{
		ExceptionInfo(_T("Recordset interface cann't be NULL"));
		return CreateSQL;
	}

	if(pTableName==NULL)
	{
		ExceptionInfo(_T("Table name cann't be NULL"));
		return CreateSQL;
	}

	IRecord->get_Fields(&IFields);

	DefList=_T("");
	for(Item=0;Item<IFields->GetCount();Item++)
	{
		IFields->get_Item(_variant_t(Item),&IField);
		DefList+=IField->GetName();
		DefList+=_T(" ");

		switch(IField->GetType())
		{
		case adVarWChar:
			{
				CString Temp;
				Temp.Format(_T("varchar(%d)"),IField->GetDefinedSize());
				DefList+=Temp;
				break;
			}

		case adLongVarWChar:
			DefList+=_T("text");
			break;

		case adVarBinary:
			{
				CString Temp;
				Temp.Format(_T("varbinary(%d)"),IField->GetDefinedSize());
				DefList+=Temp;
				break;
			}

//		case adNumeric:
//		case adGUID:
//		case adLongVarBinary:
//			{
//				ExceptionInfo(_T("不支持LongVarBinary"));
//				return CreateSQL;
//			}
//			break;

		case adInteger:
			DefList+=_T("int");
			break;

		case adUnsignedTinyInt:
		case adSmallInt:
			DefList+=_T("smallint");
			break;

		case adSingle:
		case adDouble:
			DefList+=_T("float");
			break;

		case adDBTimeStamp:
		case adDate:
			DefList+=_T("date");
			break;

//		case adBoolean:
//			break;

		default:
			{
				ExceptionInfo(_T("不支持此类型"));
				return CreateSQL;
			}

		}

		if(Item < IFields->GetCount()-1)
		{
			DefList+=_T(",");
		}

		IField.Release();
	}

	CreateSQL.Format(_T("CREATE TABLE %s(%s)"),pTableName,DefList);
	return CreateSQL;
}

BOOL CFoxBase::CopyData(_RecordsetPtr &IRecordS, _RecordsetPtr &IRecordD)
{
	_variant_t TempValue;
	short Item;
	FieldsPtr IFields;
	FieldPtr IField;

	if(IRecordS==NULL || IRecordD==NULL)
	{
		ExceptionInfo(_T("Source Recordset or destination Recordset cann't be empty"));
		return FALSE;
	}

	if(IRecordS->adoEOF && IRecordS->BOF)
	{
		return TRUE;
	}

	if(!IRecordD->adoEOF || !IRecordD->BOF)
	{
		try
		{
			IRecordD->MoveLast();
		}
		catch(_com_error &e)
		{
			ExceptionInfo(e);
			return FALSE;
		}
	}

	try
	{
		IRecordS->MoveFirst();
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}

	try
	{
		while(!IRecordS->adoEOF)
		{
			IRecordS->get_Fields(&IFields);
			IRecordD->AddNew();
			for(Item=0;Item<IFields->GetCount();Item++)
			{
				IFields->get_Item(_variant_t(Item),&IField);

				TempValue=IRecordS->GetCollect(_variant_t(IField->GetName()));
				IRecordD->PutCollect(_variant_t(IField->GetName()),TempValue);

				IField.Release();
			}
			IRecordD->Update();
			IFields.Release();
			IRecordS->MoveNext();
		}
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}


	return TRUE;
}

BOOL CFoxBase::CopyTo(LPCTSTR pSourceTable, LPCTSTR pDestinationTable)
{
	CString SourceTable,DestinationTable;
	CString SQL;
	_RecordsetPtr IRecordS,IRecordD;
	FieldsPtr IFields;
	FieldPtr IField;

	if(pSourceTable==NULL || pDestinationTable==NULL)
	{
		ExceptionInfo(_T("Source table or destination table cann't be empty"));
		return FALSE;
	}

	SourceTable=pSourceTable;
	DestinationTable=pDestinationTable;

	SourceTable.TrimLeft();
	SourceTable.TrimRight();
	DestinationTable.TrimLeft();
	DestinationTable.TrimRight();

	if(SourceTable.IsEmpty() || DestinationTable.IsEmpty())
	{
		ExceptionInfo(_T("Source table or destination table cann't be empty"));
		return FALSE;
	}

	try
	{
		IRecordS.CreateInstance(__uuidof(Recordset));

		IRecordS->Open(_variant_t(pSourceTable),_variant_t((IDispatch*)GetConnect()),	
					  adOpenKeyset,adLockOptimistic,adCmdTable);
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}
	
	SQL=CreateTableSQL(IRecordS,pDestinationTable);

	if(SQL.IsEmpty())
	{
		return FALSE;
	}

	try
	{
		GetConnect()->Execute(_bstr_t(SQL),NULL,adExecuteNoRecords);
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}

	try
	{
		IRecordD.CreateInstance(__uuidof(Recordset));

		IRecordD->Open(_variant_t(pDestinationTable),_variant_t((IDispatch*)GetConnect()),	
					  adOpenKeyset,adLockOptimistic,adCmdTable);
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}

	if(!CopyData(IRecordS,IRecordD))
	{
		return FALSE;
	}

	return TRUE;
}

/////////////////////////////////////////////////////
//
// 在记录集的当前位置插入一条新记录
//
// IRecord[in]	记录集智能指针
//
// 函数成功返回TRUE，否则返回FALSE
// 如果有异常将调用ExceptionInfo函数
//
// 函数成功后记录集将指向新插入的空记录
//
BOOL CFoxBase::InsertNew(_RecordsetPtr &IRecord,int after)
{
	int pos;
	CMap<short,short&,_variant_t,_variant_t&> FieldMap;
	CMap<short,short&,_variant_t,_variant_t&> newFieldMap;		//将新增加的记录的值记下来
	_variant_t TempValue;
	short Item;
	FieldsPtr IFields;
	FieldPtr IField;
    int pos1;
	int	nCount;		//字段个数.

	if(IRecord==NULL)
	{
		ExceptionInfo(_T("Recordset cann't be empty"));
		return FALSE;
	}

	//如果记录集为空将插入一条新记录
	if(IRecord->adoEOF && IRecord->BOF)
	{
		try
		{
			IRecord->AddNew();
			IRecord->Update();
		}
		catch(_com_error &e)
		{
			ExceptionInfo(e);
			return FALSE;
		}
		return TRUE;
	}
	else if(IRecord->adoEOF)
	{
		IRecord->AddNew();
		IRecord->Update();
		IRecord->MoveLast();
		return true;
	}

	for(pos=0; !IRecord->adoEOF ;pos++)
	{
		try
		{
			IRecord->MoveNext();
		}
		catch(_com_error &e)
		{
			ExceptionInfo(e);
			return FALSE;
		}
	}

	//
	// 查如一条新记录，并将刚开始所指的以后的记录向后移
	//
	try
	{
		IRecord->AddNew();
		IRecord->Update();
		IRecord->MoveLast();

		//将新增加的记录的值记下来.
		IFields = IRecord->GetFields();
		nCount = IFields->GetCount();   //字段个数

		for (Item = 0; Item < nCount; Item++)
		{
			newFieldMap[Item] = IRecord->GetCollect(_variant_t(Item));
		}
		IFields.Release();
		//		

		IRecord->MovePrevious();
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}
	if(after==1)
    pos--;  //使pos再减一，变成在指针的当前位置上加入一条新记录，insert blank
	pos1=pos;
	while(pos>0)
	{
		try
		{
			for(Item=0;Item<nCount;Item++)
			{
				TempValue=IRecord->GetCollect(_variant_t(Item));
				FieldMap[Item]=TempValue;
			}

			IRecord->MoveNext();

			for(Item=0;Item<nCount;Item++)
			{
				IRecord->PutCollect(_variant_t(Item),FieldMap[Item]);
			}
			IRecord->Update();

			IRecord->MovePrevious();

			pos--;
			if(pos>0)
				IRecord->MovePrevious();
		}
		catch(_com_error &e)
		{
			ExceptionInfo(e);
			return FALSE;
		}
	}

	// 使最开始所指的记录内容为空
	if(pos1==0 && after==1) {IRecord->MoveNext();}
	else
	{

	try
	{
		TempValue.Clear();

		for(Item=0;Item<nCount;Item++)
		{
			IRecord->PutCollect(_variant_t(Item),newFieldMap[Item]);
		}
		IRecord->Update();
	}
	catch(_com_error &e)
	{
		ExceptionInfo(e);
		return FALSE;
	}
	}

	return TRUE;
}

//替换指定范围的某字段的值.
bool CFoxBase::ReplaceAreaValue(_RecordsetPtr pRs, CString strField, CString strValue, int startRow, int endRow)
{
	if(startRow > endRow)
	{
		endRow = startRow;
	}
	if( endRow <= 0 )
	{
		return false;
	}
	if( endRow > pRs->GetRecordCount() )
	{
		return false;
	}
	if( pRs->GetRecordCount() <= 0 )
	{
		return false;
	}
	pRs->MoveFirst();
	for(int i=1; i<=endRow && !pRs->adoEOF; i++, pRs->MoveNext())
	{
		if( i >= startRow )
		{
			pRs->PutCollect(_variant_t(strField), _variant_t(strValue));
			pRs->Update();
		}
	}
	return true;
}
