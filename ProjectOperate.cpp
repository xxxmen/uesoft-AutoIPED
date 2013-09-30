// ProjectOperate.cpp: implementation of the CProjectOperate class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "autoiped.h"
#include "ProjectOperate.h"

#include "excel9.h"
#include "vtot.h"
#include "mobject.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CProjectOperate::CProjectOperate()
{
	m_bAuto = false;
}

CProjectOperate::~CProjectOperate()
{

}

//////////////////////////////////////////////////////////
//
// 设置工程的ID
//
// strProjectID	[in]工程的ID
//
void CProjectOperate::SetProjectID(LPCTSTR strProjectID)
{
	m_strProjectID=strProjectID; 
}

//////////////////////////////////////////////////////////////
//
// 返回工程的ID
//
CString CProjectOperate::GetProjectID()
{
	return m_strProjectID;
}

///////////////////////////////////////////////////////////////////////
//
// 设置工程数据库的连接
//
// IConnection[in]	数据库的连接
//
void CProjectOperate::SetProjectDbConnect(_ConnectionPtr IConnection)
{
	m_ProjectDbConnect=IConnection;
}

////////////////////////////////////////////////////////////
//
// 返回工程数据库的连接
//
_ConnectionPtr CProjectOperate::GetProjectDbConnect()
{
	return m_ProjectDbConnect;
}

//////////////////////////////////////////////////////////////////////////////////////
//
// 在ImportFromXLSStruct结构的pElement链中寻找匹配的单元
//
// szSource[in]	与ImportFromXLSElement的SourceFieldName匹配的字符串
// szDest[in]	与ImportFromXLSElement的DestinationName匹配的字符串
//
// 如果找到返回指向此单元的指针，如果没找到返回NULL
// 
// szSource,szDest可为空，如果都为NULL将返回NULL
// 大小写不敏感
//
CProjectOperate::ImportFromXLSElement* 
CProjectOperate::CImportFromXLSStruct::FindInElement(LPCTSTR szSource, LPCTSTR szDest)
{
	ImportFromXLSElement *pElement;
	CString strTemp;
	int iIsFind=0;

	if(szSource==NULL && szDest==NULL)
		return NULL;

	if(this->pElement==NULL || this->ElementNum<=0)
		return NULL;

	for(int i=0;i<this->ElementNum;i++)
	{
		iIsFind=1;
		pElement=&this->pElement[i];

		if(szSource)
		{
			strTemp=szSource;
			strTemp.MakeUpper();
			pElement->SourceFieldName.MakeUpper();

			if(pElement->SourceFieldName==strTemp)
			{
				iIsFind &= 1;
			}
			else
			{
				iIsFind &=0;
			}
		}

		if(szDest)
		{
			strTemp=szDest;
			strTemp.MakeUpper();
			pElement->DestinationName.MakeUpper();

			if(pElement->DestinationName==strTemp)
			{
				iIsFind &= 1;
			}
			else
			{
				iIsFind &=0;
			}
		}

		if(iIsFind==1)
			break;
	}

	if(i==this->ElementNum)
		return NULL;

	return &this->pElement[i];
}

/////////////////////////////////////////////////////////////////////////
//
// 从Excel的XLS文件导入数据
//
// pImportStruct[in]	ImportFromXLSStruct结构，内涵导入的信息
//
// throw(_com_error)
// throw(COleDispatchException*)
// throw(COleException*)
//
void CProjectOperate::ImportFromXLS(ImportFromXLSStruct *pImportStruct)
{
	BOOL bRet;
	LPDISPATCH pDispatch;
	_Application Application;
	Workbooks workbooks;
	_Workbook workbook;
	Worksheets worksheets;
	_Worksheet worksheet;
	Range range;

	_RecordsetPtr IRecordset;
	HRESULT hr;
	CString SQL;
	

	if(pImportStruct==NULL || pImportStruct->pElement==NULL || 
	   pImportStruct->ElementNum<=0 || GetProjectDbConnect()==NULL)
	{
		_com_error e(E_INVALIDARG);

		ReportExceptionErrorLV2(e);
		throw e;
	}

	bRet=Application.CreateDispatch(_T("Excel.Application"));

	if(!bRet)
	{
		ReportExceptionErrorLV1(_T("请先安装Excel!"));
		return;
	}

	//获得工作薄的集合
	pDispatch=Application.GetWorkbooks();
	workbooks.AttachDispatch(pDispatch);

	try
	{
		//打开指定的EXCEL文件
		pDispatch=workbooks.Open(pImportStruct->XLSFilePath,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing);
		workbook.AttachDispatch(pDispatch);

		//获得指定工作簿
		pDispatch=workbook.GetWorksheets();
		worksheets.AttachDispatch(pDispatch);

		//获得指定的工作表
		pDispatch=worksheets.GetItem(_variant_t(pImportStruct->XLSTableName));
		worksheet.AttachDispatch(pDispatch);
	}
	catch(COleDispatchException *e)
	{
		Application.Quit();

		CString str;
		str.Format(_T("打开%s文件或打开%s工作表出错"),pImportStruct->XLSFilePath,pImportStruct->XLSTableName);
		Exception::SetAdditiveInfo(_T(str));

		ReportExceptionErrorLV2(e);
		throw;
	}
	catch(COleException *e)
	{
		Application.Quit();

		_com_error e2(e->m_sc);
		e->Delete();
		ReportExceptionErrorLV2(e2);
		throw e2;
	}


	hr=IRecordset.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		Application.Quit();

		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	//
	// 打开指定工程的表
	// 
	//
	if(pImportStruct->ProjectID.IsEmpty())
	{
		SQL.Format(_T("SELECT * FROM %s"),pImportStruct->ProjectDBTableName);
	}
	else
	{
		SQL.Format(_T("SELECT * FROM %s WHERE EnginID='%s'"),pImportStruct->ProjectDBTableName,pImportStruct->ProjectID);
	}

	try
	{
		IRecordset->CursorLocation=adUseClient;

		IRecordset->Open(_variant_t(SQL),
						 _variant_t((IDispatch*)this->GetProjectDbConnect()),
					     adOpenDynamic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		Application.Quit();

		ReportExceptionErrorLV2(e);
		throw;
	}

	int iRow;	// 开始导入的行号
	int Count;	// 已导入的记录数
	CString strTemp,strSourceElementName;
	_variant_t varTemp;

	//开始事务
	this->GetProjectDbConnect()->BeginTrans();

	//
	// 从指定的行开始，逐行逐行从EXCEL文件导数据到数据库
	//
	for(iRow=pImportStruct->BeginRow,Count=0;Count<pImportStruct->RowCount;Count++,iRow++)
	{
		try
		{
			IRecordset->AddNew();

			if(!pImportStruct->ProjectID.IsEmpty())
			{
				IRecordset->PutCollect(_variant_t("EnginId"),_variant_t(pImportStruct->ProjectID));
			}
		}
		catch(_com_error &e)
		{
			Application.Quit();

			//回滚
			this->GetProjectDbConnect()->RollbackTrans();

			ReportExceptionErrorLV2(e);
			throw;
		}

		//
		// 在一行记录中从EXCEL单元中到对应的工程数据库的字段
		//
		for(int i=0;i<pImportStruct->ElementNum;i++)
		{
			strSourceElementName.Format(_T("%s%d"),pImportStruct->pElement[i].SourceFieldName,iRow);

			try
			{
				pDispatch=worksheet.GetRange(_variant_t(strSourceElementName),vtMissing);
				range.AttachDispatch(pDispatch);

				varTemp=range.GetValue();

				range.ReleaseDispatch();
			}
			catch(COleDispatchException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();
		
				strTemp.Format(_T("当从Excel的%s单元导入数据时出错"),strSourceElementName);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV2(e);
				throw;
			}
			catch(COleException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();

				_com_error e2(e->m_sc);
				e->Delete();
				ReportExceptionErrorLV2(e2);
				throw e2;
			}

			try
			{
				IRecordset->PutCollect(_variant_t(pImportStruct->pElement[i].DestinationName),varTemp);
			}
			catch(_com_error &e)
			{
				Application.Quit();

				this->GetProjectDbConnect()->RollbackTrans();


				//
				// 给异常给出更具体的信息
				//
				strTemp.Format(_T("在%s%d的数据有错误"),
							   pImportStruct->pElement[i].SourceFieldName,
							   iRow);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV2(e);
				throw;
			}
		}
	}

	try
	{
		IRecordset->Update();
	}
	catch(_com_error &e)
	{
		Application.Quit();

		this->GetProjectDbConnect()->RollbackTrans();

		ReportExceptionErrorLV2(e);
		throw;
	}

	//保存所有更改并结束当前事务
	this->GetProjectDbConnect()->CommitTrans();

	workbook.Close(_variant_t(false),vtMissing,vtMissing);
	workbooks.Close();
	Application.Quit();
}

/////////////////////////////////////////////////////////////////////////
//
// 从Excel的XLS文件导入数据到PRE_CALC表
//
// pImportStruct[in]	ImportFromXLSStruct结构，内涵导入的信息
// IsAutoMakeID[in]		是否自动边ID号字段，TRUE为自动编写，此时
//						在ImportFromXLSStruct的结构中不应该函有ID字段的信息
// KeyField[in]			关键字段的名称
//
//
// throw(_com_error)
// throw(COleDispatchException*)
// throw(COleException*)
//
//
// 注意:实际上此函数还可以导入其他的表，不光是PRE_CALC，
// 但如果是其他的表IsAutoMakeID应该在KeyField字段的值
// 可以++，既是数值型变量才可为TRUE，否则为FALSE。
// 
// 当IsAutoMakeID为FALSE时，如果数据库中有与KeyField相同的值时
// 那条记录将会被这条新记录替换
//
void CProjectOperate::ImportPre_CalcFromXLS(ImportFromXLSStruct *pImportStruct, BOOL IsAutoMakeID,LPCTSTR KeyField)
{
	BOOL bRet;
	LPDISPATCH pDispatch;
	_Application Application;
	Workbooks workbooks;
	_Workbook workbook;
	Worksheets worksheets;
	_Worksheet worksheet;
	Range range;

	_RecordsetPtr IRecordset;
	HRESULT hr;
	CString SQL;
	

	if(pImportStruct==NULL || pImportStruct->pElement==NULL || 
	   pImportStruct->ElementNum<=0 || GetProjectDbConnect()==NULL)
	{
		_com_error e(E_INVALIDARG);

		ReportExceptionErrorLV2(e);
		throw e;
	}

	bRet=Application.CreateDispatch(_T("Excel.Application"));

	if(!bRet)
	{
		ReportExceptionErrorLV1(_T("请先安装Excel!"));
		return;
	}

	//获得工作薄的集合
	pDispatch=Application.GetWorkbooks();
	workbooks.AttachDispatch(pDispatch);

	try
	{
		//打开指定的EXCEL文件
		pDispatch=workbooks.Open(pImportStruct->XLSFilePath,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing);
		workbook.AttachDispatch(pDispatch);

		//获得指定工作簿
		pDispatch=workbook.GetWorksheets();
		worksheets.AttachDispatch(pDispatch);

		//获得指定的工作表
		pDispatch=worksheets.GetItem(_variant_t(pImportStruct->XLSTableName));
		worksheet.AttachDispatch(pDispatch);
	}
	catch(COleDispatchException *e)
	{
		Application.Quit();

		CString str;
		str.Format(_T("打开%s文件或打开%s工作表出错"),pImportStruct->XLSFilePath,pImportStruct->XLSTableName);
		Exception::SetAdditiveInfo(_T(str));

		ReportExceptionErrorLV2(e);
		throw;
	}
	catch(COleException *e)
	{
		Application.Quit();

		_com_error e2(e->m_sc);
		e->Delete();
		ReportExceptionErrorLV2(e2);
		throw e2;
	}


	hr=IRecordset.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		Application.Quit();

		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	//
	// 打开指定工程的表
	// 
	//
	if(pImportStruct->ProjectID.IsEmpty())
	{
		SQL.Format(_T("SELECT * FROM %s ORDER BY %s"),pImportStruct->ProjectDBTableName,KeyField);
	}
	else
	{
		SQL.Format(_T("SELECT * FROM %s WHERE EnginID='%s' ORDER BY %s"),
				   pImportStruct->ProjectDBTableName,pImportStruct->ProjectID,KeyField);
	}

	try
	{
	//	IRecordset->CursorLocation=adUseClient;

		IRecordset->Open(_variant_t(SQL),
						 _variant_t((IDispatch*)this->GetProjectDbConnect()),
					     adOpenStatic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		Application.Quit();

		ReportExceptionErrorLV2(e);
		throw;
	}

	int iRow;	// 开始导入的行号
	int Count;	// 已导入的记录数
	long ID;
	CString strTemp,strSourceElementName;
	_variant_t varTemp;

	//开始事务
	this->GetProjectDbConnect()->BeginTrans();

	//
	// 如果自动加ID则返回最大的ID号做为下一条记录的ID
	//
	if(IsAutoMakeID)
	{
		if(IRecordset->adoEOF && IRecordset->BOF)
		{
			ID=0;
		}
		else
		{
			IRecordset->MoveLast();

			varTemp=IRecordset->GetCollect(_variant_t(KeyField));

			ID = vtoi( varTemp );
		}

		ID++;
	}
	else
	{
		//
		for(int i=0;i<pImportStruct->ElementNum;i++)
		{
			pImportStruct->pElement[i].DestinationName.MakeUpper();

			if(pImportStruct->pElement[i].DestinationName==KeyField)
			{				
				break;
			}
		}

		if(i==pImportStruct->ElementNum)
		{
			CString strTemp;
			Application.Quit();

			this->GetProjectDbConnect()->RollbackTrans();

			strTemp.Format(_T("没有找到与%s序号对应的Excel列"),KeyField);

			ReportExceptionErrorLV1(_T("没有找到与ID序号对应的Excel列"));
			
			return;
		}
	
		//
		// 不是自动编号则，会删除原来有相同ID号的记录
		//
		for(iRow=pImportStruct->BeginRow,Count=0;Count<pImportStruct->RowCount;Count++,iRow++)
		{
			strSourceElementName.Format(_T("%s%d"),pImportStruct->pElement[i].SourceFieldName,iRow);

			pDispatch=worksheet.GetRange(_variant_t(strSourceElementName),vtMissing);
			range.AttachDispatch(pDispatch);

			varTemp=range.GetValue();

			range.ReleaseDispatch();
			
			if ( vtos(varTemp).IsEmpty() )
			{
				continue;	// 关键字的值为空
			}

			strTemp.Format(_T("%s=%s"),KeyField,(LPCTSTR)(_bstr_t)varTemp);

			try
			{
				IRecordset->Filter = _variant_t(strTemp);

				// 有可能有多条过滤后的记录
				for (; !IRecordset->adoEOF; IRecordset->MoveNext())
				{
					IRecordset->Delete( adAffectCurrent );//
				}
			}
			catch(_com_error &e)
			{
				strSourceElementName+=_T(" 序号有错误");
				Exception::SetAdditiveInfo(strSourceElementName);
				ReportExceptionErrorLV2(e);
				GetProjectDbConnect()->RollbackTrans();
				throw;
			}
		}

	}

	//
	// 从指定的行开始，逐行逐行从EXCEL文件导数据到数据库
	//
	for(iRow=pImportStruct->BeginRow,Count=0;Count<pImportStruct->RowCount;Count++,iRow++)
	{
		try
		{
			IRecordset->AddNew();

			//
			// 自动加ID
			//
			if(IsAutoMakeID)
			{
				IRecordset->PutCollect(_variant_t(KeyField),_variant_t(ID));
				ID++;
			}

			if(!pImportStruct->ProjectID.IsEmpty())
			{
				IRecordset->PutCollect(_variant_t("EnginId"),_variant_t(pImportStruct->ProjectID));
			}
		}
		catch(_com_error &e)
		{
			Application.Quit();

			//回滚
			this->GetProjectDbConnect()->RollbackTrans();

			ReportExceptionErrorLV2(e);
			throw;
		}

		//
		// 在一行记录中从EXCEL单元中到对应的工程数据库的字段
		//
		for(int i=0;i<pImportStruct->ElementNum;i++)
		{
			strSourceElementName.Format(_T("%s%d"),pImportStruct->pElement[i].SourceFieldName,iRow);

			try
			{
				pDispatch=worksheet.GetRange(_variant_t(strSourceElementName),vtMissing);
				range.AttachDispatch(pDispatch);

				varTemp=range.GetValue();

				range.ReleaseDispatch();
			}
			catch(COleDispatchException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();
		
				strTemp.Format(_T("当从Excel的%s单元导入数据时出错"),strSourceElementName);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV2(e);
				throw;
			}
			catch(COleException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();

				_com_error e2(e->m_sc);
				e->Delete();
				ReportExceptionErrorLV2(e2);
				throw e2;
			}

			try
			{
				IRecordset->PutCollect(_variant_t(pImportStruct->pElement[i].DestinationName),varTemp);
			}
			catch(_com_error &e)
			{
				Application.Quit();

				this->GetProjectDbConnect()->RollbackTrans();

				//
				// 给异常给出更具体的信息
				//
				strTemp.Format(_T("在%s%d的数据有错误"),
							   pImportStruct->pElement[i].SourceFieldName,
							   iRow);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV2(e);

				throw;
			}
		}
	}

	try
	{
		IRecordset->Update();
	}
	catch(_com_error &e)
	{
		Application.Quit();

		this->GetProjectDbConnect()->RollbackTrans();

		ReportExceptionErrorLV2(e);
		throw;
	}

	//保存所有更改并结束当前事务
	this->GetProjectDbConnect()->CommitTrans();

	workbook.Close(_variant_t(false),vtMissing,vtMissing);
	workbooks.Close();
	Application.Quit();
}

/////////////////////////////////////////////////////////////////////////
//
// 从Excel的XLS文件导入数据到VOLUME表
//
// pImportStruct[in]	ImportFromXLSStruct结构，内涵导入的信息
//
void CProjectOperate::ImportA_DirFromXLS(ImportFromXLSStruct *pImportStruct)
{
	BOOL bRet;
	LPDISPATCH pDispatch;
	_Application Application;
	Workbooks workbooks;
	_Workbook workbook;
	Worksheets worksheets;
	_Worksheet worksheet;
	Range range;

	_RecordsetPtr IRecordset;
	HRESULT hr;
	CString SQL;
	CString strSourceJcmc;
	CString strSourceJcdm;
	ImportFromXLSElement *pElement;
	long VolumeID;	//
	_variant_t varTemp;


	if(pImportStruct==NULL || pImportStruct->pElement==NULL || 
	   pImportStruct->ElementNum<=0 || GetProjectDbConnect()==NULL)
	{
		_com_error e(E_INVALIDARG);

		ReportExceptionErrorLV2(e);
		throw e;
	}

	//工程号不能为空
	if(pImportStruct->ProjectID.IsEmpty())
	{
		_com_error e(E_FAIL);
		Exception::SetAdditiveInfo(_T("工程号不能为空"));

		ReportExceptionErrorLV2(e);

		throw e;
	}


	//
	// jcdm,jcmc必须有对应的EXCEL列
	//
	pElement=((CImportFromXLSStruct*)pImportStruct)->FindInElement(NULL,_T("jcdm"));
	if(pElement==NULL)
	{
		ReportExceptionErrorLV1(_T("必须要有与jcdm(卷册代号)字段对应的列"));
		return;
	}
	strSourceJcdm=pElement->SourceFieldName;

	pElement=((CImportFromXLSStruct*)pImportStruct)->FindInElement(NULL,_T("jcmc"));
	if(pElement==NULL)
	{
		ReportExceptionErrorLV1(_T("必须要有与jcmc(卷册名称)字段对应的列"));
		return;
	}
	strSourceJcmc=pElement->SourceFieldName;

	//
	// 建立Excel对象
	//
	bRet=Application.CreateDispatch(_T("Excel.Application"));

	if(!bRet)
	{
		ReportExceptionErrorLV1(_T("请先安装Excel!"));
		return;
	}

	//获得工作薄的集合
	pDispatch=Application.GetWorkbooks();
	workbooks.AttachDispatch(pDispatch);

	try
	{
		//打开指定的EXCEL文件
		pDispatch=workbooks.Open(pImportStruct->XLSFilePath,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing);
		workbook.AttachDispatch(pDispatch);

		//获得指定工作簿
		pDispatch=workbook.GetWorksheets();
		worksheets.AttachDispatch(pDispatch);

		//获得指定的工作表
		pDispatch=worksheets.GetItem(_variant_t(pImportStruct->XLSTableName));
		worksheet.AttachDispatch(pDispatch);
	}
	catch(COleDispatchException *e)
	{
		Application.Quit();

		CString str;
		str.Format(_T("打开%s文件或打开%s工作表出错"),pImportStruct->XLSFilePath,pImportStruct->XLSTableName);
		Exception::SetAdditiveInfo(_T(str));

		ReportExceptionErrorLV2(e);
		throw;
	}
	catch(COleException *e)
	{
		Application.Quit();

		_com_error e2(e->m_sc);
		e->Delete();
		ReportExceptionErrorLV2(e2);
		throw e2;
	}


	hr=IRecordset.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		Application.Quit();

		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		throw e;
	}

	//
	// 打开指定工程的表
	// 
	//
	SQL.Format(_T("SELECT * FROM %s ORDER BY VolumeID ASC"),pImportStruct->ProjectDBTableName);

	try
	{
		IRecordset->CursorLocation=adUseClient;

		IRecordset->Open(_variant_t(SQL),
						 _variant_t((IDispatch*)this->GetProjectDbConnect()),
					     adOpenDynamic,adLockOptimistic,adCmdText);

		//如果无记录从1开始
		if(IRecordset->adoEOF && IRecordset->BOF)
		{
			VolumeID=1;
		}
		else
		{
			//
			// 获得下一条记录的ID
			//
			IRecordset->MoveLast();

			varTemp=IRecordset->GetCollect(_variant_t("VolumeID"));

			VolumeID=varTemp;

			VolumeID++;
		}
	}
	catch(_com_error &e)
	{
		Application.Quit();

		ReportExceptionErrorLV2(e);
		throw;
	}

	int iRow;	// 开始导入的行号
	int Count;	// 已导入的记录数
	CString strTemp,strSourceElementName;

	//开始事务
	this->GetProjectDbConnect()->BeginTrans();

	//
	// 从指定的行开始，逐行逐行从EXCEL文件导数据到数据库
	//
	for(iRow=pImportStruct->BeginRow,Count=0;Count<pImportStruct->RowCount;Count++,iRow++)
	{
		strSourceElementName.Format(_T("%s%d"),strSourceJcdm,iRow);

		try
		{
			pDispatch=worksheet.GetRange(_variant_t(strSourceElementName),vtMissing);
			range.AttachDispatch(pDispatch);

			varTemp=range.GetValue();

			range.ReleaseDispatch();
		}
		catch(COleDispatchException *e)
		{
			Application.Quit();
			this->GetProjectDbConnect()->RollbackTrans();
	
			strTemp.Format(_T("当从Excel的%s单元导入数据时出错"),strSourceElementName);
			Exception::SetAdditiveInfo(strTemp);

			ReportExceptionErrorLV2(e);
			throw;
		}
		catch(COleException *e)
		{
			Application.Quit();
			this->GetProjectDbConnect()->RollbackTrans();

			_com_error e2(e->m_sc);
			e->Delete();
			ReportExceptionErrorLV2(e2);
			throw e2;
		}


		if(varTemp.vt==VT_NULL || varTemp.vt==VT_EMPTY)
		{
			Application.Quit();
			this->GetProjectDbConnect()->RollbackTrans();

			_com_error e(E_FAIL);

			strTemp.Format(_T("%s单元不能为空"),strSourceElementName);
			Exception::SetAdditiveInfo(strTemp);

			ReportExceptionErrorLV2(e);
			throw e;
		}

		SQL.Format(_T("jcdm='%s' AND EnginID='%s'"),(LPCTSTR)(_bstr_t)varTemp,pImportStruct->ProjectID);


		try
		{
			//
			// 如果有同样的卷册浩jcdm将会用现在的数据替换原来的数据并保留VolumeID
			// 如果没有则新增数据.
			//
			IRecordset->Filter=_variant_t(SQL);

			if(IRecordset->adoEOF && IRecordset->BOF)
			{
				IRecordset->AddNew();

				IRecordset->PutCollect(_variant_t("VolumeID"),_variant_t(VolumeID));

				VolumeID++;
			}
			else
			{
				varTemp=IRecordset->GetCollect(_variant_t("VolumeID"));

				IRecordset->Delete(adAffectCurrent);
				IRecordset->MoveNext();

				IRecordset->AddNew();

				IRecordset->PutCollect(_variant_t("VolumeID"),varTemp);
			}

			if(!pImportStruct->ProjectID.IsEmpty())
			{
				IRecordset->PutCollect(_variant_t("EnginId"),_variant_t(pImportStruct->ProjectID));
			}

			//
			// 设置sjhyid,sjjdid,zyid默认的值为0，4，3
			//
			IRecordset->PutCollect(_variant_t("SJHYID"),_variant_t((long)0));
			IRecordset->PutCollect(_variant_t("SJJDID"),_variant_t((long)4));
			IRecordset->PutCollect(_variant_t("ZYID"),_variant_t((long)3));
		}
		catch(_com_error &e)
		{
			Application.Quit();

			//回滚
			this->GetProjectDbConnect()->RollbackTrans();

			ReportExceptionErrorLV2(e);
			throw;
		}

		//
		// 在一行记录中从EXCEL单元中到对应的工程数据库的字段
		//
		for(int i=0;i<pImportStruct->ElementNum;i++)
		{
			strSourceElementName.Format(_T("%s%d"),pImportStruct->pElement[i].SourceFieldName,iRow);

			try
			{
				pDispatch=worksheet.GetRange(_variant_t(strSourceElementName),vtMissing);
				range.AttachDispatch(pDispatch);

				varTemp=range.GetValue();

				range.ReleaseDispatch();
			}
			catch(COleDispatchException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();
		
				strTemp.Format(_T("当从Excel的%s单元导入数据时出错"),strSourceElementName);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV2(e);
				throw;
			}
			catch(COleException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();

				_com_error e2(e->m_sc);
				e->Delete();
				ReportExceptionErrorLV2(e2);
				throw e2;
			}

			try
			{
				IRecordset->PutCollect(_variant_t(pImportStruct->pElement[i].DestinationName),varTemp);
			}
			catch(_com_error &e)
			{
				Application.Quit();

				this->GetProjectDbConnect()->RollbackTrans();


				//
				// 给异常给出更具体的信息
				//
				strTemp.Format(_T("在%s%d的数据有错误"),
							   pImportStruct->pElement[i].SourceFieldName,
							   iRow);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV2(e);
				throw;
			}
		}
	}

	try
	{
		IRecordset->Update();
	}
	catch(_com_error &e)
	{
		Application.Quit();

		this->GetProjectDbConnect()->RollbackTrans();

		ReportExceptionErrorLV2(e);
		return;
	}

	//保存所有更改并结束当前事务
	this->GetProjectDbConnect()->CommitTrans();

	workbook.Close(_variant_t(false),vtMissing,vtMissing);
	workbooks.Close();
	Application.Quit();
}



bool CProjectOperate::ImportTblEnginIDXLS(ImportFromXLSStruct *pImportStruct, BOOL IsAutoMakeID, LPCTSTR KeyField, CString EnginID)
{
	BOOL bRet;
	LPDISPATCH pDispatch;
	_Application Application;
	Workbooks workbooks;
	_Workbook workbook;
	Worksheets worksheets;
	_Worksheet worksheet;
	Range		range;
	CString strTemp;

//	CMObject  ExcelApp;

	_RecordsetPtr IRecordset, pRs;
	HRESULT hr;
	CString SQL;
	LIST_ID_ENGIN listID[10];  //不同工程对应的最大ID号
	
	if(pImportStruct==NULL || pImportStruct->pElement==NULL || 
	   pImportStruct->ElementNum<=0 || GetProjectDbConnect()==NULL)
	{
		_com_error e(E_INVALIDARG);		
		ReportExceptionErrorLV2(e);
		return FALSE;
	}
	try
	{
/*		if( FAILED(ExcelApp.GetActiveObject(_T("Excel.Application"))) )  //zsy
		{
			if( FAILED(ExcelApp.CreateObject(_T("Excel.Application"))) )
			{
				AfxMessageBox("不能启动Excel或没有安装Excel !");
				return false;
			}
		}
		Application.AttachDispatch( ExcelApp );
*/

		//重新启动EXCEL 
		bRet=Application.CreateDispatch(_T("Excel.Application"));		
		if(!bRet)
		{
			ReportExceptionErrorLV1(_T("不能启动Excel, 请确认是否安装Excel !"));
			return false;
		}
		
		//获得工作薄的集合
		pDispatch=Application.GetWorkbooks();
		workbooks.AttachDispatch(pDispatch);
	}
	catch (COleDispatchException * e)
	{
		e->Delete();
		return FALSE;
	}

	try
	{
		//打开指定的EXCEL文件
		pDispatch=workbooks.Open(pImportStruct->XLSFilePath,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,
								 vtMissing);
		workbook.AttachDispatch(pDispatch);
		//获得指定工作簿
		pDispatch=workbook.GetWorksheets();
		worksheets.AttachDispatch(pDispatch);
		//获得指定的工作表
		pDispatch=worksheets.GetItem(_variant_t(pImportStruct->XLSTableName));
		worksheet.AttachDispatch(pDispatch);		
	}
	catch(COleDispatchException *e)
	{
		Application.Quit();
		strTemp.Format(_T("打开%s文件或打开%s工作表出错"),pImportStruct->XLSFilePath,pImportStruct->XLSTableName);
		Exception::SetAdditiveInfo( strTemp );
		ReportExceptionErrorLV1(e);
		return FALSE;
	}
	catch(COleException *e)
	{
		Application.Quit();
		_com_error e2(e->m_sc);
		e->Delete();
		ReportExceptionErrorLV2(e2);
		return FALSE ;
	}
	hr=IRecordset.CreateInstance(__uuidof(Recordset));
	pRs.CreateInstance(__uuidof(Recordset));
	if(FAILED(hr))
	{
		Application.Quit();
		_com_error e(hr);
		ReportExceptionErrorLV2(e);
		return FALSE;
	}
	//
	// 打开指定工程的表
	// 
	if(pImportStruct->ProjectID.IsEmpty())
	{
		SQL.Format(_T("SELECT * FROM %s ORDER BY %s"),pImportStruct->ProjectDBTableName,KeyField);
	}
	else
	{
		SQL.Format(_T("SELECT * FROM %s WHERE EnginID='%s' ORDER BY %s"),
				   pImportStruct->ProjectDBTableName,pImportStruct->ProjectID,KeyField);
	}
	try
	{
		IRecordset->Open(_variant_t(SQL),
						 _variant_t((IDispatch*)this->GetProjectDbConnect()),
					     adOpenDynamic,adLockOptimistic,adCmdText);
	}
	catch(_com_error &e)
	{
		Application.Quit();
		ReportExceptionErrorLV2(e);
		return FALSE;
	}

	int iRow;	// 开始导入的行号
	int Count;	// 已导入的记录数
	long ID;
	CString strSourceElementName;
	_variant_t varTemp;
	int ProjNum = 1;
	int ProjIndex = 0;
	int Rj=0;			// 工程号在Excel中的列
	for(; Rj < pImportStruct->ElementNum; Rj++)
	{
		pImportStruct->pElement[Rj].DestinationName.MakeUpper();
		if( !EnginID.CompareNoCase(pImportStruct->pElement[Rj].DestinationName) )
			break;
	}
	//开始事务
	this->GetProjectDbConnect()->BeginTrans();

	if(IsAutoMakeID)
	{
		if(IRecordset->adoEOF && IRecordset->BOF)
			ID=0;
		else
		{
			IRecordset->MoveLast();
			varTemp=IRecordset->GetCollect(_variant_t(KeyField));
			ID = vtoi( varTemp );
		}
		ID++;
	}
	else
	{
		ID = -1;	// 如果不是自动编号时,使用文件中的序号
		for(int i = 0; i < pImportStruct->ElementNum; i++)
		{
			pImportStruct->pElement[i].DestinationName.MakeUpper();
			if ( pImportStruct->pElement[i].DestinationName == KeyField )
				break;
		}
		if ( i == pImportStruct->ElementNum )
		{
			Application.Quit();
			this->GetProjectDbConnect()->RollbackTrans();
			strTemp.Format(_T("没有找到与%s序号对应的Excel列"),KeyField);
			ReportExceptionErrorLV1( strTemp );			
			return FALSE;
		}
		CString strNum;		//序号
		CString strEng;		//工程号
		// 不是自动编号则，会删除原来有相同ID号的记录(当前工程号的记录)
		for(iRow=pImportStruct->BeginRow, Count=0; Count<pImportStruct->RowCount;Count++,iRow++)
		{
			if ( Rj < pImportStruct->ElementNum )
			{
				strSourceElementName.Format(_T("%s%d"), pImportStruct->pElement[Rj].SourceFieldName,iRow);
				pDispatch = worksheet.GetRange(_variant_t( strSourceElementName ), vtMissing);
				range.AttachDispatch( pDispatch );
				varTemp = range.GetValue();
				strEng = vtos( varTemp );
				range.ReleaseDispatch();
				if ( !strEng.IsEmpty() && strEng.CompareNoCase( pImportStruct->ProjectID ) )
					continue;
			}
			strSourceElementName.Format(_T("%s%d"), pImportStruct->pElement[i].SourceFieldName,iRow);
			pDispatch = worksheet.GetRange(_variant_t(strSourceElementName),vtMissing);
			range.AttachDispatch( pDispatch );
			varTemp = range.GetValue();
			strNum = vtos( varTemp );
			range.ReleaseDispatch();
			if ( strNum.IsEmpty() )
				continue;	// 关键字的值为空

			strTemp.Format(_T("%s=%s"), KeyField, strNum );
			try
			{
				IRecordset->Filter = _variant_t(strTemp);
				// 有可能有多条过滤后的记录
				for (; !IRecordset->adoEOF; IRecordset->MoveNext())
				{
					IRecordset->Delete( adAffectCurrent );//
				}
			}
			catch(_com_error &e)
			{
				strSourceElementName+=_T(" 序号有错误");
				Exception::SetAdditiveInfo(strSourceElementName);
				ReportExceptionErrorLV1(e);
				GetProjectDbConnect()->RollbackTrans();
				return FALSE;
			}
		}
	}
	//在1位设置当前工程
	listID[0].ID = ID;
	listID[0].EnginID = pImportStruct->ProjectID;
	//
	// 从指定的行开始，逐行逐行从EXCEL文件导数据到数据库
	//
	for(iRow=pImportStruct->BeginRow,Count=0;Count<pImportStruct->RowCount;Count++,iRow++)
	{
		try
		{
			IRecordset->AddNew();
		}
		catch(_com_error &e)
		{
			Application.Quit();

			//回滚
			this->GetProjectDbConnect()->RollbackTrans();

			ReportExceptionErrorLV2(e);
			return FALSE;
		}

		//
		// 在一行记录中从EXCEL单元中到对应的工程数据库的字段
		//
		for(int i=0;i<pImportStruct->ElementNum;i++)
		{
			strSourceElementName.Format(_T("%s%d"),pImportStruct->pElement[i].SourceFieldName,iRow);

			try
			{
				pDispatch=worksheet.GetRange(_variant_t(strSourceElementName),vtMissing);
				range.AttachDispatch(pDispatch);

				varTemp=range.GetValue();

				range.ReleaseDispatch();
			}
			catch(COleDispatchException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();
		
				strTemp.Format(_T("当从Excel的%s单元导入数据时出错"),strSourceElementName);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV1(e);
				return FALSE;
			}
			catch(COleException *e)
			{
				Application.Quit();
				this->GetProjectDbConnect()->RollbackTrans();

				_com_error e2(e->m_sc);
				e->Delete();
				ReportExceptionErrorLV2(e2);
				return FALSE;
			}

			try
			{
				IRecordset->PutCollect(_variant_t(pImportStruct->pElement[i].DestinationName),varTemp);

				//{{2不同的工程名称，加不同的ID号
				if( i == Rj /*工程名所在的列*/)
				{
					strTemp = vtos(varTemp);
					for(int c=0; c<ProjNum; c++)
					{
						if( !strTemp.CompareNoCase(listID[c].EnginID) || (c==0 && strTemp.IsEmpty()))
						{
							ProjIndex = c;
							break;
						}
					}
					if( c == ProjNum )
					{
						listID[c].EnginID = strTemp;
						SQL = "SELECT * FROM ["+pImportStruct->ProjectDBTableName+"]  \
							WHERE EnginID='"+listID[c].EnginID+"' ORDER BY "+KeyField+" ";
						if(pRs->State == adStateOpen)
						{
							pRs->Close();
						}

						pRs->Open(_bstr_t(SQL), (IDispatch*)m_ProjectDbConnect, 
								adOpenStatic, adLockOptimistic, adCmdText);
						if( pRs->adoEOF && pRs->BOF )
						{
							ID = 0;
						}
						else
						{
							pRs->MoveLast();
							varTemp = pRs->GetCollect(_bstr_t(KeyField));
							ID = vtoi(varTemp);
						}
						//最大的序号,将记录加到末尾.
						listID[c].ID = ++ID;
						ProjIndex = c;
						ProjNum++;
					}
				}
				//2}}
			}
			catch(_com_error &e)
			{
				Application.Quit();

				this->GetProjectDbConnect()->RollbackTrans();

				//
				// 给异常给出更具体的信息
				//
				strTemp.Format(_T("在%s%d的数据有错误"),
							   pImportStruct->pElement[i].SourceFieldName,
							   iRow);
				Exception::SetAdditiveInfo(strTemp);

				ReportExceptionErrorLV1(e);

				return FALSE;
			}
		}
		try
		{
			if ( listID[ProjIndex].ID > 0 )
			{   // 不为自动同时序号大于零时,进行编号
				strTemp.Format("%d",listID[ProjIndex].ID);
				IRecordset->PutCollect(_variant_t(KeyField),_variant_t(strTemp));
				listID[ProjIndex].ID++;
			}
			IRecordset->PutCollect(_variant_t("EnginId"),_variant_t(listID[ProjIndex].EnginID));
		}catch(_com_error& e)
		{
			AfxMessageBox(e.Description());
			return false;
		}
	}

	try
	{
		IRecordset->Update();
	}
	catch(_com_error &e)
	{
		Application.Quit();

		this->GetProjectDbConnect()->RollbackTrans();

		ReportExceptionErrorLV1(e);
		return FALSE;
	}

	//保存所有更改并结束当前事务
	this->GetProjectDbConnect()->CommitTrans();


	workbook.Close(_variant_t(false),vtMissing,vtMissing);
	workbooks.Close();	
	Application.Quit();
	Application.ReleaseDispatch();

	return true;
}
//导入AutoPHS的支吊架油漆数据.
bool CProjectOperate::ImportAutoPHSPaint(CString strFilePath)
{
	CString strSQL, strValue;
	try
	{
		if( !ConnectExcelFile( strFilePath ) )
		{
			CString strMessage;
			strMessage = "不能连接文件,请确认 ' "+strFilePath+" ' 该文件是否存在!";
			AfxMessageBox( strMessage );
			return false;
		}
		CRecordset pRs;
		pRs.m_pDatabase = &m_db;//SUM(CLzz) as CLzzs
	
		strSQL.Format("SELECT SUM(CLzz) as CLzzs FROM [管部汇总表] ");
	//	m_db.ExecuteSQL(strSQL);
		pRs.Open(AFX_DB_USE_DEFAULT_TYPE, strSQL, CRecordset::readOnly);
		if( !pRs.IsBOF() )
		{
			pRs.GetFieldValue((short)0, strValue );
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
	}
	return true;
	

}

//连接Excel文件.
bool CProjectOperate::ConnectExcelFile(CString strFileName)
{
	//新建文件	
	CString strDriver = "MICROSOFT EXCEL DRIVER (*.XLS)"; // Excel安装驱动	
	CString strCon, strSQL;
	
	try
	{
		// 创建进行存取的字符串
		strCon.Format("DRIVER={%s};DSN='''';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"%s\";DBQ=%s",
			strDriver, strFileName, strFileName);
		
		// 创建数据库 (既Excel表格文件)
		if( !m_db.OpenEx(strCon,CDatabase::noOdbcDialog) )
		{
			return false;
		}
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
	}
	return TRUE;

}
//////////////////////////////////
//功能:
//联接一个指定的数据库
bool CProjectOperate::ConnectDB(_ConnectionPtr& pCon,CString& strManageTbl,CString& strExcelTbl,CString& strDataTbl,int& nStart,int& nRowNum, CString strDBFile,int nIndex)
{
	_variant_t var;
	CString strSQL, strTMP;
	HRESULT hr;
	_RecordsetPtr pRs;

	hr = pCon.CreateInstance(__uuidof(Connection));
	if( FAILED(hr) )
	{
		_com_error e(hr);
		ReportExceptionErrorLV1( e );
		return false;
	}
	strSQL = CONNECTSTRING + strDBFile;
	try
	{   //连接数据库
		pCon->Open(_bstr_t(strSQL),"","",-1);

		hr = pRs.CreateInstance(__uuidof(Recordset));
		if( FAILED(hr) )
		{
			_com_error e(hr);
			ReportExceptionErrorLV1(e);				
			return false;
		}
		//打开管理表.获得控制变量.
		strTMP.Format("%d",nIndex);
		strSQL = "SELECT * FROM [TableExcelToAccess] WHERE ID="+strTMP+" ";
		pRs->Open(_bstr_t(strSQL), pCon.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		if( pRs->GetRecordCount() <= 0 )
		{
			return false;
		}
		//管理表,储蓄字段的对应关系.
		var = pRs->GetCollect("TableName");
		strManageTbl = vtos(var);
		//Excel表名
		var = pRs->GetCollect("ExcelTblName");
		strExcelTbl = vtos(var);
		//Access原始数据表
		var = pRs->GetCollect("AccessTblName");
		strDataTbl = vtos(var);

		//Excel中开始导入的行号
		var = pRs->GetCollect("StartRow");
		nStart = vtoi(var);
		//总共导入数据的行数
		var = pRs->GetCollect("RecordNum");
		nRowNum = vtoi(var);
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}	
	return true;
}

////////////////////////////////
//功能:
//使用一个管理数据库,导入excel数据到AutoIPED.
//(通用的一个函数.)
bool CProjectOperate::ImportOriginalData(CString strDataPath/*数据库路径*/, CString strManageDB, int nIndex)
{
	_ConnectionPtr pCon;
	_RecordsetPtr pRs, pRsOrig;
	HRESULT hr;
	_variant_t var;
	CString strSQL,			//SQL语句
			strTMP,			//临时变量.
			strManageTbl,   //管理表,储蓄字段的对应关系.
			strExcelTbl,	//Excel表名
			strDataTbl		//Access原始数据表.
			;
	int		nStart,   //Excel中开始导入的行号(从1开始)
			nRowNum;  //总共导入数据的行数.
	try
	{
		//连接数据库,并获得管理表中的值.
		if( !ConnectDB(pCon,strManageTbl,strExcelTbl,strDataTbl,nStart,nRowNum,strDataPath,nIndex) )
		{
			return false;
		}
		CProjectOperate::ImportFromXLSElement  element[30]; //对应
		int elementNum = 0 ;//导入的列数.

		GetFieldAsCol(pCon,element,elementNum, strManageTbl);
		CProjectOperate::ImportFromXLSStruct   structPrj;
		structPrj.pElement = element;

		hr = pRs.CreateInstance(__uuidof(Recordset));
		if( FAILED(hr) )
		{
			_com_error e(hr);
			AfxMessageBox(e.Description());
			return false;
		}
		strSQL = "SELECT * FROM ["+strManageTbl+"] ";
		pRs->Open(_bstr_t(strSQL), pCon.GetInterfacePtr(), 
			adOpenStatic, adLockOptimistic, adCmdText);

	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}

	return true;
}
//获得Access数据库中的字段与Excel表中的列的对应关系.
//
bool CProjectOperate::GetFieldAsCol(_ConnectionPtr& pCon,CProjectOperate::ImportFromXLSElement element[],int& elementNum, CString& strTblName)
{
	_RecordsetPtr pRs;
	HRESULT hr;
	CString strSQL;
	_variant_t var;

	hr = pRs.CreateInstance(__uuidof(Recordset));
	if( FAILED(hr) )
	{
		_com_error e(hr);
		AfxMessageBox(e.Description());
		return false;
	}
	strSQL = "SELECT * FROM ["+strTblName+"] WHERE ExcelColNO IS NOT NULL ORDER BY SEQ";
	try
	{
		pRs->Open(_bstr_t(strSQL), pCon.GetInterfacePtr(),
					adOpenStatic, adLockOptimistic, adCmdText);
		for(elementNum=0; !pRs->adoEOF; pRs->MoveNext(), elementNum++ )
		{
			var = pRs->GetCollect("FileName");		//数据库中的字段名
			element[elementNum].DestinationName = vtos(var);
			var = pRs->GetCollect("ExcelColNO");	//与Excel对应的列
			element[elementNum].SourceFieldName = vtos(var);
			var = pRs->GetCollect("ExcelColNum");	//Excel中对应一个字段的列数.
			element[elementNum].ExcelColNum = vtoi(var);
		}
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
		return false;
	}

	return true;
}
