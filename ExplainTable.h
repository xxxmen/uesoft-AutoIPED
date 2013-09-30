// ExplainTable.h: interface for the CExplainTable class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_EXPLAINTABLE_H__BCE65130_607A_46A3_8082_5CB133E8E287__INCLUDED_)
#define AFX_EXPLAINTABLE_H__BCE65130_607A_46A3_8082_5CB133E8E287__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

//对说明表的操作
class CExplainTable  
{
public:
	bool EditPaint();
	bool WritePass(CString pass_mode, CString strMode, CString pass_mark);
	bool IsPass(CString strField, CString strMode, CString pass_Last, CString pass_mark, bool flg=false, CString strCur="" );
	bool EditBWTable();
	int pos;
	_ConnectionPtr m_pConnectionCODE;
//	bool moduleCOLOR_S(double& dw,double& dCode, double& dLength, CString& strFace, double& dFace_1, CString& strRing, double& dRing_1, double& dSRing, CString& strArrow, double& dArrow_1, double& dS_arrow, double& dS_word, double& dev_area);
//	bool modC_RING();
	//生成油漆说明表。
	bool CreatePaintTable();
	//数据库连接指针(AutoIPED.mdb)
	_ConnectionPtr m_Pconnection;
	//生成保温说明表。
	bool CreateBWTable();
	CExplainTable();
	virtual ~CExplainTable();

private:
	BOOL IsTableExists(_ConnectionPtr pCon, CString tb);
	BOOL moduleT120();
	BOOL EmptyT120Table() const;
};

#endif // !defined(AFX_EXPLAINTABLE_H__BCE65130_607A_46A3_8082_5CB133E8E287__INCLUDED_)
