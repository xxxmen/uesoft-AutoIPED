#if !defined MOBJECT_H
#define MOBJECT_H


//用于创建和操作以自动化为基的类。（即它所操作的对象必需支持自动化）
class CMObject
{
public:
	CMObject(_variant_t& tvar);
	
	CMObject()
	{
		p = NULL;
	}

	CMObject(IDispatch* lp)
	{
		if((p = lp) != NULL)
			p->AddRef();
	}

	CMObject(IUnknown* lp)
	{
		p=NULL;
		if (lp != NULL)
			lp->QueryInterface(IID_IDispatch, (void **)&p);
	}

	~CMObject()
	{ 
		try
		{
			if (p) p->Release();
		}
		catch(...)
		{
		}
		p=NULL;
	}
	void Release() 
	{
		try
		{
			if (p) p->Release();
		}
		catch(...){}
		p=NULL;
	}

//Member Functions
public:
	HRESULT CreateObject(const CLSID clsid);
	HRESULT GetActiveObject(const CLSID clsid);
	HRESULT CreateObject(LPCTSTR objName);
	HRESULT GetActiveObject(LPCTSTR objName);

	HRESULT GetIDOfName(/*LPCOLESTR*/LPCTSTR lpsz, DISPID* pdispid);

	operator IDispatch*() {return p;}
	IDispatch& operator*() { return *p; }
	IDispatch** operator&() { return &p; }
	IDispatch* operator->() {return p; }
	IDispatch* operator=(IDispatch* lp);
	IDispatch* operator=(IUnknown* lp);
	IDispatch* operator=(_variant_t& newVar);
	BOOL operator!(){return (p == NULL) ? TRUE : FALSE;}

	_variant_t GetPropertyByName(/*LPCOLESTR*/LPCTSTR lpsz);
	_variant_t GetProperty(DISPID dwDispID);
	HRESULT PutPropertyByName(/*LPCOLESTR*/LPCTSTR lpsz, VARIANT* pVar);
	HRESULT PutProperty(DISPID dwDispID, VARIANT* pVar);
	static HRESULT GetProperty(IDispatch* pDisp, DISPID dwDispID,
		VARIANT* pVar);
	static HRESULT PutProperty(IDispatch* pDisp, DISPID dwDispID,
		VARIANT* pVar);

	
	_variant_t Invoke0(DISPID dispid);
	_variant_t Invoke0(/*LPCOLESTR*/LPCTSTR lpszName);
	_variant_t Invoke(DISPID dispid, long nParamCount, VARIANT * ...);
	_variant_t Invoke(LPCTSTR sName,long nParamCount,VARIANT* ...);

//Member Variant
public:
	IDispatch* p;
protected:
	_variant_t InvokeV(DISPID dispid,long nParamCount,va_list argList);
};
#endif