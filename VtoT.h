#ifndef VTOT_H              //防止多次包括头文件
#define VTOT_H

//将_variant_t的数据转化到int
extern inline int  vtoi(_variant_t & v);
//将_variant_t的数据转化到CString
extern inline CString  vtos(_variant_t& v);
extern inline CString  vtos(COleVariant& v);
//将_variant_t的数据转化到double;
extern inline double  vtof(_variant_t &v);
//将_variant_t的数据转化到bool
extern inline bool  vtob( _variant_t &v);
extern inline bool  Btob( BOOL Bn);
extern inline bool  Dtob( double Dn);

#endif
