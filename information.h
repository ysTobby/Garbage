#pragma once
#include "stdafx.h"
class Cinformation
{
public:
	Cinformation();
	double cp,ct,ss,kb,a1,gjss,bc,fi,pminp,yszya,Assl,b1;
	void CTpress(CString type);
	void CTtension(CString type);
	void ST(CString type);
	void GST(CString type);
	void kesaib(CString Ctype,CString Stype);
	void arfa1(CString type);
	void init(CString Ctype,CString Stype);
	void Bc(CString Ctype);
	void initV(CString Ctype,CString ZJtype,CString GJtype);
	void initZY(CString Ctype,CString Stype,int con,double l0,double i);
	void initPY(CString Ctype,CString Stype);
	void initYSZY(CString Ctype,CString ZJtype,CString GJtype,double l0,double i,double gd);
	void calfi(int con,double l0,double i);
	void calpminp(CString Ctype,CString Stype);
	double xxcz(double x,double x1,double x2,double y1,double y2);
	void calculayszya(CString Ctype);
	void calb1(CString Ctype);
};//·ÖºÅ²»ÄÜµô