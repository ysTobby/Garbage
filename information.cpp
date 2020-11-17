#include "information.h"
#include "stdafx.h"//需要将头文件放入这里面进行预编译
Cinformation::Cinformation()
{
	cp = 0,ct = 0,ss = 0,kb = 0,b1 = 0,gjss = 0,bc = 0,yszya = 0,Assl = 0,fi = 0,b1 = 0;
}

void Cinformation::CTpress(CString type)
{
	if(type == "C15")
		cp = 7.2;
	else if(type == "C20")
		cp = 9.6;
	else if(type == "C25")
		cp = 11.9;
	else if(type == "C30")
		cp = 14.3;
	else if(type == "C35")
		cp = 16.7;
	else if(type == "C40")
		cp = 19.1;
	else if(type == "C45")
		cp = 21.1;
	else if(type == "C50")
		cp = 23.1;
	else if(type == "C55")
		cp = 25.3;
	else if(type == "C60")
		cp = 27.5;
	else if(type == "C65")
		cp = 29.7;
	else if(type == "C70")
		cp = 31.8;
	else if(type == "C75")
		cp = 33.8;
	else if(type == "C80")
		cp = 35.9;
}

void Cinformation::CTtension(CString type)
{
	if(type == "C15")
		ct = 0.91;
	else if(type == "C20")
		ct = 1.10;
	else if(type == "C25")
		ct = 1.27;
	else if(type == "C30")
		ct = 1.43;
	else if(type == "C35")
		ct = 1.57;
	else if(type == "C40")
		ct = 1.71;
	else if(type == "C45")
		ct = 1.80;
	else if(type == "C50")
		ct = 1.89;
	else if(type == "C55")
		ct = 1.96;
	else if(type == "C60")
		ct = 2.04;
	else if(type == "C65")
		ct = 2.09;
	else if(type == "C70")
		ct = 2.14;
	else if(type == "C75")
		ct = 2.18;
	else if(type == "C80")
		ct = 2.22;
}

void Cinformation::ST(CString type)
{
	if(type == "HPB300")
		ss = 270.0;
	else if(type == "HRB335"||type == "HRBF335")
		ss = 300.0;
	else if(type == "HRB400"||type == "HRBF400"||type == "RRB400")
		ss = 360.0;
	else if(type == "HRB500"||type == "HRBF500")
		ss = 435.0;
}

void Cinformation::GST(CString type)
{
	if(type == "HPB300")
		gjss = 270.0;
	else if(type == "HRB335"||type == "HRBF335")
		gjss = 300.0;
	else if(type == "HRB400"||type == "HRBF400"||type == "RRB400")
		gjss = 360.0;
	else if(type == "HRB500"||type == "HRBF500")
		gjss = 435.0;
}

void Cinformation::kesaib(CString Ctype,CString Stype)
{
	if(Ctype == "C15"||Ctype == "C20"||Ctype == "C25"||Ctype == "C30"||Ctype == "C35"||Ctype == "C40"||Ctype == "C45"||Ctype == "C50")
	{
		if(Stype == "HPB300")
			kb = 0.576;
		else if(Stype == "HRB335"||Stype == "HRBF335")
			kb = 0.550;
		else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
			kb = 0.518;
		else if(Stype == "HRB500"||Stype == "HRBF500")
			kb = 0.482;
	}
	else if(Ctype == "C55")
	{
		if(Stype == "HPB300")
			kb = 0.569;
		else if(Stype == "HRB335"||Stype == "HRBF335")
			kb = 0.543;
		else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
			kb = 0.511;
		else if(Stype == "HRB500"||Stype == "HRBF500")
			kb = 0.476;
	}
	else if(Ctype == "C60")
	{
		if(Stype == "HPB300")
			kb = 0.557;
		else if(Stype == "HRB335"||Stype == "HRBF335")
			kb = 0.531;
		else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
			kb = 0.499;
		else if(Stype == "HRB500"||Stype == "HRBF500")
			kb = 0.464;
	}
	else if(Ctype == "C65")
	{
		if(Stype == "HPB300")
			kb = 0.554;
		else if(Stype == "HRB335"||Stype == "HRBF335")
			kb = 0.529;
		else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
			kb = 0.498;
		else if(Stype == "HRB500"||Stype == "HRBF500")
			kb = 0.464;
	}
	else if(Ctype == "C70")
	{
		if(Stype == "HPB300")
			kb = 0.537;
		else if(Stype == "HRB335"||Stype == "HRBF335")
			kb = 0.512;
		else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
			kb = 0.481;
		else if(Stype == "HRB500"||Stype == "HRBF500")
			kb = 0.447;
	}
	else if(Ctype == "C75")
	{
		if(Stype == "HPB300")
			kb = 0.528;
		else if(Stype == "HRB335"||Stype == "HRBF335")
			kb = 0.503;
		else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
			kb = 0.472;
		else if(Stype == "HRB500"||Stype == "HRBF500")
			kb = 0.438;
	}
	else if(Ctype == "C80")
	{
		if(Stype == "HPB300")
			kb = 0.518;
		else if(Stype == "HRB335"||Stype == "HRBF335")
			kb = 0.493;
		else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
			kb = 0.463;
		else if(Stype == "HRB500"||Stype == "HRBF500")
			kb = 0.429;
	}
}

void Cinformation::arfa1(CString type)
{
	if(type == "C15")
		a1 = 1.0;
	else if(type == "C20")
		a1 = 1.0;
	else if(type == "C25")
		a1 = 1.0;
	else if(type == "C30")
		a1 = 1.0;
	else if(type == "C35")
		a1 = 1.0;
	else if(type == "C40")
		a1 = 1.0;
	else if(type == "C45")
		a1 = 1.0;
	else if(type == "C50")
		a1 = 1.0;
	else if(type == "C55")
		a1 = 0.99;
	else if(type == "C60")
		a1 = 0.98;
	else if(type == "C65")
		a1 = 0.97;
	else if(type == "C70")
		a1 = 0.96;
	else if(type == "C75")
		a1 = 0.95;
	else if(type == "C80")
		a1 = 0.94;
}

void Cinformation::init(CString Ctype,CString Stype)
{
	cp = 0,ct = 0,ss = 0,kb = 0,b1 = 0,gjss = 0,bc = 0;
	this->CTtension(Ctype);
	this->CTpress(Ctype);
	this->ST(Stype);
	this->kesaib(Ctype,Stype);
	this->arfa1(Ctype);
}//ctrl K F

void Cinformation::Bc(CString type)
{
	if(type == "C15"||type == "C20"||type == "C25"||type == "C30"||type == "C35"||type == "C40"||type == "C45"||type == "C50")
		bc = 1.0;
	else if(type == "C55")
		bc = 0.97;
	else if(type == "C60")
		bc = 0.93;
	else if(type == "C65")
		bc = 0.9;
	else if(type == "C70")
		bc = 0.87;
	else if(type == "C75")
		bc = 0.83;
	else if(type == "C80")
		bc = 0.8;
}

void Cinformation::initV(CString Ctype,CString ZJtype,CString GJtype)
{
	this->CTtension(Ctype);
	this->CTpress(Ctype);
	this->ST(ZJtype);
	this->GST(GJtype);
	this->Bc(Ctype);
}

double Cinformation::xxcz(double x,double x1,double x2,double y1,double y2)
{
	return (x - x2)/(x1 - x2)*y1 + (x - x1)/(x2 - x1)*y2;
}

void Cinformation::calfi(int con,double l0,double i)
{
	if(con == 1)//方
	{
		if(l0/i <= 8)
		{
			fi = 1.0;
		}
		else if(l0/i <= 10)
		{
			fi = xxcz(l0/i,8.0,10.0,1.0,0.98);
		}
		else if(l0/i <= 12)
		{
			fi = xxcz(l0/i,10.0,12.0,0.98,0.95);
		}
		else if(l0/i <= 14)
		{
			fi = xxcz(l0/i,12.0,14.0,0.95,0.92);
		}
		else if(l0/i <= 16)
		{
			fi = xxcz(l0/i,14.0,16.0,0.92,0.87);
		}
		else if(l0/i <= 18)
		{
			fi = xxcz(l0/i,16.0,18.0,0.87,0.81);
		}
		else if(l0/i <= 20)
		{
			fi = xxcz(l0/i,18.0,20.0,0.81,0.75);
		}
		else if(l0/i <= 22)
		{
			fi = xxcz(l0/i,20.0,22.0,0.75,0.70);
		}
		else if(l0/i <= 24)
		{
			fi = xxcz(l0/i,22.0,24.0,0.70,0.65);
		}
		else if(l0/i <= 26)
		{
			fi = xxcz(l0/i,24.0,26.0,0.65,0.60);
		}
		else if(l0/i <= 28)
		{
			fi = xxcz(l0/i,26.0,28.0,0.60,0.56);
		}
		else if(l0/i <= 30)
		{
			fi = xxcz(l0/i,28.0,30.0,0.56,0.52);
		}
		else if(l0/i <= 32)
		{
			fi = xxcz(l0/i,30.0,32.0,0.52,0.48);
		}
		else if(l0/i <= 34)
		{
			fi = xxcz(l0/i,32.0,34.0,0.48,0.44);
		}
		else if(l0/i <= 36)
		{
			fi = xxcz(l0/i,34.0,36.0,0.44,0.40);
		}
		else if(l0/i <= 38)
		{
			fi = xxcz(l0/i,36.0,38.0,0.40,0.36);
		}
		else if(l0/i <= 40)
		{
			fi = xxcz(l0/i,38.0,40.0,0.36,0.32);
		}
		else if(l0/i <= 42)
		{
			fi = xxcz(l0/i,40.0,42.0,0.32,0.29);
		}
		else if(l0/i <= 44)
		{
			fi = xxcz(l0/i,42.0,44.0,0.29,0.26);
		}
		else if(l0/i <= 46)
		{
			fi = xxcz(l0/i,44.0,46.0,0.26,0.23);
		}
		else if(l0/i <= 48)
		{
			fi = xxcz(l0/i,46.0,48.0,0.23,0.21);
		}
		else if(l0/i <= 50)
		{
			fi = xxcz(l0/i,48.0,50.0,0.21,0.19);
		}
	}
	else if(con == 2)//圆
	{
		if(l0/i <= 7)
		{
			fi = 1.0;
		}
		else if(l0/i <= 8.5)
		{
			fi = xxcz(l0/i,7.0,8.5,1.0,0.98);
		}
		else if(l0/i <= 10.5)
		{
			fi = xxcz(l0/i,8.5,10.5,0.98,0.95);
		}
		else if(l0/i <= 12)
		{
			fi = xxcz(l0/i,10.5,12.0,0.95,0.92);
		}
		else if(l0/i <= 14)
		{
			fi = xxcz(l0/i,12.0,14.0,0.92,0.87);
		}
		else if(l0/i <= 15.5)
		{
			fi = xxcz(l0/i,14.0,15.5,0.87,0.81);
		}
		else if(l0/i <= 17)
		{
			fi = xxcz(l0/i,15.5,17.0,0.81,0.75);
		}
		else if(l0/i <= 19)
		{
			fi = xxcz(l0/i,17.0,19.0,0.75,0.70);
		}
		else if(l0/i <= 21)
		{
			fi = xxcz(l0/i,19.0,20.8,0.70,0.65);
		}
		else if(l0/i <= 22.5)
		{
			fi = xxcz(l0/i,20.8,22.5,0.65,0.60);
		}
		else if(l0/i <= 24)
		{
			fi = xxcz(l0/i,22.5,24.0,0.60,0.56);
		}
		else if(l0/i <= 26)
		{
			fi = xxcz(l0/i,24.0,26.0,0.56,0.52);
		}
		else if(l0/i <= 28)
		{
			fi = xxcz(l0/i,26.0,28.0,0.52,0.48);
		}
		else if(l0/i <= 29.5)
		{
			fi = xxcz(l0/i,28.0,29.5,0.48,0.44);
		}
		else if(l0/i <= 31)
		{
			fi = xxcz(l0/i,29.5,30.8,0.44,0.40);
		}
		else if(l0/i <= 33)
		{
			fi = xxcz(l0/i,30.8,33.0,0.40,0.36);
		}
		else if(l0/i <= 34.5)
		{
			fi = xxcz(l0/i,33.0,34.5,0.36,0.32);
		}
		else if(l0/i <= 36.5)
		{
			fi = xxcz(l0/i,34.5,36.5,0.32,0.29);
		}
		else if(l0/i <= 38)
		{
			fi = xxcz(l0/i,36.5,38.0,0.29,0.26);
		}
		else if(l0/i <= 40)
		{
			fi = xxcz(l0/i,38.0,40.0,0.26,0.23);
		}
		else if(l0/i <= 41.5)
		{
			fi = xxcz(l0/i,40.0,41.5,0.23,0.21);
		}
		else if(l0/i <= 43)
		{
			fi = xxcz(l0/i,41.5,43.0,0.21,0.19);
		}
	}
}

void Cinformation::calpminp(CString Ctype,CString Stype)
{
	if(Stype == "HPB300"||Stype == "HRB335"||Stype == "HRBF335")
	{
		if(Ctype == "C65"||Ctype == "C70"||Ctype == "C75"||Ctype == "C80")
			pminp = 0.0070;
		else
			pminp = 0.0060;
	}
	else if(Stype == "HRB400"||Stype == "HRBF400"||Stype == "RRB400")
	{
		if(Ctype == "C65"||Ctype == "C70"||Ctype == "C75"||Ctype == "C80")
			pminp = 0.0065;
		else
			pminp = 0.0055;
	}
	else if(Stype == "HRB500"||Stype == "HRBF500")
	{
		if(Ctype == "C65"||Ctype == "C70"||Ctype == "C75"||Ctype == "C80")
			pminp = 0.0060;
		else
			pminp = 0.0050;
	}
}

void Cinformation::initZY(CString Ctype,CString Stype,int con,double l0,double i)
{
	this->CTtension(Ctype);
	this->CTpress(Ctype);
	this->ST(Stype);
	this->kesaib(Ctype,Stype);
	this->arfa1(Ctype);
	this->calfi(con,l0,i);
	this->calpminp(Ctype,Stype);
}

void Cinformation::calculayszya(CString Ctype)
{
	int num;
	Ctype.TrimLeft(_T("C"));
	num = _ttoi(Ctype);
	if(num <= 50)
		yszya = 1.0;
	else if(num >= 80)
		yszya = 0.85;
	else
		yszya = xxcz(num,50,80,1.0,0.85);
}

void Cinformation::initYSZY(CString Ctype,CString ZJtype,CString GJtype,double l0,double i,double gd)
{
	this->CTtension(Ctype);
	this->CTpress(Ctype);
	this->ST(ZJtype);
	this->GST(GJtype);
	this->Bc(Ctype);
	this->kesaib(Ctype,ZJtype);
	this->arfa1(Ctype);
	this->calfi(2,l0,i);
	this->calpminp(Ctype,ZJtype);
	this->calculayszya(Ctype);
	Assl = 0.25 * PI * gd * gd;
}

void Cinformation::initPY(CString Ctype,CString Stype)
{
	this->CTtension(Ctype);
	this->CTpress(Ctype);
	this->ST(Stype);
	this->kesaib(Ctype,Stype);
	this->arfa1(Ctype);
	this->calpminp(Ctype,Stype);
	this->calb1(Ctype);
}
void Cinformation::calb1(CString Ctype)
{
	if(Ctype == "C15")
		b1 = 0.8;
	else if(Ctype == "C20")
		b1 = 0.8;
	else if(Ctype == "C25")
		b1 = 0.8;
	else if(Ctype == "C30")
		b1 = 0.8;
	else if(Ctype == "C35")
		b1 = 0.8;
	else if(Ctype == "C40")
		b1 = 0.8;
	else if(Ctype == "C45")
		b1 = 0.8;
	else if(Ctype == "C50")
		b1 = 0.8;
	else if(Ctype == "C55")
		b1 = 0.79;
	else if(Ctype == "C60")
		b1 = 0.78;
	else if(Ctype == "C65")
		b1 = 0.77;
	else if(Ctype == "C70")
		b1 = 0.76;
	else if(Ctype == "C75")
		b1 = 0.75;
	else if(Ctype == "C80")
		b1 = 0.74;
}