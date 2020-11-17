#include "StdAfx.h"
#include "Operate.h"


COperate::COperate(void)
{
}


COperate::~COperate(void)
{
}

CString COperate::fraction(CString up,CString dn)
{
	CString result;

	result = _T("\\f(") + up + _T(",") + dn + _T(")");
	return result;
}

CString COperate::fraction(double up, double dn)
{
	CString result, ups, dns;
	ups.Format(_T("%.3lf"),up);
	dns.Format(_T("%.3lf"),dn);
	result = fraction(ups,dns);

	return result;
}

CString COperate::radical(CString up,CString dn)
{
	CString result;

	result = _T("\\r(") + up + _T(",") + dn + _T(")");

	return result;
}

CString COperate::radical(int num1, double num2)
{
	CString result, ups, dns;
	ups.Format(_T("%d"),num1);
	dns.Format(_T("%.3lf"),num2);
	result = radical(ups,dns);

	return result;
}

CString COperate::radical(CString dn)
{
	dn = _T("\\r(") + dn + _T(")");
	return dn;
}

CString COperate::radical(double num)
{
	CString result;

	result.Format(_T("%.3lf"),num);
	result = radical(result);
	return result;
}

CString COperate::sbkg(double num1, int num2)
{
	CString result;
	result.Format(_T("%.1lf"),num1);
	result = sbkg(result,num2);

	return result;
}

CString COperate::sbkg(CString dn, int num2)
{
	CString ups;
	ups.Format(_T("%d"),num2);//⁰ ¹ ² ³ ⁴ ⁵ ⁶ ⁷ ⁸ ⁹

	switch(num2)
	{
	case 0:
		ups = _T("⁰");
		break;
	case 1:
		ups = _T("¹");
		break;
	case 2:
		ups = _T("²");
		break;
	case 3:
		ups = _T("³");
		break;
	case 4:
		ups = _T("⁴");
		break;
	case 5:
		ups = _T("⁵");
		break;
	case 6:
		ups = _T("⁶");
		break;
	case 7:
		ups = _T("⁷");
		break;
	case 8:
		ups = _T("⁸");
		break;
	case 9:
		ups = _T("⁹");
		break;
	case 10:
		ups = _T("¹⁰");
		break;
	}
	
	return sbkg(dn,ups);
}

CString COperate::sbkg(CString dn, CString up)
{
	//dn = _T("\\i\\vc\\") + dn + _T("\\in( ,") + ups + _T(", )");
	dn = dn + _T("\\s(") + up + _T(")");
	
	return dn;
}

CString COperate::xbkg(CString up, int num2)
{
	CString dn;//₀ ₁ ₂ ₃ ₄ ₅ ₆ ₇ ₈ ₉
	dn.Format(_T("%d"),num2);

	switch (num2)
	{
	case 0:
		dn = _T("₀");
		break;
	case 1:
		dn = _T("₁");
		break;

	}

	return xbkg(up,dn);
}

CString COperate::xbkg(CString up, CString dn)
{
	//up + _T("\\s\\do1(") + dn +_T(")");
	return up + _T("\\s\\do3(") + dn +_T(")");
}

CString COperate::sxbkg(CString center, int num1, int num2)
{
	CString up, dn;
	up.Format(_T("%d"),num1);
	dn.Format(_T("%d"),num2);

	switch (num2)
	{
	case 0:
		dn = _T("₀");
		break;
	case 1:
		dn = _T("₁");
		break;

	}

	return sxbkg(center,up,dn);
}

CString COperate::sxbkg(CString center, CString up, CString dn)
{
	//center = _T("\\i\\vc\\") + center + _T("\\in(") + dn + _T(",") + up + _T(", )");//采用积分的方式下、上、中后\s()
	center = center + _T("\\s(") + up + _T(",") + dn + _T(")");
	return center;
}

CString COperate::single(CString center)
{
	return xbkg(center,_T(" "));
}

CString COperate::single(double center)
{
	CString result;

	result.Format(_T("%.3lf"),center);

	return single(result);
}

CString COperate::trans(double num)
{
	CString result;
	result.Format(_T("%.3lf"),num);
	return result;
}

CString COperate::trans(int num)
{
	CString result;
	result.Format(_T("%d"),num);
	return result;
}