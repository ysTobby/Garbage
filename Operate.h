#pragma once
class COperate
{
public:
	COperate(void);
	~COperate(void);

	//分数开关
	static CString fraction(CString up,CString dn);
	static CString fraction(double up, double dn);

	//根号开关
	static CString radical(CString up,CString dn);
	static CString radical(CString dn);
	static CString radical(int num1, double num2);
	static CString radical(double num);

	//上标开关
	static CString sbkg(double num1, int num2);
	static CString sbkg(CString dn, int num2);
	static CString sbkg(CString dn, CString up);

	//下标开关
	static CString xbkg(CString up, int num2);
	static CString xbkg(CString up, CString dn);

	//上、下标开关
	static CString sxbkg(CString center, int num1, int num2);
	static CString sxbkg(CString center, CString up, CString dn);

	//单个字符转换
	static CString single(CString center);
	static CString single(double center);

	//数字转字符串
	static CString trans(double num);
	static CString trans(int num);
};

