#pragma once
class COperate
{
public:
	COperate(void);
	~COperate(void);

	//��������
	static CString fraction(CString up,CString dn);
	static CString fraction(double up, double dn);

	//���ſ���
	static CString radical(CString up,CString dn);
	static CString radical(CString dn);
	static CString radical(int num1, double num2);
	static CString radical(double num);

	//�ϱ꿪��
	static CString sbkg(double num1, int num2);
	static CString sbkg(CString dn, int num2);
	static CString sbkg(CString dn, CString up);

	//�±꿪��
	static CString xbkg(CString up, int num2);
	static CString xbkg(CString up, CString dn);

	//�ϡ��±꿪��
	static CString sxbkg(CString center, int num1, int num2);
	static CString sxbkg(CString center, CString up, CString dn);

	//�����ַ�ת��
	static CString single(CString center);
	static CString single(double center);

	//����ת�ַ���
	static CString trans(double num);
	static CString trans(int num);
};

