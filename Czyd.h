#pragma once

// Czyd �Ի���

class Czyd : public CDialogEx
{
	DECLARE_DYNAMIC(Czyd)

public:
	Czyd(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~Czyd();

// �Ի�������
	enum { IDD = IDD_ZYD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtoncalcula();
	afx_msg void OnBnClickedRadiofang();
	afx_msg void OnBnClickedRadioyuan();
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	double m_chang;
	double m_kuan;
	double m_zhijing;
	double m_H,m_l0;
	double m_N;
	double m_A,m_pp;
	CString m_ST;
	CString m_CT;
	Cinformation info;
	double m_Asp;
	double m_p;
	double m_pmin;
	afx_msg void OnBnClickedOut();
	double namda,m_Aspf,m_ppf;
	int con1;//1,fang,2yuan
	int con2;//1,��������С�����,2����������3����5%��4����3С��5
	CString g_exePATH;
};
