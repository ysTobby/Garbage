#pragma once


// Czyyd �Ի���

class Czyyd : public CDialogEx
{
	DECLARE_DYNAMIC(Czyyd)

public:
	Czyyd(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~Czyyd();

// �Ի�������
	enum { IDD = IDD_ZYYD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	CString m_CT;
	double m_As;
	double m_d;
	double m_Gd;
	double m_H;
	double m_N;
	double m_pas;
	double m_s;
	CString m_GST;
	CString m_ST;
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedCalcula();
	CString g_exePATH;
	double namda,m_l0,A,p,dcor,Acor,Ass0;
	CString temp;
	Cinformation info;
	afx_msg void OnBnClickedOut();
	double m_sp,m_fAss0,m_Nu1,m_Nu2;
	int con1;//���� 3%Ϊ1����Ϊ2
};
