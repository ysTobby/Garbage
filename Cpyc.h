#pragma once


// Cpyc �Ի���

class Cpyc : public CDialogEx
{
	DECLARE_DYNAMIC(Cpyc)

public:
	Cpyc(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~Cpyc();

// �Ի�������
	enum { IDD = IDD_PYC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	void reset();
	CString m_CS;
	CString m_ST;
	double m_as;
	double m_B;
	double m_e;
	double m_e0;
	double m_ei;
	double m_ep;
	double m_H;
	double m_iks;
	double m_mks;
	double m_Mu;
	double m_N;
	double m_Nb;
	CString m_PYT;
	afx_msg void OnBnClickedButtoncal();
	Cinformation info;
	double m_SAs;
	double m_SAsp;
	int exportcon;
	double x, tempkesai, h0,m_ea,m_kscy;
	afx_msg void OnBnClickedButtonexport();
	CString g_exePATH;
	virtual BOOL OnInitDialog();
};
