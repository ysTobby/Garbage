#pragma once


// Cpyd2 对话框

class Cpyd2 : public CDialogEx
{
	DECLARE_DYNAMIC(Cpyd2)

public:
	Cpyd2(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Cpyd2();

// 对话框数据
	enum { IDD = IDD_PYD2 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	CString g_exePATH;
	CString m_CT;
	double m_Aspl;
	double m_Aspy;
	double m_chang;
	double m_H;
	double m_kuan;
	double m_M1;
	double m_M2;
	double m_N;
	double m_p;
	double m_as;
	double m_pmin;
	CString m_ST;
	afx_msg void OnBnClickedButtoncalcula();
	double namda,m_l0,m_h0,m_i,m_A,M1,M2,M1cyM2;
	double zitac,Cm,ea,yitans,M,e0,ei,e,x,ep;
	Cinformation info;
	double yzAspy(double fAsy);
};
