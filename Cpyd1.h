#pragma once


// Cpyd1 对话框

class Cpyd1 : public CDialogEx
{
	DECLARE_DYNAMIC(Cpyd1)

public:
	Cpyd1(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Cpyd1();

// 对话框数据
	enum { IDD = IDD_PYD1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtoncalcula();
//	afx_msg void OnBnClickedRadiofang();
//	afx_msg void OnBnClickedRadioyuan();
	virtual BOOL OnInitDialog();
	CString g_exePATH;
	CString m_CT;
	double m_Aspl;
	double m_Aspy;
	double m_chang;
	double m_H,m_l0;
	double m_kuan;
	double m_M1;
	double m_M2;
	double m_N;
	double m_as;
	double namda,m_h0,M1cyM2,m_i,m_A;
	//二阶效应
	double zitac,Cm,ea,yitans,M1,M2,M,e0,ei,e,x;
	CString m_ST;
	Cinformation info;
	double m_p;
	double m_pmin;
	virtual void OnOK();
};
