#pragma once


// CDcpyd 对话框

class CDcpyd : public CDialogEx
{
	DECLARE_DYNAMIC(CDcpyd)

public:
	CDcpyd(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CDcpyd();

// 对话框数据
	enum { IDD = IDD_DCPYD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	afx_msg void OnBnClickedRadioRec();
	afx_msg void OnBnClickedRadioG();
	CString g_exePATH;
	afx_msg void OnBnClickedButtoncalcula();
	CString m_CT;
	double m_Asp;
	double m_bf;
	double m_chang;
	double m_gchang;
	double m_gkuan;
	double m_hf;
	double m_kuan;
	double m_l0;
	double m_M1;
	double m_M2;
	double m_N;
	double m_p;
	double m_pmin;
	double m_pas;
	CString m_ST;
	Cinformation info;
	double m_h0,m_i,m_A,M1,M2,M1cyM2,zitac,Cm,ea,yitans,M,e0,ei,e,Nub,x,ep,kesai;
};
