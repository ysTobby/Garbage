#pragma once


// Czyc 对话框

class Czyc : public CDialogEx
{
	DECLARE_DYNAMIC(Czyc)

public:
	Czyc(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Czyc();

// 对话框数据
	enum { IDD = IDD_ZYC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedRadiofang();
	afx_msg void OnBnClickedRadioyuan();
	virtual BOOL OnInitDialog();
	CString g_exePATH;
	double m_Asp;
	double m_chang;
	double m_H;
	double m_kuan;
	double m_N;
	double m_p;
	double m_pmin;
	double m_zhijing;
	CString m_ST;
	CString m_CT;
	virtual void OnOK();
	afx_msg void OnBnClickedButtoncalcula();
	double namda,m_l0,m_A;
	Cinformation info;
	CString temp;
	int con1;//1,fang,2yuan
	int con2;//1,不满足最小配筋率,2满足正常，3超过5%，4大于3小于5
	afx_msg void OnBnClickedOut();
};
