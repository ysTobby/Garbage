#pragma once

// Czyd 对话框

class Czyd : public CDialogEx
{
	DECLARE_DYNAMIC(Czyd)

public:
	Czyd(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Czyd();

// 对话框数据
	enum { IDD = IDD_ZYD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

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
	int con2;//1,不满足最小配筋率,2满足正常，3超过5%，4大于3小于5
	CString g_exePATH;
};
