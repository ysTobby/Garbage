#pragma once

// Czyyc 对话框

class Czyyc : public CDialogEx
{
	DECLARE_DYNAMIC(Czyyc)

public:
	Czyyc(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Czyyc();

// 对话框数据
	enum { IDD = IDD_ZYYC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CString m_CT;
	double m_Asp;

	double m_d;
	double m_GD;
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
	double namda,l0,dcor,con,p,Ass0;//con=1不修正，con=2修正
	Cinformation info;
	int con1;//1,不考虑间接配筋；2考虑间接配筋
	int con2;//1,
	double N1,N2,N3;
	afx_msg void OnBnClickedOut();
};
