#pragma once


// CEJXY 对话框

class CEJXY : public CDialogEx
{
	DECLARE_DYNAMIC(CEJXY)

public:
	CEJXY(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CEJXY();

// 对话框数据
	enum { IDD = IDD_EJXY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	double m_as;
	double m_b;
	double m_fc;
	double m_h;
	double m_l0;
//	CEdit m_M1;
//	CEdit m_M2;
//	CEdit m_N;
//	CEdit m_r1;
//	CEdit m_r2;
//	CEdit m_r3;
//	CEdit m_r4;
	double m_M1;
	double m_M2;
	double m_N;
	double m_r1;
	double m_r2;
	double m_r3;
	double m_r4;
	double m_r5;
	double m_r6;
	double m_r7;
	double m_r8;
	virtual void OnOK();
};
