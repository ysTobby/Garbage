#pragma once


// CSBTabD 对话框

class CSBTabD : public CDialogEx
{
	DECLARE_DYNAMIC(CSBTabD)

public:
	CSBTabD(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CSBTabD();

// 对话框数据
	enum { IDD = IDD_SD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	double m_M;
	double m_As;
	double m_b;
	CString m_CS;
	double m_h;
	double m_MAs;
	double m_SAs;
	CString m_ST;
	double m_x;
	double m_xb;
	Cinformation info;//根据混凝土强度和钢筋种类获取参数
	afx_msg void OnBnClickedSdCalcula();
	afx_msg void OnBnClickedSdReset();
	afx_msg void OnBnClickedOut();

	int con;
	double h0,arfas,pmin,rAs;
	double m_sres;
	afx_msg void OnBnClickedSdScal();
	CString m_snum;
	CString m_sdia;
//	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnOK();
	virtual void OnCancel();
};
