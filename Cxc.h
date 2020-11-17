#pragma once


// Cxc 对话框

class Cxc : public CDialogEx
{
	DECLARE_DYNAMIC(Cxc)

public:
	Cxc(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Cxc();

// 对话框数据
	enum { IDD = IDD_XC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedRadiojgb();
	afx_msg void OnBnClickedRadiojg();
	virtual BOOL OnInitDialog();
	afx_msg void OnCbnSelchangeCzxs();
	CString m_czxs;
	double m_a;
	double m_as;
	double m_b;
	CString m_CT;
	double m_gjjj;
	CString m_gjtp;
	CString m_gjzj;
	CString m_gjzs;
	double m_h,k;
	double m_namda;
	double acv;
	CString m_type;
	double m_v;
	double h0;
	double temp;
	Cinformation info;
	CString m_wqd;
	CString m_wqn;
	CString m_zjtp;
	afx_msg void OnCbnSelchangeSteeltype();
	virtual void OnOK();
	afx_msg void OnBnClickedReset();
	virtual void OnCancel();
	afx_msg void OnBnClickedCalcula();
	afx_msg void OnBnClickedOut();
private:
	CApplication app;
	CDocuments docs;
	CDocument0 doc;
	CParagraphs paragraphs;
	CParagraph paragraph;
	CFields fields;
	CRange range;
	CSelection sel;
	CFont0 font;
	void xb(CString center,CString x);
	void sb(CString center,CString s);
	void sxb(CString center,CString s,CString x);
	int con,com,con1,con2,con3;
};
