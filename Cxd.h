#pragma once


// Cxd 对话框

class Cxd : public CDialogEx
{
	DECLARE_DYNAMIC(Cxd)

public:
	Cxd(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Cxd();

// 对话框数据
	enum { IDD = IDD_XD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CString m_type;
	afx_msg void OnCbnSelchangeCombo1();
	virtual BOOL OnInitDialog();
	afx_msg void OnCbnSelchangeCzxs();
	CString m_czxs;
	virtual void OnOK();
	double m_a;
	double m_as;
	double m_b;
	CString m_CT;
	CString m_gjtp;
	CString m_gjzj;
	CString m_gjzs;
	double m_h;
	double m_result,m_realresult;
	double m_v;
	double h0,temp;
	CString m_wqd;
	CString m_wqn;
	CString m_zjtp;
	Cinformation info;
	afx_msg void OnBnClickedCalcula();
	afx_msg void OnBnClickedReset();
	afx_msg void OnBnClickedOut();
	afx_msg void OnBnClickedCalculalength();
	double m_gjjj;
	afx_msg void OnBnClickedRadiojg();
	afx_msg void OnBnClickedRadiojgb();
	double m_namda;
	virtual void OnCancel();
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
	double acv;
	int con,com,con1,con2,con3;
	double temp1;
	CString s_gjjjjmax;
};
