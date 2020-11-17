#pragma once


// CTBTabd 对话框

class CTBTabd : public CDialogEx
{
	DECLARE_DYNAMIC(CTBTabd)

public:
	CTBTabd(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CTBTabd();

// 对话框数据
	enum { IDD = IDD_TD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
	virtual void OnOK();
public:
	afx_msg void OnBnClickedDd2Reset();
	afx_msg void OnBnClickedDd2Calcula();
	afx_msg void OnBnClickedOut();
	double m_as;
	double m_b;
	double m_bf;
	CString m_CS;
	double m_h;
	double m_hf;
	double m_M,m_Mu1,m_Mu2;
	double m_SAs,m_rAs;
	double m_SAsmin;
	CString m_ST;
	double m_x;
	double m_xb;
	double h0;
	double arfas;
	double kesai;
	double pmin,p,m_Mp;
	Cinformation info;
	int con;
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
};
