#pragma once


// CDBTabD1 对话框

class CDBTabD1 : public CDialogEx
{
	DECLARE_DYNAMIC(CDBTabD1)

public:
	CDBTabD1(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CDBTabD1();

// 对话框数据
	enum { IDD = IDD_DD1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	double m_As;
	double m_M;
	double m_b;
	CString m_CS;
	double m_h;
	double m_MAs;
	double m_SAs;
	double m_SAsp;
	CString m_sdia;
	CString m_snum;
	double m_sres;
	CString m_ST;
	double m_x;
	double m_xb;
	Cinformation info;
	Cinformation info2;
	double h0,arfas,pmin,rAs;
	afx_msg void OnBnClickedDd1Scal();
	afx_msg void OnBnClickedDd1Calcula();
	virtual void OnOK();
	afx_msg void OnBnClickedDd1Reset();
	afx_msg void OnBnClickedOut();
	double m_Asp;
	CString m_STp;
	int con;
	virtual void OnCancel();
	void xb(CString center,CString x);
	void sb(CString center,CString s);
	void sxb(CString center,CString s,CString x);
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
};
