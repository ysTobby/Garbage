#pragma once


// CDBTabD2 对话框

class CDBTabD2 : public CDialogEx
{
	DECLARE_DYNAMIC(CDBTabD2)

public:
	CDBTabD2(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CDBTabD2();

// 对话框数据
	enum { IDD = IDD_DD2 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
	virtual void OnOK();
public:
	double m_As;
	double m_b;
	CString m_CS;
	double m_h;
	double m_M;
//	double m_MAs;
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
	double h0,arfas,Mu1,Mu2,As1,As2,kesai;
	afx_msg void OnBnClickedDd2Scal();
	afx_msg void OnBnClickedDd2Calcula();
	afx_msg void OnBnClickedDd2Reset();
	double m_Asp;
	CString m_STp;
	afx_msg void OnBnClickedOut();
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
