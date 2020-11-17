#pragma once


// CDBTabC 对话框

class CDBTabC : public CDialogEx
{
	DECLARE_DYNAMIC(CDBTabC)

public:
	CDBTabC(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CDBTabC();

// 对话框数据
	enum { IDD = IDD_DC };

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
	double m_SAs;
	double m_SAsp;
	CString m_ST;
	double m_x;
	double m_xb;
	afx_msg void OnBnClickedDd2Calcula();
	afx_msg void OnBnClickedDd2Reset();
	Cinformation info;
	Cinformation info2;
	double h0;
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
