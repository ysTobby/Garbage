#pragma once


// CSBTabD �Ի���

class CSBTabD : public CDialogEx
{
	DECLARE_DYNAMIC(CSBTabD)

public:
	CSBTabD(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CSBTabD();

// �Ի�������
	enum { IDD = IDD_SD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

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
	Cinformation info;//���ݻ�����ǿ�Ⱥ͸ֽ������ȡ����
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
