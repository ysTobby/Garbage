#pragma once


// CSBTabC �Ի���

class CSBTabC : public CDialogEx
{
	DECLARE_DYNAMIC(CSBTabC)

public:
	CSBTabC(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CSBTabC();

// �Ի�������
	enum { IDD = IDD_SC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	double m_As;
	double m_b;
	CString m_CS;
	double m_h;
//	double m_M;
	double m_MAs;
	double m_Mu;
	double m_SAs;
	CString m_ST;
	double m_x,rm_x;
	double m_xb;
	afx_msg void OnBnClickedSdReset();
	afx_msg void OnBnClickedSdCalcula();
	afx_msg void OnBnClickedScOut();
	double h0,p,pmin;
	Cinformation info;
	int control;
//	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnOK();
//	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	virtual void OnCancel();
};
