#pragma once
#include "afxcmn.h"
#include "SBTabC.h"
#include "SBTabD.h"


// Csb �Ի���

class Csb : public CDialogEx
{
	DECLARE_DYNAMIC(Csb)

public:
	Csb(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~Csb();

// �Ի�������
	enum { IDD = IDD_SBDC };

protected:
	CSBTabD SDDlg;
	CSBTabC SCDlg;
	virtual BOOL OnInitDialog();

	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CTabCtrl m_tab_sb;
	
	virtual void OnOK();
};
