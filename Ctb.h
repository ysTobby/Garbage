#pragma once
#include "TBTabc.h"
#include "TBTabd.h"
#include "afxcmn.h"

// Ctb �Ի���

class Ctb : public CDialogEx
{
	DECLARE_DYNAMIC(Ctb)

public:
	Ctb(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~Ctb();

// �Ի�������
	enum { IDD = ID_TBDC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	CTBTabc TCdlg;
	CTBTabd TDdlg;
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CTabCtrl m_tab_tb;
	virtual void OnOK();
};
