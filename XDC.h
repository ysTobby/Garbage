#pragma once
#include "Cxd.h"
#include "Cxc.h"

// CXDC �Ի���

class CXDC : public CDialogEx
{
	DECLARE_DYNAMIC(CXDC)

public:
	CXDC(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CXDC();

// �Ի�������
	enum { IDD = ID_XDC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	virtual void OnOK();
	Cxd xdDlg;
	Cxc xcDlg;
	CTabCtrl m_tab_xj;
};
