#pragma once
#include "Cxd.h"
#include "Cxc.h"

// CXDC 对话框

class CXDC : public CDialogEx
{
	DECLARE_DYNAMIC(CXDC)

public:
	CXDC(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CXDC();

// 对话框数据
	enum { IDD = ID_XDC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	virtual void OnOK();
	Cxd xdDlg;
	Cxc xcDlg;
	CTabCtrl m_tab_xj;
};
