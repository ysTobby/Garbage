#pragma once
#include "TBTabc.h"
#include "TBTabd.h"
#include "afxcmn.h"

// Ctb 对话框

class Ctb : public CDialogEx
{
	DECLARE_DYNAMIC(Ctb)

public:
	Ctb(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Ctb();

// 对话框数据
	enum { IDD = ID_TBDC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	CTBTabc TCdlg;
	CTBTabd TDdlg;
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CTabCtrl m_tab_tb;
	virtual void OnOK();
};
