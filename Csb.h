#pragma once
#include "afxcmn.h"
#include "SBTabC.h"
#include "SBTabD.h"


// Csb 对话框

class Csb : public CDialogEx
{
	DECLARE_DYNAMIC(Csb)

public:
	Csb(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Csb();

// 对话框数据
	enum { IDD = IDD_SBDC };

protected:
	CSBTabD SDDlg;
	CSBTabC SCDlg;
	virtual BOOL OnInitDialog();

	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CTabCtrl m_tab_sb;
	
	virtual void OnOK();
};
