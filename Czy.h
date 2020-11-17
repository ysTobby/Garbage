#pragma once
#include "Czyc.h"
#include "Czyd.h"
#include "Czyyd.h"
#include "Czyyc.h"


// Czy 对话框

class Czy : public CDialogEx
{
	DECLARE_DYNAMIC(Czy)

public:
	Czy(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Czy();

// 对话框数据
	enum { IDD = ID_ZXSY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	virtual BOOL OnInitDialog();
	CTabCtrl m_tab_zy;
	Czyc zycDlg;
	Czyd zydDlg;
	Czyyd zyydDlg;
	Czyyc zyycDlg;
	virtual void OnOK();
};
