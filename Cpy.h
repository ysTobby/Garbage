#pragma once
#include "Cpyc.h"
#include "Cpyd1.h"
#include "Cpyd2.h"


// Cpy 对话框

class Cpy : public CDialogEx
{
	DECLARE_DYNAMIC(Cpy)

public:
	Cpy(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~Cpy();

// 对话框数据
	enum { IDD = ID_PXSY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
	virtual void OnOK();
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	Cpyd1 pyd1Dlg;
	Cpyd2 pyd2Dlg;
	CTabCtrl m_tab;
};
