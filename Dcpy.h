#pragma once
#include "Dcpyc.h"
#include "Dcpyd.h"


// CDcpy 对话框

class CDcpy : public CDialogEx
{
	DECLARE_DYNAMIC(CDcpy)

public:
	CDcpy(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CDcpy();

// 对话框数据
	enum { IDD = ID_DCPY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	CTabCtrl m_tab;
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CDcpyd m_dcpyd;
	CDcpyc m_dcpyc;
};
