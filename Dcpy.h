#pragma once
#include "Dcpyc.h"
#include "Dcpyd.h"


// CDcpy �Ի���

class CDcpy : public CDialogEx
{
	DECLARE_DYNAMIC(CDcpy)

public:
	CDcpy(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CDcpy();

// �Ի�������
	enum { IDD = ID_DCPY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	CTabCtrl m_tab;
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CDcpyd m_dcpyd;
	CDcpyc m_dcpyc;
};
