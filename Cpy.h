#pragma once
#include "Cpyc.h"
#include "Cpyd1.h"
#include "Cpyd2.h"


// Cpy �Ի���

class Cpy : public CDialogEx
{
	DECLARE_DYNAMIC(Cpy)

public:
	Cpy(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~Cpy();

// �Ի�������
	enum { IDD = ID_PXSY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
	virtual void OnOK();
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	Cpyd1 pyd1Dlg;
	Cpyd2 pyd2Dlg;
	CTabCtrl m_tab;
};
