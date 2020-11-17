#pragma once
#include "afxcmn.h"
#include "DBTabD1.h"
#include "DBTabD2.h"
#include "DBTabC.h"

// Cdb �Ի���

class Cdb : public CDialogEx
{
	DECLARE_DYNAMIC(Cdb)

public:
	Cdb(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~Cdb();

// �Ի�������
	enum { IDD = ID_DBDC };

protected:
	CDBTabD1 DD1Dlg;
	CDBTabD2 DD2Dlg;
	CDBTabC DCDlg;
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult);
	CTabCtrl m_tab_db;
	virtual void OnOK();
};
