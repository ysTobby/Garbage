// Cdb.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Peijin.h"
#include "Cdb.h"
#include "afxdialogex.h"


// Cdb �Ի���

IMPLEMENT_DYNAMIC(Cdb, CDialogEx)

Cdb::Cdb(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cdb::IDD, pParent)
{

}

Cdb::~Cdb()
{
}

void Cdb::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB1, m_tab_db);
}


BEGIN_MESSAGE_MAP(Cdb, CDialogEx)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, &Cdb::OnTcnSelchangeTab1)
END_MESSAGE_MAP()


// Cdb ��Ϣ�������


BOOL Cdb::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ������ר�õĴ�������
	m_tab_db.InsertItem(0,_T("δ֪As'���"),0);
	m_tab_db.InsertItem(1,_T("��֪As'���"),1);
	m_tab_db.InsertItem(2,_T("����"),2);

	DD1Dlg.Create(IDD_DD1,&m_tab_db);
	DD2Dlg.Create(IDD_DD2,&m_tab_db);
	DCDlg.Create(IDD_DC,&m_tab_db);

	DD1Dlg.CenterWindow();
	DD1Dlg.ShowWindow(SW_SHOW);

	return true;
}

void Cdb::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������

	int index = m_tab_db.GetCurSel();
	switch(index)
	{
	case 0:
		DD1Dlg.CenterWindow();
		DD1Dlg.ShowWindow(SW_SHOW);
		DD2Dlg.ShowWindow(SW_HIDE);
		DCDlg.ShowWindow(SW_HIDE);
		break;
	case 1:
		DD2Dlg.CenterWindow();
		DD1Dlg.ShowWindow(SW_HIDE);
		DD2Dlg.ShowWindow(SW_SHOW);
		DCDlg.ShowWindow(SW_HIDE);
		break;
	case 2:
		DCDlg.CenterWindow();
		DD1Dlg.ShowWindow(SW_HIDE);
		DD2Dlg.ShowWindow(SW_HIDE);
		DCDlg.ShowWindow(SW_SHOW);
		break;
	}

	*pResult = 0;
}


void Cdb::OnOK()
{
	// TODO: �ڴ����ר�ô����/����û���
	MessageBox(_T("�������Ҫ����"));

	//CDialogEx::OnOK();
}
