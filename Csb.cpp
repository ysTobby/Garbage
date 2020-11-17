// Csb.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Peijin.h"
#include "Csb.h"
#include "afxdialogex.h"


// Csb �Ի���

IMPLEMENT_DYNAMIC(Csb, CDialogEx)

Csb::Csb(CWnd* pParent /*=NULL*/)
	: CDialogEx(Csb::IDD, pParent)
{
}

Csb::~Csb()
{
}

void Csb::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB_SB, m_tab_sb);
}

BEGIN_MESSAGE_MAP(Csb, CDialogEx)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB_SB, &Csb::OnTcnSelchangeTab1)
	ON_WM_CREATE()
END_MESSAGE_MAP()


// Csb ��Ϣ�������


void Csb::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int index = m_tab_sb.GetCurSel();
	switch(index)
	{
	case 0:
		SDDlg.CenterWindow();
		SDDlg.ShowWindow(SW_SHOW);
		SCDlg.ShowWindow(SW_HIDE);
		break;
	case 1:
		SCDlg.CenterWindow();
		SCDlg.ShowWindow(SW_SHOW);
		SDDlg.ShowWindow(SW_HIDE);
		break;
	}

	*pResult = 0;
}


BOOL Csb::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ������ר�õĴ�������
	m_tab_sb.InsertItem(0,_T("���������"),0);
	m_tab_sb.InsertItem(1,_T("����������"),1);

	SDDlg.Create(IDD_SD,&m_tab_sb);
	SCDlg.Create(IDD_SC,&m_tab_sb);

	SDDlg.CenterWindow();
	SDDlg.ShowWindow(SW_SHOW);

	return true;
}


void Csb::OnOK()
{
	// TODO: �ڴ����ר�ô����/����û���
	MessageBox(_T("�������Ҫ����"));
	//CDialogEx::OnOK();
}
