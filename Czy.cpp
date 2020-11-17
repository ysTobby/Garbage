// Czy.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "peijin.h"
#include "Czy.h"
#include "afxdialogex.h"


// Czy �Ի���

IMPLEMENT_DYNAMIC(Czy, CDialogEx)

Czy::Czy(CWnd* pParent /*=NULL*/)
	: CDialogEx(Czy::IDD, pParent)
{

}

Czy::~Czy()
{
}

void Czy::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB1, m_tab_zy);
}


BEGIN_MESSAGE_MAP(Czy, CDialogEx)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, &Czy::OnTcnSelchangeTab1)
END_MESSAGE_MAP()


// Czy ��Ϣ�������


void Czy::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	int index = m_tab_zy.GetCurSel();
	switch(index)
	{
	case 0:
		zydDlg.CenterWindow();
		zydDlg.ShowWindow(SW_SHOW);
		zycDlg.ShowWindow(SW_HIDE);
		zyydDlg.ShowWindow(SW_HIDE);
		zyycDlg.ShowWindow(SW_HIDE);
		break;
	case 1:
		zycDlg.CenterWindow();
		zydDlg.ShowWindow(SW_HIDE);
		zycDlg.ShowWindow(SW_SHOW);
		zyydDlg.ShowWindow(SW_HIDE);
		zyycDlg.ShowWindow(SW_HIDE);
		break;
	case 2:
		zyydDlg.CenterWindow();
		zydDlg.ShowWindow(SW_HIDE);
		zycDlg.ShowWindow(SW_HIDE);
		zyydDlg.ShowWindow(SW_SHOW);
		zyycDlg.ShowWindow(SW_HIDE);
		break;
	case 3:
		zyycDlg.CenterWindow();
		zydDlg.ShowWindow(SW_HIDE);
		zycDlg.ShowWindow(SW_HIDE);
		zyydDlg.ShowWindow(SW_HIDE);
		zyycDlg.ShowWindow(SW_SHOW);
		break;
	}

	*pResult = 0;
}


BOOL Czy::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_tab_zy.InsertItem(0,_T("��ͨ��������ѹ���"),0);
	m_tab_zy.InsertItem(1,_T("��ͨ��������ѹ����"),1);
	m_tab_zy.InsertItem(2,_T("�����ι�������ѹ���"),2);
	m_tab_zy.InsertItem(3,_T("�����ι�������ѹ����"),3);


	zycDlg.Create(IDD_ZYC,&m_tab_zy);//�˴�ע��Ҫ���ĳɶ�Ӧ�Ի����ID
	zydDlg.Create(IDD_ZYD,&m_tab_zy);
	zyydDlg.Create(IDD_ZYYD,&m_tab_zy);
	zyycDlg.Create(IDD_ZYYC,&m_tab_zy);

	zydDlg.CenterWindow();
	zydDlg.ShowWindow(SW_SHOW);

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}


void Czy::OnOK()
{
	MessageBox(_T("�������Ҫ����"));

	//CDialogEx::OnOK();
}
