// Cpy.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "peijin.h"
#include "Cpy.h"
#include "afxdialogex.h"


// Cpy �Ի���

IMPLEMENT_DYNAMIC(Cpy, CDialogEx)

Cpy::Cpy(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cpy::IDD, pParent)
{

}

Cpy::~Cpy()
{
}

void Cpy::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB1, m_tab);
}


BEGIN_MESSAGE_MAP(Cpy, CDialogEx)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, &Cpy::OnTcnSelchangeTab1)
END_MESSAGE_MAP()


// Cpy ��Ϣ�������


void Cpy::OnOK()
{
	// TODO: �ڴ����ר�ô����/����û���
	MessageBox(_T("�������Ҫ����"));
	//CDialogEx::OnOK();
}


BOOL Cpy::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_tab.InsertItem(0,_T("��֪��ѹ�ֽ������ƫ����ѹ�������"),0);
	m_tab.InsertItem(1,_T("δ֪��ѹ�ֽ������ƫ����ѹ�������"),1);

	pyd1Dlg.Create(IDD_PYD1,&m_tab);//�˴�ע��Ҫ���ĳɶ�Ӧ�Ի����ID
	pyd2Dlg.Create(IDD_PYD2,&m_tab);

	pyd1Dlg.CenterWindow();
	pyd1Dlg.ShowWindow(SW_SHOW);

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}


void Cpy::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	int index = m_tab.GetCurSel();
	switch(index)
	{
		case 0:
		pyd1Dlg.CenterWindow();
		pyd1Dlg.ShowWindow(SW_SHOW);
		pyd2Dlg.ShowWindow(SW_HIDE);
		break;
		case 1:
		pyd2Dlg.CenterWindow();
		pyd1Dlg.ShowWindow(SW_HIDE);
		pyd2Dlg.ShowWindow(SW_SHOW);
		break;
	}
	*pResult = 0;
}
