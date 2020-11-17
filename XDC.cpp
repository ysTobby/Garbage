// XDC.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Peijin.h"
#include "XDC.h"
#include "afxdialogex.h"


// CXDC �Ի���

IMPLEMENT_DYNAMIC(CXDC, CDialogEx)

CXDC::CXDC(CWnd* pParent /*=NULL*/)
	: CDialogEx(CXDC::IDD, pParent)
{

}

CXDC::~CXDC()
{
}

void CXDC::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB1, m_tab_xj);
}


BEGIN_MESSAGE_MAP(CXDC, CDialogEx)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, &CXDC::OnTcnSelchangeTab1)
END_MESSAGE_MAP()


// CXDC ��Ϣ�������


BOOL CXDC::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_tab_xj.InsertItem(0,_T("б�����ܼ����"),0);
	m_tab_xj.InsertItem(1,_T("б�����ܼ�����"),1);


	xdDlg.Create(IDD_XD,&m_tab_xj);//�˴�ע��Ҫ���ĳɶ�Ӧ�Ի����ID
	xcDlg.Create(IDD_XC,&m_tab_xj);

	xdDlg.CenterWindow();
	xdDlg.ShowWindow(SW_SHOW);

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}


void CXDC::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	int index = m_tab_xj.GetCurSel();
	switch(index)
	{
	case 0:
		xdDlg.CenterWindow();
		xdDlg.ShowWindow(SW_SHOW);
		xcDlg.ShowWindow(SW_HIDE);
		break;
	case 1:
		xcDlg.CenterWindow();
		xdDlg.ShowWindow(SW_HIDE);
		xcDlg.ShowWindow(SW_SHOW);
		break;
	}

	*pResult = 0;
}


void CXDC::OnOK()
{
	// TODO: �ڴ����ר�ô����/����û���
	MessageBox(_T("�������Ҫ����"));
	//CDialogEx::OnOK();
}
