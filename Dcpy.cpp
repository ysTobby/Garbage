// Dcpy.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "peijin.h"
#include "Dcpy.h"
#include "afxdialogex.h"


// CDcpy �Ի���

IMPLEMENT_DYNAMIC(CDcpy, CDialogEx)

CDcpy::CDcpy(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDcpy::IDD, pParent)
{

}

CDcpy::~CDcpy()
{
}

void CDcpy::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB1, m_tab);
}


BEGIN_MESSAGE_MAP(CDcpy, CDialogEx)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, &CDcpy::OnTcnSelchangeTab1)
END_MESSAGE_MAP()


// CDcpy ��Ϣ�������


void CDcpy::OnOK()
{
	// TODO: �ڴ����ר�ô����/����û���
	MessageBox(_T("�������Ҫ����"));
	//CDialogEx::OnOK();
}


BOOL CDcpy::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	m_tab.InsertItem(0,_T("ƫ����ѹ�������"),0);
	//m_tab.InsertItem(1,_T("ƫ����ѹ��������"),1);

	m_dcpyd.Create(IDD_DCPYD,&m_tab);//�˴�ע��Ҫ���ĳɶ�Ӧ�Ի����ID
	//m_dcpyc.Create(IDD_DCPYC,&m_tab);

	m_dcpyd.CenterWindow();
	m_dcpyd.ShowWindow(SW_SHOW);

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}


void CDcpy::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	int index = m_tab.GetCurSel();
	/*
	switch(index)
	{
		case 0:
		m_dcpyd.CenterWindow();
		m_dcpyd.ShowWindow(SW_SHOW);
		m_dcpyc.ShowWindow(SW_HIDE);
		break;
		case 1:
		m_dcpyc.CenterWindow();
		m_dcpyd.ShowWindow(SW_HIDE);
		m_dcpyc.ShowWindow(SW_SHOW);
		break;
	}*/

	*pResult = 0;
}
