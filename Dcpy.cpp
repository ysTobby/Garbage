// Dcpy.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Dcpy.h"
#include "afxdialogex.h"


// CDcpy 对话框

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


// CDcpy 消息处理程序


void CDcpy::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	MessageBox(_T("请输入必要参数"));
	//CDialogEx::OnOK();
}


BOOL CDcpy::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	m_tab.InsertItem(0,_T("偏心受压构件设计"),0);
	//m_tab.InsertItem(1,_T("偏心受压构件复核"),1);

	m_dcpyd.Create(IDD_DCPYD,&m_tab);//此处注意要更改成对应对话框的ID
	//m_dcpyc.Create(IDD_DCPYC,&m_tab);

	m_dcpyd.CenterWindow();
	m_dcpyd.ShowWindow(SW_SHOW);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void CDcpy::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: 在此添加控件通知处理程序代码
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
