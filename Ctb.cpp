// Ctb.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "Ctb.h"
#include "afxdialogex.h"


// Ctb 对话框

IMPLEMENT_DYNAMIC(Ctb, CDialogEx)

Ctb::Ctb(CWnd* pParent /*=NULL*/)
	: CDialogEx(Ctb::IDD, pParent)
{

}

Ctb::~Ctb()
{
}

void Ctb::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB1, m_tab_tb);
}


BEGIN_MESSAGE_MAP(Ctb, CDialogEx)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, &Ctb::OnTcnSelchangeTab1)
END_MESSAGE_MAP()


// Ctb 消息处理程序


BOOL Ctb::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	m_tab_tb.InsertItem(0,_T("设计"),0);
	m_tab_tb.InsertItem(1,_T("复核"),1);

	TDdlg.Create(IDD_TD,&m_tab_tb);
	TCdlg.Create(IDD_TC,&m_tab_tb);

	TDdlg.CenterWindow();
	TDdlg.ShowWindow(SW_SHOW);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void Ctb::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	int index = m_tab_tb.GetCurSel();
	switch(index)
	{
	case 0:
		TDdlg.CenterWindow();
		TDdlg.ShowWindow(SW_SHOW);
		TCdlg.ShowWindow(SW_HIDE);
		break;
	case 1:
		TCdlg.CenterWindow();
		TDdlg.ShowWindow(SW_HIDE);
		TCdlg.ShowWindow(SW_SHOW);
		break;
	}

	*pResult = 0;
}


void Ctb::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	MessageBox(_T("请输入必要参数"));
	//CDialogEx::OnOK();
}
