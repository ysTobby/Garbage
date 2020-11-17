// Cdb.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "Cdb.h"
#include "afxdialogex.h"


// Cdb 对话框

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


// Cdb 消息处理程序


BOOL Cdb::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加您专用的创建代码
	m_tab_db.InsertItem(0,_T("未知As'设计"),0);
	m_tab_db.InsertItem(1,_T("已知As'设计"),1);
	m_tab_db.InsertItem(2,_T("复核"),2);

	DD1Dlg.Create(IDD_DD1,&m_tab_db);
	DD2Dlg.Create(IDD_DD2,&m_tab_db);
	DCDlg.Create(IDD_DC,&m_tab_db);

	DD1Dlg.CenterWindow();
	DD1Dlg.ShowWindow(SW_SHOW);

	return true;
}

void Cdb::OnTcnSelchangeTab1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: 在此添加控件通知处理程序代码

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
	// TODO: 在此添加专用代码和/或调用基类
	MessageBox(_T("请输入必要参数"));

	//CDialogEx::OnOK();
}
