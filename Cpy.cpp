// Cpy.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Cpy.h"
#include "afxdialogex.h"


// Cpy 对话框

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


// Cpy 消息处理程序


void Cpy::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	MessageBox(_T("请输入必要参数"));
	//CDialogEx::OnOK();
}


BOOL Cpy::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_tab.InsertItem(0,_T("已知受压钢筋面积的偏心受压构件设计"),0);
	m_tab.InsertItem(1,_T("未知受压钢筋面积的偏心受压构件设计"),1);

	pyd1Dlg.Create(IDD_PYD1,&m_tab);//此处注意要更改成对应对话框的ID
	pyd2Dlg.Create(IDD_PYD2,&m_tab);

	pyd1Dlg.CenterWindow();
	pyd1Dlg.ShowWindow(SW_SHOW);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
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
