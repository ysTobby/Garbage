// ChooseDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "ChooseDlg.h"
#include "afxdialogex.h"
#include "Csb.h"
#include "Cdb.h"
#include "Ctb.h"
#include "XDC.h"
#include "Czy.h"
#include "Cpy.h"
#include "Dcpy.h"
#include "EJXY.h"


// CChooseDlg 对话框

IMPLEMENT_DYNAMIC(CChooseDlg, CDialogEx)

CChooseDlg::CChooseDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CChooseDlg::IDD, pParent)
{

}

CChooseDlg::~CChooseDlg()
{
}

void CChooseDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CChooseDlg, CDialogEx)
	ON_BN_CLICKED(IDCANCEL, &CChooseDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_SBDC, &CChooseDlg::OnBnClickedSbdc)
	ON_BN_CLICKED(IDC_DBDC, &CChooseDlg::OnBnClickedDbdc)
	ON_BN_CLICKED(IDC_DBDC2, &CChooseDlg::OnBnClickedDbdc2)
	ON_BN_CLICKED(IDC_DBDC3, &CChooseDlg::OnBnClickedDbdc3)
	ON_BN_CLICKED(IDC_SBDC2, &CChooseDlg::OnBnClickedSbdc2)
	ON_BN_CLICKED(IDC_SBDC3, &CChooseDlg::OnBnClickedSbdc3)
	ON_BN_CLICKED(IDC_SBDC4, &CChooseDlg::OnBnClickedSbdc4)
	ON_BN_CLICKED(IDC_SBDC5, &CChooseDlg::OnBnClickedSbdc5)
	ON_BN_CLICKED(IDC_SBDC6, &CChooseDlg::OnBnClickedSbdc6)
END_MESSAGE_MAP()


// CChooseDlg 消息处理程序


void CChooseDlg::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码
	
	this->EndDialog(0);
}


void CChooseDlg::OnBnClickedSbdc()
{
	// TODO: 在此添加控件通知处理程序代码
	Csb m_sb;
	m_sb.DoModal();
}


void CChooseDlg::OnBnClickedDbdc()
{
	Cdb m_db;
	m_db.DoModal();
	// TODO: 在此添加控件通知处理程序代码
}


void CChooseDlg::OnBnClickedDbdc2()
{
	// TODO: 在此添加控件通知处理程序代码
	Ctb m_tb;
	m_tb.DoModal();
}


void CChooseDlg::OnBnClickedDbdc3()
{
	CXDC m_xdc;
	m_xdc.DoModal();
}


void CChooseDlg::OnBnClickedSbdc2()
{
	Czy m_zy;
	m_zy.DoModal();
}


void CChooseDlg::OnBnClickedSbdc3()
{
	Cpy m_py;
	m_py.DoModal();
}


void CChooseDlg::OnBnClickedSbdc4()
{
	CDcpy m_dcpy;
	m_dcpy.DoModal();
}


void CChooseDlg::OnBnClickedSbdc5()
{
	CEJXY m_ejxy;
	m_ejxy.DoModal();
}


void CChooseDlg::OnBnClickedSbdc6()
{
	Cpyc m_pyc;
	m_pyc.DoModal();
}
