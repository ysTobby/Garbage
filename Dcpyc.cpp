// Dcpyc.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Dcpyc.h"
#include "afxdialogex.h"


// CDcpyc 对话框

IMPLEMENT_DYNAMIC(CDcpyc, CDialogEx)

CDcpyc::CDcpyc(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDcpyc::IDD, pParent)
{

}

CDcpyc::~CDcpyc()
{
}

void CDcpyc::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CDcpyc, CDialogEx)
END_MESSAGE_MAP()


// CDcpyc 消息处理程序
