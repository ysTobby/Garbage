// Cpyd2.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Cpyd2.h"
#include "afxdialogex.h"


// Cpyd2 对话框

IMPLEMENT_DYNAMIC(Cpyd2, CDialogEx)

Cpyd2::Cpyd2(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cpyd2::IDD, pParent)
{

	m_CT = _T("C30");
	m_Aspl = 0.0;
	m_Aspy = 0.0;
	m_chang = 400;
	m_H = 2.4;
	m_kuan = 300;
	m_M1 = 200.56;
	m_M2 = 218;
	m_N = 396;
	m_p = 0.0;
	m_as = 40;
	m_pmin = 0.0;
	m_ST = _T("HRB400");
}

Cpyd2::~Cpyd2()
{
}

void Cpyd2::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_CBString(pDX, IDC_CT, m_CT);
	DDX_Text(pDX, IDC_EDIT_Aspl, m_Aspl);
	DDX_Text(pDX, IDC_EDIT_Aspy, m_Aspy);
	DDX_Text(pDX, IDC_EDIT_chang, m_chang);
	DDX_Text(pDX, IDC_EDIT_H, m_H);
	DDX_Text(pDX, IDC_EDIT_kuan, m_kuan);
	DDX_Text(pDX, IDC_EDIT_M1, m_M1);
	DDX_Text(pDX, IDC_EDIT_M2, m_M2);
	DDX_Text(pDX, IDC_EDIT_N, m_N);
	DDX_Text(pDX, IDC_EDIT_p, m_p);
	DDX_Text(pDX, IDC_EDIT_pas, m_as);
	DDX_Text(pDX, IDC_EDIT_pmin, m_pmin);
	DDX_CBString(pDX, IDC_ST, m_ST);
}


BEGIN_MESSAGE_MAP(Cpyd2, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_calcula, &Cpyd2::OnBnClickedButtoncalcula)
END_MESSAGE_MAP()


// Cpyd2 消息处理程序


BOOL Cpyd2::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	((CButton *)GetDlgItem(IDC_RADIO_TJ))->SetCheck(TRUE); //选上

	TCHAR path[MAX_PATH] = {0};

	GetModuleFileName(NULL,path,MAX_PATH);


	g_exePATH=path;//此时获得了EXE的目录

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void Cpyd2::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnOK();
}

double Cpyd2::yzAspy(double fAsy)
{
	if(ei > (0.3 * m_h0))
	{
		//初步判定为大偏压
		e = ei + 0.5*m_chang - m_as;
		x = m_h0 - sqrt(pow(m_h0,2)-2*(m_N*e-info.ss*fAsy*(m_h0-m_as))/(info.a1*info.cp*m_kuan));
		if(x > (2*m_as)&&x < (info.kb*m_h0))
		{
			m_Aspl = (info.a1*info.cp*m_kuan*x+info.ss*fAsy-m_N)/info.ss;
			m_p = (m_Aspl+fAsy)/(m_chang*m_kuan)*100;
		}
		else if(x < (2*m_as))
		{
			m_Aspl = (m_N*(ei - 0.5*m_chang + m_as))/(info.ss*(m_h0 - m_as));
			m_p = (m_Aspl+fAsy)/(m_chang*m_kuan)*100;
		}
		else
		{
			//小偏压
			MessageBox(_T("按照小偏压进行计算"));
			double temp = (m_N*e - info.ss*fAsy*(m_h0-m_as))/(info.a1*info.cp*m_kuan);
			x = m_h0 - sqrt(pow(m_h0,2) - 2*temp);
			double sigmas = (x/m_h0 - info.b1)/(info.kb - info.b1)*info.ss;
			m_Aspl = (info.a1*info.cp*m_kuan*x + info.ss*fAsy - m_N)/sigmas;
		}
	}
	else
	{
		//小偏压
		MessageBox(_T("按照小偏压进行计算"));
		double temp = (m_N*e - info.ss*fAsy*(m_h0-m_as))/(info.a1*info.cp*m_kuan);
		x = m_h0 - sqrt(pow(m_h0,2) - 2*temp);
		double sigmas = (x/m_h0 - info.b1)/(info.kb - info.b1)*info.ss;
		m_Aspl = (info.a1*info.cp*m_kuan*x + info.ss*fAsy - m_N)/sigmas;
	}
	return m_Aspl;
}

void Cpyd2::OnBnClickedButtoncalcula()
{
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	CWnd::UpdateData(true);

	if(((CButton *)GetDlgItem(IDC_RADIO_TJ))->GetCheck())
	{
		//两铰
		namda = 1.0;
	}
	else if(((CButton *)GetDlgItem(IDC_RADIO_TG))->GetCheck())
	{
		//两固定
		namda = 0.5;
	}
	else if(((CButton *)GetDlgItem(IDC_RADIO_YGYJ))->GetCheck())
	{
		//一固定一铰支
		namda = 0.7;
	}
	else if(((CButton *)GetDlgItem(IDC_RADIO_YGYZ))->GetCheck())
	{
		//一固定一自由
		namda = 2;
	}
	m_l0 = m_H * namda * 1000;
	m_h0 = m_chang - m_as;
	m_i = 0.289 * m_chang;
	m_A = m_chang * m_kuan;
	info.initPY(m_CT,m_ST);
	m_N *= 1000;
	M1 = (m_M1>m_M2)?m_M2:m_M1;
	M1 *= 1000000;
	M2 = (m_M1<m_M2)?m_M2:m_M1;
	M2 *= 1000000;
	M1cyM2 = M1/M2;
	m_pmin = info.pminp * 100;

	ea = (20>(m_chang/30))?20:(m_chang/30);
	if(M1cyM2 > 0.9 || m_l0/m_i > (34-12*(M1cyM2)) || m_N/(info.cp * m_A) > 0.9)
	{
		//考虑二阶效应
		zitac = ((0.5*info.cp*m_A/m_N) > 1)?1:(0.5*info.cp*m_A/m_N);
		Cm = 0.7 + 0.3*M1cyM2;
		yitans = 1 + (m_h0/(1300*(M2/m_N +ea)))*(m_l0/m_chang)*(m_l0/m_chang)*zitac;
		M = ((Cm*yitans > 1.0)?(Cm*yitans):1.0) * M2;
	}
	else
		M = M2;
	e0 = M/m_N;
	ei = e0 + ea;

	if(ei > (0.3 * m_h0))
	{
		//初步判定为大偏压
		e = ei + 0.5*m_chang - m_as;
		m_Aspy = (m_N*e-info.a1*info.cp*m_kuan*pow(m_h0,2)*info.kb*(1-0.5*info.kb))/(info.ss*(m_h0-m_as));
		if(m_Aspy < 0.002*m_chang*m_kuan)
		{
			m_Aspy = 0.002*m_chang*m_kuan;
			m_Aspl = yzAspy(m_Aspy);//当做受压已知来算
		}
		else
		{
			m_Aspl = (info.a1*info.cp*m_kuan*m_h0*info.kb + info.ss*m_Aspy - m_N)/info.ss;
		}
	}
	else
	{
		ep = 0.5*m_chang - m_as - (e0 - ea);
		if(m_N <= info.cp*m_kuan*m_chang)
		{
			m_Aspl = 0.002*m_A;
		}
		else
		{
			m_Aspl = (m_N*ep - info.a1*info.cp*m_A*(m_h0 - 0.5*m_chang))/(info.ss*(m_h0 - m_as));
			m_Aspl = (m_Aspl < (0.002*m_A))?(0.002*m_A):m_Aspl;
			double u = m_as/m_h0 + (info.ss*m_Aspl*(1-m_as/m_h0))/((info.kb - info.b1)*info.a1*info.cp*m_kuan*m_h0);
			double v = 2*m_N*ep/(info.a1*info.cp*m_kuan*pow(m_h0,2)) - 2*info.b1*info.ss*m_Aspl*(1-m_as/m_h0)/((info.kb - info.b1)*info.a1*info.cp*m_kuan*m_h0);
			double ksai = u+sqrt(pow(u,2)+v);
			double ksaicy = 2*info.b1 - info.kb;
			if(ksai > info.kb && ksai < ksaicy)
			{
				m_Aspy = (m_N + m_Aspl*info.ss*(ksai - info.b1)/(info.kb - info.b1) - info.a1*info.cp*m_kuan*ksai*m_h0) / info.ss;
			}
			else if(ksai >= ksaicy && ksai < (m_chang/m_h0))
			{
				ksai = m_as / m_h0 + sqrt(pow((m_as / m_h0),2) + 2*((m_N*ep)/(info.a1*info.cp*m_kuan*pow(m_h0,2)) - (m_Aspl*info.ss*(1 - m_as / m_h0))/(info.a1*m_kuan*m_h0*info.cp)));
				m_Aspy = (m_N + m_Aspl*info.ss*(ksai - info.b1)/(info.kb - info.b1) - info.a1*info.cp*m_kuan*ksai*m_h0) / info.ss;
			}
			else
			{
				m_Aspy = (m_N - info.ss*m_Aspl - info.a1*info.cp*m_A)/(info.ss);
			}
			if(m_Aspy < 0.002*m_A)
			{
				MessageBox(_T("计算得受压配筋面积小于最小配筋面积取为最小配筋面积"));
				m_Aspy = 0.002*m_A;
			}
		}
	}

	m_p = (m_Aspl+m_Aspy)/(m_chang*m_kuan)*100;
	if(m_p<m_pmin)
	{
		//按照比例放大
		double Ast1,Ast2;
		MessageBox(_T("计算配筋率小于最小配筋率，取最小配筋率,并进行比例放大"));
		m_p = m_pmin;
		Ast1 = (m_p/100*m_A) * (m_Aspl) / (m_Aspl+m_Aspy);
		Ast2 = (m_p/100*m_A) * (m_Aspy) / (m_Aspl+m_Aspy);
		m_Aspl = Ast1;
		m_Aspy = Ast2;
	}
	m_N /= 1000;
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
	CWnd::UpdateData(false);
}
