// Dcpyd.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Dcpyd.h"
#include "afxdialogex.h"


// CDcpyd 对话框

IMPLEMENT_DYNAMIC(CDcpyd, CDialogEx)

CDcpyd::CDcpyd(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDcpyd::IDD, pParent)
{

	m_CT = _T("C40");
	m_Asp = 0.0;
	m_bf = 0.0;
	m_chang = 700;
	m_gchang = 0.0;
	m_gkuan = 0.0;
	m_hf = 0.0;
	m_kuan = 400;
	m_l0 = 3.3;
	m_M1 = 350*0.88;
	m_M2 = 350;
	m_N = 3500;
	m_p = 0.0;
	m_pmin = 0.0;
	m_pas = 45;
	m_ST = _T("HRB400");
}

CDcpyd::~CDcpyd()
{
}

void CDcpyd::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_CBString(pDX, IDC_CT, m_CT);
	DDX_Text(pDX, IDC_EDIT_Aspl, m_Asp);
	DDX_Text(pDX, IDC_EDIT_bf, m_bf);
	DDX_Text(pDX, IDC_EDIT_chang, m_chang);
	DDX_Text(pDX, IDC_EDIT_gchang, m_gchang);
	DDX_Text(pDX, IDC_EDIT_gkuan, m_gkuan);
	DDX_Text(pDX, IDC_EDIT_hf, m_hf);
	DDX_Text(pDX, IDC_EDIT_kuan, m_kuan);
	DDX_Text(pDX, IDC_EDIT_l0, m_l0);
	DDX_Text(pDX, IDC_EDIT_M1, m_M1);
	DDX_Text(pDX, IDC_EDIT_M3, m_M2);
	DDX_Text(pDX, IDC_EDIT_N2, m_N);
	DDX_Text(pDX, IDC_EDIT_p, m_p);
	DDX_Text(pDX, IDC_EDIT_pmin, m_pmin);
	DDX_Text(pDX, IDC_EDIT_pas2, m_pas);
	DDX_CBString(pDX, IDC_ST, m_ST);
}


BEGIN_MESSAGE_MAP(CDcpyd, CDialogEx)
	ON_BN_CLICKED(IDC_RADIO_REC, &CDcpyd::OnBnClickedRadioRec)
	ON_BN_CLICKED(IDC_RADIO_G, &CDcpyd::OnBnClickedRadioG)
	ON_BN_CLICKED(IDC_BUTTON_calcula, &CDcpyd::OnBnClickedButtoncalcula)
END_MESSAGE_MAP()


// CDcpyd 消息处理程序


BOOL CDcpyd::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	((CButton *)GetDlgItem(IDC_RADIO_REC))->SetCheck(TRUE); //选上

	TCHAR path[MAX_PATH] = {0};

	GetModuleFileName(NULL,path,MAX_PATH);

	g_exePATH=path;//此时获得了EXE的目录

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void CDcpyd::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	OnBnClickedButtoncalcula();
	//CDialogEx::OnOK();
}


void CDcpyd::OnBnClickedRadioRec()
{
	// 矩形截面
	GetDlgItem(IDC_STATIC_chang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_chang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_kuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_kuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_gkuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_gkuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_gchang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_gchang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_bf)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_bf)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_hf)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_hf)->ShowWindow(SW_HIDE);
	
}


void CDcpyd::OnBnClickedRadioG()
{
	// 工字型截面
	GetDlgItem(IDC_STATIC_chang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_chang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_kuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_kuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_gkuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_gkuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_gchang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_gchang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_bf)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_bf)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_hf)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_hf)->ShowWindow(SW_SHOW);
}


void CDcpyd::OnBnClickedButtoncalcula()
{
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	CWnd::UpdateData(true);

	m_l0 = m_l0 * 1000;
	info.initPY(m_CT,m_ST);
	m_N *= 1000;
	if(((CButton *)GetDlgItem(IDC_RADIO_REC))->GetCheck())
	{
		//选中的是矩形截面
		m_h0 = m_chang - m_pas;
		m_i = 0.289 * m_chang;
		m_A = m_chang * m_kuan;

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
		e = ei + 0.5*m_chang - m_pas;
		Nub = info.a1 * info.cp * m_kuan * info.kb * m_h0;
		if(Nub > m_N && ei > (0.3 * m_h0))
		{
			//判定为大偏压
			x = m_N / (info.a1 * info.cp * m_kuan);
			if(x >= (2 * m_pas) && x <= info.kb * m_h0)
			{
				m_Asp = (m_N * e - info.a1 * info.cp * m_kuan * x * (m_h0 - 0.5 * x))/(info.ss * (m_h0 - m_pas));
			}
			else if(x < (2 * m_pas))
			{
				ep = ei - 0.5 * m_chang + m_pas;
				m_Asp = (m_N * ep)/(info.ss * (m_h0 - m_pas));
			}
		}
		else if(Nub > m_N && ei <= (0.3 * m_h0))
		{
			//构造配筋
			MessageBox(_T("构造配筋"));
			return;
		}
		else
		{
			//小偏压
			kesai = (m_N - info.kb*info.a1*info.cp*m_kuan*m_h0)/((m_N*e - 0.43*info.a1*info.cp*m_kuan*pow(m_h0,2))/((info.b1 - info.kb)*(m_h0-m_pas)) + info.a1*info.cp*m_kuan*m_h0) + info.kb;
			x = kesai * m_h0;
			m_Asp = (m_N * e - info.a1 * info.cp * m_kuan * x * (m_h0 - 0.5 * x))/(info.ss * (m_h0 - m_pas));
		}
	}
	if(((CButton *)GetDlgItem(IDC_RADIO_G))->GetCheck())
	{
		//选中的是工字型截面
		m_h0 = m_gchang - m_pas;
		m_A = m_gchang * m_gkuan + (m_bf - m_gkuan)*m_hf*2;
		m_i = sqrt((1/12*m_gkuan*pow((m_gchang - 2*m_hf),3) + 1/12*m_hf*pow(m_bf,3)+1/12*m_hf*pow(m_bf,3))/m_A);
		M1 = (m_M1>m_M2)?m_M2:m_M1;
		M1 *= 1000000;
		M2 = (m_M1<m_M2)?m_M2:m_M1;
		M2 *= 1000000;
		M1cyM2 = M1/M2;
		m_pmin = info.pminp * 100;

		ea = (20>(m_gchang/30))?20:(m_gchang/30);
		if(M1cyM2 > 0.9 || m_l0/m_i > (34-12*(M1cyM2)) || m_N/(info.cp * m_A) > 0.9)
		{
			//考虑二阶效应
			zitac = ((0.5*info.cp*m_A/m_N) > 1)?1:(0.5*info.cp*m_A/m_N);
			Cm = 0.7 + 0.3*M1cyM2;
			yitans = 1 + (m_h0/(1300*(M2/m_N +ea)))*(m_l0/m_gchang)*(m_l0/m_gchang)*zitac;
			M = ((Cm*yitans > 1.0)?(Cm*yitans):1.0) * M2;
		}
		else
			M = M2;
		e0 = M/m_N;
		ei = e0 + ea;
		e = ei + 0.5*m_gchang - m_pas;
		Nub = info.a1 * info.cp * m_gkuan * info.kb * m_h0 + info.a1 * info.cp * (m_bf - m_gkuan) * m_hf;
		if(m_N <= Nub && ei > (0.3*m_h0))
		{
			//大偏压
			x = m_N / (info.a1*info.cp*m_bf);
			if(x>=(2*m_pas) && x<=m_hf)
			{
				m_Asp = (m_N * e - info.a1 * info.cp * m_bf * x * (m_h0 - 0.5 * x))/(info.ss * (m_h0 - m_pas));
			}
			else if(x>m_hf)
			{
				x = (m_N - info.a1*info.cp*(m_bf-m_gkuan)*m_hf)/(info.a1*info.cp*m_gkuan);
				m_Asp = (m_N*e - info.a1*info.cp*m_gkuan*x*(m_h0-0.5*x)-info.a1*info.cp*(m_bf-m_gkuan)*m_hf*(m_h0-0.5*m_hf))/(info.ss*(m_h0-m_pas));
			}
			else
			{
				m_Asp = (m_N*(ei - 0.5*m_gchang + m_pas))/(info.ss*(m_h0 - m_pas));
			}
		}
		else if(m_N > Nub)
		{
			//小偏压
			kesai = (m_N - info.a1*info.cp*(m_bf-m_gkuan)*m_hf - info.kb*info.a1*info.cp*m_gkuan*m_h0)/((m_N*e - info.a1*info.cp*(m_bf - m_gkuan)*m_hf*(m_h0 - 0.5*m_hf) - 0.43*info.a1*info.cp*m_gkuan*pow(m_h0,2))/((info.b1 - info.kb)*(m_h0 - m_pas)) + info.a1*info.cp*m_gkuan*m_h0) + info.kb;
			x = kesai * m_h0;
			m_Asp = (m_N * e - info.a1 * info.cp * m_gkuan * x * (m_h0 - 0.5 * x) - info.a1*info.cp*(m_bf - m_gkuan)*m_hf*(m_h0 - 0.5*m_hf))/(info.ss * (m_h0 - m_pas));
		
			//验算垂直于弯矩平面的轴压承载力
		}
		else if(m_N <= Nub && ei <= (0.3*m_h0))
		{
			MessageBox(_T("按照构造配筋即可"));
		}
	}

	m_l0 /= 1000;
	m_N /= 1000;
	m_p = 2*100*m_Asp/m_A;
	if(m_p < m_pmin)
	{
		MessageBox(_T("计算得到的配筋率小于最小配筋率，按照构造配置钢筋"));
		m_p = m_pmin;
		m_Asp = m_p * m_A/200;
	}

	CWnd::UpdateData(false);
}
