// Cpyd1.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "peijin.h"
#include "Cpyd1.h"
#include "afxdialogex.h"


// Cpyd1 �Ի���

IMPLEMENT_DYNAMIC(Cpyd1, CDialogEx)

Cpyd1::Cpyd1(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cpyd1::IDD, pParent)
{

	m_CT = _T("C30");
	m_Aspl = 0.0;
	m_Aspy = 1520;
	m_chang = 500;
	m_H = 6;
	m_kuan = 300;
	m_M1 = 250.9;
	m_M2 = 250.9;
	m_N = 160;
	m_as = 40;
	m_ST = _T("HRB400");
	m_p = 0.0;
	m_pmin = 0.0;
}

Cpyd1::~Cpyd1()
{
}

void Cpyd1::DoDataExchange(CDataExchange* pDX)
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
	DDX_Text(pDX, IDC_EDIT_pas, m_as);
	DDX_CBString(pDX, IDC_ST, m_ST);
	DDX_Text(pDX, IDC_EDIT_p, m_p);
	DDX_Text(pDX, IDC_EDIT_pmin, m_pmin);
}


BEGIN_MESSAGE_MAP(Cpyd1, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_calcula, &Cpyd1::OnBnClickedButtoncalcula)
END_MESSAGE_MAP()


// Cpyd1 ��Ϣ�������

void Cpyd1::OnBnClickedButtoncalcula()
{
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	CWnd::UpdateData(true);

	if(((CButton *)GetDlgItem(IDC_RADIO_TJ))->GetCheck())
	{
		//����
		namda = 1.0;
	}
	else if(((CButton *)GetDlgItem(IDC_RADIO_TG))->GetCheck())
	{
		//���̶�
		namda = 0.5;
	}
	else if(((CButton *)GetDlgItem(IDC_RADIO_YGYJ))->GetCheck())
	{
		//һ�̶�һ��֧
		namda = 0.7;
	}
	else if(((CButton *)GetDlgItem(IDC_RADIO_YGYZ))->GetCheck())
	{
		//һ�̶�һ����
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
		//���Ƕ���ЧӦ
		zitac = ((0.5*info.cp*m_A/m_N) > 1)?1:(0.5*info.cp*m_A/m_N);
		Cm = 0.7 + 0.3*M1cyM2;
		yitans = 1 + (m_h0/(1300*(M2/m_N +ea)))*(m_l0/m_chang)*(m_l0/m_chang)*zitac;
		M = ((Cm*yitans > 1.0)?(Cm*yitans):1.0) * M2;
	}
	else
		M = M2;
	e0 = M/m_N;
	ei = e0 + ea;
	e = ei + 0.5*m_chang - m_as;
	double temp1 = m_Aspy;
	if(ei > (0.3 * m_h0))
	{
		//�����ж�Ϊ��ƫѹ
		x = m_h0 - sqrt(pow(m_h0,2)-2*(m_N*e-info.ss*m_Aspy*(m_h0-m_as))/(info.a1*info.cp*m_kuan));
		if(x > (2*m_as)&&x < (info.kb*m_h0))
		{
			m_Aspl = (info.a1*info.cp*m_kuan*x+info.ss*m_Aspy-m_N)/info.ss;
		}
		else if(x < (2*m_as))
		{
			m_Aspl = (m_N*(ei - 0.5*m_chang + m_as))/(info.ss*(m_h0 - m_as));
		}
		else
		{
			//
			MessageBox(_T("����Сƫѹ���м���"));
			double temp = (m_N*e - info.ss*m_Aspy*(m_h0-m_as))/(info.a1*info.cp*m_kuan);
			x = m_h0 - sqrt(pow(m_h0,2) - 2*temp);
			double sigmas = (x/m_h0 - info.b1)/(info.kb - info.b1)*info.ss;
			m_Aspl = (info.a1*info.cp*m_kuan*x + info.ss*m_Aspy - m_N)/sigmas;
		}
	}
	else
	{
		//Сƫѹ
		MessageBox(_T("����Сƫѹ���м���"));
		double temp = (m_N*e - info.ss*m_Aspy*(m_h0-m_as))/(info.a1*info.cp*m_kuan);
		x = m_h0 - sqrt(pow(m_h0,2) - 2*temp);
		double sigmas = (x/m_h0 - info.b1)/(info.kb - info.b1)*info.ss;
		m_Aspl = (info.a1*info.cp*m_kuan*x + info.ss*m_Aspy - m_N)/sigmas;
	}

	//����Ϊ0���������
	m_Aspy = 0;
	double tm_Aspl;
	if(ei > (0.3 * m_h0))
	{
		//�����ж�Ϊ��ƫѹ
		x = m_h0 - sqrt(pow(m_h0,2)-2*(m_N*e-info.ss*m_Aspy*(m_h0-m_as))/(info.a1*info.cp*m_kuan));
		if(x > (2*m_as)&&x < (info.kb*m_h0))
		{
			tm_Aspl = (info.a1*info.cp*m_kuan*x+info.ss*m_Aspy-m_N)/info.ss;
		}
		else if(x < (2*m_as))
		{
			tm_Aspl = (m_N*(ei - 0.5*m_chang + m_as))/(info.ss*(m_h0 - m_as));
		}
		else
		{
			//
			MessageBox(_T("����Сƫѹ���м���"));
			double temp = (m_N*e - info.ss*m_Aspy*(m_h0-m_as))/(info.a1*info.cp*m_kuan);
			x = m_h0 - sqrt(pow(m_h0,2) - 2*temp);
			double sigmas = (x/m_h0 - info.b1)/(info.kb - info.b1)*info.ss;
			tm_Aspl = (info.a1*info.cp*m_kuan*x + info.ss*m_Aspy - m_N)/sigmas;
		}
	}
	else
	{
		//Сƫѹ
		MessageBox(_T("����Сƫѹ���м���"));
		double temp = (m_N*e - info.ss*m_Aspy*(m_h0-m_as))/(info.a1*info.cp*m_kuan);
		x = m_h0 - sqrt(pow(m_h0,2) - 2*temp);
		double sigmas = (x/m_h0 - info.b1)/(info.kb - info.b1)*info.ss;
		m_Aspl = (info.a1*info.cp*m_kuan*x + info.ss*m_Aspy - m_N)/sigmas;
	}
	m_Aspy = temp1;
	if(tm_Aspl > m_Aspl)
	{
		MessageBox(_T("����õ��Ĳ�������ѹ�ֽ������ڿ�����ѹ�ֽ�����ȡԭ������"));
	}
	else
	{
		MessageBox(_T("����õ��Ĳ�������ѹ�ֽ���С�ڿ�����ѹ�ֽ�����ȡСֵ"));
		m_Aspl = tm_Aspl;
	}

	m_p = (m_Aspl+m_Aspy)/m_A*100;

	if(m_p < m_pmin)
	{
		MessageBox(_T("���������С����С�����,ȡ��С�����"));
		m_Aspl = ((m_pmin/100) * m_A) - m_Aspy;
	}
	if(m_p > 5)
	{
		MessageBox(_T("��������ʴ���5%,���������"));
		return;
	}

	m_N /= 1000;
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
	CWnd::UpdateData(false);
}



BOOL Cpyd1::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	((CButton *)GetDlgItem(IDC_RADIO_TJ))->SetCheck(TRUE); //ѡ��

	TCHAR path[MAX_PATH] = {0};

	GetModuleFileName(NULL,path,MAX_PATH);


	g_exePATH=path;//��ʱ�����EXE��Ŀ¼

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}


void Cpyd1::OnOK()
{
	// TODO: �ڴ����ר�ô����/����û���
	OnBnClickedButtoncalcula();
	//CDialogEx::OnOK();
}
