// EJXY.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "peijin.h"
#include "EJXY.h"
#include "afxdialogex.h"


// CEJXY �Ի���

IMPLEMENT_DYNAMIC(CEJXY, CDialogEx)

CEJXY::CEJXY(CWnd* pParent /*=NULL*/)
	: CDialogEx(CEJXY::IDD, pParent)
{

	m_as = 40;
	m_b = 300;
	m_fc = 14.3;
	m_h = 400;
	m_l0 = 2400;
	m_M1 = 200.56;
	m_M2 = 218;
	m_N = 396;
	m_r1 = 0.0;
	m_r2 = 0.0;
	m_r3 = 0.0;
	m_r4 = 0.0;
	m_r5 = 0.0;
	m_r6 = 0.0;
	m_r7 = 0.0;
	m_r8 = 0.0;
}

CEJXY::~CEJXY()
{
}

void CEJXY::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_as, m_as);
	DDX_Text(pDX, IDC_b, m_b);
	DDX_Text(pDX, IDC_fc, m_fc);
	DDX_Text(pDX, IDC_h, m_h);
	DDX_Text(pDX, IDC_l0, m_l0);
	DDX_Text(pDX, IDC_M1, m_M1);
	DDX_Text(pDX, IDC_M2, m_M2);
	DDX_Text(pDX, IDC_N, m_N);
	DDX_Text(pDX, IDC_r1, m_r1);
	DDX_Text(pDX, IDC_r2, m_r2);
	DDX_Text(pDX, IDC_r3, m_r3);
	DDX_Text(pDX, IDC_r4, m_r4);
	DDX_Text(pDX, IDC_r5, m_r5);
	DDX_Text(pDX, IDC_r6, m_r6);
	DDX_Text(pDX, IDC_r7, m_r7);
	DDX_Text(pDX, IDC_r8, m_r8);
}


BEGIN_MESSAGE_MAP(CEJXY, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &CEJXY::OnBnClickedButton1)
END_MESSAGE_MAP()


// CEJXY ��Ϣ�������


void CEJXY::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CWnd::UpdateData(true);
	m_N = m_N * 1000;
	m_M2 = m_M2 * 1000000;
	m_M1 = m_M1 * 1000000;

	double A,rm_r5,m_ea,rm_r7;
	A = m_h * m_b;
	m_ea = ((m_h/30)>20)?(m_h/30):20;//����ƫ�ľ�


	m_r1 = m_M1/m_M2;
	if(m_r1 > 0.9)
	{
		m_r4 = 0.7 + 0.3*m_r1;
		m_r5 = 0.5*m_fc*A/(m_N);//�����ĵ�λȡΪkN
		rm_r5 = (m_r5 > 1.0)?1.0:m_r5;
		m_r6 = 1+(m_l0/m_h)*(m_l0/m_h)*rm_r5/(1300*(m_M2/m_N+m_ea)/(m_h-m_as));//m_M2��λkn.m,ml0 mm
		m_r7 = m_r4 * m_r6;
		rm_r7 = (m_r7 < 1.0)?1.0:m_r7;
		m_r8 = rm_r7 * m_M2;
		m_r8 = m_r8 / 1000000;
		m_r2 = 0;
		m_r3 = 0;
	}
	else
	{
		m_r2 = m_N/(m_fc*A);
		if(m_r2 > 0.9)
		{
			m_r4 = 0.7 + 0.3*m_M1/m_M2;
			m_r5 = 0.5*m_fc*A/(m_N);//�����ĵ�λȡΪkN
			rm_r5 = (m_r5 > 1.0)?1.0:m_r5;
			m_r6 = 1+(m_l0/m_h)*(m_l0/m_h)*rm_r5/(1300*(m_M2/m_N+m_ea)/(m_h-m_as));//m_M2��λkn.m
			m_r7 = m_r4 * m_r6;
			rm_r7 = (m_r7 < 1.0)?1.0:m_r7;
			m_r8 = rm_r7 * m_M2;
			m_r8 = m_r8 / 1000000;
			m_r3 = 0;
		}
		else
		{
			double I,m_i;
			I = 1/12*m_b*m_h*m_h*m_h;
			m_i = sqrt(I/A);
			m_r3 = m_l0/m_i;
			if(m_r3 > (34-12*(m_M1/m_M2)))
			{
				m_r4 = 0.7 + 0.3*m_M1/m_M2;
				m_r5 = 0.5*m_fc*A/(m_N);//�����ĵ�λȡΪkN
				rm_r5 = (m_r5 > 1.0)?1.0:m_r5;
				m_r6 = 1+(m_l0/m_h)*(m_l0/m_h)*rm_r5/(1300*(m_M2/m_N+m_ea)/(m_h-m_as));//m_M2��λkn.m
				m_r7 = m_r4 * m_r6;
				rm_r7 = (m_r7 < 1.0)?1.0:m_r7;
				m_r8 = rm_r7 * m_M2;
				m_r8 = m_r8 / 1000000;
			}
			else
			{
				m_r8 = m_M2;
				m_r4 = 0;
				m_r5 = 0;
				m_r6 = 0;
				m_r7 = 0;
				MessageBox(_T("����Ҫ���Ƕ���ЧӦ"));
			}
		}
	}
	m_N = m_N / 1000;
	m_M2 = m_M2 / 1000000;
	m_M1 = m_M1 / 1000000;

	CWnd::UpdateData(false);
}


void CEJXY::OnOK()
{
	// TODO: �ڴ����ר�ô����/����û���
	CEJXY::OnBnClickedButton1();
}
