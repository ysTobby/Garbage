// Czyyc.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Czyyc.h"
#include "afxdialogex.h"


// Czyyc 对话框

IMPLEMENT_DYNAMIC(Czyyc, CDialogEx)

Czyyc::Czyyc(CWnd* pParent /*=NULL*/)
	: CDialogEx(Czyyc::IDD, pParent)
{

	m_CT = _T("C30");
	m_Asp = 0.0;
	m_d = 400;
	m_GD = 10;
	m_H = 4.5;
	m_N = 0.0;
	m_pas = 20;
	m_s = 40;
	m_GST = _T("HPB300");
	m_ST = _T("HRB400");
}

Czyyc::~Czyyc()
{
}

void Czyyc::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_CBString(pDX, IDC_CT, m_CT);
	DDX_Text(pDX, IDC_EDIT_As, m_Asp);
	DDX_Text(pDX, IDC_EDIT_d, m_d);
	DDX_Text(pDX, IDC_EDIT_GD, m_GD);
	DDX_Text(pDX, IDC_EDIT_H, m_H);
	DDX_Text(pDX, IDC_EDIT_N, m_N);
	DDX_Text(pDX, IDC_EDIT_pas, m_pas);
	DDX_Text(pDX, IDC_EDIT_SP, m_s);
	DDX_CBString(pDX, IDC_GST, m_GST);
	DDX_CBString(pDX, IDC_ST, m_ST);
}


BEGIN_MESSAGE_MAP(Czyyc, CDialogEx)
	ON_BN_CLICKED(IDC_Calcula, &Czyyc::OnBnClickedCalcula)
	ON_BN_CLICKED(IDC_Out, &Czyyc::OnBnClickedOut)
END_MESSAGE_MAP()


// Czyyc 消息处理程序


void Czyyc::OnOK()
{
	OnBnClickedCalcula();

	//CDialogEx::OnOK();
}


BOOL Czyyc::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	((CButton *)GetDlgItem(IDC_RADIO_TJ))->SetCheck(TRUE); //选上

	TCHAR path[MAX_PATH] = {0};

	GetModuleFileName(NULL,path,MAX_PATH);


	g_exePATH=path;//此时获得了EXE的目录

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void Czyyc::OnBnClickedCalcula()
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

	l0 = m_H * namda * 1000;//mm
	dcor = m_d - 2 * (m_GD + m_pas);
	info.initYSZY(m_CT,m_ST,m_GST,l0,m_d,m_GD);
	p = m_Asp/(0.25*PI*m_d*m_d);
	if(p>0.05)
	{
		MessageBox(_T("配筋面积大于5%，请重新选择配筋面积"));
		return;
	}

	if(l0/m_d > 12)
	{
		con1 = 1;
		//p = m_Asp/(0.25*PI*m_d*m_d);
		if(p<0.03)
		{
			//不用修正
			con = 1;
			m_N = 0.9 * info.fi * (info.cp * 0.25*PI*m_d*m_d + info.ss*m_Asp)/1000;
		}
		else
		{
			//修正
			con = 2;
			m_N = 0.9 * info.fi * (info.cp * (0.25*PI*m_d*m_d - m_Asp) + info.ss*m_Asp)/1000;
		}
	}
	else
	{
		//考虑间接配筋
		con1 = 2;
		Ass0 = (PI * dcor * info.Assl)/m_s;
		if(Ass0>(m_Asp/4))
		{
			//符合间接配筋条件
			N1 = 0.9 * (info.cp*0.25*PI*dcor*dcor + info.ss*m_Asp + 2*info.yszya*info.gjss*Ass0)/1000;
			//p = m_Asp/(0.25*PI*m_d*m_d);
			if(p<0.03)
			{
				//不用修正
				con = 1;
				N2 = 0.9 * info.fi * (info.cp * 0.25*PI*m_d*m_d + info.ss*m_Asp)/1000;
			}
			else
			{
				//修正
				con = 2;
				N2 = 0.9 * info.fi * (info.cp * (0.25*PI*m_d*m_d - m_Asp) + info.ss*m_Asp)/1000;
			}
			N3 = 1.5 * N2;
			if(N1 <= N3 && N2 <= N1)
			{
				con2 = 1;
				m_N = N1;
			}
			else if(N1 < N2)
			{
				con2 = 2;
				m_N = N2;
			}
			else
			{
				con2 = 3;
				MessageBox(_T("计算的约束承载力大于未约束承载力的1.5倍，不安全"));
				return;
			}
		}
		else
		{
			//不符合间接配筋条件,间接钢筋配置太少，约束效果不明显
			con1 = 3;
			MessageBox(_T("间接钢筋配置过少，约束混凝土效果不明显,按照未配置间接钢筋计算"));
			if(p<0.03)
			{
				//不用修正
				con = 1;
				m_N = 0.9 * info.fi * (info.cp * 0.25*PI*m_d*m_d + info.ss*m_Asp)/1000;
			}
			else
			{
				//修正
				con = 2;
				m_N = 0.9 * info.fi * (info.cp * (0.25*PI*m_d*m_d - m_Asp) + info.ss*m_Asp)/1000;
			}
		}
	}
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);

	CWnd::UpdateData(false);
}


void Czyyc::OnBnClickedOut()
{
	CApplication wordApp;
	CDocuments docs;
	CDocument0 docx;
	CBookmarks bookmarks;
	CBookmark0 bookmark;
	CRange range;
	CString temppath;
	temppath = g_exePATH;

	if(con1 == 1 && con == 1)
	{
		//不考虑间接配筋，不需要修正
		//绑定参数
		CString s_Asp,s_Assl,s_CT,s_d,s_fc,s_fi,s_fyp,s_fyv,s_GST,s_l0,s_l0cyd,
			s_l0mm,s_N,s_p,s_pjgd,s_s,s_ST,s_dcor;
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N.Format(_T("%.1lf"),m_N);
		s_p.Format(_T("%.1lf"),p*100);
		s_pjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\bxbj.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("pjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_pjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
	}
	else if(con1 == 1 && con == 2)
	{
		CString s_Asp,s_Assl,s_CT,s_d,s_fc,s_fi,s_fyp,s_fyv,s_GST,s_l0,s_l0cyd,
			s_l0mm,s_N,s_p,s_pjgd,s_s,s_ST,s_dcor;
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N.Format(_T("%.1lf"),m_N);
		s_p.Format(_T("%.1lf"),p*100);
		s_pjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\xbj.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("pjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_pjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
	}
	else if(con1 == 3 && con == 1)
	{
		//不修改，因面积导致不考虑间接 bxbjA
		CString s_Asp,s_Assl,s_CT,s_d,s_fc,s_fi,s_fyp,s_fyv,s_GST,s_l0,s_l0cyd,
			s_l0mm,s_N,s_p,s_pjgd,s_s,s_ST,s_dcor,s_Aspcy4,s_Ass0;
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N.Format(_T("%.1lf"),m_N);
		s_p.Format(_T("%.1lf"),p*100);
		s_pjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		s_Aspcy4.Format(_T("%.1lf"),(m_Asp/4));
		s_Ass0.Format(_T("%.1lf"),Ass0);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\bxbjA.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy4")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass01")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl2")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("pjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_pjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
	}
	else if(con1 == 3 && con == 2)
	{
		//修改，因面积导致不考虑间接 xbjA
		CString s_Asp,s_Assl,s_CT,s_d,s_fc,s_fi,s_fyp,s_fyv,s_GST,s_l0,s_l0cyd,
			s_l0mm,s_N,s_p,s_pjgd,s_s,s_ST,s_dcor,s_Aspcy4,s_Ass0;
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N.Format(_T("%.1lf"),m_N);
		s_p.Format(_T("%.1lf"),p*100);
		s_pjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		s_Aspcy4.Format(_T("%.1lf"),(m_Asp/4));
		s_Ass0.Format(_T("%.1lf"),Ass0);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\xbjA.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp5")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy4")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass01")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl2")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(s_N);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("pjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_pjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
	}
	else if(con1 == 2 && con == 1 && con2 == 1)
	{
		//考虑间接配筋正常不修正 - jpzcbx
		CString s_a,s_Asp,s_Aspcy4,s_Ass0,s_Assl,s_bjgd,s_CT,s_d,s_dcor,s_fc,s_fi,s_fyp,s_fyv,
			s_GST,s_l0,s_l0cyd,s_l0mm,s_N1,s_N2,s_N3,s_s,s_ST,s_p;
		s_a.Format(_T("%.1lf"),info.yszya);
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N1.Format(_T("%.1lf"),N1);
		s_N2.Format(_T("%.1lf"),N2);
		s_N3.Format(_T("%.1lf"),N3);
		s_p.Format(_T("%.1lf"),p*100);
		s_bjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		s_Aspcy4.Format(_T("%.1lf"),(m_Asp/4));
		s_Ass0.Format(_T("%.1lf"),Ass0);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jpzcbx.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("a1")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("a2")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp5")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp6")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy41")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy42")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass01")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass02")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass03")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl2")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("bjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_bjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d6")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor3")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc4")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi3")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp3")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp4")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N11")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N12")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N13")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N14")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N21")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N22")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N31")));
		range = bookmark.get_Range();
		range.put_Text(s_N3);
		bookmark = bookmarks.Item(&_variant_t(_T("N32")));
		range = bookmark.get_Range();
		range.put_Text(s_N3);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
	}
	else if(con1 == 2 && con == 2 && con2 == 1)
	{
		//考虑间接配筋正常修正 - jpzcx
		CString s_a,s_Asp,s_Aspcy4,s_Ass0,s_Assl,s_bjgd,s_CT,s_d,s_dcor,s_fc,s_fi,s_fyp,s_fyv,
			s_GST,s_l0,s_l0cyd,s_l0mm,s_N1,s_N2,s_N3,s_s,s_ST,s_p;
		s_a.Format(_T("%.1lf"),info.yszya);
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N1.Format(_T("%.1lf"),N1);
		s_N2.Format(_T("%.1lf"),N2);
		s_N3.Format(_T("%.1lf"),N3);
		s_p.Format(_T("%.1lf"),p*100);
		s_bjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		s_Aspcy4.Format(_T("%.1lf"),(m_Asp/4));
		s_Ass0.Format(_T("%.1lf"),Ass0);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jpzcx.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("a1")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("a2")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp5")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp6")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp7")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp8")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy41")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy42")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass01")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass02")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass03")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl2")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("bjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_bjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d6")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor3")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc4")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi3")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp3")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp4")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N11")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N12")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N13")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N14")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N21")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N22")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N31")));
		range = bookmark.get_Range();
		range.put_Text(s_N3);
		bookmark = bookmarks.Item(&_variant_t(_T("N32")));
		range = bookmark.get_Range();
		range.put_Text(s_N3);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
	}
	else if(con1 == 2 && con == 1 && con2 == 2)
	{
		//考虑间接配筋不正常不修正 - jpbzcbx
		CString s_a,s_Asp,s_Aspcy4,s_Ass0,s_Assl,s_bjgd,s_CT,s_d,s_dcor,s_fc,s_fi,s_fyp,s_fyv,
			s_GST,s_l0,s_l0cyd,s_l0mm,s_N1,s_N2,s_s,s_ST,s_p;
		s_a.Format(_T("%.1lf"),info.yszya);
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N1.Format(_T("%.1lf"),N1);
		s_N2.Format(_T("%.1lf"),N2);
		s_p.Format(_T("%.1lf"),p*100);
		s_bjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		s_Aspcy4.Format(_T("%.1lf"),(m_Asp/4));
		s_Ass0.Format(_T("%.1lf"),Ass0);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jpbzcbx.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("a1")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("a2")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp5")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy41")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy42")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass01")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass02")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass03")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl2")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("bjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_bjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor3")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp3")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N11")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N12")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N21")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N22")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N23")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
	}
	else if(con1 == 2 && con == 2 && con2 == 2)
	{
		//考虑间接配筋不正常修正 - jpbzcx
		CString s_a,s_Asp,s_Aspcy4,s_Ass0,s_Assl,s_bjgd,s_CT,s_d,s_dcor,s_fc,s_fi,s_fyp,s_fyv,
			s_GST,s_l0,s_l0cyd,s_l0mm,s_N1,s_N2,s_s,s_ST,s_p;
		s_a.Format(_T("%.1lf"),info.yszya);
		s_Asp.Format(_T("%.0lf"),m_Asp);
		s_Assl.Format(_T("%.1lf"),info.Assl);
		s_CT = m_CT;
		s_d.Format(_T("%.0lf"),m_d);
		s_fc.Format(_T("%.1lf"),info.cp);
		s_fi.Format(_T("%.2lf"),info.fi);
		s_fyp.Format(_T("%.0lf"),info.ss);
		s_fyv.Format(_T("%.0lf"),info.gjss);
		s_GST = m_GST;
		s_l0.Format(_T("%.1lf"),l0/1000);
		s_l0cyd.Format(_T("%.1lf"),l0/m_d);
		s_l0mm.Format(_T("%.0lf"),l0);
		s_N1.Format(_T("%.1lf"),N1);
		s_N2.Format(_T("%.1lf"),N2);
		s_p.Format(_T("%.1lf"),p*100);
		s_bjgd.Format(_T("%.0lf"),(m_pas+m_GD));
		s_s.Format(_T("%.0lf"),m_s);
		s_ST = m_ST;
		s_dcor.Format(_T("%.0lf"),dcor);
		s_Aspcy4.Format(_T("%.1lf"),(m_Asp/4));
		s_Ass0.Format(_T("%.1lf"),Ass0);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jpbzcx.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant    covZero((short)0),
			covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR),
			covDocxType((short)0),
			start_line, end_line,
			dot(temppath);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot, covOptional, covOptional, covOptional);
		bookmarks = docx.get_Bookmarks();

		bookmark = bookmarks.Item(&_variant_t(_T("a1")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("a2")));
		range = bookmark.get_Range();
		range.put_Text(s_a);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp5")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp6")));
		range = bookmark.get_Range();
		range.put_Text(s_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy41")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Aspcy42")));
		range = bookmark.get_Range();
		range.put_Text(s_Aspcy4);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass01")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass02")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Ass03")));
		range = bookmark.get_Range();
		range.put_Text(s_Ass0);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl1")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("Assl2")));
		range = bookmark.get_Range();
		range.put_Text(s_Assl);
		bookmark = bookmarks.Item(&_variant_t(_T("bjgd1")));
		range = bookmark.get_Range();
		range.put_Text(s_bjgd);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(s_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(s_d);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor3")));
		range = bookmark.get_Range();
		range.put_Text(s_dcor);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(s_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(s_fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyp3")));
		range = bookmark.get_Range();
		range.put_Text(s_fyp);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv1")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("fyv2")));
		range = bookmark.get_Range();
		range.put_Text(s_fyv);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(s_GST);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(s_l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd2")));
		range = bookmark.get_Range();
		range.put_Text(s_l0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("l0mm1")));
		range = bookmark.get_Range();
		range.put_Text(s_l0mm);
		bookmark = bookmarks.Item(&_variant_t(_T("N11")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N12")));
		range = bookmark.get_Range();
		range.put_Text(s_N1);
		bookmark = bookmarks.Item(&_variant_t(_T("N21")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N22")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("N23")));
		range = bookmark.get_Range();
		range.put_Text(s_N2);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(s_s);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(s_p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(s_ST);
	}

	range.ReleaseDispatch();
	bookmark.ReleaseDispatch();
	bookmarks.ReleaseDispatch();
	docx.ReleaseDispatch();
	docs.ReleaseDispatch();
	wordApp.ReleaseDispatch();
}
