// Czyd.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Czyd.h"
#include "afxdialogex.h"


// Czyd 对话框

IMPLEMENT_DYNAMIC(Czyd, CDialogEx)

	Czyd::Czyd(CWnd* pParent /*=NULL*/)
	: CDialogEx(Czyd::IDD, pParent)
{

	m_chang = 400;
	m_kuan = 400;
	m_zhijing = 0.0;
	m_H = 3.9;
	m_N = 3090;
	m_ST = _T("HRB400");
	m_CT = _T("C40");
	m_Asp = 0.0;
	m_p = 0.0;
	m_pmin = 0.0;
}

Czyd::~Czyd()
{
}

void Czyd::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_chang, m_chang);
	DDX_Text(pDX, IDC_EDIT_kuan, m_kuan);
	DDX_Text(pDX, IDC_EDIT_zhijing, m_zhijing);
	DDX_Text(pDX, IDC_EDIT_H, m_H);
	DDX_Text(pDX, IDC_EDIT_N, m_N);
	DDX_CBString(pDX, IDC_ST, m_ST);
	DDX_CBString(pDX, IDC_CT, m_CT);
	DDX_Text(pDX, IDC_EDIT_Asp, m_Asp);
	DDX_Text(pDX, IDC_EDIT_p, m_p);
	DDX_Text(pDX, IDC_EDIT_pmin, m_pmin);
}


BEGIN_MESSAGE_MAP(Czyd, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_calcula, &Czyd::OnBnClickedButtoncalcula)
	ON_BN_CLICKED(IDC_RADIO_fang, &Czyd::OnBnClickedRadiofang)
	ON_BN_CLICKED(IDC_RADIO_yuan, &Czyd::OnBnClickedRadioyuan)
	ON_BN_CLICKED(IDC_Out, &Czyd::OnBnClickedOut)
END_MESSAGE_MAP()


// Czyd 消息处理程序


void Czyd::OnBnClickedButtoncalcula()
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
	m_l0 = namda*m_H;

	m_l0 *= 1000;//化为了毫米
	m_N *= 1000;//化为了N
	if(((CButton *)GetDlgItem(IDC_RADIO_fang))->GetCheck())
	{
		info.initZY(m_CT,m_ST,1,m_l0,m_kuan);
		m_A = m_kuan * m_chang;
		con1 = 1;
	}
	if(((CButton *)GetDlgItem(IDC_RADIO_yuan))->GetCheck())
	{
		info.initZY(m_CT,m_ST,2,m_l0,m_zhijing);
		m_A = m_zhijing * m_zhijing * 3.1415926 /4;
		con1 = 2;
	}

	m_Asp = (m_N/(0.9*info.fi) - info.cp*m_A)/info.ss;
	m_pp = m_Asp/m_A;
	m_Aspf = m_Asp;
	m_ppf = m_pp;
	m_ppf *= 100;

	if(m_pp <= 0.03)
	{
		MessageBox(_T("纵向钢筋配筋率小于3%，公式不需要修正"));
		if(m_pp < info.pminp)
		{
			MessageBox(_T("所配钢筋不满足受压最小纵筋配筋率要求,取为最小配筋面积"));
			m_Asp = info.pminp*m_A;
			con2 = 1;
		}
		else
		{
			MessageBox(_T("所配钢筋满足受压最小纵筋配筋率要求"));
			con2 = 2;
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
		}
	}
	else if(m_pp > 0.03)
	{
		m_Asp = (m_N/(0.9*info.fi) - info.cp*m_A)/(info.ss - info.cp);
		m_pp = m_Asp/m_A;

		MessageBox(_T("纵筋配给率大于3%，需要对公式进行修正"));
		if(m_pp > 0.05)
		{
			if(con1 == 1)
				MessageBox(_T("计算配筋面积超过5%，应增大截面尺寸"));
			else
				MessageBox(_T("计算配筋面积超过5%，可增大截面或者采用螺旋形箍筋柱"));
			con2 = 3;
			return;
		}
		else
		{
			MessageBox(_T("所配钢筋面积未超过5%，可以使用"));
			con2 = 4;
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
		}
	}
	m_p = m_pp*100;
	m_pmin = info.pminp*100;

	m_l0 /= 1000;//化为了毫米
	m_N /= 1000;//化为了N
	CWnd::UpdateData(false);
}


void Czyd::OnBnClickedRadiofang()
{
	GetDlgItem(IDC_EDIT_chang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_kuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_chang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_kuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_zhijing)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_zhijing)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_chang)->SetFocus();
}


void Czyd::OnBnClickedRadioyuan()
{
	GetDlgItem(IDC_EDIT_chang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_kuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_chang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_kuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_zhijing)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_zhijing)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_zhijing)->SetFocus();
}


BOOL Czyd::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	((CButton *)GetDlgItem(IDC_RADIO_fang))->SetCheck(TRUE); //选上
	((CButton *)GetDlgItem(IDC_RADIO_TJ))->SetCheck(TRUE); //选上
	OnBnClickedRadiofang();

	TCHAR path[MAX_PATH] = {0};

	GetModuleFileName(NULL,path,MAX_PATH);


	g_exePATH=path;//此时获得了EXE的目录

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void Czyd::OnOK()
{
	OnBnClickedButtoncalcula();
	//CDialogEx::OnOK();
}




void Czyd::OnBnClickedOut()
{
	CApplication wordApp;
	CDocuments docs;
	CDocument0 docx;
	CBookmarks bookmarks;
	CBookmark0 bookmark;
	CRange range;
	CString temppath;
	temppath = g_exePATH;

	if(con1 == 1&&con2 == 2)
	{
		//绑定参数
		CString As,b,fc,fi,ft,fy,h,l0,l02,N,snamda,p,pmin,H,l0cyb;
		As.Format(_T("%.1lf"),m_Asp);
		b.Format(_T("%.0lf"),m_kuan);
		fc.Format(_T("%.1lf"),info.cp);
		fi.Format(_T("%.3lf"),info.fi);
		ft.Format(_T("%.2lf"),info.ct);
		fy.Format(_T("%.0lf"),info.ss);
		h.Format(_T("%.0lf"),m_chang);
		l0.Format(_T("%.2lf"),m_l0);
		l02.Format(_T("%.0lf"),m_l0*1000);
		N.Format(_T("%.0lf"),m_N);
		snamda.Format(_T("%.1lf"),namda);
		p.Format(_T("%.1lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		H.Format(_T("%.1lf"),m_H);
		l0cyb.Format(_T("%.2lf"),m_l0*1000/m_kuan);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jxzyzc.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("As1")));//与模板中的书签名对应
		range = bookmark.get_Range();
		range.put_Text(As);//修改的值
		bookmark = bookmarks.Item(&_variant_t(_T("As2")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("As3")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("b1")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b2")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b3")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b4")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(m_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("ft1")));
		range = bookmark.get_Range();
		range.put_Text(ft);
		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("h1")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("h2")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("h3")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("namda")));
		range = bookmark.get_Range();
		range.put_Text(snamda);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("pmin")));
		range = bookmark.get_Range();
		range.put_Text(pmin);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(m_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("zg")));
		range = bookmark.get_Range();
		range.put_Text(H);
	}
	else if(con1 == 2&&con2 == 2)
	{
		CString As,fc,fi,ft,fy,d,l0,l02,N,snamda,p,pmin,H,l0cyb;
		As.Format(_T("%.1lf"),m_Asp);
		d.Format(_T("%.0lf"),m_zhijing);
		fc.Format(_T("%.1lf"),info.cp);
		fi.Format(_T("%.3lf"),info.fi);
		ft.Format(_T("%.2lf"),info.ct);
		fy.Format(_T("%.0lf"),info.ss);
		//h.Format(_T("%.0lf"),m_chang);
		l0.Format(_T("%.2lf"),m_l0);
		l02.Format(_T("%.0lf"),m_l0*1000);
		N.Format(_T("%.0lf"),m_N);
		snamda.Format(_T("%.1lf"),namda);
		p.Format(_T("%.1lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		H.Format(_T("%.1lf"),m_H);
		l0cyb.Format(_T("%.2lf"),m_l0*1000/m_kuan);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\yxzyzc.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("As1")));//与模板中的书签名对应
		range = bookmark.get_Range();
		range.put_Text(As);//修改的值
		bookmark = bookmarks.Item(&_variant_t(_T("As2")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("As3")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(m_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("ft1")));
		range = bookmark.get_Range();
		range.put_Text(ft);
		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("namda")));
		range = bookmark.get_Range();
		range.put_Text(snamda);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("pmin")));
		range = bookmark.get_Range();
		range.put_Text(pmin);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(m_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("zg")));
		range = bookmark.get_Range();
		range.put_Text(H);
	}
	else if(con1 == 1&&con2 == 4)
	{
		//绑定参数
		CString As,b,fc,fi,ft,fy,h,l0,l02,N,snamda,p,pmin,H,l0cyb,pf,Asf;
		As.Format(_T("%.1lf"),m_Asp);
		b.Format(_T("%.0lf"),m_kuan);
		fc.Format(_T("%.1lf"),info.cp);
		fi.Format(_T("%.3lf"),info.fi);
		ft.Format(_T("%.2lf"),info.ct);
		fy.Format(_T("%.0lf"),info.ss);
		h.Format(_T("%.0lf"),m_chang);
		l0.Format(_T("%.2lf"),m_l0);
		l02.Format(_T("%.0lf"),m_l0*1000);
		N.Format(_T("%.0lf"),m_N);
		snamda.Format(_T("%.1lf"),namda);
		p.Format(_T("%.1lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		H.Format(_T("%.1lf"),m_H);
		l0cyb.Format(_T("%.2lf"),m_l0*1000/m_kuan);
		pf.Format(_T("%.1lf"),m_ppf);
		Asf.Format(_T("%.1lf"),m_Aspf);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jxzyxz.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("As1")));//与模板中的书签名对应
		range = bookmark.get_Range();
		range.put_Text(Asf);//修改的值
		bookmark = bookmarks.Item(&_variant_t(_T("As2")));
		range = bookmark.get_Range();
		range.put_Text(Asf);
		bookmark = bookmarks.Item(&_variant_t(_T("As3")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("As4")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("b1")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b2")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b3")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b4")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b5")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("b6")));
		range = bookmark.get_Range();
		range.put_Text(b);
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(m_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc4")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi3")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("ft1")));
		range = bookmark.get_Range();
		range.put_Text(ft);
		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy3")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("h1")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("h2")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("h3")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("h4")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("h5")));
		range = bookmark.get_Range();
		range.put_Text(h);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("N3")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("namda")));
		range = bookmark.get_Range();
		range.put_Text(snamda);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(pf);
		bookmark = bookmarks.Item(&_variant_t(_T("p2")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(m_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("zg")));
		range = bookmark.get_Range();
		range.put_Text(H);
	}
	else if(con1 == 2&&con2 == 4)
	{
		CString As,fc,fi,ft,fy,d,l0,l02,N,snamda,p,pmin,H,l0cyb,Asf,pf;
		//b.Format(_T("%.0lf"),m_kuan);
		//h.Format(_T("%.0lf"),m_chang);
		pf.Format(_T("%.1lf"),m_ppf);
		Asf.Format(_T("%.1lf"),m_Aspf);
		As.Format(_T("%.1lf"),m_Asp);
		d.Format(_T("%.0lf"),m_zhijing);
		fc.Format(_T("%.1lf"),info.cp);
		fi.Format(_T("%.3lf"),info.fi);
		ft.Format(_T("%.2lf"),info.ct);
		fy.Format(_T("%.0lf"),info.ss);
		//h.Format(_T("%.0lf"),m_chang);
		l0.Format(_T("%.2lf"),m_l0);
		l02.Format(_T("%.0lf"),m_l0*1000);
		N.Format(_T("%.0lf"),m_N);
		snamda.Format(_T("%.1lf"),namda);
		p.Format(_T("%.1lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		H.Format(_T("%.1lf"),m_H);
		l0cyb.Format(_T("%.2lf"),m_l0*1000/m_kuan);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\yxzyxz.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("As1")));//与模板中的书签名对应
		range = bookmark.get_Range();
		range.put_Text(Asf);//修改的值
		bookmark = bookmarks.Item(&_variant_t(_T("As2")));
		range = bookmark.get_Range();
		range.put_Text(Asf);
		bookmark = bookmarks.Item(&_variant_t(_T("As3")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("As4")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d4")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d5")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("d6")));
		range = bookmark.get_Range();
		range.put_Text(d);
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(m_CT);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc4")));
		range = bookmark.get_Range();
		range.put_Text(fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fi1")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi2")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("fi3")));
		range = bookmark.get_Range();
		range.put_Text(fi);
		bookmark = bookmarks.Item(&_variant_t(_T("ft1")));
		range = bookmark.get_Range();
		range.put_Text(ft);
		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy3")));
		range = bookmark.get_Range();
		range.put_Text(fy);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l0);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("N3")));
		range = bookmark.get_Range();
		range.put_Text(N);
		bookmark = bookmarks.Item(&_variant_t(_T("namda")));
		range = bookmark.get_Range();
		range.put_Text(snamda);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(pf);
		bookmark = bookmarks.Item(&_variant_t(_T("p2")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(m_ST);
		bookmark = bookmarks.Item(&_variant_t(_T("zg")));
		range = bookmark.get_Range();
		range.put_Text(H);
	}

	range.ReleaseDispatch();
	bookmark.ReleaseDispatch();
	bookmarks.ReleaseDispatch();
	docx.ReleaseDispatch();
	docs.ReleaseDispatch();
	wordApp.ReleaseDispatch();

	// 退出word应用
	//docx.Close(covFalse, covOptional, covOptional);
	//wordApp.Quit(covOptional, covOptional, covOptional);

}
