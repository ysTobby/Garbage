// Czyc.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Czyc.h"
#include "afxdialogex.h"


// Czyc 对话框

IMPLEMENT_DYNAMIC(Czyc, CDialogEx)

Czyc::Czyc(CWnd* pParent /*=NULL*/)
	: CDialogEx(Czyc::IDD, pParent)
{

	m_Asp = 0.0;
	m_chang = 400;
	m_H = 3.9;
	m_kuan = 400;
	m_N = 0.0;
	m_p = 0.0;
	m_pmin = 0.0;
	m_zhijing = 0.0;
	m_ST = _T("HRB400");
	m_CT = _T("C40");
}

Czyc::~Czyc()
{
}

void Czyc::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_Asp, m_Asp);
	DDX_Text(pDX, IDC_EDIT_chang, m_chang);
	DDX_Text(pDX, IDC_EDIT_H, m_H);
	DDX_Text(pDX, IDC_EDIT_kuan, m_kuan);
	DDX_Text(pDX, IDC_EDIT_N, m_N);
	DDX_Text(pDX, IDC_EDIT_p, m_p);
	DDX_Text(pDX, IDC_EDIT_pmin, m_pmin);
	DDX_Text(pDX, IDC_EDIT_zhijing, m_zhijing);
	DDX_CBString(pDX, IDC_ST, m_ST);
	DDX_CBString(pDX, IDC_CT, m_CT);
}


BEGIN_MESSAGE_MAP(Czyc, CDialogEx)
	ON_BN_CLICKED(IDC_RADIO_fang, &Czyc::OnBnClickedRadiofang)
	ON_BN_CLICKED(IDC_RADIO_yuan, &Czyc::OnBnClickedRadioyuan)
	ON_BN_CLICKED(IDC_BUTTON_calcula, &Czyc::OnBnClickedButtoncalcula)
	ON_BN_CLICKED(IDC_Out, &Czyc::OnBnClickedOut)
END_MESSAGE_MAP()


// Czyc 消息处理程序


void Czyc::OnBnClickedRadiofang()
{
	GetDlgItem(IDC_EDIT_chang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_kuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_chang)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_kuan)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_zhijing)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_zhijing)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_chang)->SetFocus();
}


void Czyc::OnBnClickedRadioyuan()
{
	GetDlgItem(IDC_EDIT_chang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_kuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_chang)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_kuan)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_EDIT_zhijing)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_zhijing)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_EDIT_zhijing)->SetFocus();
}


BOOL Czyc::OnInitDialog()
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


void Czyc::OnOK()
{
	OnBnClickedButtoncalcula();

	//CDialogEx::OnOK();
}


void Czyc::OnBnClickedButtoncalcula()
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

	m_l0 *= 1000;
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
	m_p = m_Asp/(m_A);

	if(m_p < 0.03)
	{
		MessageBox(_T("纵向钢筋配筋率小于3%，公式不需要修正"));
		if(m_p < info.pminp)
		{
			temp.Format(_T("所配钢筋配筋率%.1lf%%小于最小配筋率%.1lf%%,请重新选择钢筋面积"),m_p*100,info.pminp*100);
			MessageBox(temp);
			con2 = 1;
			return;
		}
		else
		{
			MessageBox(_T("所配钢筋满足受压最小纵筋配筋率要求"));
			con2 = 2;
			m_N = 0.9*info.fi*(info.cp*m_A+info.ss*m_Asp)/1000;
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
		}
	}
	else
	{
		MessageBox(_T("纵筋配给率大于3%，需要对公式进行修正"));
		if(m_p > 0.05)
		{
			temp.Format(_T("所配钢筋配筋率%.1lf%%超过5%%,请重新选择配筋面积"),m_p*100);
			MessageBox(temp);
			con2 = 3;
			return;
		}
		else
		{
			MessageBox(_T("所配钢筋面积未超过5%，使用修正后的公式计算"));
			m_N = 0.9*info.fi*(info.cp*(m_A-m_Asp)+info.ss*m_Asp)/1000;
			con2 = 4;
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
		}
	}
	m_p*=100,m_pmin=info.pminp*100;
	CWnd::UpdateData(false);
}


void Czyc::OnBnClickedOut()
{
	CApplication wordApp;
	CDocuments docs;
	CDocument0 docx;
	CBookmarks bookmarks;
	CBookmark0 bookmark;
	CRange range;
	CString temppath;
	temppath = g_exePATH;

	if(con2 == 2 && con1 == 1)//方柱子不修正满足最小配筋
	{
		CString b,chang,CT,fc,ft,ST,fy,H,As,u,l01,l02,p,pmin,l0cyb,fi,Nu;
		b.Format(_T("%.0lf"),m_kuan);
		chang.Format(_T("%.0lf"),m_chang);
		CT = m_CT;
		fc.Format(_T("%.1lf"),info.cp);
		ft.Format(_T("%.2lf"),info.ct);
		ST = m_ST;
		fy.Format(_T("%.0lf"),info.ss);
		H.Format(_T("%.1lf"),m_H);
		As.Format(_T("%.0lf"),m_Asp);
		u.Format(_T("%.1lf"),namda);
		l01.Format(_T("%.2lf"),m_l0/1000);
		l02.Format(_T("%.0lf"),m_l0);
		p.Format(_T("%.2lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		l0cyb.Format(_T("%.2lf"),m_l0/m_kuan);
		fi.Format(_T("%.2lf"),info.fi);
		Nu.Format(_T("%.2lf"),m_N);

		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jxzyzcjh.dot"));

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
		bookmark = bookmarks.Item(&_variant_t(_T("chang1")));
		range = bookmark.get_Range();
		range.put_Text(chang);
		bookmark = bookmarks.Item(&_variant_t(_T("chang2")));
		range = bookmark.get_Range();
		range.put_Text(chang);
		bookmark = bookmarks.Item(&_variant_t(_T("chang3")));
		range = bookmark.get_Range();
		range.put_Text(chang);
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(CT);
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
		bookmark = bookmarks.Item(&_variant_t(_T("H1")));
		range = bookmark.get_Range();
		range.put_Text(H);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l01);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu1")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu2")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("pmin")));
		range = bookmark.get_Range();
		range.put_Text(pmin);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(ST);
		bookmark = bookmarks.Item(&_variant_t(_T("u1")));
		range = bookmark.get_Range();
		range.put_Text(u);
	}
	else if(con2 == 2 && con1 == 2)//圆柱子不修正满足最小配筋
	{ 
		CString As,CT,d,fc,fi,ft,fy,H,l01,l02,l0cyb,Nu,p,pmin,ST,u;
		d.Format(_T("%.0lf"),m_zhijing);
		CT = m_CT;
		fc.Format(_T("%.1lf"),info.cp);
		ft.Format(_T("%.2lf"),info.ct);
		ST = m_ST;
		fy.Format(_T("%.0lf"),info.ss);
		H.Format(_T("%.1lf"),m_H);
		As.Format(_T("%.0lf"),m_Asp);
		u.Format(_T("%.1lf"),namda);
		l01.Format(_T("%.2lf"),m_l0/1000);
		l02.Format(_T("%.0lf"),m_l0);
		p.Format(_T("%.2lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		l0cyb.Format(_T("%.2lf"),m_l0/m_zhijing);
		fi.Format(_T("%.2lf"),info.fi);
		Nu.Format(_T("%.2lf"),m_N);

		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\yxzyzcjh.dot"));

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
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(CT);
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
		bookmark = bookmarks.Item(&_variant_t(_T("H1")));
		range = bookmark.get_Range();
		range.put_Text(H);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l01);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu1")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu2")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("pmin")));
		range = bookmark.get_Range();
		range.put_Text(pmin);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(ST);
		bookmark = bookmarks.Item(&_variant_t(_T("u1")));
		range = bookmark.get_Range();
		range.put_Text(u);
	}
	else if(con2 == 4 && con1 == 2)//圆柱子修正
	{ 
		CString As,CT,d,fc,fi,ft,fy,H,l01,l02,l0cyb,Nu,p,pmin,ST,u;
		d.Format(_T("%.0lf"),m_zhijing);
		CT = m_CT;
		fc.Format(_T("%.1lf"),info.cp);
		ft.Format(_T("%.2lf"),info.ct);
		ST = m_ST;
		fy.Format(_T("%.0lf"),info.ss);
		H.Format(_T("%.1lf"),m_H);
		As.Format(_T("%.0lf"),m_Asp);
		u.Format(_T("%.1lf"),namda);
		l01.Format(_T("%.2lf"),m_l0/1000);
		l02.Format(_T("%.0lf"),m_l0);
		p.Format(_T("%.2lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		l0cyb.Format(_T("%.2lf"),m_l0/m_zhijing);
		fi.Format(_T("%.2lf"),info.fi);
		Nu.Format(_T("%.2lf"),m_N);

		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\yxzyxzjh.dot"));

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
		bookmark = bookmarks.Item(&_variant_t(_T("As4")));
		range = bookmark.get_Range();
		range.put_Text(As);
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(CT);
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
		bookmark = bookmarks.Item(&_variant_t(_T("H1")));
		range = bookmark.get_Range();
		range.put_Text(H);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l01);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu1")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu2")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(ST);
		bookmark = bookmarks.Item(&_variant_t(_T("u1")));
		range = bookmark.get_Range();
		range.put_Text(u);
	}
	else if(con2 == 4 && con1 == 1)//方柱子修正
	{
		CString b,chang,CT,fc,ft,ST,fy,H,As,u,l01,l02,p,pmin,l0cyb,fi,Nu;
		b.Format(_T("%.0lf"),m_kuan);
		chang.Format(_T("%.0lf"),m_chang);
		CT = m_CT;
		fc.Format(_T("%.1lf"),info.cp);
		ft.Format(_T("%.2lf"),info.ct);
		ST = m_ST;
		fy.Format(_T("%.0lf"),info.ss);
		H.Format(_T("%.1lf"),m_H);
		As.Format(_T("%.0lf"),m_Asp);
		u.Format(_T("%.1lf"),namda);
		l01.Format(_T("%.2lf"),m_l0/1000);
		l02.Format(_T("%.0lf"),m_l0);
		p.Format(_T("%.2lf"),m_p);
		pmin.Format(_T("%.2lf"),m_pmin);
		l0cyb.Format(_T("%.2lf"),m_l0/m_kuan);
		fi.Format(_T("%.2lf"),info.fi);
		Nu.Format(_T("%.2lf"),m_N);

		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\jxzyxzjh.dot"));

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
		bookmark = bookmarks.Item(&_variant_t(_T("chang1")));
		range = bookmark.get_Range();
		range.put_Text(chang);
		bookmark = bookmarks.Item(&_variant_t(_T("chang2")));
		range = bookmark.get_Range();
		range.put_Text(chang);
		bookmark = bookmarks.Item(&_variant_t(_T("chang3")));
		range = bookmark.get_Range();
		range.put_Text(chang);
		bookmark = bookmarks.Item(&_variant_t(_T("CT")));
		range = bookmark.get_Range();
		range.put_Text(CT);
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
		bookmark = bookmarks.Item(&_variant_t(_T("H1")));
		range = bookmark.get_Range();
		range.put_Text(H);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(l01);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(l02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyb")));
		range = bookmark.get_Range();
		range.put_Text(l0cyb);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu1")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("Nu2")));
		range = bookmark.get_Range();
		range.put_Text(Nu);
		bookmark = bookmarks.Item(&_variant_t(_T("p1")));
		range = bookmark.get_Range();
		range.put_Text(p);
		bookmark = bookmarks.Item(&_variant_t(_T("ST")));
		range = bookmark.get_Range();
		range.put_Text(ST);
		bookmark = bookmarks.Item(&_variant_t(_T("u1")));
		range = bookmark.get_Range();
		range.put_Text(u);
	}


	range.ReleaseDispatch();
	bookmark.ReleaseDispatch();
	bookmarks.ReleaseDispatch();
	docx.ReleaseDispatch();
	docs.ReleaseDispatch();
	wordApp.ReleaseDispatch();
}
