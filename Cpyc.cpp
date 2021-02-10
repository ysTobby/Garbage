// Cpyc.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Cpyc.h"
#include "afxdialogex.h"
#include "OperateWord.h"


// Cpyc 对话框

IMPLEMENT_DYNAMIC(Cpyc, CDialogEx)

Cpyc::Cpyc(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cpyc::IDD, pParent)
	, m_CS(_T("C40"))
	, m_ST(_T("HRB400"))
	, m_as(40)
	, m_B(400)
	, m_e(0)
	, m_e0(0)
	, m_ei(0)
	, m_ep(0)
	, m_H(600)
	, m_iks(0)
	, m_mks(0)
	, m_Mu(0)
	, m_N(1800)
	, m_Nb(0)
	, m_PYT(_T(""))
	, m_SAs(1964)
	, m_SAsp(763)
{
	x = 0;
	m_kscy = 0;
}

Cpyc::~Cpyc()
{
}

void Cpyc::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_CBString(pDX, IDC_DD1_CS, m_CS);
	DDX_CBString(pDX, IDC_DD1_ST, m_ST);
	DDX_Text(pDX, IDC_EDIT_As, m_as);
	DDX_Text(pDX, IDC_EDIT_B, m_B);
	DDX_Text(pDX, IDC_EDIT_e, m_e);
	DDX_Text(pDX, IDC_EDIT_e0, m_e0);
	DDX_Text(pDX, IDC_EDIT_ei, m_ei);
	DDX_Text(pDX, IDC_EDIT_ep, m_ep);
	DDX_Text(pDX, IDC_EDIT_H, m_H);
	DDX_Text(pDX, IDC_EDIT_iks, m_iks);
	DDX_Text(pDX, IDC_EDIT_mks, m_mks);
	DDX_Text(pDX, IDC_EDIT_Mu, m_Mu);
	DDX_Text(pDX, IDC_EDIT_N, m_N);
	DDX_Text(pDX, IDC_EDIT_Nb, m_Nb);
	DDX_Text(pDX, IDC_EDIT_PYT, m_PYT);
	DDX_Text(pDX, IDC_EDIT_SAs, m_SAs);
	DDX_Text(pDX, IDC_EDIT_SAsp, m_SAsp);
}


BEGIN_MESSAGE_MAP(Cpyc, CDialogEx)
	ON_BN_CLICKED(IDOK, &Cpyc::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON_cal, &Cpyc::OnBnClickedButtoncal)
	ON_BN_CLICKED(IDC_BUTTON_export, &Cpyc::OnBnClickedButtonexport)
END_MESSAGE_MAP()


// Cpyc 消息处理程序


void Cpyc::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	//CDialogEx::OnOK();
	OnBnClickedButtoncal();
}

void Cpyc::reset()
{
	m_Nb = 0;
	m_PYT = _T("");
	m_e = 0;
	m_ep = 0;
	m_ei = 0;
	m_e0 = 0;
	m_iks = 0;
	m_mks = 0;
	m_Mu = 0;
}

void Cpyc::OnBnClickedButtoncal()
{
	// TODO: 在此添加控件通知处理程序代码
	CWnd::UpdateData(true);
	reset();
	GetDlgItem(IDC_BUTTON_export)->EnableWindow(FALSE);
	h0 = m_H - m_as;
	m_N = m_N * 1e3;
	//MessageBox(m_CS);
	info.initPY(m_CS, m_ST);
	m_Nb = info.a1*info.cp*m_B*h0*info.kb + info.ss*m_SAsp - info.ss*m_SAs;
	if (m_N <= m_Nb)
	{
		//大偏压
		m_PYT = _T("大偏压");
		x = (m_N-info.ss*m_SAsp+info.ss*m_SAs) / (info.a1*info.cp*m_B);
		tempkesai = x / h0;
		if (tempkesai>2*m_as/h0)
		{
			//输出数据（1），（2），（3），（5），（6），（7），（9）；
			//输出计算书（计算书文件名DPA-FH-ZC）
			exportcon = 1;
			m_e = (info.a1*info.cp*m_B*pow(h0,2)*tempkesai*(1-0.5*tempkesai)+
				info.ss*m_SAsp*(h0-m_as)) / (m_N);
			m_ei = m_e - 0.5*m_H + m_as;
			m_ea = max(m_H / 30, 20);
			m_e0 = m_ei - m_ea;
			m_Mu = m_N * m_e0;
			GetDlgItem(IDC_BUTTON_export)->EnableWindow(TRUE);
		}
		else
		{
			//若 ，出现提示语“受压钢筋不屈服”
			//输出数据输出数据（1），（2），（4），（5），（6），（7），（9）
			//输出计算书（计算书文件名DPA-FH-XX）
			MessageBox(_T("受压钢筋不屈服"));
			m_e = 0.0;
			exportcon = 2;
			m_ep = info.ss*m_SAs*(h0 - m_as)/m_N;
			m_ei = m_ep + 0.5*m_H - m_as;
			m_ea = max(m_H / 30, 20);
			m_e0 = m_ei - m_ea;
			m_Mu = m_N * m_e0;
			GetDlgItem(IDC_BUTTON_export)->EnableWindow(TRUE);
		}
	}
	else
	{
		m_PYT = _T("小偏压");
		m_kscy = 1.6 - info.kb;
		tempkesai = (m_N-info.ss*m_SAsp-(0.8)/(info.kb-info.b1)*
			info.ss*m_SAs) / (info.a1*info.cp*m_B*h0-1/(info.kb-info.b1)*info.ss*m_SAs);
		if (tempkesai<=m_kscy)
		{
			//输出数据输出数据（1），（2），（3），（5），（6），（7），（9），
			//输出计算书（计算书文件名XPA-FH-ZC）
			exportcon = 3;
			m_e = (info.a1*info.cp*m_B*pow(h0,2)*tempkesai*(1-0.5*tempkesai)+info.ss*m_SAsp*
				(h0-m_as)) / (m_N);
			m_ei = m_e - 0.5*m_H + m_as;
			m_ea = max(m_H / 30, 20);
			m_e0 = m_ei - m_ea;
			m_Mu = m_N * m_e0;
			GetDlgItem(IDC_BUTTON_export)->EnableWindow(TRUE);
		}
		else if (tempkesai<=(m_H/h0))
		{
			//输出数据（1），（2），（3），（5），（6），（7），（8），（9），
			//无计算书输出
			m_mks = (m_N - info.ss*m_SAsp-info.ss*m_SAs) / (info.a1*info.cp*m_B*h0);
			m_e = (info.a1*info.cp*m_B*pow(h0, 2)*m_mks*(1 - 0.5*m_mks) + info.ss*m_SAsp*
				(h0 - m_as)) / (m_N);
			m_ei = m_e - 0.5*m_H + m_as;
			m_ea = max(m_H / 30, 20);
			m_e0 = m_ei - m_ea;
			m_Mu = m_N * m_e0;
		}
		else
		{
			//输出数据（1），（2），（3），（5），（6），（7），（9），无计算书输出
			m_e = info.cp*m_B*m_H*(h0 - 0.5*m_H) + info.ss*m_SAsp*(h0 - m_as);
			m_e = m_e / m_N;
			m_ei = m_e - 0.5*m_H + m_as;
			m_ea = max(m_H / 30, 20);
			m_e0 = m_ei - m_ea;
			m_Mu = m_N * m_e0;
		}
	}
	m_N /= 1000;
	m_Nb /= 1000;
	m_Mu /= 1e6;
	m_iks = tempkesai;
	CWnd::UpdateData(false);
}


void Cpyc::OnBnClickedButtonexport()
{
	//公共字符
	
	CString S_b, S_h, S_n, S_Asp, S_As, S_as, S_fc, S_kb, S_fy, S_h0;
	S_b.Format(_T("%.0lf"), m_B);
	S_h.Format(_T("%.0lf"), m_H);
	S_n.Format(_T("%.0lf"), m_N);
	S_Asp.Format(_T("%.0lf"), m_SAsp);
	S_As.Format(_T("%.0lf"), m_SAs);
	S_as.Format(_T("%.0lf"), m_as);
	S_fc.Format(_T("%.1lf"), info.cp);
	S_kb.Format(_T("%.3lf"), info.kb);
	S_fy.Format(_T("%.0lf"), info.ss);
	S_h0.Format(_T("%.0lf"), h0);

	CString S_e0, S_e, S_ea, S_ei, S_ksai, S_Mu, S_nb, S_x, S_ep, S_ksaicy;
	S_e0.Format(_T("%.1lf"), m_e0);
	S_e.Format(_T("%.1lf"), m_e);
	S_ea.Format(_T("%.1lf"), m_ea);
	S_ei.Format(_T("%.1lf"), m_ei);
	S_ksai.Format(_T("%.3lf"), m_iks);
	S_Mu.Format(_T("%.0lf"), m_Mu);
	S_nb.Format(_T("%.1lf"), m_Nb);
	S_x.Format(_T("%.0lf"), x);
	S_ep.Format(_T("%.1lf"), m_ep);
	S_ksaicy.Format(_T("%.3lf"), m_kscy);


	if (exportcon == 1)
	{
		//--------------------------------------------------------
		//输出计算书（计算书文件名DPA-FH-ZC）
		COperateWord conword(g_exePATH, _T("\\templet\\DPY_FH_ZC.dot"));
		if (!conword.IsSuccess())
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}
		conword.ReplaceBookmark(_T("as"), S_as, 3);
		conword.ReplaceBookmark(_T("Asp"), S_Asp, 4);
		conword.ReplaceBookmark(_T("Ast"), S_As, 3);
		conword.ReplaceBookmark(_T("b"), S_b, 4);
		conword.ReplaceBookmark(_T("CT"), m_CS, 1);
		conword.ReplaceBookmark(_T("e0"), S_e0, 2);
		conword.ReplaceBookmark(_T("e"), S_e, 2);
		conword.ReplaceBookmark(_T("ea"), S_ea, 2);
		conword.ReplaceBookmark(_T("ei"), S_ei, 2);
		conword.ReplaceBookmark(_T("fc"), S_fc, 4);
		conword.ReplaceBookmark(_T("fy"), S_fy, 4);
		conword.ReplaceBookmark(_T("h0"), S_h0, 5);
		conword.ReplaceBookmark(_T("h"), S_h, 2);
		conword.ReplaceBookmark(_T("ksai"), S_ksai, 3);
		conword.ReplaceBookmark(_T("ksaib"), S_kb, 2);
		conword.ReplaceBookmark(_T("Mu"), S_Mu, 1);
		conword.ReplaceBookmark(_T("N"), S_n, 4);
		conword.ReplaceBookmark(_T("Nb"), S_nb, 1);
		conword.ReplaceBookmark(_T("ST"), m_ST, 1);
		conword.ReplaceBookmark(_T("x"), S_x, 2);
	}
	else if (exportcon == 2)
	{
		//输出计算书（计算书文件名DPA-FH-XX）
		COperateWord conword(g_exePATH, _T("\\templet\\DPY_FH_XX.dot"));
		if (!conword.IsSuccess())
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}
		conword.ReplaceBookmark(_T("as"), S_as, 3);
		conword.ReplaceBookmark(_T("Asp"), S_Asp, 3);
		conword.ReplaceBookmark(_T("Ast"), S_As, 4);
		conword.ReplaceBookmark(_T("b"), S_b, 3);
		conword.ReplaceBookmark(_T("CT"), m_CS, 1);
		conword.ReplaceBookmark(_T("e0"), S_e0, 2);
		conword.ReplaceBookmark(_T("e"), S_e, 2);
		conword.ReplaceBookmark(_T("ea"), S_ea, 2);
		conword.ReplaceBookmark(_T("ei"), S_ei, 2);
		conword.ReplaceBookmark(_T("ep"), S_ep, 2);
		conword.ReplaceBookmark(_T("fc"), S_fc, 3);
		conword.ReplaceBookmark(_T("fy"), S_fy, 4);
		conword.ReplaceBookmark(_T("h0"), S_h0, 4);
		conword.ReplaceBookmark(_T("h"), S_h, 2);
		conword.ReplaceBookmark(_T("ksai"), S_ksai, 1);
		conword.ReplaceBookmark(_T("ksaib"), S_kb, 2);
		conword.ReplaceBookmark(_T("Mu"), S_Mu, 1);
		conword.ReplaceBookmark(_T("N"), S_n, 4);
		conword.ReplaceBookmark(_T("Nb"), S_nb, 1);
		conword.ReplaceBookmark(_T("ST"), m_ST, 1);
		conword.ReplaceBookmark(_T("x"), S_x, 2);
	}
	else
	{
		//输出计算书（计算书文件名XPA-FH-ZC）

		COperateWord conword(g_exePATH, _T("\\templet\\XPY_FH_ZC.dot"));
		if (!conword.IsSuccess())
		{
			AfxMessageBox(_T("本机没有安装word产品！"));
			return;
		}
		conword.ReplaceBookmark(_T("as"), S_as, 3);
		conword.ReplaceBookmark(_T("Asp"), S_Asp, 4);
		conword.ReplaceBookmark(_T("Ast"), S_As, 4);
		conword.ReplaceBookmark(_T("b"), S_b, 4);
		conword.ReplaceBookmark(_T("CT"), m_CS, 1);
		conword.ReplaceBookmark(_T("e0"), S_e0, 2);
		conword.ReplaceBookmark(_T("e"), S_e, 2);
		conword.ReplaceBookmark(_T("ea"), S_ea, 2);
		conword.ReplaceBookmark(_T("ei"), S_ei, 2);
		conword.ReplaceBookmark(_T("fc"), S_fc, 4);
		conword.ReplaceBookmark(_T("fy"), S_fy, 6);
		conword.ReplaceBookmark(_T("h0"), S_h0, 5);
		conword.ReplaceBookmark(_T("h"), S_h, 2);
		conword.ReplaceBookmark(_T("ksai"), S_ksai, 2);
		conword.ReplaceBookmark(_T("ksaib"), S_kb, 4);
		conword.ReplaceBookmark(_T("ksaicy"), S_ksaicy, 1);
		conword.ReplaceBookmark(_T("Mu"), S_Mu, 1);
		conword.ReplaceBookmark(_T("N"), S_n, 4);
		conword.ReplaceBookmark(_T("Nb"), S_nb, 1);
		conword.ReplaceBookmark(_T("ST"), m_ST, 1);
	}
}


BOOL Cpyc::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	TCHAR path[MAX_PATH] = { 0 };

	GetModuleFileName(NULL, path, MAX_PATH);


	g_exePATH = path;//此时获得了EXE的目录

	return TRUE;  // return TRUE unless you set the focus to a control
				  // 异常: OCX 属性页应返回 FALSE
}

/*
* // TODO: 在此添加控件通知处理程序代码
	CApplication wordApp;
	CDocuments docs;
	CDocument0 docx;
	CBookmarks bookmarks;
	CBookmark0 bookmark;
	CRange range;
	CString temppath;
	temppath = g_exePATH;

	//公共字符
	CString S_b, S_h, S_n, S_Asp, S_As, S_as, S_fc, S_kb, S_fy, S_h0;
	S_b.Format(_T("%.0lf"), m_B);
	S_h.Format(_T("%.0lf"), m_H);
	S_n.Format(_T("%.0lf"), m_N);
	S_Asp.Format(_T("%.0lf"), m_SAsp);
	S_As.Format(_T("%.0lf"), m_SAs);
	S_as.Format(_T("%.0lf"), m_as);
	S_fc.Format(_T("%.1lf"), info.cp);
	S_kb.Format(_T("%.3lf"), info.kb);
	S_fy.Format(_T("%.0lf"), info.ss);
	S_h0.Format(_T("%.0lf"), h0);

	CString S_e0, S_e, S_ea, S_ei, S_ksai, S_Mu, S_nb, S_x, S_ep, S_ksaicy;
	S_e0.Format(_T("%.1lf"), m_e0);
	S_e.Format(_T("%.1lf"), m_e);
	S_ea.Format(_T("%.1lf"), m_ea);
	S_ei.Format(_T("%.1lf"), m_ei);
	S_ksai.Format(_T("%.3lf"), m_iks);
	S_Mu.Format(_T("%.0lf"), m_Mu);
	S_nb.Format(_T("%.1lf"), m_Nb);
	S_x.Format(_T("%.0lf"), x);
	S_ep.Format(_T("%.1lf"), m_ep);
	S_ksaicy.Format(_T("%.3lf"), m_kscy);

	if (exportcon == 1)
	{
		//--------------------------------------------------------
		//输出计算书（计算书文件名DPA-FH-ZC）
		temppath.Replace(_T("\\peijin.exe"), _T("\\templet\\DPY_FH_ZC.dot"));

		CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
		COleVariant covZero((short)0),
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

		//-----------------------------------------------------------------------------
		bookmark = bookmarks.Item(&_variant_t(_T("as1")));
		range = bookmark.get_Range();
		range.put_Text(S_as);

		bookmark = bookmarks.Item(&_variant_t(_T("as2")));
		range = bookmark.get_Range();
		range.put_Text(S_as);

		bookmark = bookmarks.Item(&_variant_t(_T("as3")));
		range = bookmark.get_Range();
		range.put_Text(S_as);

		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);

		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);

		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);

		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);

		bookmark = bookmarks.Item(&_variant_t(_T("Ast1")));
		range = bookmark.get_Range();
		range.put_Text(S_As);

		bookmark = bookmarks.Item(&_variant_t(_T("Ast2")));
		range = bookmark.get_Range();
		range.put_Text(S_As);

		bookmark = bookmarks.Item(&_variant_t(_T("Ast3")));
		range = bookmark.get_Range();
		range.put_Text(S_As);

		bookmark = bookmarks.Item(&_variant_t(_T("b1")));
		range = bookmark.get_Range();
		range.put_Text(S_b);

		bookmark = bookmarks.Item(&_variant_t(_T("b2")));
		range = bookmark.get_Range();
		range.put_Text(S_b);

		bookmark = bookmarks.Item(&_variant_t(_T("b3")));
		range = bookmark.get_Range();
		range.put_Text(S_b);

		bookmark = bookmarks.Item(&_variant_t(_T("b4")));
		range = bookmark.get_Range();
		range.put_Text(S_b);

		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(m_CS);

		bookmark = bookmarks.Item(&_variant_t(_T("e01")));
		range = bookmark.get_Range();
		range.put_Text(S_e0);

		bookmark = bookmarks.Item(&_variant_t(_T("e02")));
		range = bookmark.get_Range();
		range.put_Text(S_e0);

		bookmark = bookmarks.Item(&_variant_t(_T("e1")));
		range = bookmark.get_Range();
		range.put_Text(S_e);

		bookmark = bookmarks.Item(&_variant_t(_T("e2")));
		range = bookmark.get_Range();
		range.put_Text(S_e);

		bookmark = bookmarks.Item(&_variant_t(_T("ea1")));
		range = bookmark.get_Range();
		range.put_Text(S_ea);

		bookmark = bookmarks.Item(&_variant_t(_T("ea2")));
		range = bookmark.get_Range();
		range.put_Text(S_ea);

		bookmark = bookmarks.Item(&_variant_t(_T("ei1")));
		range = bookmark.get_Range();
		range.put_Text(S_ei);

		bookmark = bookmarks.Item(&_variant_t(_T("ei2")));
		range = bookmark.get_Range();
		range.put_Text(S_ei);

		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);

		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc4")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);

		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy3")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy4")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);

		bookmark = bookmarks.Item(&_variant_t(_T("h01")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h02")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h03")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h04")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h05")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);

		bookmark = bookmarks.Item(&_variant_t(_T("h1")));
		range = bookmark.get_Range();
		range.put_Text(S_h);
		bookmark = bookmarks.Item(&_variant_t(_T("h2")));
		range = bookmark.get_Range();
		range.put_Text(S_h);

		bookmark = bookmarks.Item(&_variant_t(_T("ksai1")));
		range = bookmark.get_Range();
		range.put_Text(S_ksai);
		bookmark = bookmarks.Item(&_variant_t(_T("ksai2")));
		range = bookmark.get_Range();
		range.put_Text(S_ksai);
		bookmark = bookmarks.Item(&_variant_t(_T("ksai3")));
		range = bookmark.get_Range();
		range.put_Text(S_ksai);

		bookmark = bookmarks.Item(&_variant_t(_T("ksaib1")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);
		bookmark = bookmarks.Item(&_variant_t(_T("ksaib2")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);

		bookmark = bookmarks.Item(&_variant_t(_T("Mu1")));
		range = bookmark.get_Range();
		range.put_Text(S_Mu);

		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N3")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N4")));
		range = bookmark.get_Range();
		range.put_Text(S_n);

		bookmark = bookmarks.Item(&_variant_t(_T("Nb1")));
		range = bookmark.get_Range();
		range.put_Text(S_nb);

		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(m_ST);

		bookmark = bookmarks.Item(&_variant_t(_T("x1")));
		range = bookmark.get_Range();
		range.put_Text(S_x);
		bookmark = bookmarks.Item(&_variant_t(_T("x2")));
		range = bookmark.get_Range();
		range.put_Text(S_x);

	}
	else if (exportcon == 2)
	{
		//输出计算书（计算书文件名DPA-FH-XX）
		temppath.Replace(_T("\\peijin.exe"), _T("\\templet\\DPY_FH_XX.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("as1")));
		range = bookmark.get_Range();
		range.put_Text(S_as);
		bookmark = bookmarks.Item(&_variant_t(_T("as2")));
		range = bookmark.get_Range();
		range.put_Text(S_as);
		bookmark = bookmarks.Item(&_variant_t(_T("as3")));
		range = bookmark.get_Range();
		range.put_Text(S_as);

		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);

		bookmark = bookmarks.Item(&_variant_t(_T("Ast1")));
		range = bookmark.get_Range();
		range.put_Text(S_As);
		bookmark = bookmarks.Item(&_variant_t(_T("Ast2")));
		range = bookmark.get_Range();
		range.put_Text(S_As);
		bookmark = bookmarks.Item(&_variant_t(_T("Ast3")));
		range = bookmark.get_Range();
		range.put_Text(S_As);
		bookmark = bookmarks.Item(&_variant_t(_T("Ast4")));
		range = bookmark.get_Range();
		range.put_Text(S_As);

		bookmark = bookmarks.Item(&_variant_t(_T("b1")));
		range = bookmark.get_Range();
		range.put_Text(S_b);
		bookmark = bookmarks.Item(&_variant_t(_T("b2")));
		range = bookmark.get_Range();
		range.put_Text(S_b);
		bookmark = bookmarks.Item(&_variant_t(_T("b3")));
		range = bookmark.get_Range();
		range.put_Text(S_b);

		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(m_CS);

		bookmark = bookmarks.Item(&_variant_t(_T("e01")));
		range = bookmark.get_Range();
		range.put_Text(S_e0);
		bookmark = bookmarks.Item(&_variant_t(_T("e02")));
		range = bookmark.get_Range();
		range.put_Text(S_e0);

		bookmark = bookmarks.Item(&_variant_t(_T("ea1")));
		range = bookmark.get_Range();
		range.put_Text(S_ea);
		bookmark = bookmarks.Item(&_variant_t(_T("ea2")));
		range = bookmark.get_Range();
		range.put_Text(S_ea);

		bookmark = bookmarks.Item(&_variant_t(_T("ei1")));
		range = bookmark.get_Range();
		range.put_Text(S_ei);
		bookmark = bookmarks.Item(&_variant_t(_T("ei2")));
		range = bookmark.get_Range();
		range.put_Text(S_ei);

		bookmark = bookmarks.Item(&_variant_t(_T("ep1")));
		range = bookmark.get_Range();
		range.put_Text(S_ep);
		bookmark = bookmarks.Item(&_variant_t(_T("ep2")));
		range = bookmark.get_Range();
		range.put_Text(S_ep);

		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);

		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy3")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy4")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);

		bookmark = bookmarks.Item(&_variant_t(_T("h01")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h02")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h03")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h04")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);

		bookmark = bookmarks.Item(&_variant_t(_T("h1")));
		range = bookmark.get_Range();
		range.put_Text(S_h);
		bookmark = bookmarks.Item(&_variant_t(_T("h2")));
		range = bookmark.get_Range();
		range.put_Text(S_h);

		bookmark = bookmarks.Item(&_variant_t(_T("ksai1")));
		range = bookmark.get_Range();
		range.put_Text(S_ksai);

		bookmark = bookmarks.Item(&_variant_t(_T("ksaib1")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);
		bookmark = bookmarks.Item(&_variant_t(_T("ksaib2")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);

		bookmark = bookmarks.Item(&_variant_t(_T("Mu1")));
		range = bookmark.get_Range();
		range.put_Text(S_Mu);

		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N3")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N4")));
		range = bookmark.get_Range();
		range.put_Text(S_n);

		bookmark = bookmarks.Item(&_variant_t(_T("Nb1")));
		range = bookmark.get_Range();
		range.put_Text(S_nb);

		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(m_ST);

		bookmark = bookmarks.Item(&_variant_t(_T("x1")));
		range = bookmark.get_Range();
		range.put_Text(S_x);
		bookmark = bookmarks.Item(&_variant_t(_T("x2")));
		range = bookmark.get_Range();
		range.put_Text(S_x);
	}
	else
	{
		//输出计算书（计算书文件名XPA-FH-ZC）
		temppath.Replace(_T("\\peijin.exe"), _T("\\templet\\XPY_FH_ZC.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("as1")));
		range = bookmark.get_Range();
		range.put_Text(S_as);
		bookmark = bookmarks.Item(&_variant_t(_T("as2")));
		range = bookmark.get_Range();
		range.put_Text(S_as);
		bookmark = bookmarks.Item(&_variant_t(_T("as3")));
		range = bookmark.get_Range();
		range.put_Text(S_as);

		bookmark = bookmarks.Item(&_variant_t(_T("Asp1")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp2")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp3")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);
		bookmark = bookmarks.Item(&_variant_t(_T("Asp4")));
		range = bookmark.get_Range();
		range.put_Text(S_Asp);

		bookmark = bookmarks.Item(&_variant_t(_T("Ast1")));
		range = bookmark.get_Range();
		range.put_Text(S_As);
		bookmark = bookmarks.Item(&_variant_t(_T("Ast2")));
		range = bookmark.get_Range();
		range.put_Text(S_As);
		bookmark = bookmarks.Item(&_variant_t(_T("Ast3")));
		range = bookmark.get_Range();
		range.put_Text(S_As);
		bookmark = bookmarks.Item(&_variant_t(_T("Ast4")));
		range = bookmark.get_Range();
		range.put_Text(S_As);

		bookmark = bookmarks.Item(&_variant_t(_T("b1")));
		range = bookmark.get_Range();
		range.put_Text(S_b);
		bookmark = bookmarks.Item(&_variant_t(_T("b2")));
		range = bookmark.get_Range();
		range.put_Text(S_b);
		bookmark = bookmarks.Item(&_variant_t(_T("b3")));
		range = bookmark.get_Range();
		range.put_Text(S_b);
		bookmark = bookmarks.Item(&_variant_t(_T("b4")));
		range = bookmark.get_Range();
		range.put_Text(S_b);

		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(m_CS);

		bookmark = bookmarks.Item(&_variant_t(_T("e01")));
		range = bookmark.get_Range();
		range.put_Text(S_e0);
		bookmark = bookmarks.Item(&_variant_t(_T("e02")));
		range = bookmark.get_Range();
		range.put_Text(S_e0);

		bookmark = bookmarks.Item(&_variant_t(_T("e1")));
		range = bookmark.get_Range();
		range.put_Text(S_e);
		bookmark = bookmarks.Item(&_variant_t(_T("e2")));
		range = bookmark.get_Range();
		range.put_Text(S_e);

		bookmark = bookmarks.Item(&_variant_t(_T("ea1")));
		range = bookmark.get_Range();
		range.put_Text(S_ea);
		bookmark = bookmarks.Item(&_variant_t(_T("ea2")));
		range = bookmark.get_Range();
		range.put_Text(S_ea);

		bookmark = bookmarks.Item(&_variant_t(_T("ei1")));
		range = bookmark.get_Range();
		range.put_Text(S_ei);
		bookmark = bookmarks.Item(&_variant_t(_T("ei2")));
		range = bookmark.get_Range();
		range.put_Text(S_ei);

		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc3")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc4")));
		range = bookmark.get_Range();
		range.put_Text(S_fc);

		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy3")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy4")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy5")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy6")));
		range = bookmark.get_Range();
		range.put_Text(S_fy);

		bookmark = bookmarks.Item(&_variant_t(_T("h01")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h02")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h03")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h04")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);
		bookmark = bookmarks.Item(&_variant_t(_T("h05")));
		range = bookmark.get_Range();
		range.put_Text(S_h0);

		bookmark = bookmarks.Item(&_variant_t(_T("h1")));
		range = bookmark.get_Range();
		range.put_Text(S_h);
		bookmark = bookmarks.Item(&_variant_t(_T("h2")));
		range = bookmark.get_Range();
		range.put_Text(S_h);

		bookmark = bookmarks.Item(&_variant_t(_T("ksai1")));
		range = bookmark.get_Range();
		range.put_Text(S_ksai);
		bookmark = bookmarks.Item(&_variant_t(_T("ksai2")));
		range = bookmark.get_Range();
		range.put_Text(S_ksai);

		bookmark = bookmarks.Item(&_variant_t(_T("ksaib1")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);
		bookmark = bookmarks.Item(&_variant_t(_T("ksaib2")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);
		bookmark = bookmarks.Item(&_variant_t(_T("ksaib3")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);
		bookmark = bookmarks.Item(&_variant_t(_T("ksaib4")));
		range = bookmark.get_Range();
		range.put_Text(S_kb);

		bookmark = bookmarks.Item(&_variant_t(_T("ksaicy1")));
		range = bookmark.get_Range();
		range.put_Text(S_ksaicy);

		bookmark = bookmarks.Item(&_variant_t(_T("Mu1")));
		range = bookmark.get_Range();
		range.put_Text(S_Mu);

		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N3")));
		range = bookmark.get_Range();
		range.put_Text(S_n);
		bookmark = bookmarks.Item(&_variant_t(_T("N4")));
		range = bookmark.get_Range();
		range.put_Text(S_n);

		bookmark = bookmarks.Item(&_variant_t(_T("Nb1")));
		range = bookmark.get_Range();
		range.put_Text(S_nb);

		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(m_ST);
	}


	range.ReleaseDispatch();
	bookmark.ReleaseDispatch();
	bookmarks.ReleaseDispatch();
	docx.ReleaseDispatch();
	docs.ReleaseDispatch();
	wordApp.ReleaseDispatch();
*/