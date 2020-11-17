// Czyyd.cpp : 实现文件
//

#include "stdafx.h"
#include "peijin.h"
#include "Czyyd.h"
#include "afxdialogex.h"


// Czyyd 对话框

IMPLEMENT_DYNAMIC(Czyyd, CDialogEx)

Czyyd::Czyyd(CWnd* pParent /*=NULL*/)
	: CDialogEx(Czyyd::IDD, pParent)
{

	m_CT = _T("C30");
	m_As = 6382;
	m_d = 450;
	m_Gd = 10;
	m_H = 4.5;
	m_N = 4800;
	m_pas = 20;
	m_s = 0.0;
	m_GST = _T("HRB335");
	m_ST = _T("HRB400");
	m_sp = 0.0;
}

Czyyd::~Czyyd()
{
}

void Czyyd::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_CBString(pDX, IDC_CT, m_CT);
	DDX_Text(pDX, IDC_EDIT_As, m_As);
	DDX_Text(pDX, IDC_EDIT_d, m_d);
	DDX_Text(pDX, IDC_EDIT_GD, m_Gd);
	DDX_Text(pDX, IDC_EDIT_H, m_H);
	DDX_Text(pDX, IDC_EDIT_N, m_N);
	DDX_Text(pDX, IDC_EDIT_pas, m_pas);
	DDX_Text(pDX, IDC_EDIT_S, m_s);
	DDX_CBString(pDX, IDC_GST, m_GST);
	DDX_CBString(pDX, IDC_ST, m_ST);
	DDX_Text(pDX, IDC_EDIT_SP, m_sp);
}


BEGIN_MESSAGE_MAP(Czyyd, CDialogEx)
	ON_BN_CLICKED(IDC_Calcula, &Czyyd::OnBnClickedCalcula)
	ON_BN_CLICKED(IDC_Out, &Czyyd::OnBnClickedOut)
END_MESSAGE_MAP()


// Czyyd 消息处理程序


void Czyyd::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	OnBnClickedCalcula();
	//CDialogEx::OnOK();
}


BOOL Czyyd::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	((CButton *)GetDlgItem(IDC_RADIO_TJ))->SetCheck(TRUE); //选上

	TCHAR path[MAX_PATH] = {0};

	GetModuleFileName(NULL,path,MAX_PATH);


	g_exePATH=path;//此时获得了EXE的目录

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void Czyyd::OnBnClickedCalcula()
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

	m_l0 = m_H * namda *1000;
	temp.Format(_T("%.1lf"),m_l0/m_d);

	if(m_l0/m_d < 12)
		MessageBox(_T("计算长细比为:")+temp+_T("<12，可以采用螺旋箍筋柱"));
	else
	{
		MessageBox(_T("计算长细比为:")+temp+_T(">12，不能采用螺旋箍筋柱"));
		return;
	}

	A = 1.0/4 * 3.1415926 * m_d * m_d;
	p = m_As / A;
	temp.Format(_T("%.2lf"),p*100);
	info.initYSZY(m_CT,m_ST,m_GST,m_l0,m_d,m_Gd);

	if(p < 0.05)
	{
		if(p < info.pminp)
		{
			MessageBox(_T("所配纵筋不满足最小配筋要求，请重新选择"));
			return;
		}
		MessageBox(_T("计算配筋率为:")+temp+_T("%<5%,满足配筋率要求"));
	}
	else
	{
		MessageBox(_T("计算配筋率为:")+temp+_T("%>5%,请重新选择纵筋面积"));
		return;
	}

	dcor = m_d - 2 * (m_Gd + m_pas);
	Acor = 1.0/4 * 3.1415926 * dcor * dcor;
	//info.initYSZY(m_CT,m_ST,m_GST,m_l0,m_d);

	Ass0 = ((m_N * 1000 / 0.9) - (info.cp * Acor + info.ss * m_As))/(2 * info.yszya * info.gjss);

	temp.Format(_T("%.1lf"),Ass0);
	if(Ass0 > 0.25 * m_As)
		MessageBox(_T("计算得Ass0=")+temp+_T(">0.25As',满足构造要求"));
	else
	{
		MessageBox(_T("计算得Ass0=")+temp+_T("<0.25As',不满足构造要求，请重新选取纵筋面积"));
		return;
	}
	m_s = (3.1415926 * dcor * 0.25 * 3.1415926 * m_Gd * m_Gd)/(Ass0);
	temp.Format(_T("%.1lf"),m_s);
	if((m_s <= ((80>0.2*dcor)?0.2*dcor:80)) && (m_s >= 40))
		MessageBox(_T("计算得:s=")+temp+_T("mm,满足不小于40mm，并不大于80mm及0.2dcor的要求"));
	else
	{
		MessageBox(_T("计算得:s=")+temp+_T("mm,不满足不小于40mm，并不大于80mm及0.2dcor的要求"));
		return;
	}
	m_sp = ((int)m_s/5)*5;
	//复核
	m_fAss0 = (3.1415926 * dcor * 0.25 * 3.1415926 * m_Gd * m_Gd)/(m_sp);
	//Ass0 = ((m_N * 1000 / 0.9) - (info.cp * Acor + info.ss * m_As))/(2 * info.yszya * info.gjss);
	m_Nu1 = 0.9 * (info.cp*Acor+2*info.yszya*info.gjss*m_fAss0+info.ss*m_As);
	if(m_As/A > 0.03)
	{
		MessageBox(_T("纵筋配筋率大于3%，需要修正公式"));
		//大于 3%
		con1 = 1;
		//Nu2 = 0.9Ф(fcA – fcAs +f’yA’s)
		m_Nu2 = 0.9*info.fi*(info.cp*(A - m_As)+info.ss*m_As);
		//Nu1> Nu2 ，且Nu1<1.5 Nu2
		if((m_Nu1>m_Nu2) && (m_Nu1<1.5*m_Nu2))
			MessageBox(_T("验算满足要求"));
		else
		{
			MessageBox(_T("验算不满足要求"));
			return;
		}
	}
	else
	{
		MessageBox(_T("纵筋配筋率小于3%，不需要修正公式"));
		//小于 3%
		con1 = 2;
		m_Nu2 = 0.9*info.fi*(info.cp*A+info.ss*m_As);
		if((m_Nu1>m_Nu2) && (m_Nu1<1.5*m_Nu2))
			MessageBox(_T("验算满足要求"));
		else
		{
			MessageBox(_T("验算不满足要求"));
			return;
		}
	}

	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
	CWnd::UpdateData(false);
}


void Czyyd::OnBnClickedOut()
{
	CApplication wordApp;
	CDocuments docs;
	CDocument0 docx;
	CBookmarks bookmarks;
	CBookmark0 bookmark;
	CRange range;
	CString temppath;
	temppath = g_exePATH;


	if(con1 == 1)
	{
		//绑定参数
		CString snamda,sl01,sl02,sl0cyd,ssp,sAcor,sarfa,sAs,sAsl,sAsso,sCT,sd,sdcor,sfc,sfy,sGd,sGfy,sGST,sH,sN,spas,ss,ssfzyAs,sST;
		sAcor.Format(_T("%.1lf"),Acor/10000);
		sarfa.Format(_T("%.1lf"),info.yszya);
		sAs.Format(_T("%.0lf"),m_As);
		sAsl.Format(_T("%.1lf"),0.25*3.1415926*m_Gd*m_Gd);
		sAsso.Format(_T("%.0lf"),Ass0);
		sCT = m_CT;
		sd.Format(_T("%.0lf"),m_d);
		sdcor.Format(_T("%.0lf"),dcor);
		sfc.Format(_T("%.1lf"),info.cp);
		sfy.Format(_T("%.0lf"),info.ss);
		sGd.Format(_T("%.0lf"),m_Gd);
		sGfy.Format(_T("%.0lf"),info.gjss);
		sGST = m_GST;
		sH.Format(_T("%.1lf"),m_H);
		sN.Format(_T("%.0lf"),m_N);
		spas.Format(_T("%.0lf"),m_pas);
		ss.Format(_T("%.1lf"),m_s);
		ssfzyAs.Format(_T("%.1lf"),0.25*m_As);
		sST = m_ST;
		snamda.Format(_T("%.1lf"),namda);
		sl01.Format(_T("%.1lf"),m_l0/1000);
		sl02.Format(_T("%.0lf"),m_l0);
		sl0cyd.Format(_T("%.1lf"),m_l0/m_d);
		ssp.Format(_T("%.0lf"),m_sp);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\lxsjd3.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("Acor1")));
		range = bookmark.get_Range();
		range.put_Text(sAcor);
		bookmark = bookmarks.Item(&_variant_t(_T("Acor2")));
		range = bookmark.get_Range();
		range.put_Text(sAcor);
		bookmark = bookmarks.Item(&_variant_t(_T("arfa1")));
		range = bookmark.get_Range();
		range.put_Text(sarfa);
		bookmark = bookmarks.Item(&_variant_t(_T("arfa2")));
		range = bookmark.get_Range();
		range.put_Text(sarfa);
		bookmark = bookmarks.Item(&_variant_t(_T("As1")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("As2")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("As3")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("Asl1")));
		range = bookmark.get_Range();
		range.put_Text(sAsl);
		bookmark = bookmarks.Item(&_variant_t(_T("Asso1")));
		range = bookmark.get_Range();
		range.put_Text(sAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("Asso2")));
		range = bookmark.get_Range();
		range.put_Text(sAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("Asso3")));
		range = bookmark.get_Range();
		range.put_Text(sAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(sCT);
		bookmark = bookmarks.Item(&_variant_t(_T("CT2")));
		range = bookmark.get_Range();
		range.put_Text(sCT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(sd);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(sd);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(sd);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor3")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Gd1")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("Gd2")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("Gfy1")));
		range = bookmark.get_Range();
		range.put_Text(sGfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Gfy2")));
		range = bookmark.get_Range();
		range.put_Text(sGfy);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(sGST);
		bookmark = bookmarks.Item(&_variant_t(_T("H1")));
		range = bookmark.get_Range();
		range.put_Text(sH);
		bookmark = bookmarks.Item(&_variant_t(_T("H2")));
		range = bookmark.get_Range();
		range.put_Text(sH);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(sN);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(sN);
		bookmark = bookmarks.Item(&_variant_t(_T("pas1")));
		range = bookmark.get_Range();
		range.put_Text(spas);
		bookmark = bookmarks.Item(&_variant_t(_T("pas2")));
		range = bookmark.get_Range();
		range.put_Text(spas);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(ss);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("sfzyAs")));
		range = bookmark.get_Range();
		range.put_Text(ssfzyAs);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(sST);
		bookmark = bookmarks.Item(&_variant_t(_T("namda1")));
		range = bookmark.get_Range();
		range.put_Text(snamda);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(sl01);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(sl02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd")));
		range = bookmark.get_Range();
		range.put_Text(sl0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("FGd1")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("FGd2")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("Fs1")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("Fs2")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("Fs3")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("Fdcor1")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsl1")));
		range = bookmark.get_Range();
		range.put_Text(sAsl);

		CString sfAsso;
		sfAsso.Format(_T("%.0lf"),m_fAss0);

		bookmark = bookmarks.Item(&_variant_t(_T("FAsso2")));
		range = bookmark.get_Range();
		range.put_Text(sfAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsso1")));
		range = bookmark.get_Range();
		range.put_Text(sfAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffc1")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffc2")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffc3")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("FAcor1")));
		range = bookmark.get_Range();
		range.put_Text(sAcor);
		bookmark = bookmarks.Item(&_variant_t(_T("Farfa1")));
		range = bookmark.get_Range();
		range.put_Text(sarfa);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffyv1")));
		range = bookmark.get_Range();
		range.put_Text(sGfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffy1")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffy2")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsp1")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsp2")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsp3")));
		range = bookmark.get_Range();
		range.put_Text(sAs);

		CString sNu1,sNu2,sfi,sA,sbNu2;
		sNu1.Format(_T("%.0lf"),m_Nu1/1000);
		sNu2.Format(_T("%.0lf"),m_Nu2/1000);
		sfi.Format(_T("%.4lf"),info.fi);
		sA.Format(_T("%.1lf"),A/10000);
		sbNu2.Format(_T("%.0lf"),1.5*m_Nu2/1000);

		bookmark = bookmarks.Item(&_variant_t(_T("FNu11")));
		range = bookmark.get_Range();
		range.put_Text(sNu1);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffi1")));
		range = bookmark.get_Range();
		range.put_Text(sfi);
		bookmark = bookmarks.Item(&_variant_t(_T("FA1")));
		range = bookmark.get_Range();
		range.put_Text(sA);
		bookmark = bookmarks.Item(&_variant_t(_T("FNu21")));
		range = bookmark.get_Range();
		range.put_Text(sNu2);
		bookmark = bookmarks.Item(&_variant_t(_T("FNu22")));
		range = bookmark.get_Range();
		range.put_Text(sNu2);
		bookmark = bookmarks.Item(&_variant_t(_T("FbNu2")));
		range = bookmark.get_Range();
		range.put_Text(sbNu2);
	}
	else if(con1 == 2)
	{
		//绑定参数
		CString snamda,sl01,sl02,sl0cyd,ssp,sAcor,sarfa,sAs,sAsl,sAsso,sCT,sd,sdcor,sfc,sfy,sGd,sGfy,sGST,sH,sN,spas,ss,ssfzyAs,sST;
		sAcor.Format(_T("%.1lf"),Acor/10000);
		sarfa.Format(_T("%.1lf"),info.yszya);
		sAs.Format(_T("%.0lf"),m_As);
		sAsl.Format(_T("%.1lf"),0.25*3.1415926*m_Gd*m_Gd);
		sAsso.Format(_T("%.0lf"),Ass0);
		sCT = m_CT;
		sd.Format(_T("%.0lf"),m_d);
		sdcor.Format(_T("%.0lf"),dcor);
		sfc.Format(_T("%.1lf"),info.cp);
		sfy.Format(_T("%.0lf"),info.ss);
		sGd.Format(_T("%.0lf"),m_Gd);
		sGfy.Format(_T("%.0lf"),info.gjss);
		sGST = m_GST;
		sH.Format(_T("%.1lf"),m_H);
		sN.Format(_T("%.0lf"),m_N);
		spas.Format(_T("%.0lf"),m_pas);
		ss.Format(_T("%.1lf"),m_s);
		ssfzyAs.Format(_T("%.1lf"),0.25*m_As);
		sST = m_ST;
		snamda.Format(_T("%.1lf"),namda);
		sl01.Format(_T("%.1lf"),m_l0/1000);
		sl02.Format(_T("%.0lf"),m_l0);
		sl0cyd.Format(_T("%.1lf"),m_l0/m_d);
		ssp.Format(_T("%.0lf"),m_sp);
		//结束绑定
		temppath.Replace(_T("\\peijin.exe"),_T("\\templet\\lxsjx3.dot"));

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

		bookmark = bookmarks.Item(&_variant_t(_T("Acor1")));
		range = bookmark.get_Range();
		range.put_Text(sAcor);
		bookmark = bookmarks.Item(&_variant_t(_T("Acor2")));
		range = bookmark.get_Range();
		range.put_Text(sAcor);
		bookmark = bookmarks.Item(&_variant_t(_T("arfa1")));
		range = bookmark.get_Range();
		range.put_Text(sarfa);
		bookmark = bookmarks.Item(&_variant_t(_T("arfa2")));
		range = bookmark.get_Range();
		range.put_Text(sarfa);
		bookmark = bookmarks.Item(&_variant_t(_T("As1")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("As2")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("As3")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("Asl1")));
		range = bookmark.get_Range();
		range.put_Text(sAsl);
		bookmark = bookmarks.Item(&_variant_t(_T("Asso1")));
		range = bookmark.get_Range();
		range.put_Text(sAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("Asso2")));
		range = bookmark.get_Range();
		range.put_Text(sAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("Asso3")));
		range = bookmark.get_Range();
		range.put_Text(sAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("CT1")));
		range = bookmark.get_Range();
		range.put_Text(sCT);
		bookmark = bookmarks.Item(&_variant_t(_T("CT2")));
		range = bookmark.get_Range();
		range.put_Text(sCT);
		bookmark = bookmarks.Item(&_variant_t(_T("d1")));
		range = bookmark.get_Range();
		range.put_Text(sd);
		bookmark = bookmarks.Item(&_variant_t(_T("d2")));
		range = bookmark.get_Range();
		range.put_Text(sd);
		bookmark = bookmarks.Item(&_variant_t(_T("d3")));
		range = bookmark.get_Range();
		range.put_Text(sd);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor1")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor2")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("dcor3")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("fc1")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("fc2")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("fy1")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("fy2")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Gd1")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("Gd2")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("Gfy1")));
		range = bookmark.get_Range();
		range.put_Text(sGfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Gfy2")));
		range = bookmark.get_Range();
		range.put_Text(sGfy);
		bookmark = bookmarks.Item(&_variant_t(_T("GST1")));
		range = bookmark.get_Range();
		range.put_Text(sGST);
		bookmark = bookmarks.Item(&_variant_t(_T("H1")));
		range = bookmark.get_Range();
		range.put_Text(sH);
		bookmark = bookmarks.Item(&_variant_t(_T("H2")));
		range = bookmark.get_Range();
		range.put_Text(sH);
		bookmark = bookmarks.Item(&_variant_t(_T("N1")));
		range = bookmark.get_Range();
		range.put_Text(sN);
		bookmark = bookmarks.Item(&_variant_t(_T("N2")));
		range = bookmark.get_Range();
		range.put_Text(sN);
		bookmark = bookmarks.Item(&_variant_t(_T("pas1")));
		range = bookmark.get_Range();
		range.put_Text(spas);
		bookmark = bookmarks.Item(&_variant_t(_T("pas2")));
		range = bookmark.get_Range();
		range.put_Text(spas);
		bookmark = bookmarks.Item(&_variant_t(_T("s1")));
		range = bookmark.get_Range();
		range.put_Text(ss);
		bookmark = bookmarks.Item(&_variant_t(_T("s2")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("sfzyAs")));
		range = bookmark.get_Range();
		range.put_Text(ssfzyAs);
		bookmark = bookmarks.Item(&_variant_t(_T("ST1")));
		range = bookmark.get_Range();
		range.put_Text(sST);
		bookmark = bookmarks.Item(&_variant_t(_T("namda1")));
		range = bookmark.get_Range();
		range.put_Text(snamda);
		bookmark = bookmarks.Item(&_variant_t(_T("l01")));
		range = bookmark.get_Range();
		range.put_Text(sl01);
		bookmark = bookmarks.Item(&_variant_t(_T("l02")));
		range = bookmark.get_Range();
		range.put_Text(sl02);
		bookmark = bookmarks.Item(&_variant_t(_T("l0cyd")));
		range = bookmark.get_Range();
		range.put_Text(sl0cyd);
		bookmark = bookmarks.Item(&_variant_t(_T("FGd1")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("FGd2")));
		range = bookmark.get_Range();
		range.put_Text(sGd);
		bookmark = bookmarks.Item(&_variant_t(_T("Fs1")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("Fs2")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("Fs3")));
		range = bookmark.get_Range();
		range.put_Text(ssp);
		bookmark = bookmarks.Item(&_variant_t(_T("Fdcor1")));
		range = bookmark.get_Range();
		range.put_Text(sdcor);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsl1")));
		range = bookmark.get_Range();
		range.put_Text(sAsl);

		CString sfAsso;
		sfAsso.Format(_T("%.0lf"),m_fAss0);

		bookmark = bookmarks.Item(&_variant_t(_T("FAsso2")));
		range = bookmark.get_Range();
		range.put_Text(sfAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsso1")));
		range = bookmark.get_Range();
		range.put_Text(sfAsso);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffc1")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffc2")));
		range = bookmark.get_Range();
		range.put_Text(sfc);
		bookmark = bookmarks.Item(&_variant_t(_T("FAcor1")));
		range = bookmark.get_Range();
		range.put_Text(sAcor);
		bookmark = bookmarks.Item(&_variant_t(_T("Farfa1")));
		range = bookmark.get_Range();
		range.put_Text(sarfa);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffyv1")));
		range = bookmark.get_Range();
		range.put_Text(sGfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffy1")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffy2")));
		range = bookmark.get_Range();
		range.put_Text(sfy);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsp1")));
		range = bookmark.get_Range();
		range.put_Text(sAs);
		bookmark = bookmarks.Item(&_variant_t(_T("FAsp3")));
		range = bookmark.get_Range();
		range.put_Text(sAs);

		CString sNu1,sNu2,sfi,sA,sbNu2;
		sNu1.Format(_T("%.0lf"),m_Nu1/1000);
		sNu2.Format(_T("%.0lf"),m_Nu2/1000);
		sfi.Format(_T("%.4lf"),info.fi);
		sA.Format(_T("%.1lf"),A/10000);
		sbNu2.Format(_T("%.0lf"),1.5*m_Nu2/1000);

		bookmark = bookmarks.Item(&_variant_t(_T("FNu11")));
		range = bookmark.get_Range();
		range.put_Text(sNu1);
		bookmark = bookmarks.Item(&_variant_t(_T("Ffi1")));
		range = bookmark.get_Range();
		range.put_Text(sfi);
		bookmark = bookmarks.Item(&_variant_t(_T("FA1")));
		range = bookmark.get_Range();
		range.put_Text(sA);
		bookmark = bookmarks.Item(&_variant_t(_T("FNu21")));
		range = bookmark.get_Range();
		range.put_Text(sNu2);
		bookmark = bookmarks.Item(&_variant_t(_T("FNu22")));
		range = bookmark.get_Range();
		range.put_Text(sNu2);
		bookmark = bookmarks.Item(&_variant_t(_T("FbNu2")));
		range = bookmark.get_Range();
		range.put_Text(sbNu2);
	}
	

	range.ReleaseDispatch();
	bookmark.ReleaseDispatch();
	bookmarks.ReleaseDispatch();
	docx.ReleaseDispatch();
	docs.ReleaseDispatch();
	wordApp.ReleaseDispatch();
}
