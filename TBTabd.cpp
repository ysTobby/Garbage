// TBTabd.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "TBTabd.h"
#include "afxdialogex.h"


// CTBTabd 对话框

IMPLEMENT_DYNAMIC(CTBTabd, CDialogEx)

	CTBTabd::CTBTabd(CWnd* pParent /*=NULL*/)
	: CDialogEx(CTBTabd::IDD, pParent)
{

	m_as = 40;
	m_b = 200;
	m_bf = 400;
	m_CS = _T("C30");
	m_h = 500;
	m_hf = 100;
	m_M = 0.0;
	m_SAs = 0.0;
	m_SAsmin = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
}

CTBTabd::~CTBTabd()
{
}

void CTBTabd::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_TD_as, m_as);
	DDX_Text(pDX, IDC_TD_b, m_b);
	DDX_Text(pDX, IDC_TD_bf, m_bf);
	DDX_CBString(pDX, IDC_TD_CS, m_CS);
	DDX_Text(pDX, IDC_TD_h, m_h);
	DDX_Text(pDX, IDC_TD_hf, m_hf);
	DDX_Text(pDX, IDC_TD_M, m_M);
	DDX_Text(pDX, IDC_TD_SAs, m_SAs);
	DDX_Text(pDX, IDC_TD_SAsmin, m_SAsmin);
	DDX_CBString(pDX, IDC_TD_ST, m_ST);
	DDX_Text(pDX, IDC_TD_x, m_x);
	DDX_Text(pDX, IDC_TD_xb, m_xb);
}


BEGIN_MESSAGE_MAP(CTBTabd, CDialogEx)
	ON_BN_CLICKED(IDC_DD2_Reset, &CTBTabd::OnBnClickedDd2Reset)
	ON_BN_CLICKED(IDC_DD2_Calcula, &CTBTabd::OnBnClickedDd2Calcula)
	ON_BN_CLICKED(IDC_Out, &CTBTabd::OnBnClickedOut)
END_MESSAGE_MAP()


// CTBTabd 消息处理程序


void CTBTabd::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	OnBnClickedDd2Calcula();
	//CDialogEx::OnOK();
}


void CTBTabd::OnBnClickedDd2Reset()
{
	m_as = 40;
	m_b = 200;
	m_bf = 400;
	m_CS = _T("C30");
	m_h = 500;
	m_hf = 100;
	m_M = 0.0;
	m_SAs = 0.0;
	m_SAsmin = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	CWnd::UpdateData(false);
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
}


void CTBTabd::OnBnClickedDd2Calcula()
{
	CWnd::UpdateData(true);
	if(m_b==0.0||m_h==0.0||m_M==0.0||m_as==0.0||m_CS==""||m_ST==""||m_hf==0.0||m_bf==0.0)
	{
		MessageBox(_T("请输入必要参数"));
		return ;
	}
	m_M *= 1000000;
	h0 = m_h - m_as;
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
	info.init(m_CS,m_ST);
	m_xb = info.kb*h0;
	m_Mp = info.a1*info.cp*m_bf*m_hf*(h0-m_hf/2);
	if(m_M <= m_Mp)
	{
		con = 1;
		arfas = m_M / (info.a1*info.cp*m_bf*h0*h0);
		kesai = 1 - sqrt(1-2*arfas);
		m_x = kesai * h0;
		m_SAs = info.a1*info.cp*m_bf*m_x/info.ss;
		m_rAs = m_SAs;
		pmin = 0.002;
		if(pmin < 0.45*info.ct/info.ss)
			pmin = 0.45*info.ct/info.ss;
		m_SAsmin = pmin*m_h*m_b;
		p = m_SAs/(m_b*m_h);
		if(p < pmin*m_h/h0)
		{
			con = 2;
			MessageBox(_T("最小配筋率不满足,采用构造配筋"));
			m_SAs = m_SAsmin;
		}
	}
	else
	{
		con = 3;
		m_Mu1 = info.a1*info.cp*(m_bf-m_b)*m_hf*(h0-m_hf/2);
		m_Mu2 = m_M - m_Mu1;
		arfas = m_Mu2 / (info.a1*info.cp*m_b*h0*h0);
		if((1-2*arfas)<0)
		{
			m_M /= 1000000;
			m_Mp /= 1000000;
			con = 5;
			MessageBox(_T("计算as为负值，请重新设计截面尺寸"));
			return;
		}
		kesai = 1 - sqrt(1-2*arfas);
		m_x = kesai * h0;
		if(m_x>m_xb)
		{
			m_M /= 1000000;
			m_Mp /= 1000000;
			con = 4;
			MessageBox(_T("截面超筋，请加大截面尺寸"));
			return;
		}
		m_SAs = (info.a1*info.cp*m_b*m_x+info.a1*info.cp*(m_bf-m_b)*m_hf)/info.ss;
		m_rAs = m_SAs;
	}
	m_M /= 1000000;
	m_Mp /= 1000000;
	CWnd::UpdateData(false);
}
void CTBTabd::xb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Subscript(true),sel.TypeText(x),font.put_Subscript(false);
}
void CTBTabd::sb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Superscript(true),sel.TypeText(x),font.put_Superscript(false);
}

void CTBTabd::sxb(CString center,CString s,CString x)
{
	xb(center,x);
	sb(_T(""),s);
}

void CTBTabd::OnBnClickedOut()
{
	int i = 0;
	CString s_Mu1,s_Mu2,s_fy,s_kesai,
		s_xb,s_ft,s_fc,s_a1,s_kb,s_arfas,s_x,s_b,s_M,s_Mp,
		s_h,s_as,s_h0,s_SAs,content,szBuf,s_bf,s_hf,s_SAsmin,s_p,s_pmin;

	//参数绑定
	s_arfas.Format(_T("%.3lf"),arfas);
	s_h0.Format(_T("%.0lf"),h0);
	s_fc.Format(_T("%.1lf"),info.cp);
	s_a1.Format(_T("%.1lf"),info.a1);
	s_kb.Format(_T("%.3lf"),info.kb);
	s_fy.Format(_T("%.0lf"),info.ss);
	s_ft.Format(_T("%.2lf"),info.ct); 
	s_kesai.Format(_T("%.3lf"),kesai);
	s_as.Format(_T("%.0lf"),m_as);
	s_b.Format(_T("%.0lf"),m_b);
	s_bf.Format(_T("%.0lf"),m_bf);
	s_h.Format(_T("%.0lf"),m_h);
	s_hf.Format(_T("%.0lf"),m_hf);
	s_M.Format(_T("%.1lf"),m_M);
	s_Mu1.Format(_T("%.1lf"),m_Mu1/1000000);
	s_Mu2.Format(_T("%.1lf"),m_Mu2/1000000);
	s_SAs.Format(_T("%.0lf"),m_rAs);
	s_SAsmin.Format(_T("%.0lf"),m_SAsmin);
	s_x.Format(_T("%.0lf"),m_x);
	s_xb.Format(_T("%.0lf"),(info.kb * h0)); 
	s_p.Format(_T("%.4lf"),p); 
	s_pmin.Format(_T("%.4lf"),pmin);
	s_Mp.Format(_T("%.0lf"),m_Mp);
	//绑定结束


	CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应
	if(!app.CreateDispatch(_T("word.application"))) //启动WORD
	{
		AfxMessageBox(_T("居然你连OFFICE都没有安装吗?"));
		return;
	}
	//创建word进程默认是不可见的
	app.put_Visible(TRUE); //设置WORD可见。
	docs = app.get_Documents();
	docs.Add(new CComVariant(_T("")),new CComVariant(FALSE),new CComVariant(0),new CComVariant());//创建新文档
	doc = app.get_ActiveDocument();
	COleVariant vTrue((short)TRUE),vFalse((short)FALSE),vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	COleVariant formula;
	sel=app.get_Selection();
	font =sel.get_Font();
	font.put_Name(_T("Times New Roman"));
	font.put_Size(18);
	font.put_Color(WdColor::wdColorBlack);
	font.put_Bold(0);

	//1
	sel.TypeText(_T("T形截面梁受弯设计"));
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(1);
	paragraph.put_Alignment(1);
	sel.TypeParagraph();

	font =sel.get_Font();
	font.put_Size(10);
	//2
	content = _T("    已知条件：弯矩设计值:M=") + s_M + _T("kN∙m,截面宽度:b=") 
		+ s_b + _T("mm,截面高度:h=") + s_h
		+ _T("mm,");
	sel.TypeText(content);
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(2);
	paragraph.put_Alignment(0);

	sel.TypeText(_T("翼缘宽度:"));

	sxb(_T("b"),_T("'"),_T("f"));
	sel.TypeText(_T("=")+s_bf + _T("mm,"));

	sel.TypeText(_T("翼缘厚度:"));

	sxb(_T("h"),_T("'"),_T("f"));
	sel.TypeText(_T("=")+s_hf + _T("mm,"));

	sel.TypeText(_T("混凝土强度等级:")+m_CS+_T("."));

	xb(_T("f"),_T("c"));
	sel.TypeText(_T("=")+s_fc + _T("N/mm²,"));

	xb(_T("f"),_T("t"));
	sel.TypeText(_T("=")+s_ft + _T("N/mm²."));

	content = _T("受拉钢筋型号:") + m_ST + _T(",");
	sel.TypeText(content);

	xb(_T("f"),_T("y"));
	sel.TypeText(_T("=")+s_fy + _T("N/mm²."));

	xb(_T("α"),_T("1"));
	sel.TypeText(_T("=")+s_a1 + _T(","));

	xb(_T("ξ"),_T("b"));
	sel.TypeText(_T("=")+s_kb + _T("."));

	content = _T(",受拉钢筋合力点到受拉边缘距离:");
	sel.TypeText(content);
	xb(_T("a"),_T("s"));
	sel.TypeText(_T("=")+s_as + _T("mm."));
	sel.TypeParagraph();

	fields.ToggleShowCodes();
	fields = sel.get_Fields();
	range = sel.get_Range();
	formula = _T("");
	fields.Add(range,vOpt,&formula,vFalse);
	sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdExtend));
	CString a = sel.get_Text();
	if(a == _T("."))
	{
		i = 1;
		sel.MoveRight(COleVariant((short)3),COleVariant((short)3),COleVariant((short)wdExtend));
		fields.ToggleShowCodes();
		paragraphs = doc.get_Paragraphs();
		paragraph = paragraphs.Item(3);
		range = paragraph.get_Range();
		sel.Cut();
		sel.TypeParagraph();
	}
	else
	{
		sel.MoveLeft(COleVariant((short)1),COleVariant((short)3),COleVariant((short)wdMove));
		sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdExtend));
		sel.Cut();
		sel.TypeParagraph();
	}

	//3
	sel.TypeText(_T("    解:"));
	sel.TypeParagraph();
	//4
	xb(_T("h"),_T("0"));
	xb(_T(" = h - a"),_T("s"));
	sel.TypeText(_T(" = ")+s_h+_T(" - ")+s_as+_T(" = ")+s_h0+_T(" (mm)"));
	sel.TypeParagraph();
	//5
	xb(_T("x"),_T("b"));
	xb(_T(" = ξ"),_T("b"));
	xb(_T("h"),_T("0"));
	sel.TypeText(_T(" = ")+s_kb+_T("×")+s_h0+_T(" = ")+s_xb+_T(" (mm)"));
	sel.TypeParagraph();
	if(con == 1||con == 2)
	{
		//6
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ M = "),_T(""));
			sel.TypeText(s_M+ _T(" (kN∙m)"));
			xb(_T("  ≤ α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("b"),_T("'"),_T("f"));
			sxb(_T("h"),_T("'"),_T("f"));
			xb(_T("(h"),_T("0"));
			sxb(_T(" - \\f(h"),_T("'"),_T("f"));
			xb(_T(",2)) = "),_T(""));
			sel.TypeText(s_a1+_T("×")+s_fc+_T("×")+s_bf+_T("×")+s_hf+_T("×(")+s_h0+_T(" - ")+COperate::fraction(s_hf,_T("2"))+_T(") = ")+
				s_Mp+_T(" (kN∙m)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//7
		sel.TypeText(_T("故属于第一种类型T型截面梁"));
		sel.TypeParagraph();
		//8
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ α"),_T("s"));
			xb(_T(" = \\f(M,α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("b"),_T("'"),_T("f"));
			sxb(_T("h"),_T("2"),_T("0"));
			xb(_T(") = "),_T(""));
			sel.TypeText(COperate::fraction(s_M+_T("×10⁶"),s_a1+_T("×")+s_fc+_T("×")+s_bf+_T("×")+s_h0+_T("²"))+
				_T(" = ")+s_arfas);
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//9
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ ξ = 1 - \\r(1 - 2α"),_T("s"));
			xb(_T(") = 1 - "),_T(""));
			sel.TypeText(COperate::radical(_T("1 - 2×")+s_arfas)+_T(" = ")+s_kesai);
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//10
		xb(_T("x = ξh"),_T("0"));
		xb(_T(" = "),_T(""));
		sel.TypeText(s_kesai+_T("×")+s_h0+_T(" = ")+s_x+_T(" (mm)"));
		sel.TypeParagraph();
		//11
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ A"),_T("s"));
			xb(_T(" = \\f(α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("b"),_T("'"),_T("f"));
			xb(_T("x,f"),_T("y"));
			xb(_T(") = "),_T(""));
			sel.TypeText(COperate::fraction(s_a1+_T("×")+s_fc+_T("×")+s_bf+_T("×")+s_x,s_fy)
				+_T(" = ")+s_SAs+_T(" (mm²)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//12
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ ρ"),_T("min"));
			xb(_T(" = max{0.2%,0.45\\f(f"),_T("t"));
			xb(_T(", f"),_T("y"));
			xb(_T(")} = "),_T(""));
			sel.TypeText(s_pmin);
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//13
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ ρ = \\f(A"),_T("s"));
			xb(_T(",bh) = "),_T(""));
			sel.TypeText(COperate::fraction(s_SAs,s_b+_T("×")+s_h)
				+_T(" = ")+s_p);
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(con == 1)
		{
			//14
			xb(_T("ρ ≥ ρ"),_T("min"));
			sel.TypeText(_T("最小配筋率满足要求."));
		}
		else
		{
			//14
			xb(_T("ρ < ρ"),_T("min"));
			sel.TypeText(_T("最小配筋率不满足要求，按构造要求配筋:"));
			sel.TypeParagraph();
			//15
			xb(_T("A"),_T("s"));
			xb(_T(" = A"),_T("smin"));
			xb(_T(" = ρ"),_T("min"));
			xb(_T("hb = "),_T(""));
			sel.TypeText(s_pmin+_T("×")+s_h+_T("×")+s_b
				+_T(" = ")+s_SAsmin+_T(" (mm²)"));
		}
	}
	else
	{
		//6
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ M = "),_T(""));
			sel.TypeText(s_M+ _T(" (kN∙m)"));
			xb(_T("  > α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("b"),_T("'"),_T("f"));
			sxb(_T("h"),_T("'"),_T("f"));
			xb(_T("(h"),_T("0"));
			sxb(_T(" - \\f(h"),_T("'"),_T("f"));
			xb(_T(",2)) = "),_T(""));
			sel.TypeText(s_a1+_T("×")+s_fc+_T("×")+s_bf+_T("×")+s_hf+_T("×(")+s_h0+_T(" - ")+COperate::fraction(s_hf,_T("2"))+_T(") = ")+
				s_Mp+_T(" (kN∙m)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//7
		sel.TypeText(_T("故属于第二种类型T型截面梁"));
		sel.TypeParagraph();
		//8
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ M"),_T("u1"));
			xb(_T(" = α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("(b"),_T("'"),_T("f"));
			sxb(_T(" - b)h"),_T("'"),_T("f"));
			xb(_T("(h"),_T("0"));
			sxb(_T(" - \\f(h"),_T(";"),_T("f"));
			xb(_T(",2)) = "),_T(""));
			sel.TypeText(s_a1+_T("×")+s_fc+_T("×")+_T("(")+s_bf+_T(" - ")+s_b+_T(")")+_T("×")+s_hf+_T("×")+_T("(")+s_h0+_T(" - ")+COperate::fraction(s_hf,_T("2"))+_T(")")+_T(" = ")+s_Mu1+_T(" (kN∙m)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//9
		xb(_T("M"),_T("u2"));
		xb(_T(" = M - M"),_T("u1"));
		xb(_T(" = "),_T(""));
		sel.TypeText(s_M+_T(" - ")+s_Mu1+_T(" = ")+s_Mu2+_T(" (kN∙m)"));
		sel.TypeParagraph();
		//10 α
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ α"),_T("s"));
			xb(_T(" = \\f(M"),_T("u2"));
			xb(_T(",α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("bh"),_T("2"),_T("0"));
			xb(_T(") = "),_T(""));
			sel.TypeText(COperate::fraction(s_Mu2+_T("×10⁶"),s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T("²"))+_T(" = ")+s_arfas);
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(con == 5)
		{
			//1-2as<0 11
			xb(_T("1 - 2α"),_T("s"));
			xb(_T(" < 0,"),_T(""));
			sel.TypeText(_T("请增大截面面积。"));

			font.ReleaseDispatch();
			paragraph.ReleaseDispatch();
			paragraphs.ReleaseDispatch();
			range.ReleaseDispatch();
			fields.ReleaseDispatch();
			doc.ReleaseDispatch();
			sel.ReleaseDispatch();
			docs.ReleaseDispatch(); 
			app.ReleaseDispatch();  
			CoUninitialize();
			return;
		}
		//11 
		fields.ToggleShowCodes();
		fields = sel.get_Fields();
		range = sel.get_Range();
		formula = _T("");
		fields.Add(range,vOpt,&formula,vFalse);

		if(i == 1)
		{
			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
		}

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			xb(_T("EQ ξ = 1 - \\r(1 - 2α"),_T("s"));
			xb(_T(") = 1 - "),_T(""));
			sel.TypeText(COperate::radical(_T("1 - 2×")+s_arfas)+_T(" = ")+s_kesai);
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//12
		xb(_T("x = ξh"),_T("0"));
		sel.TypeText(_T(" = ")+s_kesai+_T("×")+s_h0+_T(" = ")+s_x+_T(" (mm)"));
		sel.TypeParagraph();
		if(con == 3)
		{
			//13
			xb(_T("x ≤ x"),_T("b"));
			sel.TypeText(_T(",满足受压区高度限制要求"));
			sel.TypeParagraph();
			//14
			fields.ToggleShowCodes();
			fields = sel.get_Fields();
			range = sel.get_Range();
			formula = _T("");
			fields.Add(range,vOpt,&formula,vFalse);

			if(i == 1)
			{
				sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
				sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
				fields.ToggleShowCodes();
				sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
			}

			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			{
				xb(_T("EQ A"),_T("s"));
				xb(_T(" = \\f(α"),_T("1"));
				xb(_T("f"),_T("c"));
				xb(_T("bx + α"),_T("1"));
				xb(_T("f"),_T("c"));
				sxb(_T("(b"),_T("'"),_T("f"));
				sxb(_T(" - b)h"),_T("'"),_T("f"));
				xb(_T(",f"),_T("y"));
				xb(_T(")"),_T(""));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
			//15
			sel.TypeText(_T("  "));
			fields.ToggleShowCodes();
			fields = sel.get_Fields();
			range = sel.get_Range();
			formula = _T("");
			fields.Add(range,vOpt,&formula,vFalse);

			if(i == 1)
			{
				sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
				sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
				fields.ToggleShowCodes();
				sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdMove));
			}

			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			{
				sel.TypeText(_T("EQ = ")+COperate::fraction(s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_x+_T(" + ")+s_a1
					+_T("×")+s_fc+_T("×")+_T("(")+s_bf+_T(" - ")+s_b+_T(")")+_T("×")+s_hf,s_fy)+_T(" = ")+s_SAs+_T(" (mm²)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		}
		else
		{
			//13
			xb(_T("x > x"),_T("b"));
			sel.TypeText(_T("截面超筋，请重新设计截面尺寸"));
		}
	}




	font.ReleaseDispatch();
	paragraph.ReleaseDispatch();
	paragraphs.ReleaseDispatch();
	range.ReleaseDispatch();
	fields.ReleaseDispatch();
	doc.ReleaseDispatch();
	sel.ReleaseDispatch();
	docs.ReleaseDispatch(); 
	app.ReleaseDispatch();  
	CoUninitialize();
}


void CTBTabd::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
