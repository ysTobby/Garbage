// TBTabc.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "TBTabc.h"
#include "afxdialogex.h"


// CTBTabc 对话框

IMPLEMENT_DYNAMIC(CTBTabc, CDialogEx)

	CTBTabc::CTBTabc(CWnd* pParent /*=NULL*/)
	: CDialogEx(CTBTabc::IDD, pParent)
{

	m_as = 40;
	m_b = 200;
	m_bf = 600;
	m_CS = _T("C30");
	m_h = 500;
	m_hf = 100;
	m_M = 0.0;
	m_SAs = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	m_SAsmin = 0.0;
}

CTBTabc::~CTBTabc()
{
}

void CTBTabc::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_TC_as, m_as);
	DDX_Text(pDX, IDC_TC_b, m_b);
	DDX_Text(pDX, IDC_TC_bf, m_bf);
	DDX_CBString(pDX, IDC_TC_CS, m_CS);
	DDX_Text(pDX, IDC_TC_h, m_h);
	DDX_Text(pDX, IDC_TC_hf, m_hf);
	DDX_Text(pDX, IDC_TC_M, m_M);
	DDX_Text(pDX, IDC_TC_SAs, m_SAs);
	DDX_CBString(pDX, IDC_TC_ST, m_ST);
	DDX_Text(pDX, IDC_TC_x, m_x);
	DDX_Text(pDX, IDC_TC_xb, m_xb);
	DDX_Text(pDX, IDC_TC_SAsmin, m_SAsmin);
}


BEGIN_MESSAGE_MAP(CTBTabc, CDialogEx)
	ON_BN_CLICKED(IDC_DD2_Reset, &CTBTabc::OnBnClickedDd2Reset)
	ON_BN_CLICKED(IDC_DD2_Calcula, &CTBTabc::OnBnClickedDd2Calcula)
	ON_BN_CLICKED(IDC_Out, &CTBTabc::OnBnClickedOut)
END_MESSAGE_MAP()


// CTBTabc 消息处理程序


void CTBTabc::OnBnClickedDd2Reset()
{
	m_as = 40;
	m_b = 200;
	m_bf = 600;
	m_CS = _T("C30");
	m_h = 500;
	m_hf = 100;
	m_M = 0.0;
	m_SAs = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	m_SAsmin = 0.0;
	CWnd::UpdateData(false);
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
}


void CTBTabc::OnBnClickedDd2Calcula()
{
	CWnd::UpdateData(true);

	if(m_b==0.0||m_h==0.0||m_SAs==0.0||m_as==0.0||m_CS==""||m_ST==""||m_hf==0.0||m_bf==0.0)
	{
		MessageBox(_T("请输入必要参数"));
		return ;
	}
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);

	h0 = m_h - m_as;
	info.init(m_CS,m_ST);
	m_xb = info.kb*h0;
	p = m_SAs/(m_b*m_h);
	pmin = 0.002;
	if(pmin < 0.45*info.ct/info.ss)
		pmin = 0.45*info.ct/info.ss;
	m_SAsmin = pmin*m_h*m_b;
	if(info.ss*m_SAs <= info.a1*info.cp*m_bf*m_hf)
	{
		con = 1;
		if(p < pmin)
		{
			con = 2;
			m_M /= 1000000;
			MessageBox(_T("输入钢筋面积不满足最小配筋面积要求，最小配筋面积见结果!"));
			CWnd::UpdateData(false);
			return ;
		}
		m_x = (info.ss*m_SAs)/(info.a1*info.cp*m_bf);
		m_M = info.ss*m_SAs*(h0 - m_x/2);
	}
	else
	{
		con = 3;
		m_SAs1 = info.a1*info.cp*(m_bf - m_b)*m_hf/info.ss;
		m_SAs2 = m_SAs - m_SAs1;
		m_x = (info.ss*m_SAs2)/(info.a1*info.cp*m_b);
		if(m_x>m_xb)
		{
			con = 4;
			MessageBox(_T("钢筋配置过多,取x=xb计算极限弯矩!"));
			m_Mu1 = info.ss*m_SAs1*(h0 - m_hf/2);
			m_Mu2 = info.a1*info.cp*m_b*m_xb*(h0 - m_xb/2);
			m_M = m_Mu1 + m_Mu2;
		}
		else
		{
			m_Mu1 = info.ss*m_SAs1*(h0 - m_hf/2);
			m_Mu2 = info.a1*info.cp*m_b*m_x*(h0 - m_x/2);
			m_M = m_Mu1 + m_Mu2;
		}
		m_Mu1 /= 1000000;
		m_Mu2 /= 1000000;
	}
	m_M /= 1000000;
	CWnd::UpdateData(false);
}

void CTBTabc::xb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Subscript(true),sel.TypeText(x),font.put_Subscript(false);
}
void CTBTabc::sb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Superscript(true),sel.TypeText(x),font.put_Superscript(false);
}

void CTBTabc::sxb(CString center,CString s,CString x)
{
	xb(center,x);
	sb(_T(""),s);
}

void CTBTabc::OnBnClickedOut()
{
	int i = 0;
	CString s_Mu1,s_Mu2,s_fy,s_SAs1,s_SAs2,
		s_xb,s_ft,s_fc,s_a1,s_kb,s_x,s_b,s_M,s_Mp,s_SF,s_CF,
		s_h,s_as,s_h0,s_SAs,content,szBuf,s_bf,s_hf,s_SAsmin,s_pmin;

	//参数绑定
	s_h0.Format(_T("%.0lf"),h0);
	s_fc.Format(_T("%.1lf"),info.cp);
	s_a1.Format(_T("%.1lf"),info.a1);
	s_kb.Format(_T("%.3lf"),info.kb);
	s_fy.Format(_T("%.0lf"),info.ss);
	s_ft.Format(_T("%.2lf"),info.ct); 
	s_as.Format(_T("%.0lf"),m_as);
	s_b.Format(_T("%.0lf"),m_b);
	s_bf.Format(_T("%.0lf"),m_bf);
	s_h.Format(_T("%.0lf"),m_h);
	s_hf.Format(_T("%.0lf"),m_hf);
	s_M.Format(_T("%.0lf"),m_M);
	s_Mu1.Format(_T("%.0lf"),m_Mu1);
	s_Mu2.Format(_T("%.0lf"),m_Mu2);
	s_SAs.Format(_T("%.0lf"),m_SAs);
	s_SAs1.Format(_T("%.0lf"),m_SAs1);
	s_SAs2.Format(_T("%.0lf"),m_SAs2);
	s_SAsmin.Format(_T("%.0lf"),m_SAsmin);
	s_x.Format(_T("%.0lf"),m_x);
	s_xb.Format(_T("%.0lf"),(info.kb * h0)); 
	s_pmin.Format(_T("%.4lf"),pmin);
	s_SF.Format(_T("%.1lf"),info.ss*m_SAs/1000);
	s_CF.Format(_T("%.1lf"),info.a1*info.cp*m_bf*m_hf/1000);
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
	sel.TypeText(_T("T形截面梁受弯复核"));
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(1);
	paragraph.put_Alignment(1);
	sel.TypeParagraph();

	font =sel.get_Font();
	font.put_Size(10);
	//2
	content = _T("    已知条件:截面宽度:b=") 
		+ s_b + _T("mm,截面高度:h=") + s_h
		+ _T("mm,");
	sel.TypeText(content);
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(2);
	paragraph.put_Alignment(0);

	sel.TypeText(_T("翼缘宽度:"));

	sxb(_T("b"),_T("'"),_T("f"));
	xb(_T("="),_T(""));
	sel.TypeText(s_bf + _T("mm,"));

	sel.TypeText(_T("翼缘厚度:"));

	sxb(_T("h"),_T("'"),_T("f"));
	xb(_T("="),_T(""));
	sel.TypeText(s_hf + _T("mm,"));

	sel.TypeText(_T("混凝土强度等级:")+m_CS+_T("."));

	xb(_T("f"),_T("c"));
	xb(_T("="),_T(""));
	sel.TypeText(s_fc + _T("N/mm²,"));

	xb(_T("f"),_T("t"));
	xb(_T("="),_T(""));
	sel.TypeText(s_ft + _T("N/mm²."));

	content = _T("受拉钢筋型号:") + m_ST + _T(",");
	sel.TypeText(content);

	xb(_T("f"),_T("y"));
	xb(_T("="),_T(""));
	sel.TypeText(s_fy + _T("N/mm²."));

	xb(_T("α"),_T("1"));
	xb(_T("="),_T(""));
	sel.TypeText(s_a1 + _T(","));

	xb(_T("ξ"),_T("b"));
	xb(_T("="),_T(""));
	sel.TypeText(s_kb + _T("."));

	content = _T("受拉钢筋合力点到受拉边缘距离:");
	sel.TypeText(content);

	xb(_T("a"),_T("s"));
	xb(_T("="),_T(""));
	sel.TypeText(s_as + _T("mm,"));

	content = _T("受拉钢筋面积:");
	sel.TypeText(content);

	xb(_T("A"),_T("s"));
	xb(_T("="),_T(""));
	sel.TypeText(s_SAs + _T("mm²."));
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
	xb(_T(" = "),_T(""));
	sel.TypeText(s_h+_T(" - ")+s_as+_T(" = ")+s_h0+_T(" (mm)"));
	sel.TypeParagraph();
	//5
	xb(_T("f"),_T("y"));
	xb(_T("A"),_T("s"));
	xb(_T(" = "),_T(""));
	sel.TypeText(s_fy+_T("×")+s_SAs+_T(" = ") + s_SF + _T(" (kN)"));
	sel.TypeParagraph();
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
		xb(_T("EQ α"),_T("1"));
		xb(_T("f"),_T("c"));
		sxb(_T("b"),_T("'"),_T("f"));
		sxb(_T("h"),_T("'"),_T("f"));
		xb(_T(" = "),_T(""));
		sel.TypeText(s_a1+_T("×")+s_fc+_T("×")+s_bf+_T("×")+s_hf+_T(" = ") + s_CF + _T(" (kN)"));
	}
	fields.ToggleShowCodes();
	sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.TypeParagraph();
	if(con == 1||con ==2)
	{
		//7
		xb(_T("f"),_T("y"));
		xb(_T("A"),_T("s"));
		xb(_T(" ≤ α"),_T("1"));
		xb(_T("f"),_T("c"));
		sxb(_T("b"),_T("'"),_T("f"));
		sxb(_T("h"),_T("'"),_T("f"));
		sel.TypeText(_T("属于第一种类型"));
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
			xb(_T("EQ A"),_T("smin"));
			xb(_T(" = max{0.2%,0.45\\f(f"),_T("t"));
			xb(_T(", f"),_T("y"));
			xb(_T(")}hb = "),_T(""));
			sel.TypeText(s_pmin+_T("×")+s_h+_T("×")+s_b+_T(" = ")+s_SAsmin+_T(" (mm²)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(con == 2)
		{
			//9
			xb(_T("A"),_T("s"));
			xb(_T(" < A"),_T("smin"));
			sel.TypeText(_T("钢筋面积未满足最小配筋要求，建议增大受拉钢筋面积."));
			return;
		}
		//9
		xb(_T("A"),_T("s"));
		xb(_T(" > A"),_T("smin"));
		sel.TypeText(_T("满足最小配筋要求"));
		sel.TypeParagraph();
		//10
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
			xb(_T("EQ x = \\f(f"),_T("y"));
			xb(_T("A"),_T("s"));
			xb(_T(",α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("b"),_T("'"),_T("f"));
			xb(_T(") = "),_T(""));
			sel.TypeText(COperate::fraction(s_fy+_T("×")+s_SAs,s_a1+_T("×")+s_fc+_T("×")+s_bf)+_T(" = ")+s_x+_T(" (mm)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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
			xb(_T("EQ M"),_T("u"));
			xb(_T(" = f"),_T("y"));
			xb(_T("A"),_T("s"));
			xb(_T("(h"),_T("0"));
			xb(_T(" - \\f(x,2)) = "),_T(""));
			sel.TypeText(s_fy+_T("×")+s_SAs+_T("×")+_T("(")+s_h0+_T(" - ")+COperate::fraction(s_x,_T("2"))+_T(")")+_T(" = ")+s_M+_T(" (kN∙m)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
	}
	else
	{
		//7
		xb(_T("f"),_T("y"));
		xb(_T("A"),_T("s"));
		xb(_T(" > α"),_T("1"));
		xb(_T("f"),_T("c"));
		sxb(_T("b"),_T("'"),_T("f"));
		sxb(_T("h"),_T("'"),_T("f"));
		sel.TypeText(_T("属于第二种类型"));
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
			xb(_T("EQ A"),_T("s1"));
			xb(_T(" = \\f(α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("h"),_T("'"),_T("f"));
			sxb(_T("(b"),_T("'"),_T("f"));
			xb(_T(" - b),f"),_T("y"));
			xb(_T(") = "),_T(""));
			sel.TypeText(COperate::fraction(s_a1+_T("×")+s_fc+_T("×")+s_hf+_T("×")+_T("(")+s_bf+_T(" - ")+s_b+_T(")"),s_fy)+_T(" = ")+s_SAs1+_T(" (mm²)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//9
		xb(_T("A"),_T("s2"));
		xb(_T(" = A"),_T("s"));
		xb(_T(" - A"),_T("s1"));
		xb(_T(" = "),_T(""));
		sel.TypeText(s_SAs+_T(" - ")+s_SAs1+_T(" = ")+s_SAs2+_T(" (mm²)"));
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
			xb(_T("EQ x = \\f(f"),_T("y"));
			xb(_T("A"),_T("s2"));
			xb(_T(",α"),_T("1"));
			xb(_T("f"),_T("c"));
			xb(_T("b) = "),_T(""));
			sel.TypeText(COperate::fraction(s_fy+_T("×")+s_SAs2,s_a1+_T("×")+s_fc+_T("×")+s_b)+_T(" = ")+s_x+_T(" (mm)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//10
		xb(_T("x"),_T("b"));
		xb(_T(" = ξ"),_T("b"));
		xb(_T("h"),_T("0"));
		xb(_T(" = "),_T(""));
		sel.TypeText(s_kb+_T("×")+s_h0+_T(" = ")+s_xb+_T(" (mm)"));
		sel.TypeParagraph();
		if(con ==3)
		{
			//11
			xb(_T("x ≤ x"),_T("b"));
			sel.TypeText(_T("计算受压区高度未超过界限受压区高度."));
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
				xb(_T("EQ M"),_T("u1"));
				xb(_T(" = f"),_T("y"));
				xb(_T("A"),_T("s1"));
				xb(_T("(h"),_T("0"));
				sxb(_T(" - \\f(h"),_T("'"),_T("f"));
				xb(_T(",2)) = "),_T(""));
				sel.TypeText(s_fy+_T("×")+s_SAs1+_T("×")+_T("(")+s_h0+_T(" - ")+COperate::fraction(s_hf,_T("2"))+_T(")")+_T(" = ")+s_Mu1+_T(" (kN∙m)"));
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
				xb(_T("EQ M"),_T("u2"));
				xb(_T(" = α"),_T("1"));
				xb(_T("f"),_T("c"));
				xb(_T("bx(h"),_T("0"));
				xb(_T(" - \\f(x,2)) = "),_T(""));
				sel.TypeText(s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_x+_T("×")+_T("(")+s_h0+_T(" - ")+COperate::fraction(s_x,_T("2"))+_T(")")+_T(" = ")+s_Mu2+_T(" (kN∙m)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
			//14
			xb(_T("M"),_T("u"));
			xb(_T(" = M"),_T("u1"));
			xb(_T(" + M"),_T("u2"));
			xb(_T(" = "),_T(""));
			sel.TypeText(s_Mu1+_T(" + ")+s_Mu2+_T(" = ")+s_M+_T(" (kN∙m)"));
		}
		else
		{
			//11
			xb(_T("x > x"),_T("b"));
			sel.TypeText(_T("计算受压区高度超过界限受压区高度,取界限受压区高度计算弯矩极限值."));
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
				xb(_T("EQ M"),_T("u1"));
				xb(_T(" = f"),_T("y"));
				xb(_T("A"),_T("s1"));
				xb(_T("(h"),_T("0"));
				sxb(_T(" - \\f(h"),_T("'"),_T("f"));
				xb(_T(",2)) = "),_T(""));
				sel.TypeText(s_fy+_T("×")+s_SAs1+_T("×")+_T("(")+s_h0+_T(" - ")+COperate::fraction(s_hf,_T("2"))+_T(")")+_T(" = ")+s_Mu1+_T(" (kN∙m)"));
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
				xb(_T("EQ M"),_T("u2"));
				xb(_T(" = α"),_T("1"));
				xb(_T("f"),_T("c"));
				xb(_T("bx"),_T("b"));
				xb(_T("(h"),_T("0"));
				xb(_T(" - \\f(x"),_T("b"));
				xb(_T(",2)) = "),_T(""));
				sel.TypeText(s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_xb+_T("×")+_T("(")+s_h0+_T(" - ")+COperate::fraction(s_xb,_T("2"))+_T(")")+_T(" = ")+s_Mu2+_T(" (kN∙m)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
			//14
			xb(_T("M"),_T("u"));
			xb(_T(" = M"),_T("u1"));
			xb(_T(" + M"),_T("u2"));
			xb(_T(" = "),_T(""));
			sel.TypeText(s_Mu1+_T(" + ")+s_Mu2+_T(" = ")+s_M+_T(" (kN∙m)"));
			sel.TypeParagraph();
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


void CTBTabc::OnOK()
{
	OnBnClickedDd2Calcula();

	//CDialogEx::OnOK();
}


void CTBTabc::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
