// DBTabC.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "DBTabC.h"
#include "afxdialogex.h"


// CDBTabC 对话框

IMPLEMENT_DYNAMIC(CDBTabC, CDialogEx)

CDBTabC::CDBTabC(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDBTabC::IDD, pParent)
{

	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_M = 0.0;
	m_SAs = 0.0;
	m_SAsp = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	m_Asp = 40;
	m_STp = _T("HRB400");
}

CDBTabC::~CDBTabC()
{
}

void CDBTabC::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_DD2_As, m_As);
	DDX_Text(pDX, IDC_DD2_b, m_b);
	DDX_CBString(pDX, IDC_DD2_CS, m_CS);
	DDX_Text(pDX, IDC_DD2_h, m_h);
	DDX_Text(pDX, IDC_DD2_M, m_M);
	DDX_Text(pDX, IDC_DD2_SAs, m_SAs);
	DDX_Text(pDX, IDC_DD2_SAsp, m_SAsp);
	DDX_CBString(pDX, IDC_DD2_ST, m_ST);
	DDX_Text(pDX, IDC_DD2_x, m_x);
	DDX_Text(pDX, IDC_DD2_xb, m_xb);
	DDX_Text(pDX, IDC_DD1_Asp, m_Asp);
	DDX_CBString(pDX, IDC_DD2_STp, m_STp);
}


BEGIN_MESSAGE_MAP(CDBTabC, CDialogEx)
	ON_BN_CLICKED(IDC_DD2_Calcula, &CDBTabC::OnBnClickedDd2Calcula)
	ON_BN_CLICKED(IDC_DD2_Reset, &CDBTabC::OnBnClickedDd2Reset)
	ON_BN_CLICKED(IDC_Out, &CDBTabC::OnBnClickedOut)
END_MESSAGE_MAP()


// CDBTabC 消息处理程序


void CDBTabC::OnOK()
{
	CDBTabC::OnBnClickedDd2Calcula();
	//CDialogEx::OnOK();
}


void CDBTabC::OnBnClickedDd2Calcula()
{
	CWnd::UpdateData(true);

	if(m_b==0.0||m_h==0.0||m_As==0.0||m_CS==""||m_ST==""||m_STp==""||m_SAs == 0.0||m_SAsp == 0.0)
	{
		MessageBox(_T("请输入必要参数"));
		return ;
	}

	h0 = m_h - m_As;
	info.init(m_CS,m_ST);
	info2.init(m_CS,m_STp);

	m_x = (info.ss*m_SAs-info2.ss*m_SAsp)/(info.a1*info.cp*m_b);
	m_xb = info.kb*h0;
	if(m_x < (2*m_As))
	{
		con = 1;
		MessageBox(_T("受压区高度小于2as’，受压钢筋不屈服，假定受压区合力点在受压钢筋重心处计算极限承载力"));
		m_M = m_SAs *  info.ss * (h0 - m_Asp);
	}
	else if(m_x > (info.kb*h0))
	{
		con = 2;
		MessageBox(_T("受拉区钢筋未屈服，取混凝土受压区高度等于极限受压区高度进行计算"));
		m_M = info.a1*info.cp*m_b*m_xb*(h0-m_xb/2)+info2.ss*m_SAsp*(h0-m_Asp);
	}
	else
	{
		con = 3;
		m_M = info.a1*info.cp*m_b*m_x*(h0-m_x/2)+info2.ss*m_SAsp*(h0-m_Asp);
	}
	m_M/= 1000000;
	CWnd::UpdateData(false);
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
	
}


void CDBTabC::OnBnClickedDd2Reset()
{
	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_M = 0.0;
	m_SAs = 0.0;
	m_SAsp = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	m_Asp = 40;
	m_STp = _T("HRB400");
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	CWnd::UpdateData(false);
}

void CDBTabC::xb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Subscript(true),sel.TypeText(x),font.put_Subscript(false);
}
void CDBTabC::sb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Superscript(true),sel.TypeText(x),font.put_Superscript(false);
}

void CDBTabC::sxb(CString center,CString s,CString x)
{
	xb(center,x);
	sb(_T(""),s);
}	
void CDBTabC::OnBnClickedOut()
{
	int i = 0;
	CString s_fy,s_SAsp,
		s_asp,s_xb,s_ft,s_fc,s_a1,s_kb,s_x,s_b,s_M,
		s_h,s_as,s_h0,s_SAs,content,szBuf,s_fyp;

	//参数绑定
	s_h0.Format(_T("%.0lf"),h0);
	s_fc.Format(_T("%.1lf"),info.cp);
	s_a1.Format(_T("%.1lf"),info.a1);
	s_kb.Format(_T("%.3lf"),info.kb);
	s_fy.Format(_T("%.0lf"),info.ss);
	s_ft.Format(_T("%.2lf"),info.ct); 
	s_fyp.Format(_T("%.0lf"),info2.ss);
	s_as.Format(_T("%.0lf"),m_As);
	s_asp.Format(_T("%.0lf"),m_Asp);
	s_b.Format(_T("%.0lf"),m_b);
	s_h.Format(_T("%.0lf"),m_h);
	s_M.Format(_T("%.0lf"),m_M);	
	s_SAs.Format(_T("%.0lf"),m_SAs);
	s_SAsp.Format(_T("%.0lf"),m_SAsp);
	s_x.Format(_T("%.0lf"),m_x);
	s_xb.Format(_T("%.0lf"),(info.kb * h0)); 
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
	sel.TypeText(_T("双筋矩形截面梁复核"));
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(1);
	paragraph.put_Alignment(1);
	sel.TypeParagraph();

	font =sel.get_Font();
	font.put_Size(10);
	//2
	content = _T("    已知条件：截面宽度:b=") 
		+ s_b + _T("mm,截面高度:h=") + s_h
		+ _T("mm,混凝土强度等级:") + m_CS + _T(",");
	sel.TypeText(content);
	paragraphs = doc.get_Paragraphs();//获取当前文档的全部段落
	paragraph = paragraphs.Item(2);//选中第三段为
	paragraph.put_Alignment(0);

	xb(_T("f"),_T("c"));
	sel.TypeText(_T("=")+s_fc + _T("N/mm²,"));

	xb(_T("f"),_T("t"));
	sel.TypeText(_T("=")+s_ft + _T("N/mm²."));

	content = _T("受拉钢筋型号:") + m_ST + _T(",");
	sel.TypeText(content);

	xb(_T("f"),_T("y"));
	sel.TypeText(_T("=")+s_fy + _T("N/mm²."));

	content = _T("受压钢筋型号:") + m_STp + _T(",");
	sel.TypeText(content);

	sxb(_T("f"),_T("'"),_T("y"));
	sel.TypeText(_T("=")+s_fyp + _T("N/mm²."));

	xb(_T("α"),_T("1"));
	sel.TypeText(_T("=")+s_a1 + _T(","));

	xb(_T("ξ"),_T("b"));
	sel.TypeText(_T("=")+s_kb + _T("."));

	content = _T("受拉钢筋合力点到受拉边缘距离:");
	sel.TypeText(content);
	xb(_T("a"),_T("s"));
	sel.TypeText(_T("=")+s_as + _T("mm,"));

	content = _T("受压钢筋合力点到受拉边缘距离:");
	sel.TypeText(content);
	sxb(_T("a"),_T("'"),_T("s"));
	sel.TypeText(_T("=")+s_asp + _T("mm."));

	content = _T("受拉钢筋截面积:");
	sel.TypeText(content);
	xb(_T("A"),_T("s"));
	sel.TypeText(_T("=")+s_SAs + _T("mm²,"));

	content = _T("受压钢筋截面积:");
	sel.TypeText(content);
	sxb(_T("A"),_T("'"),_T("s"));
	sel.TypeText(_T("=")+s_SAsp + _T("mm²."));
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
		sxb(_T(" - f"),_T("'"),_T("y"));
		sxb(_T("A"),_T("'"),_T("s"));
		xb(_T(",α"),_T("1"));
		xb(_T("f"),_T("c"));
		xb(_T("b) = "),_T(""));
		sel.TypeText(COperate::fraction(s_fy+_T("×")+s_SAs+_T(" - ")+s_fyp+_T("×")+s_SAsp,
		s_a1+_T("×")+s_fc+_T("×")+s_b)+_T(" = ")+s_x+_T(" (mm)"));
	}
	fields.ToggleShowCodes();
	sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.TypeParagraph();
	//6
	xb(_T("x"),_T("b"));
	xb(_T(" = ξ"),_T("b"));
	xb(_T(" = h"),_T("0"));
	sel.TypeText(_T(" = ")+s_kb+_T("×")+s_h0+_T(" = ")+s_xb+_T(" (mm)"));
	sel.TypeParagraph();
	if(con == 1)
	{
		//7
		sel.TypeText(_T("计算得:"));
		sxb(_T("2a"),_T("'"),_T("s"));
		xb(_T(" > x,"),_T(""));
		sel.TypeText(_T("受压区高度小于2as’，受压钢筋不屈服，假定受压区合力点在受压钢筋重心处计算极限承载力:"));
		sel.TypeParagraph();
		//8
		xb(_T("M = A"),_T("s"));
		xb(_T("f"),_T("y"));
		xb(_T("(h"),_T("0"));
		sxb(_T(" - a"),_T("'"),_T("s"));
		sel.TypeText(_T(") = ")+s_SAs+_T("×")+s_fy+_T("×")+_T("(")+s_h0+_T(" - ")+s_asp+_T(")")+_T(" = ")+s_M+_T(" (kN∙m)"));
	}
	else if(con == 2)
	{
		//7
		sel.TypeText(_T("计算得:"));
		xb(_T("x > x"),_T("b"));
		sel.TypeText(_T("受拉区钢筋未屈服，取混凝土受压区高度等于极限受压区高度进行计算:"));
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
			xb(_T("EQ M = α"),_T("1"));
			xb(_T("f"),_T("c"));
			xb(_T("bx"),_T("b"));
			xb(_T("(h"),_T("0"));
			xb(_T(" - \\f(x"),_T("b"));
			sxb(_T(",2)) + f"),_T("'"),_T("y"));
			sxb(_T("A"),_T("'"),_T("s"));
			xb(_T("(h"),_T("0"));
			sxb(_T(" - a"),_T("'"),_T("s"));
			sel.TypeText(_T(")"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//9
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
			sel.TypeText(_T("EQ = ")+s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_xb+_T("×")+
			_T("(")+s_h0+_T(" - ")+COperate::fraction(s_xb,_T("2"))+_T(") + ")
			+s_fyp+_T("×")+s_SAsp+_T("×")+_T("(")+s_h0+_T(" - ")+s_asp+_T(") = ")+s_M+_T(" (kN∙m)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	}
	else
	{
		//7
		sel.TypeText(_T("计算得:"));
		sxb(_T("2a"),_T("'"),_T("s"));
		xb(_T(" < x <x"),_T("b"));
		sel.TypeText(_T(",表明受压区高度处于正常位置"));
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
			xb(_T("EQ M = α"),_T("1"));
			xb(_T("f"),_T("c"));
			xb(_T("bx"),_T(""));
			xb(_T("(h"),_T("0"));
			xb(_T(" - \\f(x"),_T(""));
			sxb(_T(",2)) + f"),_T("'"),_T("y"));
			sxb(_T("A"),_T("'"),_T("s"));
			xb(_T("(h"),_T("0"));
			sxb(_T(" - a"),_T("'"),_T("s"));
			sel.TypeText(_T(")"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//9
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
			sel.TypeText(_T("EQ = ")+s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_x+_T("×")+
			_T("(")+s_h0+_T(" - ")+COperate::fraction(s_x,_T("2"))+_T(") + ")
			+s_fyp+_T("×")+s_SAsp+_T("×")+_T("(")+s_h0+_T(" - ")+s_asp+_T(") = ")+s_M+_T(" (kN∙m)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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


void CDBTabC::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
