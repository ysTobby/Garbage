// DBTabD2.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "DBTabD2.h"
#include "afxdialogex.h"


// CDBTabD2 对话框

IMPLEMENT_DYNAMIC(CDBTabD2, CDialogEx)

	CDBTabD2::CDBTabD2(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDBTabD2::IDD, pParent)
{

	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_M = 0.0;
	//  m_MAs = 0.0;
	m_SAs = 0.0;
	m_SAsp = 0.0;
	m_sdia = _T("");
	m_snum = _T("");
	m_sres = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	m_Asp = 40;
	m_STp = _T("HRB400");
}

CDBTabD2::~CDBTabD2()
{
}

void CDBTabD2::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_DD2_As, m_As);
	DDX_Text(pDX, IDC_DD2_b, m_b);
	DDX_CBString(pDX, IDC_DD2_CS, m_CS);
	DDX_Text(pDX, IDC_DD2_h, m_h);
	DDX_Text(pDX, IDC_DD2_M, m_M);
	//  DDX_Text(pDX, IDC_DD2_MAs, m_MAs);
	DDX_Text(pDX, IDC_DD2_SAs, m_SAs);
	DDX_Text(pDX, IDC_DD2_SAsp, m_SAsp);
	DDX_CBString(pDX, IDC_DD2_Sdia, m_sdia);
	DDX_CBString(pDX, IDC_DD2_Snum, m_snum);
	DDX_Text(pDX, IDC_DD2_Sres, m_sres);
	DDX_CBString(pDX, IDC_DD2_ST, m_ST);
	DDX_Text(pDX, IDC_DD2_x, m_x);
	DDX_Text(pDX, IDC_DD2_xb, m_xb);
	DDX_Text(pDX, IDC_DD1_Asp, m_Asp);
	DDX_CBString(pDX, IDC_DD2_STp, m_STp);
}


BEGIN_MESSAGE_MAP(CDBTabD2, CDialogEx)
	ON_BN_CLICKED(IDC_DD2_Scal, &CDBTabD2::OnBnClickedDd2Scal)
	ON_BN_CLICKED(IDC_DD2_Calcula, &CDBTabD2::OnBnClickedDd2Calcula)
	ON_BN_CLICKED(IDC_DD2_Reset, &CDBTabD2::OnBnClickedDd2Reset)
	ON_BN_CLICKED(IDC_Out, &CDBTabD2::OnBnClickedOut)
END_MESSAGE_MAP()


// CDBTabD2 消息处理程序


void CDBTabD2::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	OnBnClickedDd2Calcula();
	//CDialogEx::OnOK();
}


void CDBTabD2::OnBnClickedDd2Scal()
{
	// TODO: 在此添加控件通知处理程序代码
	CWnd::UpdateData(true);
	m_sres = (double)1/4 * (_ttoi(m_sdia)*_ttoi(m_sdia)) * 3.14159 * _ttoi(m_snum);
	CWnd::UpdateData(false);
}


void CDBTabD2::OnBnClickedDd2Calcula()
{
	CWnd::UpdateData(true);

	if(m_b==0.0||m_h==0.0||m_M==0.0||m_As==0.0||m_CS==""||m_ST==""||m_SAsp==0.0||m_STp=="")
	{
		MessageBox(_T("请输入必要参数"));
		return ;
	}

	info.init(m_CS,m_ST);
	info2.init(m_CS,m_STp);
	h0 = m_h - m_As;

	m_M = m_M * 1000000;
	Mu1 = info2.ss * m_SAsp * (h0 - m_Asp);
	Mu2  = m_M - Mu1;

	arfas  = Mu2 / (info.a1*info.cp*m_b*h0*h0);
	if((1-2*arfas)<0)
	{
		con = 0;
		m_M /= 1000000;
		GetDlgItem(IDC_Out)->EnableWindow(TRUE);
		MessageBox(_T("受压区高度超过界限受压区高度，超筋，表明受压钢筋面积不够，可以按照受压钢筋未知重新进行配筋"));
		return;
	}
	else
	{
		kesai = 1 - sqrt((1-2*arfas));
		m_x  = kesai * h0;
		m_xb = info.kb * h0;
		if(((2*m_As)<=m_x) && (m_x<=(info.kb*h0)))
		{
			con = 1;
			As2 = (info.a1*info.cp*m_b*m_x) / (info.ss);
			m_SAs = As2 + m_SAsp*info2.ss/info.ss;
			m_M /= 1000000;
			CWnd::UpdateData(false);
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
		}
		else if((2*m_As)>m_x)
		{
			con = 2;
			MessageBox(_T("受压区高度小于2as’，说明受压钢筋不屈服，假定受压区合力点在受压钢筋重心处计算受拉钢筋面积"));
			m_SAs=m_M/(info.ss*(h0-m_Asp));
			m_M /= 1000000;
			CWnd::UpdateData(false);
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
		}
		else
		{
			con = 3;
			m_M /= 1000000;
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
			MessageBox(_T("受压区高度超过界限受压区高度，超筋，表明受压钢筋面积不够，可以按照受压钢筋未知重新进行配筋"));
		}
	}
}


void CDBTabD2::OnBnClickedDd2Reset()
{
	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_M = 0.0;
	//  m_MAs = 0.0;
	m_SAs = 0.0;
	m_SAsp = 0.0;
	m_sdia = _T("");
	m_snum = _T("");
	m_sres = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	m_Asp = 40;
	m_STp = _T("HRB400");
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	CWnd::UpdateData(false);
}

void CDBTabD2::xb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Subscript(true),sel.TypeText(x),font.put_Subscript(false);
}
void CDBTabD2::sb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Superscript(true),sel.TypeText(x),font.put_Superscript(false);
}

void CDBTabD2::sxb(CString center,CString s,CString x)
{
	xb(center,x);
	sb(_T(""),s);
}	


void CDBTabD2::OnBnClickedOut()
{
	int i = 0;
	CString s_Mu1,s_Mu2,s_fy,s_kesai,s_SAs1,s_SAs2,s_SAsp,
		s_asp,s_xb,s_ft,s_fc,s_a1,s_kb,s_arfas,s_x,s_b,s_M,
		s_h,s_as,s_h0,s_SAs,content,szBuf,s_fyp;

	//参数绑定
	s_arfas.Format(_T("%.3lf"),arfas);
	s_SAs1.Format(_T("%.1lf"),As1);
	s_SAs2.Format(_T("%.1lf"),As2);
	s_h0.Format(_T("%.0lf"),h0);
	s_fc.Format(_T("%.1lf"),info.cp);
	s_a1.Format(_T("%.1lf"),info.a1);
	s_kb.Format(_T("%.3lf"),info.kb);
	s_fy.Format(_T("%.0lf"),info.ss);
	s_ft.Format(_T("%.2lf"),info.ct); 
	s_fyp.Format(_T("%.0lf"),info2.ss);
	s_kesai.Format(_T("%.0lf"),kesai);
	s_as.Format(_T("%.0lf"),m_As);
	s_asp.Format(_T("%.0lf"),m_Asp);
	s_b.Format(_T("%.0lf"),m_b);
	s_h.Format(_T("%.0lf"),m_h);
	s_M.Format(_T("%.0lf"),m_M);	
	s_SAs.Format(_T("%.0lf"),m_SAs);
	s_SAsp.Format(_T("%.0lf"),m_SAsp);
	s_x.Format(_T("%.0lf"),m_x);
	s_xb.Format(_T("%.0lf"),(info.kb * h0)); 
	s_Mu1.Format(_T("%.2lf"),Mu1/1000000); 
	s_Mu2.Format(_T("%.2lf"),Mu2/1000000); 
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
	sel.TypeText(_T("双筋矩形截面梁第二类受弯设计"));
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(1);
	paragraph.put_Alignment(1);
	sel.TypeParagraph();

	font =sel.get_Font();
	font.put_Size(10);
	//2
	content = _T("    已知条件：弯矩设计值:M=") + s_M + _T("kN∙m,截面宽度:b=") 
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
	xb(_T("  = h - a"),_T("s"));
	sel.TypeText(_T(" = ")+s_h + _T(" - ") + s_as + _T(" = ") + s_h0 + _T("（mm）"));
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
		xb(_T("EQ M"),_T("u1"));
		sxb(_T(" = f"),_T("'"),_T("y"));
		sxb(_T("A"),_T("'"),_T("s"));
		xb(_T("( h"),_T("0"));
		sxb(_T(" - a"),_T("'"),_T("s"));
		sel.TypeText(_T(" ) = ")+s_fyp + _T("×") + s_SAsp + _T("×(") + s_h0 + 
			_T(" - ") + s_asp + _T(") = ") +
			s_Mu1 + _T("（kN∙m）"));
	}
	fields.ToggleShowCodes();
	sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.TypeParagraph();
	//6
	xb(_T("M"),_T("u2"));
	xb(_T(" = M - M"),_T("u1"));
	sel.TypeText(_T(" = ")+s_M + _T(" - ") + s_Mu1 + _T(" = ") +  
		s_Mu2 + _T("（kN∙m）"));
	sel.TypeParagraph();
	//7
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
		xb(_T(")"),_T(""));
		sel.TypeText(_T(" = ")+COperate::fraction(s_Mu2+_T("×10⁶"),s_a1+_T("×")+s_fc+_T("×")+s_b
			+_T("×")+s_h0+_T("²")) + _T(" = ")+s_arfas);
	}
	fields.ToggleShowCodes();
	sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.TypeParagraph();
	if(con == 0)
	{
		//8
		xb(_T("计算得:1 - 2α"),_T("s"));
		xb(_T(" < 0,"),_T(""));
		sel.TypeText(_T("受压区高度超过界限受压区高度，超筋，表明受压钢筋面积不够，可以按照受压钢筋未知重新进行配筋"));
	}
	else 
	{
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
			xb(_T("EQ x = (1 - \\r(1 - 2α"),_T("s"));
			xb(_T("))h"),_T("0"));
			sel.TypeText(_T(" = (1 - ") + COperate::radical(_T("1 - 2×")+ s_arfas)+_T(")×")+s_h0+_T(" = ")+s_x + _T(" (mm)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();

		//9
		xb(_T("x"),_T("b"));
		xb(_T(" = ξ"),_T("b"));
		xb(_T("h"),_T("0"));
		sel.TypeText(_T(" = ")+s_kb+_T("×")+s_h0+_T(" = ")+s_xb+_T(" (mm)"));
		sel.TypeParagraph();
		if(con == 1)
		{
			//10
			sel.TypeText(_T("求得:"));
			xb(_T("2a"),_T("s"));
			xb(_T(" ≤ x ≤ x"),_T("b"));
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
				xb(_T("EQ A"),_T("s2"));
				xb(_T(" = \\f(α"),_T("1"));
				xb(_T("f"),_T("c"));
				xb(_T("bx,f"),_T("y"));
				xb(_T(") = "),_T(""));
				sel.TypeText(COperate::fraction(s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_x,s_fy)+_T(" = ")+s_SAs2 + _T(" (mm²)"));
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
				xb(_T("EQ A"),_T("s"));
				xb(_T(" = A"),_T("s2"));
				sxb(_T(" + \\f(f"),_T("'"),_T("y"));
				xb(_T(",f"),_T("y"));
				sxb(_T(")A"),_T("'"),_T("s"));
				sel.TypeText(_T(" = ")+s_SAs2 +_T(" + ") + COperate::fraction(s_fyp,s_fy)+_T("×")+s_SAsp+_T(" = ")+s_SAs + _T(" (mm²)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
		}
		else if(con == 2)
		{
			//10
			xb(_T("计算得:2a"),_T("s"));
			xb(_T(" > x"),_T(""));
			sel.TypeText(_T(",说明受压钢筋不屈服，假定受压区合力点在受压钢筋重心处计算受拉钢筋面积."));
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
				xb(_T(" = \\f(M,f"),_T("y"));
				xb(_T("(h"),_T("0"));
				sxb(_T(" - a"),_T("'"),_T("s"));
				xb(_T(")) = "),_T(""));
				sel.TypeText(COperate::fraction(s_M+_T("×10⁶"),s_fy+_T("×(")+s_h0+_T(" - ")+s_asp+_T(")"))+_T(" = ")+s_SAs+_T(" (mm²)."));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		}
		else
		{
			//10
			xb(_T("x > x"),_T("b"));
			sel.TypeText(_T("受压区高度超过界限受压区高度，超筋，表明受压钢筋面积不够，可以按照受压钢筋未知重新进行配筋"));
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


void CDBTabD2::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
