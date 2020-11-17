// DBTabD1.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "DBTabD1.h"
#include "afxdialogex.h"


// CDBTabD1 对话框

IMPLEMENT_DYNAMIC(CDBTabD1, CDialogEx)

CDBTabD1::CDBTabD1(CWnd* pParent /*=NULL*/)
	: CDialogEx(CDBTabD1::IDD, pParent)
{

	m_As = 40;
	m_M = 0.0;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_MAs = 0.0;
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

CDBTabD1::~CDBTabD1()
{
}

void CDBTabD1::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_DD1_As, m_As);
	DDX_Text(pDX, IDC_DD1_M, m_M);
	DDX_Text(pDX, IDC_DD1_b, m_b);
	DDX_CBString(pDX, IDC_DD1_CS, m_CS);
	DDX_Text(pDX, IDC_DD1_h, m_h);
	DDX_Text(pDX, IDC_DD1_MAs, m_MAs);
	DDX_Text(pDX, IDC_DD1_SAs, m_SAs);
	DDX_Text(pDX, IDC_DD1_SAsp, m_SAsp);
	DDX_CBString(pDX, IDC_DD1_Sdia, m_sdia);
	DDX_CBString(pDX, IDC_DD1_Snum, m_snum);
	DDX_Text(pDX, IDC_DD1_Sres, m_sres);
	DDX_CBString(pDX, IDC_DD1_ST, m_ST);
	DDX_Text(pDX, IDC_DD1_x, m_x);
	DDX_Text(pDX, IDC_DD1_xb, m_xb);
	DDX_Text(pDX, IDC_DD1_Asp, m_Asp);
	DDX_CBString(pDX, IDC_DD1_STp, m_STp);
}


BEGIN_MESSAGE_MAP(CDBTabD1, CDialogEx)
	ON_BN_CLICKED(IDC_DD1_Scal, &CDBTabD1::OnBnClickedDd1Scal)
	ON_BN_CLICKED(IDC_DD1_Calcula, &CDBTabD1::OnBnClickedDd1Calcula)
//	ON_WM_KEYDOWN()
	ON_BN_CLICKED(IDC_DD1_Reset, &CDBTabD1::OnBnClickedDd1Reset)
	ON_BN_CLICKED(IDC_Out, &CDBTabD1::OnBnClickedOut)
END_MESSAGE_MAP()


// CDBTabD1 消息处理程序


void CDBTabD1::OnBnClickedDd1Scal()
{
	// TODO: 在此添加控件通知处理程序代码
	CWnd::UpdateData(true);//将数据从对话框的控件中传送到对应的数据成员中
	m_sres = (double)1/4 * (_ttoi(m_sdia)*_ttoi(m_sdia)) * 3.14159 * _ttoi(m_snum);
	CWnd::UpdateData(false);//
}


void CDBTabD1::OnBnClickedDd1Calcula()
{
	CWnd::UpdateData(true);//将数据从对话框的控件中传送到对应的数据成员中

	if(m_b==0.0||m_h==0.0||m_M==0.0||m_As==0.0||m_Asp==0.0||m_CS==""||m_ST==""||m_STp=="")
	{
		MessageBox(_T("请输入必要参数"));
		return ;
	}

	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
	con = 0;

	info.init(m_CS,m_ST);
	info2.init(m_CS,m_STp);
	h0 = m_h - m_As;

	m_M = m_M * 1000000;
	m_xb = h0 * info.kb;
	m_x = m_xb;
	m_SAsp = (m_M-info.a1*info.cp*h0*h0*m_b*info.kb*(1-0.5*info.kb))/(info2.ss*(h0-m_Asp));
	if(m_SAsp<0)
	{
		con = 1;
		m_M /= 1000000;
		MessageBox(_T("计算得As'为负值！\n1.可按单筋梁设计.\n2.仍采用双筋梁设计时，As'可按构造配筋，然后视As'为已知进行设计."));
		return;
	}
	m_SAs = info.a1*info.cp*m_b*h0*info.kb/info.ss + m_SAsp*info2.ss/info.ss;
	rAs = m_SAs;
	pmin = 0.002;
	if(pmin < 0.45*info.ct/info.ss)
		pmin = 0.45*info.ct/info.ss;
	m_MAs = pmin*m_b*m_h;
	if(m_SAs<m_MAs)
	{
		con = 2;
		MessageBox(_T("最小配筋率不满足要求按最小配筋面积配筋"));
		m_SAs = m_MAs;
	}
	m_M = m_M / 1000000;
	CWnd::UpdateData(false);
}

void CDBTabD1::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	OnBnClickedDd1Calcula();
	//CDialogEx::OnOK();
}


void CDBTabD1::OnBnClickedDd1Reset()
{
	m_As = 40;
	m_M = 0.0;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_MAs = 0.0;
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
	m_STp = _T("HRB400");
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);//用来改变不可点击按钮的状态
	CWnd::UpdateData(false);
}

void CDBTabD1::xb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Subscript(true),sel.TypeText(x),font.put_Subscript(false);
}
void CDBTabD1::sb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Superscript(true),sel.TypeText(x),font.put_Superscript(false);
}

void CDBTabD1::sxb(CString center,CString s,CString x)
{
	xb(center,x);
	sb(_T(""),s);
}

void CDBTabD1::OnBnClickedOut()
{
	int i = 0;
	CString s_fy,s_SAsp,s_rAs,s_asp,s_pmin,s_xb,s_ft,s_fc,s_a1,s_kb,s_arfas,s_x,s_b,s_M,s_Asmin,s_h,s_as,s_h0,s_SAs,resultM,content,szBuf,s_fyp;

	//参数绑定
	s_arfas.Format(_T("%.3lf"),arfas);
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
	s_Asmin.Format(_T("%.0lf"),m_MAs); 
	s_rAs.Format(_T("%.0lf"),rAs); //用来记录直接计算得到的钢筋面积，在没有进行最小配筋面积验算处理前的结果
	s_SAs.Format(_T("%.0lf"),m_SAs);
	s_SAsp.Format(_T("%.0lf"),m_SAsp);
	s_x.Format(_T("%.0lf"),m_x);
	s_xb.Format(_T("%.0lf"),(info.kb * h0)); 
	s_pmin.Format(_T("%.4lf"),pmin); 
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
	doc = app.get_ActiveDocument();//活动文档
	COleVariant vTrue((short)TRUE),vFalse((short)FALSE),vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	COleVariant formula;

	sel=app.get_Selection();//Selection表示输入点，即光标闪烁的那个地方
	font =sel.get_Font();
	font.put_Name(_T("Times New Roman"));//设置字体
	font.put_Size(18);
	font.put_Color(WdColor::wdColorBlack);
	font.put_Bold(0);
	//1
	sel.TypeText(_T("双筋矩形截面梁第一类受弯设计"));//调用函数Selection::TypeText 向WORD发送字符
	paragraphs = doc.get_Paragraphs();//获取当前文档的全部段落
	paragraph = paragraphs.Item(1);//选中第三段为
	paragraph.put_Alignment(1);//设置文本格式：0左对齐，1居中对齐，2右对齐
	sel.TypeParagraph();

	font =sel.get_Font();
	font.put_Size(10);
	//2
	content = _T("    已知条件：弯矩设计值:M=") + s_M + _T("kN∙m,截面宽度:b=") + s_b + _T("mm,截面高度:h=") + s_h
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
	sel.TypeText(_T("=")+s_a1 + _T("."));

	xb(_T("ξ"),_T("b"));
	sel.TypeText(_T("=")+s_kb + _T("."));

	content = _T("受拉钢筋合力点到受拉边缘距离:");
	sel.TypeText(content);
	xb(_T("a"),_T("s"));
	sel.TypeText(_T("=")+s_as + _T("mm"));

	content = _T(",受压钢筋合力点到受拉边缘距离:");
	sel.TypeText(content);
	sxb(_T("a"),_T("'"),_T("s"));
	sel.TypeText(_T("=")+s_asp + _T("mm."));
	sel.TypeParagraph();

	fields.ToggleShowCodes();
	fields = sel.get_Fields();
	range = sel.get_Range();
	formula = _T("");
	fields.Add(range,vOpt,&formula,vFalse);
	sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdExtend));
	CString a = sel.get_Text();//a为空正常a为点多运行一次
	if(a == _T("."))
	{
		i = 1;
		sel.MoveRight(COleVariant((short)3),COleVariant((short)3),COleVariant((short)wdExtend));
		fields.ToggleShowCodes();
		paragraphs = doc.get_Paragraphs();//获取当前文档的全部段落
		paragraph = paragraphs.Item(3);//选中第三段为
		range = paragraph.get_Range();
		sel.Cut();
		sel.TypeParagraph();
	}
	else
	{
		sel.MoveLeft(COleVariant((short)1),COleVariant((short)3),COleVariant((short)wdMove));
		sel.MoveRight(COleVariant((short)1),COleVariant((short)4),COleVariant((short)wdExtend));//第一个参数可能是多少格多少格一次，第二个可能是多少次
		sel.Cut();
		sel.TypeParagraph();
	}

	//3
	sel.TypeText(_T("    解:"));
	sel.TypeParagraph();

	//4
	xb(_T("h"),_T("0"));
	xb(_T(" = h - a"),_T("s"));
	sel.TypeText(_T(" = ")+s_h + _T(" - ") + s_as + _T(" = ") + s_h0 + _T("（mm）"));
	sel.TypeParagraph();
	//5
	xb(_T("x"),_T("b"));
	xb(_T(" = ξ"),_T("b"));
	xb(_T("∙h"),_T("0"));
	sel.TypeText(_T(" = ")+s_kb + _T("×") + s_h0 + _T(" = ")+s_xb + _T(" (mm)"));
	sel.TypeParagraph();
	//6
	sel.TypeText(_T("取:"));
	xb(_T("x = x"),_T("b"));
	sel.TypeText(_T(" = ")+s_xb + _T(" (mm)"));
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
		sxb(_T("EQ A"),_T("'"),_T("s"));
		xb(_T(" = \\f(M - α"),_T("1"));
		xb(_T("f"),_T("c"));
		xb(_T("bx(h"),_T("0"));
		sxb(_T(" - \\f(x,2)),f"),_T("'"),_T("y"));
		xb(_T("(h"),_T("0"));
		sxb(_T(" - a"),_T("'"),_T("s"));
		xb(_T(")) = "),_T(""));
		sel.TypeText(COperate::fraction((s_M + _T("×10⁶")+_T("-")+s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_x+_T("×(")+s_h0+_T(" - \\f(")+s_x+_T(",2))")),(s_fyp+_T("×(")+s_h0+_T(" - ")+s_asp+_T(")")))+_T(" = ") + s_SAsp + _T(" (mm²)"));
	}
	fields.ToggleShowCodes();
	sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.TypeParagraph();
	
	if(con != 1)
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
			xb(_T("EQ A"),_T("s"));
			xb(_T(" = \\f( α"),_T("1"));
			xb(_T("f"),_T("c"));
			sxb(_T("bx + f"),_T("'"),_T("y"));
			sxb(_T("A"),_T("'"),_T("s"));
			xb(_T(",f"),_T("y)"));
			sel.TypeText(_T(" = ")+COperate::fraction((s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_x+_T(" + ")
			+s_fyp+_T("×")+s_SAsp
			),(s_fy))+_T(" = ")+s_SAs+_T(" (mm²)"));
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
			xb(_T("EQ A"),_T("smin"));
			xb(_T(" = max{0.2%,0.45\\f(f"),_T("t"));
			xb(_T(", f"),_T("y"));
			xb(_T(")}∙b∙h = "),_T(""));
			sel.TypeText(s_pmin + _T("×") + s_b + _T("×") + s_h + _T(" = ") + s_Asmin+_T(" (mm²)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(con == 0)
		{
			//10
			xb(_T("因: A"),_T("s"));
			xb(_T(" ≥ A"),_T("smin"));
			sel.TypeText(_T("，满足最小配筋率要求"));
			sel.TypeParagraph();
			//11
			sel.TypeText(_T("综上所述:"));
			xb(_T("A"),_T("s"));
			sel.TypeText(_T(" = ")+s_SAs + _T("(mm²),"));
			sxb(_T("A"),_T("'"),_T("s"));
			sel.TypeText(_T(" = ")+s_SAsp + _T("(mm²),"));
		}
		else
		{
			//10
			sel.TypeText(_T("因: "));

			xb(_T("A"),_T("s"));
			xb(_T(" ≤ A"),_T("smin"));
			sel.TypeText(_T("，不满足最小配筋率要求，取:"));
			xb(_T("A"),_T("s"));
			xb(_T(" = A"),_T("smin"));
			sel.TypeParagraph();

			//11
			sel.TypeText(_T("综上所述:"));
			xb(_T("A"),_T("s"));
			sel.TypeText(_T(" = ")+s_SAs + _T("(mm²),"));
			sxb(_T("A"),_T("'"),_T("s"));
			sel.TypeText(_T(" = ")+s_SAsp + _T("(mm²)."));
		}
	}
	else
	{
		//计算得As'为负值！\n1.可按单筋梁设计.\n2.仍采用双筋梁设计时，As'可按构造配筋，然后视As'为已知进行设计
		//8
		sel.TypeText(_T("计算得的受压钢筋面积为负值，可按照单筋梁进行设计，若采用双筋梁设计时应按照构造配置受压钢筋，并将受压钢筋面积看做已知"));
	}




	font.ReleaseDispatch();//注意：所有东西用完之后一定要ReleaseDispatch，否则报错；不过最好像例子2中，在最后集中ReleaseDispatch
	paragraph.ReleaseDispatch();
	paragraphs.ReleaseDispatch();
	range.ReleaseDispatch();
	fields.ReleaseDispatch();
	doc.ReleaseDispatch();
	sel.ReleaseDispatch();
	docs.ReleaseDispatch(); 

	//调用Quit退出WORD应用程序。当然不调用也可以，那样的话WORD还在运行着那
	//app.Quit(new CComVariant(FALSE),new CComVariant(),new CComVariant());
	app.ReleaseDispatch();   //释放对象指针。切记，必须调用
	//卸载com

	CoUninitialize();//对应CoInitialize
}


void CDBTabD1::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
