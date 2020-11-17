// SBTabD.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "SBTabD.h"
#include "afxdialogex.h"
//#include "information.h"


// CSBTabD 对话框

IMPLEMENT_DYNAMIC(CSBTabD, CDialogEx)

	CSBTabD::CSBTabD(CWnd* pParent /*=NULL*/)
	: CDialogEx(CSBTabD::IDD, pParent)
	, m_snum("0")
{
	m_M = 0.0;
	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_MAs = 0.0;
	m_SAs = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	//  m_sdia = 0;
	m_sres = 0.0;
	m_snum = _T("");
	m_sdia = _T("");
}

CSBTabD::~CSBTabD()
{
}

void CSBTabD::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_SD_M, m_M);
	DDX_Text(pDX, IDC_SD_As, m_As);
	DDX_Text(pDX, IDC_SD_b, m_b);
	DDX_CBString(pDX, IDC_SD_CS, m_CS);
	DDX_Text(pDX, IDC_SD_h, m_h);
	DDX_Text(pDX, IDC_SD_MAs, m_MAs);
	DDX_Text(pDX, IDC_SD_SAs, m_SAs);
	DDX_CBString(pDX, IDC_SD_ST, m_ST);
	DDX_Text(pDX, IDC_SD_x, m_x);
	DDX_Text(pDX, IDC_SD_xb, m_xb);
	//  DDX_CBIndex(pDX, IDC_SD_Snum, m_snum);
	//  DDX_CBIndex(pDX, IDC_SD_Sdia, m_sdia);
	DDX_Text(pDX, IDC_SD_Sres, m_sres);
	DDX_CBString(pDX, IDC_SD_Snum, m_snum);
	DDX_CBString(pDX, IDC_SD_Sdia, m_sdia);
}


BEGIN_MESSAGE_MAP(CSBTabD, CDialogEx)
	ON_BN_CLICKED(IDC_SD_Calcula, &CSBTabD::OnBnClickedSdCalcula)
	ON_BN_CLICKED(IDC_SD_Reset, &CSBTabD::OnBnClickedSdReset)
	ON_BN_CLICKED(IDC_Out, &CSBTabD::OnBnClickedOut)
	ON_BN_CLICKED(IDC_SD_Scal, &CSBTabD::OnBnClickedSdScal)
END_MESSAGE_MAP()


// CSBTabD 消息处理程序


void CSBTabD::OnBnClickedSdCalcula()
{
	// TODO: 在此添加控件通知处理程序代码
	CWnd::UpdateData(true);//将数据从对话框的控件中传送到对应的数据成员中

	if(m_b==0.0||m_h==0.0||m_M==0.0||m_As==0.0||m_CS==""||m_ST=="")
	{
		MessageBox(_T("请输入必要参数"));
		return ;
	}

	//Cinformation info;//根据混凝土强度和钢筋种类获取参数
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);//用来改变不可点击按钮的状态

	info.init(m_CS,m_ST);

	h0 = m_h - m_As;

	if((m_M * 1000000) >= (info.a1 * info.cp * m_b * info.kb * h0 * (h0 - info.kb * h0 / 2)))
	{
		con = 0;
		MessageBox(_T("此为超筋梁，建议可以增大截面尺寸，提高混凝土强度和配置双筋梁"));
	}
	else
	{
		con = 1;
		arfas = m_M * 1000000 / (info.a1 * info.cp * m_b * h0 *h0);
		m_x = (1 - sqrt(1 - 2 * arfas)) * h0;
		m_SAs = m_M * 1000000 / (info.ss * h0 * (1 - (m_x / h0) / 2));
		rAs = m_SAs;
		m_xb = info.kb * h0;

		m_MAs = 0.002 * m_b * m_h;
		pmin = 0.002;
		if(m_MAs < ((0.45 * info.ct / info.ss) * m_b * m_h))
		{
			m_MAs = (0.45 * info.ct / info.ss) * m_b * m_h;
			pmin = (0.45 * info.ct / info.ss);
		}

		if(m_SAs < m_MAs)
		{
			MessageBox(_T("计算面积小于最小配筋面积，按照最小配筋面积配筋"));
			con = 2;
			m_SAs = m_MAs;
		}

		CWnd::UpdateData(false);
	}
	/*
	m_SAs = info.cp;
	m_x = info.ct;
	m_xb = info.ss;
	CWnd::UpdateData(false);
	*/
}

void CSBTabD::OnBnClickedSdReset()
{
	// TODO: 在此添加控件通知处理程序代码
	m_M = 0.0;
	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_MAs = 0.0;
	m_SAs = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	//  m_sdia = 0;
	m_sres = 0.0;
	m_snum = _T("");
	m_sdia = _T("");
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);//用来改变不可点击按钮的状态
	CWnd::UpdateData(false);
}


void CSBTabD::OnBnClickedOut()
{
	int i = 0;
	CString s_fy,s_pmin,s_xb,s_ft,s_fc,s_a1,s_kb,s_arfas,s_x,s_b,s_M,s_Asmin,s_h,s_as,s_h0,s_SAs,resultM,content,szBuf;
	CoInitialize(NULL);//初始化COM，与最后一行CoUninitialize对应

	CApplication app;
	if(!app.CreateDispatch(_T("word.application"))) //启动WORD
	{
		AfxMessageBox(_T("居然你连OFFICE都没有安装吗?"));
		return;
	}
	//创建word进程默认是不可见的
	app.put_Visible(TRUE); //设置WORD可见。

	CDocuments docs = app.get_Documents();
	docs.Add(new CComVariant(_T("")),new CComVariant(FALSE),new CComVariant(0),new CComVariant());//创建新文档
	CDocument0 doc = app.get_ActiveDocument();//活动文档
	CParagraphs paragraphs;
	CParagraph paragraph;
	CFields fields;
	CRange range;
	COleVariant vTrue((short)TRUE),vFalse((short)FALSE),vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	COleVariant formula;

	//参数绑定
	s_as.Format(_T("%.0lf"),m_As);
	s_h.Format(_T("%.0lf"),m_h);
	s_h0.Format(_T("%.0lf"),h0);
	s_M.Format(_T("%.0lf"),m_M);
	s_a1.Format(_T("%.1lf"),info.a1);
	s_fc.Format(_T("%.1lf"),info.cp);
	s_b.Format(_T("%.0lf"),m_b);
	s_kb.Format(_T("%.3lf"),info.kb);
	s_fy.Format(_T("%.0lf"),info.ss);
	s_arfas.Format(_T("%.3lf"),arfas);
	s_x.Format(_T("%.3lf"),m_x);
	s_Asmin.Format(_T("%.0lf"),m_MAs); 
	s_SAs.Format(_T("%.0lf"),rAs); 
	s_ft.Format(_T("%.2lf"),info.ct); 
	s_xb.Format(_T("%.0lf"),(info.kb * h0)); 
	s_pmin.Format(_T("%.4lf"),pmin); 
	resultM.Format(_T("%.1lf"),(info.a1 * info.cp * m_b * info.kb * h0 * (h0 - info.kb * h0 / 2))/1000000);
	//绑定结束

	CSelection sel=app.get_Selection();//Selection表示输入点，即光标闪烁的那个地方
	CFont0 font =sel.get_Font();
	font.put_Name(_T("Times New Roman"));//设置字体
	font.put_Size(18);
	font.put_Color(WdColor::wdColorBlack);
	font.put_Bold(0);
	//1
	sel.TypeText(_T("单筋矩形截面梁受弯设计"));//调用函数Selection::TypeText 向WORD发送字符
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

	sel.TypeText(_T("f"));
	font.put_Subscript(true);//激活下标
	sel.TypeText(_T("c"));
	font.put_Subscript(false);//反激活下标
	sel.TypeText(_T(" = ") + s_fc + _T("N/mm²."));

	sel.TypeText(_T("f"));
	font.put_Subscript(true);
	sel.TypeText(_T("t"));
	font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_ft + _T("N/mm²."));

	content = _T("钢筋型号:") + m_ST + _T(",");
	sel.TypeText(content);

	sel.TypeText(_T("f"));
	font.put_Subscript(true);
	sel.TypeText(_T("y"));
	font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_fy + _T("N/mm²."));

	sel.TypeText(_T("α"));
	font.put_Subscript(true);
	sel.TypeText(_T("1"));
	font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_a1 + _T(","));

	sel.TypeText(_T("ξ"));
	font.put_Subscript(true);
	sel.TypeText(_T("b"));
	font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_kb + _T("."));

	sel.TypeText(_T("受拉钢筋合力点到受拉边缘距离:"));

	sel.TypeText(_T("a"));
	font.put_Subscript(true);
	sel.TypeText(_T("s"));
	font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_as + _T("mm."));

	sel.TypeParagraph();//已知条件中不需要使用域代码


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
	sel.TypeText(_T("h"));
	font.put_Subscript(true);
	sel.TypeText(_T("0"));
	font.put_Subscript(false);
	sel.TypeText(_T(" = h - a"));
	font.put_Subscript(true);
	sel.TypeText(_T("s"));
	font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_h + _T(" - ") + s_as + _T(" = ") + s_h0 + _T("（mm）"));
	sel.TypeParagraph();

	if(con == 0)
	{

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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ M = ") + s_M + _T("(kN∙m) ≥ α"));
			font.put_Subscript(true),sel.TypeText(_T("1")),font.put_Subscript(false);
			sel.TypeText(_T("f"));
			font.put_Subscript(true),sel.TypeText(_T("c")),font.put_Subscript(false);
			sel.TypeText(_T("bξ"));
			font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
			sel.TypeText(_T("h"));
			font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
			sel.TypeText(_T("(h"));
			font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
			sel.TypeText(_T(" - \\f(ξ"));
			font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
			sel.TypeText(_T("h"));
			font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
			sel.TypeText(_T(",2))"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ = ") + s_a1 + _T("×") + s_fc + _T("×") + s_b + _T("×") + s_kb + _T("×") + s_h0 + _T("×(") + s_h0 + _T(" - ") + COperate::fraction((s_kb+_T("×")+s_h0),_T("2")) + _T(") = ") +  resultM + _T("(kN∙m)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//7
		sel.TypeText(_T("    故该梁为超筋梁."));
	}
	else
	{
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ α"));
			font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
			sel.TypeText(_T(" = \\f(M,α"));
			font.put_Subscript(true),sel.TypeText(_T("1")),font.put_Subscript(false);
			sel.TypeText(_T("f"));
			font.put_Subscript(true),sel.TypeText(_T("c")),font.put_Subscript(false);
			sel.TypeText(_T("bh"));
			font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
			font.put_Superscript(true),sel.TypeText(_T("2")),font.put_Superscript(false);
			sel.TypeText(_T(") = \\f(") + s_M + _T("×") + _T("10"));
			font.put_Superscript(true),sel.TypeText(_T("6")),font.put_Superscript(false);
			sel.TypeText(_T(",") + s_a1 +_T("×")+s_fc+_T("×")+s_b+_T("×")+s_h0);
			font.put_Superscript(true),sel.TypeText(_T("2")),font.put_Superscript(false);
			sel.TypeText(_T(") = ") + s_arfas);
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ x = ( 1 - \\r(1 - 2α"));
			font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
			sel.TypeText(_T(") )∙h"));
			font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
			sel.TypeText(_T(" = ( 1 - \\r(1 - 2×") + s_arfas + _T(") )×") + s_h0 + _T(" = ") + s_x + _T(" (mm)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ x"));
			font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
			sel.TypeText(_T(" = ξ"));
			font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
			sel.TypeText(_T("∙h"));
			font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
			sel.TypeText(_T(" = ") + s_kb + _T("×") + s_h0 + _T(" = ") + s_xb + _T(" (mm)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ x < x"));
			font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
			sel.TypeText(_T(",不超筋."));
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ A"));
			font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
			sel.TypeText(_T(" = \\f(M,f"));
			font.put_Subscript(true),sel.TypeText(_T("y")),font.put_Subscript(false);
			sel.TypeText(_T("(h"));
			font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
			sel.TypeText(_T(" - \\f(x,2))) = \\f(")+s_M+_T("×10"));
			font.put_Superscript(true),sel.TypeText(_T("6")),font.put_Superscript(false);
			sel.TypeText(_T(",")+s_fy + _T("×(") + s_h0 + _T(" - \\f(")+ s_x + _T(",2)))") + _T(" = ") + s_SAs + _T(" (mm²)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ A"));
			font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);
			sel.TypeText(_T(" = max{0.2%,0.45\\f(f"));
			font.put_Subscript(true),sel.TypeText(_T("t")),font.put_Subscript(false);
			sel.TypeText(_T(", f"));
			font.put_Subscript(true),sel.TypeText(_T("y")),font.put_Subscript(false);
			sel.TypeText(_T(")}∙b∙h = ")+s_pmin + _T("×") + s_b + _T("×") + s_h + _T(" = ") + s_Asmin+_T(" (mm²)"));	
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(con == 1)
		{
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
			}//用来解决开关问题

			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			{
				sel.TypeText(_T("EQ A"));
				font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);
				sel.TypeText(_T(" < A"));
				font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
				sel.TypeText(_T(",满足最小配筋率要求.故配筋面积为:A"));
				font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
				sel.TypeText(_T(" = ")+ s_SAs + _T(" (mm²)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		}
		else
		{
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
			}//用来解决开关问题

			sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			{
				sel.TypeText(_T("EQ A"));
				font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);
				sel.TypeText(_T(" > A"));
				font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
				sel.TypeText(_T(",不满足最小配筋率要求.故配筋面积为:A"));
				font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
				sel.TypeText(_T(" = A"));
				font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);
				sel.TypeText(_T(" = ")+ s_Asmin + _T(" (mm²)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		}

	}



	/*
	//保存文件
	//CDocument0 doc = app.get_ActiveDocument();//活动文档
	CString FileName(_T("E:\\doc.doc")); //文件名
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	COleVariant varZero((short)0); 
	COleVariant varTrue(short(1),VT_BOOL); 
	COleVariant varFalse(short(0),VT_BOOL); 
	doc.SaveAs(COleVariant(FileName), varZero, varFalse,
	COleVariant(_T("")), varTrue, COleVariant(_T("")),
	varFalse, varFalse, varFalse, varFalse, varFalse,
	covOptional,covOptional,covOptional,covOptional,
	covOptional);
	*/
	//释放对应的对象

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


void CSBTabD::OnBnClickedSdScal()
{
	// TODO: 在此添加控件通知处理程序代码_ttoi(m_sdia);
	CWnd::UpdateData(true);//将数据从对话框的控件中传送到对应的数据成员中
	m_sres = (double)1/4 * (_ttoi(m_sdia)*_ttoi(m_sdia)) * 3.14159 * _ttoi(m_snum);
	CWnd::UpdateData(false);//
}




void CSBTabD::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	OnBnClickedSdCalcula();
	//CDialogEx::OnOK();
}


void CSBTabD::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
