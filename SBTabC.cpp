// SBTabC.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "SBTabC.h"
#include "afxdialogex.h"


// CSBTabC 对话框

IMPLEMENT_DYNAMIC(CSBTabC, CDialogEx)

	CSBTabC::CSBTabC(CWnd* pParent /*=NULL*/)
	: CDialogEx(CSBTabC::IDD, pParent)
{

	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_MAs = 0.0;
	m_Mu = 0.0;
	m_SAs = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
}

CSBTabC::~CSBTabC()
{
}

void CSBTabC::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_SD_As, m_As);
	DDX_Text(pDX, IDC_SD_b, m_b);
	DDX_CBString(pDX, IDC_SD_CS, m_CS);
	DDX_Text(pDX, IDC_SD_h, m_h);
	DDX_Text(pDX, IDC_SD_MAs, m_MAs);
	DDX_Text(pDX, IDC_SD_Mu, m_Mu);
	DDX_Text(pDX, IDC_SD_SAs, m_SAs);
	DDX_CBString(pDX, IDC_SD_ST, m_ST);
	DDX_Text(pDX, IDC_SD_x, m_x);
	DDX_Text(pDX, IDC_SD_xb, m_xb);
}


BEGIN_MESSAGE_MAP(CSBTabC, CDialogEx)
	ON_BN_CLICKED(IDC_SD_Reset, &CSBTabC::OnBnClickedSdReset)
	ON_BN_CLICKED(IDC_SD_Calcula, &CSBTabC::OnBnClickedSdCalcula)
	ON_BN_CLICKED(IDC_SC_Out, &CSBTabC::OnBnClickedScOut)
	//	ON_WM_KEYDOWN()
END_MESSAGE_MAP()


// CSBTabC 消息处理程序


void CSBTabC::OnBnClickedSdReset()
{
	// TODO: 在此添加控件通知处理程序代码
	m_As = 40;
	m_b = 200;
	m_CS = _T("C30");
	m_h = 500;
	m_MAs = 0.0;
	m_Mu = 0.0;
	m_SAs = 0.0;
	m_ST = _T("HRB400");
	m_x = 0.0;
	m_xb = 0.0;
	CWnd::UpdateData(false);
	GetDlgItem(IDC_SC_Out)->EnableWindow(FALSE);
}


void CSBTabC::OnBnClickedSdCalcula()
{
	// TODO: 在此添加控件通知处理程序代码
	CWnd::UpdateData(true);

	if(m_b==0.0||m_h==0.0||m_As==0.0||m_CS==""||m_ST==""||m_SAs == 0.0)
	{
		MessageBox(_T("请输入必要参数"));
		return ;
	}
	GetDlgItem(IDC_SC_Out)->EnableWindow(TRUE);
	info.init(m_CS,m_ST);

	h0 = m_h - m_As;
	m_x = (info.ss * m_SAs) /(info.a1 * info.cp * m_b);
	rm_x = m_x;
	p = (m_SAs) / (m_b * m_h);
	pmin = 0.002;
	if(pmin <= 0.45 * (info.ct) / (info.ss))
		pmin = 0.45 * (info.ct) / (info.ss);

	if((m_x <= (info.kb * h0)) && (p >= pmin))
	{
		m_Mu = info.ss * m_SAs * (h0 - m_x / 2);
		control = 1;
	}
	else
	{
		if(m_x > (info.kb * h0))
		{
			MessageBox(_T("该梁为超筋梁，取受压区高度为界限受压区高度进行计算"));
			m_x = (info.kb * h0);
			m_Mu = info.a1 * info.cp * m_b * m_x * (h0 - m_x / 2);
			control = 2;
		}
		else
		{
			MessageBox(_T("该梁为少筋梁，需按照素混凝土模型进行计算，最小配筋面积见输出结果"));	
			control = 3;
		}
	}
	m_Mu = m_Mu / 1000000;

	m_MAs = pmin * m_b * m_h;
	m_xb = info.kb * h0;
	CWnd::UpdateData(false);
}


void CSBTabC::OnBnClickedScOut()
{
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
	CDocument0 doc = app.get_ActiveDocument();
	CParagraphs paragraphs;
	CParagraph paragraph;
	CFields fields;
	CRange range;
	COleVariant vTrue((short)TRUE),vFalse((short)FALSE),vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	COleVariant formula;

	CSelection sel=app.get_Selection();
	CFont0 font =sel.get_Font();
	font.put_Name(_T("Times New Roman"));
	font.put_Size(18);
	font.put_Color(WdColor::wdColorBlack);
	font.put_Bold(0);

	CString s_fy,s_pmin,s_xb,s_ft,s_Mu,s_fc,s_a1,s_kb,s_x,s_b,s_Asmin,s_h,s_as,s_h0,s_SAs,resultM,content,szBuf;
	int i = 0;
	//参数绑定
	s_as.Format(_T("%.0lf"),m_As);
	s_h.Format(_T("%.0lf"),m_h);
	s_h0.Format(_T("%.0lf"),h0);
	//s_M.Format(_T("%.0lf"),m_M);
	s_a1.Format(_T("%.1lf"),info.a1);
	s_fc.Format(_T("%.1lf"),info.cp);
	s_b.Format(_T("%.0lf"),m_b);
	s_kb.Format(_T("%.3lf"),info.kb);
	s_fy.Format(_T("%.0lf"),info.ss);
	//s_arfas.Format(_T("%.3lf"),arfas);
	s_x.Format(_T("%.3lf"),rm_x);
	s_Asmin.Format(_T("%.0lf"),m_MAs); 
	s_SAs.Format(_T("%.0lf"),m_SAs); 
	s_ft.Format(_T("%.2lf"),info.ct); 
	s_xb.Format(_T("%.0lf"),(info.kb * h0)); 
	s_pmin.Format(_T("%.4lf"),pmin); 
	s_Mu.Format(_T("%.1lf"),m_Mu); 
	resultM.Format(_T("%.1lf"),(info.a1 * info.cp * m_b * info.kb * h0 * (h0 - info.kb * h0 / 2))/1000000);
	//绑定结束

	//1
	sel.TypeText(_T("单筋矩形截面梁受弯复核"));//调用函数Selection::TypeText 向WORD发送字符
	paragraphs = doc.get_Paragraphs();//获取当前文档的全部段落
	paragraph = paragraphs.Item(1);//选中第三段为
	paragraph.put_Alignment(1);//设置文本格式：0左对齐，1居中对齐，2右对齐
	sel.TypeParagraph();

	font =sel.get_Font();
	font.put_Size(10);

	//2
	content = _T("    已知条件：截面宽度:b=") + s_b + _T("mm,截面高度:h=") + s_h
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
	font.put_Subscript(true),sel.TypeText(_T("t")),font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_fc + _T("N/mm²."));

	content = _T("钢筋型号:") + m_ST + _T(":");
	sel.TypeText(content);

	sel.TypeText(_T("f"));
	font.put_Subscript(true),sel.TypeText(_T("y")),font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_fy + _T("N/mm²."));

	sel.TypeText(_T("α"));
	font.put_Subscript(true),sel.TypeText(_T("1")),font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_a1 + _T(","));

	sel.TypeText(_T("ξ"));
	font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_kb + _T("."));

	content = _T("受拉钢筋合力点到受拉边缘距离:");
	sel.TypeText(content);
	sel.TypeText(_T("a"));
	font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_as + _T("mm."));

	content = _T("受拉钢筋截面积:");
	sel.TypeText(content);
	sel.TypeText(_T("A"));
	font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
	sel.TypeText(_T(" = ") + s_SAs + _T("mm²."));
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


	if(control == 1||control == 2)
	{
		//正常
		//4
		sel.TypeText(_T("h"));
		font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
		sel.TypeText(_T(" = h - a"));
		font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
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
		}//用来解决开关问题

		sel.MoveLeft(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		{
			sel.TypeText(_T("EQ x = \\f(f"));
			font.put_Subscript(true),sel.TypeText(_T("y")),font.put_Subscript(false);
			sel.TypeText(_T("A"));
			font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
			sel.TypeText(_T(",a"));
			font.put_Subscript(true),sel.TypeText(_T("1")),font.put_Subscript(false);
			sel.TypeText(_T("f"));
			font.put_Subscript(true),sel.TypeText(_T("c")),font.put_Subscript(false);
			sel.TypeText(_T("b) = ") + COperate::fraction(s_fy+_T("×")+s_SAs,s_a1+_T("×")+s_fc+_T("×")+s_b)+_T(" = ")+s_x+_T("(mm)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();

		//6
		sel.TypeText(_T("x"));
		font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
		sel.TypeText(_T(" = ξ"));
		font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
		sel.TypeText(_T(" = h"));
		font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
		sel.TypeText(_T(" = ")+s_kb+_T("×")+s_h0+_T(" = ")+s_xb+_T("(mm)"));
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
			sel.TypeText(_T("EQ A"));
			font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);
			sel.TypeText(_T(" = max{0.002,0.45\\f(f"));
			font.put_Subscript(true),sel.TypeText(_T("t")),font.put_Subscript(false);
			sel.TypeText(_T(",f"));
			font.put_Subscript(true),sel.TypeText(_T("y")),font.put_Subscript(false);
			sel.TypeText(_T(")}∙bh = ")+s_pmin+_T("×")+s_b+_T("×")+s_h+ _T(" = ")+s_Asmin+_T(" (mm²)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(control == 1)
		{
			//8
			sel.TypeText(_T("    因:A"));
			font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
			sel.TypeText(_T(" > A"));
			font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);
			sel.TypeText(_T(",且:"));
			sel.TypeText(_T("x"));
			font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
			sel.TypeText(_T(" > x"));
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
				sel.TypeText(_T("EQ M"));
				font.put_Subscript(true),sel.TypeText(_T("u")),font.put_Subscript(false);
				sel.TypeText(_T(" = α"));
				font.put_Subscript(true),sel.TypeText(_T("1")),font.put_Subscript(false);
				sel.TypeText(_T("f"));
				font.put_Subscript(true),sel.TypeText(_T("c")),font.put_Subscript(false);
				sel.TypeText(_T("bx( h"));
				font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
				sel.TypeText(_T(" - \\f(x,2) ) = ") + s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_x+_T("×")_T("(")+s_h0+_T(" - ")+COperate::fraction(s_x,_T("2"))+_T(")") + _T(" = ")+s_Mu+_T(" (kN∙m)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		}
		else
		{
			//8
			sel.TypeText(_T("    因:A"));
			font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
			sel.TypeText(_T(" > A"));
			font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);
			sel.TypeText(_T("且:x"));
			font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
			sel.TypeText(_T(" < x"));

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
				sel.TypeText(_T("EQ M"));
				font.put_Subscript(true),sel.TypeText(_T("u")),font.put_Subscript(false);
				sel.TypeText(_T(" = a"));
				font.put_Subscript(true),sel.TypeText(_T("1")),font.put_Subscript(false);
				sel.TypeText(_T("f"));
				font.put_Subscript(true),sel.TypeText(_T("c")),font.put_Subscript(false);
				sel.TypeText(_T("bx"));
				font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
				sel.TypeText(_T("( h"));
				font.put_Subscript(true),sel.TypeText(_T("0")),font.put_Subscript(false);
				sel.TypeText(_T(" - \\f(x"));
				font.put_Subscript(true),sel.TypeText(_T("b")),font.put_Subscript(false);
				sel.TypeText(_T(",2) ) = ")+s_a1+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_xb+_T("×")_T("( ")+s_h0+_T(" - ")+COperate::fraction(s_xb,_T("2"))+_T(" )")+_T(" = ")+s_Mu+_T(" (kN∙m)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		}
	}
	else
	{
		//少筋
		//4
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
			sel.TypeText(_T(" = max{0.002,0.45\\f(f"));
			font.put_Subscript(true),sel.TypeText(_T("t")),font.put_Subscript(false);
			sel.TypeText(_T(",f"));
			font.put_Subscript(true),sel.TypeText(_T("y")),font.put_Subscript(false);
			sel.TypeText(_T(")}∙bh = ")+s_pmin+_T("×")+s_b+_T("×")+s_h+ _T(" = ")+s_Asmin+_T(" (mm²)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();

		//5
		sel.TypeText(_T("    因:A"));
		font.put_Subscript(true),sel.TypeText(_T("s")),font.put_Subscript(false);
		sel.TypeText(_T(" < A"));
		font.put_Subscript(true),sel.TypeText(_T("smin")),font.put_Subscript(false);

		sel.TypeText(_T(",应按照素混凝土计算该受弯构件承载力"));
	}


	font.ReleaseDispatch();//注意：所有东西用完之后一定要ReleaseDispatch，否则报错；不过最好像例子2中，在最后集中ReleaseDispatch
	paragraph.ReleaseDispatch();
	paragraphs.ReleaseDispatch();
	range.ReleaseDispatch();
	fields.ReleaseDispatch();
	doc.ReleaseDispatch();
	sel.ReleaseDispatch();
	docs.ReleaseDispatch(); 
	app.ReleaseDispatch();   //释放对象指针。切记，必须调用
	CoUninitialize();//对应CoInitialize
}



void CSBTabC::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类
	OnBnClickedSdCalcula();
	//CDialogEx::OnOK();
}



void CSBTabC::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
