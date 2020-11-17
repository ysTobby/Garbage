// Cxc.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "Cxc.h"
#include "afxdialogex.h"


// Cxc 对话框

IMPLEMENT_DYNAMIC(Cxc, CDialogEx)

Cxc::Cxc(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cxc::IDD, pParent)
{

	m_czxs = _T("0.7");
	m_a = 0.0;
	m_as = 40;
	m_b = 200;
	m_CT = _T("C30");
	m_gjjj = 0.0;
	m_gjtp = _T("HRB335");
	m_gjzj = _T("8");
	m_gjzs = _T("2");
	m_h = 500;
	m_namda = 0.0;
	m_type = _T("1.仅配置箍筋");
	m_v = 0.0;
	m_wqd = _T("25");
	m_wqn = _T("1");
	m_zjtp = _T("HRB400");
}

Cxc::~Cxc()
{
}

void Cxc::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_CBString(pDX, IDC_CZXS, m_czxs);
	DDX_Text(pDX, IDC_a, m_a);
	DDX_Text(pDX, IDC_as, m_as);
	DDX_Text(pDX, IDC_b, m_b);
	DDX_CBString(pDX, IDC_CT, m_CT);
	DDX_Text(pDX, IDC_GJJJ, m_gjjj);
	DDX_CBString(pDX, IDC_GJTP, m_gjtp);
	DDX_CBString(pDX, IDC_GJZJ, m_gjzj);
	DDX_CBString(pDX, IDC_GJZS, m_gjzs);
	DDX_Text(pDX, IDC_h, m_h);
	DDX_Text(pDX, IDC_namda, m_namda);
	DDX_CBString(pDX, IDC_STEELTYPE, m_type);
	DDX_Text(pDX, IDC_V, m_v);
	DDX_CBString(pDX, IDC_WQD, m_wqd);
	DDX_CBString(pDX, IDC_WQN, m_wqn);
	DDX_CBString(pDX, IDC_ZJTP, m_zjtp);
}


BEGIN_MESSAGE_MAP(Cxc, CDialogEx)
	ON_BN_CLICKED(IDC_RADIO_jgb, &Cxc::OnBnClickedRadiojgb)
	ON_BN_CLICKED(IDC_RADIO_jg, &Cxc::OnBnClickedRadiojg)
	ON_CBN_SELCHANGE(IDC_CZXS, &Cxc::OnCbnSelchangeCzxs)
	ON_CBN_SELCHANGE(IDC_STEELTYPE, &Cxc::OnCbnSelchangeSteeltype)
	ON_BN_CLICKED(IDC_Reset, &Cxc::OnBnClickedReset)
	ON_BN_CLICKED(IDC_Calcula, &Cxc::OnBnClickedCalcula)
	ON_BN_CLICKED(IDC_Out, &Cxc::OnBnClickedOut)
END_MESSAGE_MAP()


// Cxc 消息处理程序


void Cxc::OnBnClickedRadiojgb()
{
	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_namda)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_SHOW);
}


void Cxc::OnBnClickedRadiojg()
{
	GetDlgItem(IDC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_a)->ShowWindow(SW_SHOW);
}


BOOL Cxc::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_jgxxsr)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jg)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jgb)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_ZJTP)->ShowWindow(SW_HIDE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void Cxc::OnCbnSelchangeCzxs()
{
	CWnd::UpdateData(true);
	if(m_czxs == _T("0.7"))
	{
		GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_a)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_jgxxsr)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_RADIO_jg)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_RADIO_jgb)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_namda)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_HIDE);
	}
	else
	{
		GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_a)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_jgxxsr)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_RADIO_jg)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_RADIO_jgb)->ShowWindow(SW_SHOW);
		
		((CButton *)GetDlgItem(IDC_RADIO_jg))->SetCheck(TRUE);//选上
		((CButton *)GetDlgItem(IDC_RADIO_jgb))->SetCheck(FALSE);
	}
}


void Cxc::OnCbnSelchangeSteeltype()
{
	CWnd::UpdateData(true);
	if(m_type == _T("1.仅配置箍筋"))
	{
		GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_WQD)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_WQN)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_ZJTP)->ShowWindow(SW_HIDE);
	}
	else
	{
		GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_WQD)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_WQN)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_ZJTP)->ShowWindow(SW_SHOW);
	}
}


void Cxc::OnOK()
{
	OnBnClickedCalcula();
	//CDialogEx::OnOK();
}


void Cxc::OnBnClickedReset()
{
	m_czxs = _T("0.7");
	m_a = 0.0;
	m_as = 40;
	m_b = 200;
	m_CT = _T("C30");
	m_gjjj = 0.0;
	m_gjtp = _T("HRB335");
	m_gjzj = _T("8");
	m_gjzs = _T("2");
	m_h = 500;
	m_namda = 0.0;
	m_type = _T("1.仅配置箍筋");
	m_v = 0.0;
	m_wqd = _T("25");
	m_wqn = _T("1");
	m_zjtp = _T("HRB400");
	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_jgxxsr)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jg)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jgb)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_ZJTP)->ShowWindow(SW_HIDE);
	CWnd::UpdateData(false);
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
}


void Cxc::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}


void Cxc::OnBnClickedCalcula()
{
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	CWnd::UpdateData(true);

	if(m_gjjj == 0.0)
	{
		OnBnClickedReset();
		MessageBox(_T("请输入箍筋间距"));
		return;
	}

	h0 = m_h - m_as;
	info.initV(m_CT,m_zjtp,m_gjtp);

	if((h0/m_b)<=4)
		k = 0.25;
	else if((h0/m_b)>=6)
		k = 0.2;
	else
		k = (-0.025*(h0/m_b)+0.35);

	m_v = k * info.bc * info.cp * m_b * h0;

	if(m_czxs == _T("0.7"))
	{
		con = 1;
		acv = 0.7;
	}
	else
	{
		con = 2;
		if(((CButton *)GetDlgItem(IDC_RADIO_jgb))->GetCheck())
		{
			if(m_namda > 3)
			{
				MessageBox(_T("输入剪跨比大于3，取为3计算"));
				m_namda = 3;
			}
			else if(m_namda < 1.5)
			{
				MessageBox(_T("输入剪跨比小于1.5，取为1.5计算"));
				m_namda = 1.5;
			}
			acv = 1.75/(m_namda + 1);
		}
		else
		{
			m_namda = (m_a/h0);
			if(m_namda > 3)
			{
				MessageBox(_T("输入剪跨比大于3，取为3计算"));
				m_namda = 3;
			}
			else if(m_namda < 1.5)
			{
				MessageBox(_T("输入剪跨比小于1.5，取为1.5计算"));
				m_namda = 1.5;
			}
			acv = 1.75/(m_namda + 1);
		}
	}

	if(m_type == _T("1.仅配置箍筋"))
	{
		com = 1;
	}
	else
	{
		com = 2;
	}

	temp = 0.25 * 3.14159 * _ttof(m_gjzj) * _ttof(m_gjzj) * _ttof(m_gjzs);//1.0是必要的要不当成int算了
	temp = temp/(m_b * m_gjjj);
	
	if(temp < (0.24*info.ct/info.gjss))
	{
		con1 = 1;
		MessageBox(_T("最小配箍率未满足要求,忽略箍筋的贡献"));
		if(m_type == _T("1.仅配置箍筋"))
		{
			if(acv*info.ct*m_b*h0 < m_v)
			{
				m_v = acv*info.ct*m_b*h0;
				con2 = 1;
			}
			else
			{
				con2 = 2;
				MessageBox(_T("计算结果 小于斜压破坏发生时的承载力，取发生斜压破坏的值作为极限承载力"));
			}
		}
		else
		{
			if((acv*info.ct*m_b*h0 + 0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8) < m_v)
			{
				con2 = 1;
				m_v = acv*info.ct*m_b*h0 + 0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8;
			}
			else
			{
				MessageBox(_T("计算结果 小于斜压破坏发生时的承载力，取发生斜压破坏的值作为极限承载力"));
				con2 = 2;
			}
		}
	}
	else
	{
		con1 = 2;
		if(m_type == _T("1.仅配置箍筋"))
		{
			if((acv*info.ct*m_b*h0 + info.gjss*h0*(1.0/4*3.14159*_ttof(m_gjzj)*_ttof(m_gjzj)*_ttof(m_gjzs))/(m_gjjj)) < m_v)
			{
				m_v = acv*info.ct*m_b*h0 + info.gjss*h0*(1.0/4*3.14159*_ttof(m_gjzj)*_ttof(m_gjzj)*_ttof(m_gjzs))/(m_gjjj);
				con2 = 1;
			}
			else
			{
				MessageBox(_T("计算结果 小于斜压破坏发生时的承载力，取发生斜压破坏的值作为极限承载力"));
				con2 = 2;
			}
		}
		else
		{
			if((acv*info.ct*m_b*h0 + info.gjss*h0*(1.0/4*3.14159*_ttof(m_gjzj)*_ttof(m_gjzj)*_ttof(m_gjzs))/(m_gjjj) + 0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8) < m_v)
			{
				m_v = acv*info.ct*m_b*h0 + info.gjss*h0*(1.0/4*3.14159*_ttof(m_gjzj)*_ttof(m_gjzj)*_ttof(m_gjzs))/(m_gjjj) + 0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8;
				con2 = 1;
			}
			else
			{
				con2 = 2;
				MessageBox(_T("计算结果 小于斜压破坏发生时的承载力，取发生斜压破坏的值作为极限承载力"));
			}
		}
	}
	m_v /= 1000;
	CWnd::UpdateData(false);
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
}

void Cxc::xb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Subscript(true),sel.TypeText(x),font.put_Subscript(false);
}
void Cxc::sb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Superscript(true),sel.TypeText(x),font.put_Superscript(false);
}
void Cxc::sxb(CString center,CString s,CString x)
{
	xb(center,x);
	sb(_T(""),s);
}

void Cxc::OnBnClickedOut()
{
	int i = 0;
	CString s_fy,s_fvy,s_acv,s_ft,s_fc,s_a1,s_kb,s_b,s_V,
		s_h,s_as,s_h0,content,szBuf,s_p,s_gjjj,s_k;

	//参数绑定
	s_h0.Format(_T("%.0lf"),h0);
	s_fc.Format(_T("%.1lf"),info.cp);
	s_a1.Format(_T("%.1lf"),info.a1);
	s_kb.Format(_T("%.3lf"),info.kb);
	s_fvy.Format(_T("%.0lf"),info.gjss);
	s_fy.Format(_T("%.0lf"),info.ss);
	s_ft.Format(_T("%.2lf"),info.ct); 
	s_acv.Format(_T("%.2lf"),acv);
	s_as.Format(_T("%.0lf"),m_as);
	s_b.Format(_T("%.0lf"),m_b);
	s_h.Format(_T("%.0lf"),m_h);
	s_V.Format(_T("%.1lf"),m_v);
	s_gjjj.Format(_T("%.0lf"),m_gjjj);
	s_k.Format(_T("%.2lf"),k);
	//绑定结束

	CoInitialize(NULL);
	if(!app.CreateDispatch(_T("word.application")))
	{
		AfxMessageBox(_T("居然你连OFFICE都没有安装吗?"));
		return;
	}
	app.put_Visible(TRUE); 
	docs = app.get_Documents();
	docs.Add(new CComVariant(_T("")),new CComVariant(FALSE),new CComVariant(0),new CComVariant());
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
	sel.TypeText(_T("斜截面受剪设计"));
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(1);
	paragraph.put_Alignment(1);
	sel.TypeParagraph();

	font =sel.get_Font();
	font.put_Size(10);
	//2
	content = _T("    已知条件：截面宽度:b=") 
		+ s_b + _T("mm,截面高度:h=") + s_h
		+ _T("mm,");
	sel.TypeText(content);
	paragraphs = doc.get_Paragraphs();
	paragraph = paragraphs.Item(2);
	paragraph.put_Alignment(0);

	sel.TypeText(_T("受拉钢筋到混凝土边缘距离:"));

	xb(_T("a"),_T("s"));
	sel.TypeText(_T("=")+s_as + _T("mm,"));

	sel.TypeText(_T("混凝土强度为:") + m_CT);

	xb(_T(",f"),_T("c"));
	sel.TypeText(_T("=")+s_fc+ _T("N/mm²,"));

	xb(_T("f"),_T("t"));
	sel.TypeText(_T("=")+s_ft+ _T("N/mm²."));

	sel.TypeText(_T("箍筋型号是:") + m_gjtp);

	xb(_T(",f"),_T("yv"));
	sel.TypeText(_T("=")+s_fvy+ _T("N/mm²."));

	sel.TypeText(_T("箍筋肢数:")+m_gjzs+_T(",箍筋直径:")+m_gjzj+_T("mm,箍筋间距:") + s_gjjj+_T("mm."));

	sel.TypeText(_T("受荷形式:"));
	if(con == 1)
	{
		sel.TypeText(_T("均布荷载,"));
		xb(_T("α"),_T("cv"));
		sel.TypeText(_T("=")+s_acv+ _T("."));
	}
	else
	{
		sel.TypeText(_T("集中荷载,"));
		xb(_T("α"),_T("cv"));
		sel.TypeText(_T("=")+s_acv+ _T("."));
	}

	sel.TypeText(_T("配筋方式:"));
	if(com == 1)
		sel.TypeText(_T("仅配置箍筋."));
	else
		sel.TypeText(_T("兼配弯起钢筋."));
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
	CString s_vmax;
	s_vmax.Format(_T("%.1lf"),k * info.bc * info.cp * m_b * h0/1000);
	xb(_T("V"),_T("max"));
	xb(_T(" = ")+s_k+_T("f"),_T("c"));
	xb(_T("bh"),_T("0"));
	sel.TypeText(_T(" = ")+s_k+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+s_vmax+_T(" (kN)"));
	sel.TypeParagraph();
	//6
	CString s_vc;
	s_vc.Format(_T("%.1lf"),acv*info.ct*m_b*h0/1000);
	xb(_T("V"),_T("c"));
	xb(_T(" = α"),_T("cv"));
	xb(_T("f"),_T("t"));
	xb(_T("bh"),_T("0"));
	xb(_T(" = ")+s_acv+_T("×")+s_ft+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+s_vc+_T(" (kN)"),_T(""));
	sel.TypeParagraph();
	//(7)
	CString s_Asb,s_Vsb;
	s_Asb.Format(_T("%.0lf"),_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)/4);
	s_Vsb.Format(_T("%.1lf"),(0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8)/1000);
	if(com == 2)
	{
		xb(_T("V"),_T("sb"));
		xb(_T(" = 0.8f"),_T("y"));
		xb(_T("A"),_T("s"));
		xb(_T("sin 45º = "),_T(""));
		sel.TypeText(_T("0.8×"));
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
			xb(_T("EQ \\f(\\r(2),2)"),_T(""));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));

		sel.TypeText(_T("×")+s_fy+_T("×")+s_Asb+_T(" = ")+s_Vsb+_T(" (kN)"));

		sel.TypeParagraph();
	}
	//7
	CString s_Asv;
	s_Asv.Format(_T("%.1lf"),0.25 * 3.14159 * _ttof(m_gjzj) * _ttof(m_gjzj) * _ttof(m_gjzs));
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
		xb(_T("EQ A"),_T("sv"));
		xb(_T(" = \\f(1,4)πd²n = "),_T(""));
		sel.TypeText(_T(""));
		sel.TypeText(_T("\\f(1,4)×π×")+m_gjzj+_T("²")+_T("×")+m_gjzs+_T(" = ")+s_Asv+_T(" (mm²)"));
	}
	fields.ToggleShowCodes();
	sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.TypeParagraph();

	//7
	CString p;
	p.Format(_T("%.4lf"),0.25 * 3.14159 * _ttof(m_gjzj) * _ttof(m_gjzj) * _ttof(m_gjzs)/(m_b * m_gjjj));
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
		xb(_T("EQ \\f(A"),_T("sv"));
		xb(_T(",bs) = "),_T(""));
		sel.TypeText(COperate::fraction(s_Asv,s_b+_T("×")+s_gjjj)+_T(" = ")+p);
	}
	fields.ToggleShowCodes();
	sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
	sel.TypeParagraph();
	CString pmin;
	pmin.Format(_T("%.4lf"),0.24*info.ct/info.gjss);
	if(con1 == 1)
	{
		//最小配筋率未满足，忽略箍筋作用
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
			xb(_T("EQ \\f(A"),_T("sv"));
			xb(_T(",bs) = ")+p+_T(" < 0.24\\f(f"),_T("t"));
			xb(_T(",f"),_T("yv"));
			sel.TypeText(_T(") = 0.24×")+COperate::fraction(s_ft,s_fvy)+_T(" = ")+pmin+_T(",不满足最小配箍率要求,忽略箍筋的作用。"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(com == 1)
		{
			//9
			xb(_T("V = V"),_T("c"));
			sel.TypeText(_T(" = ")+s_vc+_T(" (kN)"));
			sel.TypeParagraph();
			//10
			if(con2 == 1)
			{
				xb(_T("V = ")+s_vc+_T(" (kN) < V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_vc+_T(" (kN)"));
			}
			else
			{
				xb(_T("V = ")+s_vc+_T(" (kN) > V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_vmax+_T(" (kN)"));
			}
		}
		else
		{
			//9
			CString s_Vcs;
			s_Vcs.Format(_T("%.1lf"),(acv*info.ct*m_b*h0+0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8)/1000);
			xb(_T("V = V"),_T("c"));
			xb(_T(" + V"),_T("sb"));
			sel.TypeText(_T(" = ")+s_vc+_T(" + ")+s_Vsb+_T(" = ")+s_Vcs+_T(" (kN)"));
			sel.TypeParagraph();
			//10
			if(con2 == 1)
			{
				xb(_T("V = ")+s_Vcs+_T(" (kN) < V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_Vcs+_T(" (kN)"));
			}
			else
			{
				xb(_T("V = ")+s_Vcs+_T(" (kN) > V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_vmax+_T(" (kN)"));
			}
		}
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
			xb(_T("EQ \\f(A"),_T("sv"));
			xb(_T(",bs) = ")+p+_T(" > 0.24\\f(f"),_T("t"));
			xb(_T(",f"),_T("yv"));
			sel.TypeText(_T(") = 0.24×")+COperate::fraction(s_ft,s_fvy)+_T(" = ")+pmin+_T(",满足最小配箍率要求，考虑箍筋对剪力贡献。"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		//9
		CString s_Vg;
		s_Vg.Format(_T("%.1lf"),(info.gjss*h0*(1.0/4*3.14159*_ttof(m_gjzj)*_ttof(m_gjzj)*_ttof(m_gjzs))/(m_gjjj))/1000);
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
			xb(_T("EQ V"),_T("sv"));
			xb(_T(" = f"),_T("yv"));
			xb(_T("h"),_T("0"));
			xb(_T("\\f(A"),_T("sv"));
			xb(_T(",s) = "),_T(""));
			sel.TypeText(s_fvy+_T("×")+s_h0+_T("×")+COperate::fraction(s_Asv,s_gjjj));
			sel.TypeText(_T(" = ")+s_Vg+_T(" (kN)"));
		}
		fields.ToggleShowCodes();
		sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		sel.TypeParagraph();
		if(com == 1)
		{
			//10
			CString s_Vcg;
			s_Vcg.Format(_T("%.1lf"),(acv*info.ct*m_b*h0 + info.gjss*h0*(1.0/4*3.14159*_ttof(m_gjzj)*_ttof(m_gjzj)*_ttof(m_gjzs))/(m_gjjj))/1000);
			xb(_T("V = V"),_T("c"));
			xb(_T(" + V"),_T("sv"));
			sel.TypeText(_T(" = ")+s_vc+_T(" + ")+s_Vg+_T(" = ")+s_Vcg+_T(" (kN)"));
			sel.TypeParagraph();
			//11
			if(con2 == 1)
			{
				xb(_T("V = ")+s_Vcg+_T(" (kN) < V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_Vcg+_T(" (kN)"));
			}
			else
			{
				xb(_T("V = ")+s_Vcg+_T(" (kN) > V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_vmax+_T(" (kN)"));
			}
		}
		else
		{
			//10
			CString s_Vcsg;
			s_Vcsg.Format(_T("%.1lf"),(acv*info.ct*m_b*h0 + info.gjss*h0*(1.0/4*3.14159*_ttof(m_gjzj)*_ttof(m_gjzj)*_ttof(m_gjzs))/(m_gjjj) + 0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8)/1000);
			xb(_T("V = V"),_T("c"));
			xb(_T(" + V"),_T("sb"));
			xb(_T(" + V"),_T("sv"));
			sel.TypeText(_T(" = ")+s_vc+_T(" + ")+s_Vsb+_T(" + ")+s_Vg+_T(" = ")+s_Vcsg+_T(" (kN)"));
			sel.TypeParagraph();
			//11
			if(con2 == 1)
			{
				xb(_T("V = ")+s_Vcsg+_T(" (kN) < V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_Vcsg+_T(" (kN)"));
			}
			else
			{
				xb(_T("V = ")+s_Vcsg+_T(" (kN) > V"),_T("max"));
				sel.TypeText(_T(" = ")+s_vmax+_T(" (kN),取V = ")+s_vmax+_T(" (kN)"));
			}
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
