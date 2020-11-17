// Cxd.cpp : 实现文件
//

#include "stdafx.h"
#include "Peijin.h"
#include "Cxd.h"
#include "afxdialogex.h"


// Cxd 对话框

IMPLEMENT_DYNAMIC(Cxd, CDialogEx)

	Cxd::Cxd(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cxd::IDD, pParent)
{

	m_type = _T("1.仅配置箍筋");
	m_czxs = _T("0.7");
	m_a = 0.0;
	m_as = 40;
	m_b = 200;
	m_CT = _T("C30");
	m_gjtp = _T("HRB335");
	m_gjzj = _T("8");
	m_gjzs = _T("2");
	m_h = 500;
	m_result = 0.0;
	m_v = 0.0;
	m_wqd = _T("25");
	m_wqn = _T("1");
	m_zjtp = _T("HRB400");
	m_gjjj = 0.0;
	m_namda = 0.0;
}

Cxd::~Cxd()
{
}

void Cxd::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_CBString(pDX, IDC_STEELTYPE, m_type);
	DDX_CBString(pDX, IDC_CZXS, m_czxs);
	DDX_Text(pDX, IDC_a, m_a);
	DDX_Text(pDX, IDC_as, m_as);
	DDX_Text(pDX, IDC_b, m_b);
	DDX_CBString(pDX, IDC_CT, m_CT);
	DDX_CBString(pDX, IDC_GJTP, m_gjtp);
	DDX_CBString(pDX, IDC_GJZJ, m_gjzj);
	DDX_CBString(pDX, IDC_GJZS, m_gjzs);
	DDX_Text(pDX, IDC_h, m_h);
	DDX_Text(pDX, IDC_RESULT, m_result);
	DDX_Text(pDX, IDC_V, m_v);
	DDX_CBString(pDX, IDC_WQD, m_wqd);
	DDX_CBString(pDX, IDC_WQN, m_wqn);
	DDX_CBString(pDX, IDC_ZJTP, m_zjtp);
	DDX_Text(pDX, IDC_GJJJ, m_gjjj);
	DDX_Text(pDX, IDC_namda, m_namda);
}


BEGIN_MESSAGE_MAP(Cxd, CDialogEx)
	ON_CBN_SELCHANGE(IDC_STEELTYPE, &Cxd::OnCbnSelchangeCombo1)
	ON_CBN_SELCHANGE(IDC_CZXS, &Cxd::OnCbnSelchangeCzxs)
	ON_BN_CLICKED(IDC_Calcula, &Cxd::OnBnClickedCalcula)
	ON_BN_CLICKED(IDC_Reset, &Cxd::OnBnClickedReset)
	ON_BN_CLICKED(IDC_Out, &Cxd::OnBnClickedOut)
	ON_BN_CLICKED(IDC_Calculalength, &Cxd::OnBnClickedCalculalength)
	ON_BN_CLICKED(IDC_RADIO_jg, &Cxd::OnBnClickedRadiojg)
	ON_BN_CLICKED(IDC_RADIO_jgb, &Cxd::OnBnClickedRadiojgb)
END_MESSAGE_MAP()


// Cxd 消息处理程序


void Cxd::OnCbnSelchangeCombo1()
{
	CWnd::UpdateData(true);
	if(m_type == _T("1.仅配置箍筋"))
	{
		GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_WQD)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_WQN)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_ZJTP)->ShowWindow(SW_HIDE);
	}
	else
	{
		GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_WQD)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_WQN)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_ZJTP)->ShowWindow(SW_SHOW);
	}
}


BOOL Cxd::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_jgxxsr)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jg)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jgb)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_ZJTP)->ShowWindow(SW_HIDE);


	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}


void Cxd::OnCbnSelchangeCzxs()
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
		((CButton *)GetDlgItem(IDC_RADIO_jgb))->SetCheck(FALSE);//选上
	}
}


void Cxd::OnOK()
{
	OnBnClickedCalcula();
	//CDialogEx::OnOK();
}


void Cxd::OnBnClickedCalcula()
{
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	CString strMsg;
	CWnd::UpdateData(true);
	if(m_v == 0||m_b == 0||m_h == 0||m_as == 0)
	{
		MessageBox(_T("请输入必要参数进行计算"));
		return;
	}
	info.initV(m_CT,m_zjtp,m_gjtp);

	temp = 0;
	h0 = m_h - m_as;
	m_v *= 1000;

	if(m_czxs == _T("0.7"))
	{
		con = 1;
		temp = 0.7 * info.ct * m_b * h0;
		acv = 0.7;
	}
	else
	{
		con = 2;
		if(((CButton *)GetDlgItem(IDC_RADIO_jgb))->GetCheck())//判断是否被选中
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
			temp = 1.75/(m_namda + 1) * info.ct * m_b * h0;
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
			temp = 1.75/(m_namda + 1) * info.ct * m_b * h0;
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

	if((h0/m_b)<=4)
	{
		con1 = 1,con2 = 1;
		if(m_v > 0.25 * info.bc * info.cp * m_b * h0)
		{
			con2 = 2;
			m_v /= 1000;
			MessageBox(_T("截面尺寸不满足要求，需增大截面高度"));
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
			return;
		}
	}
	else if((h0/m_b)>=6)
	{
		con1 = 2,con2 = 1;
		if(m_v > 0.2 * info.bc * info.cp * m_b * h0)
		{
			con2 = 2;
			m_v /= 1000;
			MessageBox(_T("截面尺寸不满足要求，需增大截面高度"));
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
			return;
		}
	}
	else
	{
		con1 = 3,con2 = 1;
		temp = -0.025*(h0/m_b)+0.35;
		temp1 = temp;
		if(m_v > temp * info.bc * info.cp * m_b * h0)
		{
			m_v /= 1000;
			con2 = 2;
			MessageBox(_T("截面尺寸不满足要求，需增大截面高度"));
			GetDlgItem(IDC_Out)->EnableWindow(TRUE);
			return;
		}
	}
	MessageBox(_T("截面尺寸验算符合要求"));

	if(m_type == _T("1.仅配置箍筋"))
	{
		con3 = 1;
		if(m_v < temp)
		{
			con3 = 2;
			strMsg.Format(_T("V≤Vc=%.1lf,按照构造配置箍筋"),temp/1000);
			MessageBox(strMsg);
			m_v /= 1000;
			GetDlgItem(IDC_Calculalength)->EnableWindow(TRUE);
			m_result = 0.24*info.ct*m_b/info.gjss;
			CWnd::UpdateData(false);
			m_realresult = m_result;
			return;
		}
		else
		{
			strMsg.Format(_T("V>Vc=%.1lf,需计算配置箍筋"),temp/1000);
			MessageBox(strMsg);
			m_result = (m_v - temp)/(info.gjss * h0);
		}
	}
	else
	{
		con3 = 1;
		temp += 0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8;
		if(m_v < temp)
		{
			con3 = 2;
			strMsg.Format(_T("V≤Vc+Vsb=%.1lf,按照构造配置箍筋"),temp/1000);
			MessageBox(strMsg);
			m_v /= 1000;
			GetDlgItem(IDC_Calculalength)->EnableWindow(TRUE);
			m_result = 0.24*info.ct*m_b/info.gjss;
			CWnd::UpdateData(false);
			m_realresult = m_result;
			return;
		}
		else
		{
			strMsg.Format(_T("V>Vc+Vsb=%.1lf,需计算配置箍筋"),temp/1000);
			MessageBox(strMsg);
			m_result = (m_v - temp)/(info.gjss * h0);
		}
	}
	m_realresult = m_result;
	if((m_result > 0)&&(m_result/m_b<(0.24*info.ct/info.gjss)))
	{
		MessageBox(_T("箍筋最小的配筋率未满足要求，取最小配筋率进行计算"));
		m_result = 0.24*info.ct*m_b/info.gjss;
	}
	m_v /= 1000;
	CWnd::UpdateData(false);
	GetDlgItem(IDC_Calculalength)->EnableWindow(TRUE);
}


void Cxd::OnBnClickedReset()
{
	m_type = _T("1.仅配置箍筋");
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
	m_result = 0.0;
	m_v = 0.0;
	m_wqd = _T("");
	m_wqn = _T("");
	m_zjtp = _T("HRB400");
	m_namda = 0.0;
	GetDlgItem(IDC_STATIC_WQI)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQD)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_WQN)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_jgxxsr)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jg)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_RADIO_jgb)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_wq)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_ZJTP)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_Out)->EnableWindow(FALSE);
	GetDlgItem(IDC_Calculalength)->EnableWindow(FALSE);
	CWnd::UpdateData(false);
}

void Cxd::xb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Subscript(true),sel.TypeText(x),font.put_Subscript(false);
}
void Cxd::sb(CString center,CString x)
{
	sel.TypeText(center);
	font.put_Superscript(true),sel.TypeText(x),font.put_Superscript(false);
}
void Cxd::sxb(CString center,CString s,CString x)
{
	xb(center,x);
	sb(_T(""),s);
}

void Cxd::OnBnClickedOut()
{
	int i = 0;
	CString s_fy,s_fvy,s_acv,s_ft,s_fc,s_a1,s_kb,s_b,s_V,
		s_h,s_as,s_h0,content,szBuf,s_p,s_result;

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
	s_result.Format(_T("%.3lf"),m_realresult);
	//s_p.Format(_T("%.4lf"),p); 
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
	content = _T("    已知条件：剪力设计值:V=") + s_V + _T("kN,截面宽度:b=") 
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
	if(con2 == 2)
	{
		//5
		if(con1 == 1)
		{
			CString bc;
			bc.Format(_T("%.1lf"),0.25 * info.bc * info.cp * m_b * h0/1000);
			sel.TypeText(_T("V = ")+s_V+_T("(kN) > "));
			xb(_T("0.25f"),_T("c"));
			xb(_T("bh"),_T("0"));
			sel.TypeText(_T(" = 0.25×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+ bc + _T(" (kN)，截面选取过小，需重新设计截面。"));
		}
		else if(con1 == 2)
		{
			CString bc;
			bc.Format(_T("%.1lf"),0.2 * info.bc * info.cp * m_b * h0/1000);
			sel.TypeText(_T("V = ")+s_V+_T("(kN) > "));
			xb(_T("0.2f"),_T("c"));
			xb(_T("bh"),_T("0"));
			sel.TypeText(_T(" = 0.2×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+ bc + _T(" (kN)，截面选取过小，需重新设计截面。"));
		}
		else
		{
			CString bc,ak;
			ak.Format(_T("%.2lf"),temp1);
			bc.Format(_T("%.1lf"),temp1 * info.bc * info.cp * m_b * h0/1000);
			sel.TypeText(_T("V = ")+s_V+_T("(kN) > "));
			xb(ak+_T("f"),_T("c"));
			xb(_T("bh"),_T("0"));
			sel.TypeText(_T(" = ")+ak+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+ bc + _T(" (kN)，截面选取过小，需重新设计截面。"));
		}
	}
	else
	{
		//5
		if(con1 == 1)
		{
			CString bc;
			bc.Format(_T("%.1lf"),0.25 * info.bc * info.cp * m_b * h0/1000);
			sel.TypeText(_T("V = ")+s_V+_T("(kN) < "));
			xb(_T("0.25f"),_T("c"));
			xb(_T("bh"),_T("0"));
			sel.TypeText(_T(" = 0.25×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+ bc + _T(" (kN)，满足截面的最小尺寸要求。"));
		}
		else if(con1 == 2)
		{
			CString bc;
			bc.Format(_T("%.1lf"),0.2 * info.bc * info.cp * m_b * h0/1000);
			sel.TypeText(_T("V = ")+s_V+_T("(kN) < "));
			xb(_T("0.2f"),_T("c"));
			xb(_T("bh"),_T("0"));
			sel.TypeText(_T(" = 0.2×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+ bc + _T(" (kN)，满足截面的最小尺寸要求。"));
		}
		else
		{
			CString bc,ak;
			ak.Format(_T("%.2lf"),temp1);
			bc.Format(_T("%.1lf"),temp1 * info.bc * info.cp * m_b * h0/1000);
			sel.TypeText(_T("V = ")+s_V+_T("(kN) < "));
			xb(ak+_T("f"),_T("c"));
			xb(_T("bh"),_T("0"));
			sel.TypeText(_T(" = ")+ak+_T("×")+s_fc+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+ bc + _T(" (kN)，满足截面的最小尺寸要求。"));
		}
		sel.TypeParagraph();
		if(con3 == 1)
		{
			//按计算配筋
			if(com == 1)
			{
				//6仅配置箍筋
				CString fcv;
				fcv.Format(_T("%.1lf"),acv*info.ct*m_b*h0/1000);
				xb(_T("V"),_T("c"));
				xb(_T(" = α"),_T("cv"));
				xb(_T("f"),_T("t"));
				xb(_T("bh"),_T("0"));
				sel.TypeText(_T(" = ")+s_acv+_T("×")+s_ft+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+fcv+ _T(" (kN)"));
				sel.TypeParagraph();
				//7
				xb(_T("V = ")+s_V+_T("(kN) > V"),_T("c"));
				sel.TypeText(_T(" = ")+fcv+_T("(kN)，应通过计算配置箍筋"));
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
					xb(_T("EQ \\f(A"),_T("sv"));
					xb(_T(",s) = \\f(V - V"),_T("c"));
					xb(_T(",f"),_T("yv"));
					xb(_T("h"),_T("0"));
					xb(_T(") = "),_T(""));
					sel.TypeText(COperate::fraction(_T("(")+s_V +_T(" - ")+fcv+_T(")×10³")
						,s_fvy+_T("×")+s_h0)+_T(" = ")+s_result+_T (" (mm²/mm)"));
				}
				fields.ToggleShowCodes();
				sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
				sel.TypeParagraph();
			}
			else
			{
				//6
				CString fcv,s_Asb,s_Vc,s_Vsb;
				s_Vc.Format(_T("%.1lf"),(acv*info.ct*m_b*h0)/1000);
				s_Vsb.Format(_T("%.1lf"),(0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8)/1000);
				fcv.Format(_T("%.1lf"),(acv*info.ct*m_b*h0+0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8)/1000);
				s_Asb.Format(_T("%.0lf"),_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)/4);
				xb(_T("V"),_T("c"));
				xb(_T(" = α"),_T("cv"));
				xb(_T("f"),_T("t"));
				xb(_T("bh"),_T("0"));
				sel.TypeText(_T(" = ")+s_acv+_T("×")+s_ft+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+s_Vc+ _T(" (kN)"));
				sel.TypeParagraph();
				//7
				xb(_T("V"),_T("sb"));
				xb(_T(" = 0.8f"),_T("y"));
				xb(_T("A"),_T("s"));
				xb(_T("sin 45º"),_T(""));
				sel.TypeText(_T(" = 0.8×"));
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
				sel.TypeText(_T("×")+s_fy+_T("×")+s_Asb+_T(" = ")+s_Vsb+ _T(" (kN)"));
				sel.TypeParagraph();
				//8
				xb(_T("V = ")+s_V+_T("(kN) > V"),_T("c"));
				xb(_T(" + V"),_T("sb"));
				sel.TypeText(_T(" = ")+s_Vc+_T(" + ")+s_Vsb+_T(" = ")+fcv+_T("(kN)，应通过计算配置箍筋"));
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
					xb(_T("EQ \\f(A"),_T("sv"));
					xb(_T(",s) = \\f(V - V"),_T("c"));
					xb(_T(" - V"),_T("sb"));
					xb(_T(",f"),_T("yv"));
					xb(_T("h"),_T("0"));
					xb(_T(") = "),_T(""));
					sel.TypeText(COperate::fraction(_T("(")+s_V +_T(" - ")+s_Vc+_T(" - ")+s_Vsb+_T(")×10³"),s_fvy+_T("×")+s_h0)+_T(" = ")+s_result+_T (" (mm²/mm)"));
				}
				fields.ToggleShowCodes();
				sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
				sel.TypeParagraph();
			}
			//9
			CString s_gjjj;
			s_gjjj.Format(_T("%.0lf"),m_gjjj);
			sel.TypeText(_T("选取箍筋:肢数n=")+m_gjzs+_T(",直径d=")+m_gjzj+_T("(mm)"));
			sel.TypeParagraph();
			//10
			CString s_Asv1;
			s_Asv1.Format(_T("%.1lf"),(_ttof(m_gjzj)*_ttof(m_gjzj)*3.14/4));
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
				xb(_T("EQ A"),_T("sv1"));
				xb(_T(" = \\f(1,4)πd² = "),_T(""));
				sel.TypeText(_T("\\f(1,4)×π×")+m_gjzj+_T("² = ")+s_Asv1+_T(" (mm²)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
			//11
			CString s_calresult,s_smax;
			s_calresult.Format(_T("%.3lf"),(_ttof(m_gjzs)*_ttof(m_gjzj)*_ttof(m_gjzj)*3.14/(4*_ttof(s_gjjj))));
			s_smax.Format(_T("%.0lf"),(_ttof(m_gjzs)*_ttof(m_gjzj)*_ttof(m_gjzj)*3.14/(4*m_realresult)));
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
				xb(_T("EQ s ≤ \\f(nA"),_T("sv1"));
				sel.TypeText(_T(",")+s_result+_T(") = "));
				sel.TypeText(COperate::fraction(m_gjzs+_T("×")+s_Asv1,s_result)+_T(" = ")+s_smax+_T(" (mm),最大箍筋间距:")+s_gjjjjmax+_T(" (mm),取s = ")+s_gjjj+_T("(mm)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
			//12
			s_calresult.Format(_T("%.3lf"),(_ttof(m_gjzs)*_ttof(m_gjzj)*_ttof(m_gjzj)*3.14/(4*_ttof(s_gjjj))));
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
				CString p,pmin;
				p.Format(_T("%.3lf"),m_result/m_b);
				pmin.Format(_T("%.3lf"),0.24*info.ct/info.gjss);
				xb(_T("EQ \\f(A"),_T("sv"));
				xb(_T(",bs) = "),_T(""));
				sel.TypeText(COperate::fraction(m_gjzs+_T("×π×")+m_gjzj+_T("²"),_T("4×")+s_gjjj+_T("×")+
					s_b)+_T(" = ")+p+_T(" > "));
				xb(_T("0.24\\f(f"),_T("t"));
				xb(_T(",f"),_T("yv"));
				sel.TypeText(_T(") = 0.24×")+COperate::fraction(s_ft,s_fvy)+_T(" = ")+pmin+_T(",满足最小配箍率要求。"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
		}
		else
		{
			//按构造配筋
			if(com == 1)
			{
				//6仅配置箍筋
				CString fcv;
				fcv.Format(_T("%.1lf"),acv*info.ct*m_b*h0/1000);
				xb(_T("V"),_T("c"));
				xb(_T(" = α"),_T("cv"));
				xb(_T("f"),_T("t"));
				xb(_T("bh"),_T("0"));
				sel.TypeText(_T(" = ")+s_acv+_T("×")+s_ft+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+fcv+ _T(" (kN)"));
				sel.TypeParagraph();
				//7
				xb(_T("V = ")+s_V+_T("(kN) < V"),_T("c"));
				sel.TypeText(_T(" = ")+fcv+_T("(kN)，应通过构造配置箍筋"));
				sel.TypeParagraph();
			}
			else
			{
				//6
				CString fcv,s_Asb,s_Vc,s_Vsb;
				s_Vc.Format(_T("%.1lf"),(acv*info.ct*m_b*h0)/1000);
				s_Vsb.Format(_T("%.1lf"),(0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8)/1000);
				fcv.Format(_T("%.1lf"),(acv*info.ct*m_b*h0+0.8*_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)*info.ss*1.414/8)/1000);
				s_Asb.Format(_T("%.0lf"),_ttof(m_wqd)*_ttof(m_wqd)*3.14159*_ttof(m_wqn)/4);
				xb(_T("V"),_T("c"));
				xb(_T(" = α"),_T("cv"));
				xb(_T("f"),_T("t"));
				xb(_T("bh"),_T("0"));
				sel.TypeText(_T(" = ")+s_acv+_T("×")+s_ft+_T("×")+s_b+_T("×")+s_h0+_T(" = ")+s_Vc+ _T(" (kN)"));
				sel.TypeParagraph();
				//7
				xb(_T("V"),_T("sb"));
				xb(_T(" = 0.8f"),_T("y"));
				xb(_T("A"),_T("s"));
				xb(_T("sin 45º"),_T(""));
				sel.TypeText(_T(" = 0.8×"));
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
				sel.TypeText(_T("×")+s_fy+_T("×")+s_Asb+_T(" = ")+s_Vsb+ _T(" (kN)"));
				sel.TypeParagraph();
				//8
				xb(_T("V = ")+s_V+_T("(kN) < V"),_T("c"));
				xb(_T(" + V"),_T("sb"));
				sel.TypeText(_T(" = ")+s_Vc+_T(" + ")+s_Vsb+_T(" = ")+fcv+_T("(kN)，应通过构造配置箍筋"));
				sel.TypeParagraph();
			}
			//9
			CString gcalresult;
			gcalresult.Format(_T("%.3lf"),0.24*info.ct*m_b/info.gjss);
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
				xb(_T(",s) = 0.24\\f(f"),_T("t"));
				xb(_T("b,f"),_T("yv"));
				xb(_T(") = "),_T(""));
				sel.TypeText(_T("0.24×")+COperate::fraction(s_ft+_T("×")+s_b,
					s_fvy)+_T(" = ")+gcalresult+_T(" (mm²/mm)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
			//10
			CString s_gjjj;
			s_gjjj.Format(_T("%.0lf"),m_gjjj);
			sel.TypeText(_T("选取箍筋:肢数n=")+m_gjzs+_T(",直径d=")+m_gjzj+_T("(mm)"));
			sel.TypeParagraph();
			//11
			CString s_Asv1;
			s_Asv1.Format(_T("%.1lf"),(_ttof(m_gjzj)*_ttof(m_gjzj)*3.14/4));
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
				xb(_T("EQ A"),_T("sv1"));
				xb(_T(" = \\f(1,4)πd² = "),_T(""));
				sel.TypeText(_T("\\f(1,4)×π×")+m_gjzj+_T("² = ")+s_Asv1+_T(" (mm²)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
			sel.TypeParagraph();
			//12
			CString s_calresult,s_smax;
			s_calresult.Format(_T("%.3lf"),(_ttof(m_gjzs)*_ttof(m_gjzj)*_ttof(m_gjzj)*3.14/(4*_ttof(s_gjjj))));
			s_smax.Format(_T("%.0lf"),(_ttof(m_gjzs)*_ttof(m_gjzj)*_ttof(m_gjzj)*3.14/(4*m_realresult)));
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
				xb(_T("EQ s ≤ \\f(nA"),_T("sv1"));
				sel.TypeText(_T(",")+s_result+_T(") = "));
				sel.TypeText(COperate::fraction(m_gjzs+_T("×")+s_Asv1,s_result)+_T(" = ")+s_smax+_T(" (mm),最大箍筋间距:")+s_gjjjjmax+_T(" (mm),取s = ")+s_gjjj+_T("(mm)"));
			}
			fields.ToggleShowCodes();
			sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)wdMove));
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


void Cxd::OnBnClickedCalculalength()
{
	CWnd::UpdateData(true);
	m_gjjj = _ttof(m_gjzj) * _ttof(m_gjzj) * _ttof(m_gjzs) * 3.14159 / (4 * m_result);
	m_gjjj = (int)m_gjjj/10*10;
	if(m_v * 1000 > 0.7 * info.ct * m_b * h0)
	{
		if(m_h>150&&m_h<=300)
		{
			s_gjjjjmax = _T("150");
			if(m_gjjj>150)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 150;
			}
		}
		else if(m_h<=500)
		{
			s_gjjjjmax = _T("200");
			if(m_gjjj>200)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 200;
			}
		}
		else if(m_h<=800)
		{
			s_gjjjjmax = _T("250");
			if(m_gjjj>250)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 250;
			}
		}
		else
		{
			s_gjjjjmax = _T("300");
			if(m_gjjj>300)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 300;
			}
		}
	}
	else
	{
		if(m_h>150&&m_h<=300)
		{
			s_gjjjjmax = _T("200");
			if(m_gjjj>200)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 200;
			}
		}
		else if(m_h<=500)
		{
			s_gjjjjmax = _T("300");
			if(m_gjjj>300)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 300;
			}
		}
		else if(m_h<=800)
		{
			s_gjjjjmax = _T("350");
			if(m_gjjj>350)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 350;
			}
		}
		else
		{
			s_gjjjjmax = _T("400");
			if(m_gjjj>400)
			{
				MessageBox(_T("超过箍筋配置最大距离按照最大间距配置"));
				m_gjjj = 400;
			}
		}
	}
	GetDlgItem(IDC_Out)->EnableWindow(TRUE);
	CWnd::UpdateData(false);
}


void Cxd::OnBnClickedRadiojg()
{
	GetDlgItem(IDC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_a)->ShowWindow(SW_SHOW);
}


void Cxd::OnBnClickedRadiojgb()
{
	GetDlgItem(IDC_STATIC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_a)->ShowWindow(SW_HIDE);
	GetDlgItem(IDC_namda)->ShowWindow(SW_SHOW);
	GetDlgItem(IDC_STATIC_namda)->ShowWindow(SW_SHOW);
}


void Cxd::OnCancel()
{
	// TODO: 在此添加专用代码和/或调用基类

	//CDialogEx::OnCancel();
}
