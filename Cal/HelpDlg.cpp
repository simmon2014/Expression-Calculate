// HelpDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Cal.h"
#include "HelpDlg.h"


// CHelpDlg 对话框

IMPLEMENT_DYNAMIC(CHelpDlg, CDialog)
CHelpDlg::CHelpDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CHelpDlg::IDD, pParent)
	, m_Text(_T(""))
{
}

CHelpDlg::~CHelpDlg()
{
}

void CHelpDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, m_Text);
}


BEGIN_MESSAGE_MAP(CHelpDlg, CDialog)
END_MESSAGE_MAP()


// CHelpDlg 消息处理程序

BOOL CHelpDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	m_Text="";
	CString str="";
	str="表达式计算\r\n表达式由数字及运算符号+-*/以及()组成\r\n";
	str=str+"还包括一下函数sin,asin,cos,acos,tan,atan,sqrt,exp,log,fabs\r\n这些函数只有一个参数，并需用()\r\n";
	str=str+"以下都是可以的\r\nsin(5+6),exp(103-21+43),tan(6)\r\n还有pow,max,min这几个函数都是两个参数，如下所示:\r\n";
	str=str+"max(6,1+8),pow(2,103-24)\r\n还有常数PI表示圆周率，E表示常数e\r\n表达式示例\r\n";
	str=str+"3+sin(PI/6)-6*5*pow(2,3)\r\n\r\n";
	str=str+"方程输入为如下所示，分号后是注释\r\n";
	str=str+"\r\n";
	str=str+"[参数开始]\r\n";
	str=str+"n=2730;齿轮1轴转速,单位r/min\r\n";
	str=str+"Mn=6;法向模数,单位mm\r\n";
	str=str+"belta=14;螺旋角,单位度\r\n";
	str=str+"Han'=1;法面齿顶高系数\r\n";
	str=str+"[参数结束]\r\n";
	str=str+"[公式开始]\r\n";
	str=str+"Rn=pow(Mn,0.51);\r\n";
	str=str+"Cn=Rn+pow(2,8);\r\n";
	str=str+"Max=max(Rn,Cn);\r\n";
	str=str+"Min=min(Rn,Cn);\r\n";
	str=str+"BB=2+max(Rn,Cn)*7+sin(min(Rn,Cn));\r\n";
	str=str+"[公式结束]\r\n";
	str=str+"\r\n";
	str=str+"[参数开始]和[参数结束]里面是已知参数定义，=左边为参数\r\n";
	str=str+"=右边是不含参数的表达式可以用分号注释\r\n";
	str=str+"[公式开始][公式结束]里面是要计算的参数和方程\r\n";
	str=str+"方程右边的未知参数，一定要能在他的上面找到\r\n";
	str=str+"\r\n";
	str=str+"通过文件计算方程是将输入以文件的形式提供,具体格式同上\r\n";




		m_Text=str;

		UpdateData(FALSE);
	

	// TODO:  在此添加额外的初始化

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常：OCX 属性页应返回 FALSE
}
