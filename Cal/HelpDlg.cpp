// HelpDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Cal.h"
#include "HelpDlg.h"


// CHelpDlg �Ի���

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


// CHelpDlg ��Ϣ�������

BOOL CHelpDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	m_Text="";
	CString str="";
	str="���ʽ����\r\n���ʽ�����ּ��������+-*/�Լ�()���\r\n";
	str=str+"������һ�º���sin,asin,cos,acos,tan,atan,sqrt,exp,log,fabs\r\n��Щ����ֻ��һ��������������()\r\n";
	str=str+"���¶��ǿ��Ե�\r\nsin(5+6),exp(103-21+43),tan(6)\r\n����pow,max,min�⼸��������������������������ʾ:\r\n";
	str=str+"max(6,1+8),pow(2,103-24)\r\n���г���PI��ʾԲ���ʣ�E��ʾ����e\r\n���ʽʾ��\r\n";
	str=str+"3+sin(PI/6)-6*5*pow(2,3)\r\n\r\n";
	str=str+"��������Ϊ������ʾ���ֺź���ע��\r\n";
	str=str+"\r\n";
	str=str+"[������ʼ]\r\n";
	str=str+"n=2730;����1��ת��,��λr/min\r\n";
	str=str+"Mn=6;����ģ��,��λmm\r\n";
	str=str+"belta=14;������,��λ��\r\n";
	str=str+"Han'=1;����ݶ���ϵ��\r\n";
	str=str+"[��������]\r\n";
	str=str+"[��ʽ��ʼ]\r\n";
	str=str+"Rn=pow(Mn,0.51);\r\n";
	str=str+"Cn=Rn+pow(2,8);\r\n";
	str=str+"Max=max(Rn,Cn);\r\n";
	str=str+"Min=min(Rn,Cn);\r\n";
	str=str+"BB=2+max(Rn,Cn)*7+sin(min(Rn,Cn));\r\n";
	str=str+"[��ʽ����]\r\n";
	str=str+"\r\n";
	str=str+"[������ʼ]��[��������]��������֪�������壬=���Ϊ����\r\n";
	str=str+"=�ұ��ǲ��������ı��ʽ�����÷ֺ�ע��\r\n";
	str=str+"[��ʽ��ʼ][��ʽ����]������Ҫ����Ĳ����ͷ���\r\n";
	str=str+"�����ұߵ�δ֪������һ��Ҫ�������������ҵ�\r\n";
	str=str+"\r\n";
	str=str+"ͨ���ļ����㷽���ǽ��������ļ�����ʽ�ṩ,�����ʽͬ��\r\n";




		m_Text=str;

		UpdateData(FALSE);
	

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣��OCX ����ҳӦ���� FALSE
}
