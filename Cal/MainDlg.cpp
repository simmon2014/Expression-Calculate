// MainDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Cal.h"
#include "MainDlg.h"
#include "MyUtil.h"
#include "HelpDlg.h"
#include "CalDlg.h"
#include "Input.h"


// CMainDlg 对话框

IMPLEMENT_DYNAMIC(CMainDlg, CDialog)
CMainDlg::CMainDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMainDlg::IDD, pParent)
	, m_Result(_T(""))
	, m_Expression(_T(""))
	, m_EInput(_T(""))
	, m_EOutput(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

CMainDlg::~CMainDlg()
{
}

void CMainDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_RESULT, m_Result);
	DDX_Text(pDX, IDC_EDIT1, m_Expression);
	DDX_Text(pDX, IDC_EDIT2, m_EInput);
	DDX_Text(pDX, IDC_EDIT5, m_EOutput);
}


BEGIN_MESSAGE_MAP(CMainDlg, CDialog)
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON5, OnBnClickedButton5)
	ON_BN_CLICKED(IDC_BUTTON_HELP, OnBnClickedButtonHelp)
	ON_BN_CLICKED(IDC_BUTTON_FILECAL, OnBnClickedButtonFilecal)
	ON_WM_KEYDOWN()
END_MESSAGE_MAP()


// CMainDlg 消息处理程序

BOOL CMainDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// TODO:  在此添加额外的初始化
	CString str;
	str="[参数开始]\r\n";
	str=str+"[参数结束]\r\n";
	str=str+"[公式开始]\r\n";
	str=str+"[公式结束]\r\n";
	m_EInput=str;
	UpdateData(FALSE);

	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常：OCX 属性页应返回 FALSE
}

void CMainDlg::OnBnClickedButton1()
{
	UpdateData(TRUE);
	double result;
	CMyUtil util;
   CStringArray tempArray;
   int status=util.CheckExpression(m_Expression,&tempArray,&tempArray);
   if(status==0) return;
	result=util.CalString(m_Expression);
	m_Result.Format("%s=%f",m_Expression,result);
	UpdateData(FALSE);
}

void CMainDlg::OnBnClickedButton5()///计算方程
{
	UpdateData(TRUE);
    CInput cal;
	cal.CalExEqual(m_EInput,m_EOutput);
	UpdateData(FALSE);
}

void CMainDlg::OnBnClickedButtonHelp()
{
	CHelpDlg dlg;
	dlg.DoModal();
}

void CMainDlg::OnBnClickedButtonFilecal()
{
	CCalDlg dlg;
	dlg.DoModal();
}

void CMainDlg::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	CDialog::OnKeyDown(nChar, nRepCnt, nFlags);
}
