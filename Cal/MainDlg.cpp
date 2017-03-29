// MainDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Cal.h"
#include "MainDlg.h"
#include "MyUtil.h"
#include "HelpDlg.h"
#include "CalDlg.h"
#include "Input.h"


// CMainDlg �Ի���

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


// CMainDlg ��Ϣ�������

BOOL CMainDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	CString str;
	str="[������ʼ]\r\n";
	str=str+"[��������]\r\n";
	str=str+"[��ʽ��ʼ]\r\n";
	str=str+"[��ʽ����]\r\n";
	m_EInput=str;
	UpdateData(FALSE);

	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣��OCX ����ҳӦ���� FALSE
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

void CMainDlg::OnBnClickedButton5()///���㷽��
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
