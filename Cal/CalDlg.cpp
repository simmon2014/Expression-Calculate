// CalDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Cal.h"
#include "CalDlg.h"
#include "MyUtil.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CCalDlg 对话框



CCalDlg::CCalDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CCalDlg::IDD, pParent)
	, m_Text1(_T(""))
	, m_Text2(_T(""))
	, m_Text3(_T(""))
	, m_Text4(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCalDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, m_Text1);
	DDX_Text(pDX, IDC_EDIT2, m_Text2);
	DDX_Text(pDX, IDC_EDIT3, m_Text3);
	DDX_Text(pDX, IDC_EDIT4, m_Text4);
}

BEGIN_MESSAGE_MAP(CCalDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_CAL, OnBnClickedButtonCal)
	ON_BN_CLICKED(IDC_BUTTON_Output, OnBnClickedButtonOutput)
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, OnBnClickedButton4)
	ON_WM_CTLCOLOR()
END_MESSAGE_MAP()


// CCalDlg 消息处理程序

BOOL CCalDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将\“关于...\”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO：在此添加额外的初始化代码


	CFont * f; 

	f = new CFont; 


	f->CreateFont(16, // nHeight 

		0, // nWidth 

		0, // nEscapement 

		0, // nOrientation 

		FW_BOLD, // nWeight 

		TRUE, // bItalic 

		TRUE, // bUnderline 

		0, // cStrikeOut 

		ANSI_CHARSET, // nCharSet 

		OUT_DEFAULT_PRECIS, // nOutPrecision 

		CLIP_DEFAULT_PRECIS, // nClipPrecision 

		DEFAULT_QUALITY, // nQuality 

		DEFAULT_PITCH | FF_SWISS, // nPitchAndFamily 

		_T("Arial")); // lpszFac 

	this->GetDlgItem(IDC_BUTTON_CAL)->SetFont(f);

   
	
	return TRUE;  // 除非设置了控件的焦点，否则返回 TRUE
}

void CCalDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CCalDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标显示。
HCURSOR CCalDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CCalDlg::OnBnClickedButton1()
{
	UpdateData(TRUE);
	static char BASED_CODE szFilter[] = "文本文件 (*.txt)|*.txt|All Files (*.*)|*.*||";
	CFileDialog  fdlg(TRUE,NULL,NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,szFilter,NULL,0);
	if(fdlg.DoModal()==IDOK)
	{
		m_Text1=fdlg.GetPathName();
	}
	UpdateData(FALSE);
}

void CCalDlg::OnBnClickedButton2() 
{
	UpdateData(TRUE);
	static char BASED_CODE szFilter[] = "文本文件 (*.txt)|*.txt|All Files (*.*)|*.*||";
	CFileDialog  fdlg(FALSE,"txt",NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,szFilter,NULL,0);
	if(fdlg.DoModal()==IDOK)
	{
		m_Text2=fdlg.GetPathName();
	}
	UpdateData(FALSE);
}



void CCalDlg::OnBnClickedButton3() 
{
	UpdateData(TRUE);
	static char BASED_CODE szFilter[] = "doc文件 (*.doc)|*.doc|All Files (*.*)|*.*||";
	CFileDialog  fdlg(TRUE,NULL,NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,szFilter,NULL,0);
	if(fdlg.DoModal()==IDOK)
	{
		m_Text3=fdlg.GetPathName();
	}
	UpdateData(FALSE);
}

void CCalDlg::OnBnClickedButton4()
{
	UpdateData(TRUE);
	static char BASED_CODE szFilter[] = "doc文件 (*.doc)|*.doc|All Files (*.*)|*.*||";
	CFileDialog  fdlg(FALSE,"doc",NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,szFilter,NULL,0);
	if(fdlg.DoModal()==IDOK)
	{
		m_Text4=fdlg.GetPathName();
	}
	UpdateData(FALSE);
}




void CCalDlg::OnBnClickedButtonCal()
{
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(TRUE);
	if(m_Text1=="") return;
	int status;
	status=m_Compute.ReadFile(m_Text1);
	if(status==0) 
	{
		MessageBox("输入文件格式错误，计算失败!","提示",MB_OK);
		return ;
	}
	status=m_Compute.Check();
	if(status==0) 
	{
		MessageBox("输入文件格式错误，计算失败!","提示",MB_OK);
		return ;
	}
	status=m_Compute.Compute();
	if(status==1) 
	{
		MessageBox("计算成功","恭喜",MB_OK);
	}
	else
	{
		MessageBox("输入文件格式错误，计算失败!","提示",MB_OK);
	}

	

}

void CCalDlg::OnBnClickedButtonOutput()
{
	UpdateData(TRUE);
	int status;
	status=m_Compute.WriteOutput(m_Text2);

	ShellExecute(this->GetSafeHwnd(),"open",m_Text2,NULL,NULL,SW_SHOWNORMAL);
	
	if(m_Text3==""||m_Text4=="") return;
	m_Compute.WriteDoc(m_Text3,m_Text4);
	UpdateData(FALSE);
}





HBRUSH CCalDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);


	//// TODO: Change any attributes of the DC here  
	//switch(pWnd->GetDlgCtrlID())  
	//{  
	//case IDC_BUTTON_CAL://控件定义的资源名  
	//	pDC->SetTextColor(RGB(255,0,0));  
	//	pDC->SetBkColor(RGB(0,255,0));  
	//	break;  
	//case IDC_EDIT1:  
	//	pDC->SetTextColor(RGB(0,0,255));  
	//	break;  
	//case IDC_EDIT2:  
	//	pDC->SetTextColor(RGB(0,0,255));  
	//	break;  
	//case IDC_EDIT3:  
	//	pDC->SetTextColor(RGB(0,0,255));  
	//	break;  
	//}  




	// TODO:  如果默认的不是所需画笔，则返回另一个画笔
	return hbr;
}
