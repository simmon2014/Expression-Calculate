// CalDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Cal.h"
#include "CalDlg.h"
#include "MyUtil.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CCalDlg �Ի���



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


// CCalDlg ��Ϣ�������

BOOL CCalDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��\������...\���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO���ڴ���Ӷ���ĳ�ʼ������


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

   
	
	return TRUE;  // ���������˿ؼ��Ľ��㣬���򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CCalDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ��������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù����ʾ��
HCURSOR CCalDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CCalDlg::OnBnClickedButton1()
{
	UpdateData(TRUE);
	static char BASED_CODE szFilter[] = "�ı��ļ� (*.txt)|*.txt|All Files (*.*)|*.*||";
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
	static char BASED_CODE szFilter[] = "�ı��ļ� (*.txt)|*.txt|All Files (*.*)|*.*||";
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
	static char BASED_CODE szFilter[] = "doc�ļ� (*.doc)|*.doc|All Files (*.*)|*.*||";
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
	static char BASED_CODE szFilter[] = "doc�ļ� (*.doc)|*.doc|All Files (*.*)|*.*||";
	CFileDialog  fdlg(FALSE,"doc",NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,szFilter,NULL,0);
	if(fdlg.DoModal()==IDOK)
	{
		m_Text4=fdlg.GetPathName();
	}
	UpdateData(FALSE);
}




void CCalDlg::OnBnClickedButtonCal()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);
	if(m_Text1=="") return;
	int status;
	status=m_Compute.ReadFile(m_Text1);
	if(status==0) 
	{
		MessageBox("�����ļ���ʽ���󣬼���ʧ��!","��ʾ",MB_OK);
		return ;
	}
	status=m_Compute.Check();
	if(status==0) 
	{
		MessageBox("�����ļ���ʽ���󣬼���ʧ��!","��ʾ",MB_OK);
		return ;
	}
	status=m_Compute.Compute();
	if(status==1) 
	{
		MessageBox("����ɹ�","��ϲ",MB_OK);
	}
	else
	{
		MessageBox("�����ļ���ʽ���󣬼���ʧ��!","��ʾ",MB_OK);
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
	//case IDC_BUTTON_CAL://�ؼ��������Դ��  
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




	// TODO:  ���Ĭ�ϵĲ������軭�ʣ��򷵻���һ������
	return hbr;
}
