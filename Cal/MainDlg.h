#pragma once


// CMainDlg 对话框

class CMainDlg : public CDialog
{
	DECLARE_DYNAMIC(CMainDlg)

public:
	CMainDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CMainDlg();

// 对话框数据
	enum { IDD = IDD_DIALOG_MAIN };

protected:
	HICON m_hIcon;
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	CString m_Result;
	CString m_Expression;
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton5();
	afx_msg void OnBnClickedButtonHelp();
	afx_msg void OnBnClickedButtonFilecal();
	CString m_EInput;
	CString m_EOutput;
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
};
