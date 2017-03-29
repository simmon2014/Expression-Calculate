// CalDlg.h : ͷ�ļ�
//

#pragma once
#include "Input.h"

// CCalDlg �Ի���
class CCalDlg : public CDialog
{
// ����
public:
	CCalDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_CAL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CInput m_Compute;
public:
	CString m_Text1;
	CString m_Text2;
	CString m_Text3;
	CString m_Text4;
	afx_msg void OnBnClickedButtonCal();
	afx_msg void OnBnClickedButtonOutput();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
};
