#pragma once


// CChooseDlg �Ի���

class CChooseDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CChooseDlg)

public:
	CChooseDlg(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CChooseDlg();

// �Ի�������
	enum { IDD = IDD_DIALOG_CHOOSE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedSbdc();
	afx_msg void OnBnClickedDbdc();
	afx_msg void OnBnClickedDbdc2();
	afx_msg void OnBnClickedDbdc3();
	afx_msg void OnBnClickedSbdc2();
	afx_msg void OnBnClickedSbdc3();
	afx_msg void OnBnClickedSbdc4();
	afx_msg void OnBnClickedSbdc5();
	afx_msg void OnBnClickedSbdc6();
};
