#pragma once


// CDcpyc �Ի���

class CDcpyc : public CDialogEx
{
	DECLARE_DYNAMIC(CDcpyc)

public:
	CDcpyc(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CDcpyc();

// �Ի�������
	enum { IDD = IDD_DCPYC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
};
