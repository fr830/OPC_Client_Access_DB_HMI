// TEST_2_HMIDlg.h : header file
//

#pragma once
#include "OPCClient.h"
#include "ItfDB.h"
#include "afxwin.h"

// CTEST_2_HMIDlg dialog
class CTEST_2_HMIDlg : public CDialog
{
// Construction
public:
	CTEST_2_HMIDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_TEST_2_HMI_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonStart();

	/* Custom members */
	typRecipe m_pRecipe;
	COPCClient m_cOPCClient;
	CItfDB m_cIftDB;

	//Critical section for the protection of the thread
	CComCriticalSection	m_csSecure;


	CEdit m_EditStatus;
	CEdit m_EditOrder;
	afx_msg void OnBnClickedButtonStop();
	afx_msg void OnBnClickedButtonUpdate();
};
