// TEST_2_HMIDlg.cpp : implementation file
//

#include "stdafx.h"
#include "TEST_2_HMI.h"
#include "TEST_2_HMIDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
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


// CTEST_2_HMIDlg dialog




CTEST_2_HMIDlg::CTEST_2_HMIDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTEST_2_HMIDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CTEST_2_HMIDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_STATUS, m_EditStatus);
	DDX_Control(pDX, IDC_ORDER, m_EditOrder);
}

BEGIN_MESSAGE_MAP(CTEST_2_HMIDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_START, &CTEST_2_HMIDlg::OnBnClickedButtonStart)
	ON_BN_CLICKED(IDC_BUTTON_STOP, &CTEST_2_HMIDlg::OnBnClickedButtonStop)
	ON_BN_CLICKED(IDC_BUTTON_UPDATE, &CTEST_2_HMIDlg::OnBnClickedButtonUpdate)
END_MESSAGE_MAP()


// CTEST_2_HMIDlg message handlers

BOOL CTEST_2_HMIDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
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

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CTEST_2_HMIDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CTEST_2_HMIDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CTEST_2_HMIDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CTEST_2_HMIDlg::OnBnClickedButtonStart()
{
	m_csSecure.Init();
	m_pRecipe.lngYieldStrength = 1200;
	CoInitializeEx( NULL, COINIT_MULTITHREADED );
	m_cOPCClient.SetRecipe(&m_pRecipe);

	long ret = m_cOPCClient.Connect(500);

	m_EditStatus.SetWindowTextA(m_cOPCClient.getPLCStatusMessage());
	//m_cOPCClient.Loop(&m_pRecipe);
	m_EditStatus.SetWindowTextA(m_cOPCClient.getPLCStatusMessage());
}

void CTEST_2_HMIDlg::OnBnClickedButtonStop()
{
	long ret = m_cOPCClient.Close();
	m_EditStatus.SetWindowTextA(m_cOPCClient.getPLCStatusMessage());
	CoUninitialize();
	m_csSecure.Term();

}

void CTEST_2_HMIDlg::OnBnClickedButtonUpdate()
{
/*
	m_typRecipe->strOrderName = "28112017";
	m_typRecipe->strCustomerName = "REDEX";
	m_typRecipe->strCoilNumber = "28112017_1";
	m_typRecipe->lngThickness = 6;
	m_typRecipe->lngWidth    = 800;
	m_typRecipe->lngEModul    = 210000;
	m_typRecipe->lngYieldStrength = 410;
	m_typRecipe->strCassetteName = "C2";*/
	m_cIftDB.UpdateDB(&m_pRecipe);
	
}
