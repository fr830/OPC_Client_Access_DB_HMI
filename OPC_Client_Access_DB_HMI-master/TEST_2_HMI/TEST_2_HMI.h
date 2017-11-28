// TEST_2_HMI.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CTEST_2_HMIApp:
// See TEST_2_HMI.cpp for the implementation of this class
//

class CTEST_2_HMIApp : public CWinApp
{
public:
	CTEST_2_HMIApp();

// Overrides
	public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern CTEST_2_HMIApp theApp;