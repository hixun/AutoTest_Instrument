
// PI_Serial.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CPI_SerialApp:
// See PI_Serial.cpp for the implementation of this class
//

class CPI_SerialApp : public CWinApp
{
public:
	CPI_SerialApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern CPI_SerialApp theApp;