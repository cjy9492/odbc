// ODBCTEST.h : main header file for the ODBCTEST application
//

#if !defined(AFX_ODBCTEST_H__1B069A68_12C2_44B8_80F3_D58AA9F09F4D__INCLUDED_)
#define AFX_ODBCTEST_H__1B069A68_12C2_44B8_80F3_D58AA9F09F4D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CODBCTESTApp:
// See ODBCTEST.cpp for the implementation of this class
//

class CODBCTESTApp : public CWinApp
{
public:
	CODBCTESTApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CODBCTESTApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CODBCTESTApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ODBCTEST_H__1B069A68_12C2_44B8_80F3_D58AA9F09F4D__INCLUDED_)
