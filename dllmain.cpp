// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "dllmain.h"

CAddinModule _AtlModule;

class CAddinApp : public CWinApp
{
public:

	// Overrides
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CAddinApp, CWinApp)
END_MESSAGE_MAP()

CAddinApp theApp;

BOOL CAddinApp::InitInstance()
{
	return CWinApp::InitInstance();
}

int CAddinApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
