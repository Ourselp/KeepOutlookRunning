/*!-----------------------------------------------------------------------
connect.cpp

The main implementation of the addin. It includes the implementation
for IDTExtensibility2
-----------------------------------------------------------------------!*/
#include "stdafx.h"
#include "connect.h"

const TCHAR g_tzWndClassRen[] = TEXT("rctrl_renwnd32");

int  g_verOLMajor;
bool g_fHostShutdown = false;
LONG_PTR g_origWndProc;

/*!-----------------------------------------------------------------------
CConnect implementation
-----------------------------------------------------------------------!*/

_ATL_FUNC_INFO CConnect::DispatchFuncInfo1 = {CC_STDCALL, VT_EMPTY, 1, {VT_DISPATCH}};

LRESULT CALLBACK NewWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (uMsg == WM_CLOSE) {
		uMsg = WM_SYSCOMMAND;
		wParam = SC_MINIMIZE;
		lParam = NULL;
	}

	return CallWindowProc((WNDPROC)g_origWndProc, hWnd, uMsg, wParam, lParam);
}

STDMETHODIMP CConnect::OnConnection
	( IDispatch *pApplication
	, ext_ConnectMode /* ConnectMode */
	, IDispatch* /* pAddInInst */
	, SAFEARRAY ** /* custom */ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComBSTR bstrVersion;
	int      v2, v3, v4;

	if (!pApplication)
		return E_POINTER;

	m_spApplication = pApplication;

	// Get the version of Outlook
	if (FAILED(m_spApplication->get_Version(&bstrVersion)))
		return E_FAIL;
	if (4 != swscanf_s(bstrVersion, L"%d.%d.%d.%d", &g_verOLMajor, &v2, &v3, &v4))
		return E_FAIL;

	// XXX: Multiple outlook windows instances? This only catches the first
	HWND hWndOutlook = FindWindowEx(NULL, NULL, g_tzWndClassRen, NULL);
	g_origWndProc = GetWindowLongPtr(hWndOutlook, GWLP_WNDPROC);

	WNDPROC newWndProc = NewWndProc;
	SetWindowLongPtr(hWndOutlook, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(newWndProc));

	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection
	( ext_DisconnectMode extRemoveMode
	, SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	g_fHostShutdown = (extRemoveMode == ext_dm_HostShutdown);

	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate(SAFEARRAY ** /*custom*/)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete(SAFEARRAY ** /*custom*/)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown(SAFEARRAY ** /*custom*/)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return S_OK;
}


