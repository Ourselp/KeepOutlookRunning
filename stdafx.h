// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#include <afxwin.h>
#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

using namespace ATL;

#include <algorithm>
#include <list>

using namespace std;
using namespace std::tr1;

// Import for IDTExtensibility2
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"\
	auto_rename auto_search raw_interfaces_only rename_namespace("AddinDesign")
using namespace AddinDesign;

// Office type library (i.e. mso.dll)
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Office")\
	rename("DocumentProperties", "__DocumentProperties")\
	rename("SearchPath", "__SearchPath")\
	rename("RGB", "__RGB")
using namespace Office;

// Outlook type library (i.e. msoutl.olb)
#import "libid:00062FFF-0000-0000-C000-000000000046"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Outlook")\
	rename("CopyFile", "__CopyFile")\
	rename("PlaySound", "__PlaySound")
using namespace Outlook;

// This addin type library
#if defined(_WIN64)
#import "x64\Release\KeepOutlookRunning.tlb"\
	auto_rename auto_search raw_interfaces_only rename_namespace("KeepOutlookRunningAddin")
#else
#import "Release\KeepOutlookRunning.tlb"\
	auto_rename auto_search raw_interfaces_only rename_namespace("KeepOutlookRunningAddin")
#endif

using namespace KeepOutlookRunningAddin;
extern int g_verOLMajor;
