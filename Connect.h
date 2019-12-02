/*!-----------------------------------------------------------------------
	connect.h
-----------------------------------------------------------------------!*/
#pragma once
#include "ids.h"

class CConnect;

typedef
	IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), 1, 0>
	IDTExtensibilityImpl;


/*------------------------------------------------------------------------------
	CConnect class - what Outlook uses to communicate with an add-in
------------------------------------------------------------------------------*/
class ATL_NO_VTABLE CConnect
	: public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CConnect, &__uuidof(Connect)>
	, public IDTExtensibilityImpl
{
public:
	CConnect() {}

	// Setup the registration found in addin.rgs
	static HRESULT WINAPI UpdateRegistry(BOOL bRegister) throw()
	{
		ATL::_ATL_REGMAP_ENTRY regMapEntries[] =
		{ { OLESTR("PROGID"), ADDIN_PROGID }
		, { OLESTR("CLSID"), ADDIN_CLSID_STR }
		, { OLESTR("TYPELIB"), TYPELIB_GUID_STR }
		, { NULL, NULL }
		};

		return ATL::_pAtlModule->UpdateRegistryFromResource(IDR_ADDIN, bRegister, regMapEntries);
	}

	DECLARE_NOT_AGGREGATABLE(CConnect)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY(IDTExtensibility2)
	END_COM_MAP()

	static _ATL_FUNC_INFO DispatchFuncInfo1;

	BEGIN_SINK_MAP(CConnect)
	END_SINK_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct() { return S_OK;	}
	void FinalRelease()	{ }

public:
	// IDTExtensibility2 interface
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

protected:
	_ApplicationPtr  m_spApplication;
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
