// plugin.h : Declaration of the Cplugin

#ifndef __PLUGIN_H_
#define __PLUGIN_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// Cplugin
class ATL_NO_VTABLE Cplugin : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<Cplugin, &CLSID_plugin>,
	public IDispatchImpl<Iplugin, &IID_Iplugin, &LIBID_ATLSAMPLELib>
{
public:
	Cplugin()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_PLUGIN)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(Cplugin)
	COM_INTERFACE_ENTRY(Iplugin)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

// Iplugin
public:
	STDMETHOD(StartUp)(/*[in,out]*/ int*);
	STDMETHOD(SetHost)(/*[in,out]*/ IDispatch**);
private:
	IDispatch FAR* frmMain;
};

#endif //__PLUGIN_H_
