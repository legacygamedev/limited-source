#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_ADDRESS_INCLUDED__
#define JETBYTE_ADDRESS_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Address.h $
// Version        : $Revision: 2 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 11/06/02 9:02 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/Address.h $
// 
// 2     11/06/02 9:02 Len
// Changed type library id.
// 
// 1     3/06/02 11:37 Len
// 
///////////////////////////////////////////////////////////////////////////////
//
// Copyright 1997 - 2002 JetByte Limited.
//
// JetByte Limited grants you ("Licensee") a non-exclusive, royalty free, 
// licence to use, modify and redistribute this software in source and binary 
// code form, provided that i) this copyright notice and licence appear on all 
// copies of the software; and ii) Licensee does not utilize the software in a 
// manner which is disparaging to JetByte Limited.
//
// This software is provided "as is" without a warranty of any kind. All 
// express or implied conditions, representations and warranties, including
// any implied warranty of merchantability, fitness for a particular purpose
// or non-infringement, are hereby excluded. JetByte Limited and its licensors 
// shall not be liable for any damages suffered by licensee as a result of 
// using, modifying or distributing the software or its derivatives. In no
// event will JetByte Limited be liable for any lost revenue, profit or data,
// or for direct, indirect, special, consequential, incidental or punitive
// damages, however caused and regardless of the theory of liability, arising 
// out of the use of or inability to use software, even if JetByte Limited 
// has been advised of the possibility of such damages.
//
// This software is not designed or intended for use in on-line control of 
// aircraft, air traffic, aircraft navigation or aircraft communications; or in 
// the design, construction, operation or maintenance of any nuclear 
// facility. Licensee represents and warrants that it will not use or 
// redistribute the Software for such purposes. 
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

#include "resource.h"       
#include "IAddressInit.h"

///////////////////////////////////////////////////////////////////////////////
// CAddress
///////////////////////////////////////////////////////////////////////////////

class ATL_NO_VTABLE CAddress : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAddress, &CLSID_Address>,
	public ISupportErrorInfo,
	public IDispatchImpl<IAddress, &IID_IAddress, &LIBID_JBSOCKETSERVERLib>,
   public IAddressInit
{
   public:

      CAddress();

      DECLARE_REGISTRY_RESOURCEID(IDR_ADDRESS)
      DECLARE_NOT_AGGREGATABLE(CAddress)

      BEGIN_COM_MAP(CAddress)
	      COM_INTERFACE_ENTRY(IAddress)
	      COM_INTERFACE_ENTRY(IDispatch)
	      COM_INTERFACE_ENTRY(ISupportErrorInfo)
         COM_INTERFACE_ENTRY(IAddressInit)
      END_COM_MAP()

      // ISupportsErrorInfo
	   
      STDMETHOD(InterfaceSupportsErrorInfo)(
         REFIID riid);

      // IAddress
	   
      STDMETHOD(get_Port)(
         /*[out, retval]*/ long *pVal);
	   
      STDMETHOD(get_Address)(
         /*[out, retval]*/ BSTR *pVal);

      // IAddressInit

      STDMETHOD(Init)(
         /*[in]*/ unsigned long address,
         /*[in]*/ unsigned short port);

   private :

      unsigned short m_port;
      CComBSTR m_address;
};

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_ADDRESS_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
