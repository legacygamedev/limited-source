#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_SERVER_INCLUDED__
#define JETBYTE_SERVER_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Server.h $
// Version        : $Revision: 2 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 11/06/02 9:01 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/Server.h $
// 
// 2     11/06/02 9:01 Len
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
#include "IDataInit.h"
#include "JetByteTools\COMTools\AsyncServerEventHelper.h"
#include "JetByteTools\COMTools\ExceptionToCOMError.h"

#include "COMSocketServerCP.h"

///////////////////////////////////////////////////////////////////////////////
// Classes defined in other files
///////////////////////////////////////////////////////////////////////////////

class CCOMSocketServer;

///////////////////////////////////////////////////////////////////////////////
// CServer
///////////////////////////////////////////////////////////////////////////////

class ATL_NO_VTABLE CServer : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CServer, &CLSID_Server>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CServer>,
	public IDispatchImpl<IServer, &IID_IServer, &LIBID_JBSOCKETSERVERLib>,
   public IAddressInit,
   public _AsyncServerEvent,
   public CProxy_IServerEvents<CServer>,
   public JetByteTools::COM::TExceptionToCOMError<CServer>
{
   public:
	   
      CServer();
      ~CServer();

      DECLARE_REGISTRY_RESOURCEID(IDR_SERVER)
      DECLARE_NOT_AGGREGATABLE(CServer)

      BEGIN_COM_MAP(CServer)
	      COM_INTERFACE_ENTRY(IServer)
	      COM_INTERFACE_ENTRY(IDispatch)
	      COM_INTERFACE_ENTRY(ISupportErrorInfo)
	      COM_INTERFACE_ENTRY(IConnectionPointContainer)
         COM_INTERFACE_ENTRY(IAddressInit)
         COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
      END_COM_MAP()

      BEGIN_CONNECTION_POINT_MAP(CServer)
         CONNECTION_POINT_ENTRY(DIID__IServerEvents)
      END_CONNECTION_POINT_MAP()

      // ISupportsErrorInfo
	
      STDMETHOD(InterfaceSupportsErrorInfo)(
         REFIID riid);

      // IServer
	   
      STDMETHOD(StopListening)();
	   
      STDMETHOD(StartListening)();
	   
      STDMETHOD(get_LocalAddress)(
         /*[out, retval]*/ IAddress **ppVal);

      // IAddressInit

      STDMETHOD(Init)(
         /*[in]*/ unsigned long address,
         /*[in]*/ unsigned short port);

      // _AsyncServerEvent

      STDMETHOD(OnEvent)(
         long eventID);

   private :

      HRESULT OnConnectionEstablished();

      HRESULT OnDataRecieved();
      
      HRESULT OnConnectionClosed();

      IAddress *m_pIAddress;

      IData *m_pIData;

      IDataInit *m_pIDataInit;

      CAsyncServerEventHelper m_asyncServerEventHelper;

      CCOMSocketServer *m_pSocketServer;
};

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_SERVER_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
