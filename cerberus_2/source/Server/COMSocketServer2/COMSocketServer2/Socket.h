#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_SOCKET_INCLUDED__
#define JETBYTE_SOCKET_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Socket.h $
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
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/Socket.h $
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

#include "ISocketInit.h"
#include "SocketServer.h"

#include "JetByteTools\COMTools\ExceptionToCOMError.h"

///////////////////////////////////////////////////////////////////////////////
// CSocket
///////////////////////////////////////////////////////////////////////////////

class ATL_NO_VTABLE CSocket : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSocket, &CLSID_Socket>,
	public ISupportErrorInfo,
	public IDispatchImpl<ISocket, &IID_ISocket, &LIBID_JBSOCKETSERVERLib>,
   public ISocketInit,
   public JetByteTools::COM::TExceptionToCOMError<CSocket>
{
   public:
	   
      CSocket();
      ~CSocket();

      DECLARE_REGISTRY_RESOURCEID(IDR_SOCKET)
      DECLARE_NOT_AGGREGATABLE(CSocket)

      BEGIN_COM_MAP(CSocket)
	      COM_INTERFACE_ENTRY(ISocket)
	      COM_INTERFACE_ENTRY(IDispatch)
	      COM_INTERFACE_ENTRY(ISupportErrorInfo)
         COM_INTERFACE_ENTRY(ISocketInit)
      END_COM_MAP()

      // ISupportsErrorInfo
	   
      STDMETHOD(InterfaceSupportsErrorInfo)(
         REFIID riid);

      // ISocket

      STDMETHOD(get_RemoteAddress)(
         /*[out, retval]*/ IAddress **ppVal);
	   
      STDMETHOD(RequestRead)();
	   
      STDMETHOD(Write)(
         /*[in]*/ VARIANT arrayOfBytes,
         /*[in, optional] */ VARIANT thenShutdown);
	   
      STDMETHOD(WriteString)(
         /*[in]*/ BSTR data,
         /*[in, optional] */ VARIANT sendAsUNICODE,
         /*[in, optional] */ VARIANT thenShutdown);

	   STDMETHOD(Shutdown)(
         /*[in]*/ ShutdownMethod how);

      STDMETHOD(Close)();

	   STDMETHOD(get_UserData)(
         /*[out, retval]*/ VARIANT *pVal);
	   
      STDMETHOD(put_UserData)(
         /*[in]*/ VARIANT newVal);

      // ISocketInit

      STDMETHOD(Init)(
         /*[in]*/ unsigned long address,
         /*[in]*/ unsigned short port,
         /*[in]*/ void *pSocket);

   private :

      CCOMSocketServer::Socket *m_pSocket;

      IAddress *m_pAddress;

      CComVariant m_userData;
};

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_SOCKET_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
