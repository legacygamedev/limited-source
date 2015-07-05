///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Server.cpp $
// Version        : $Revision: 1 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 3/06/02 11:37 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/Server.cpp $
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

#include "stdafx.h"
#include "COMSocketServer.h"
#include "Server.h"

#include "Address.h"
#include "Data.h"
#include "Socket.h"

#include "SocketServer.h"

#include "JetByteTools\Win32Tools\Exception.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using JetByteTools::Win32::CIOBuffer;
using JetByteTools::Win32::CException;

///////////////////////////////////////////////////////////////////////////////
// CAddress
///////////////////////////////////////////////////////////////////////////////

CServer::CServer()
   :  m_pIAddress(0),
      m_pIData(0),
      m_pIDataInit(0),
      m_asyncServerEventHelper(*this),
      m_pSocketServer(0)
{

}

CServer::~CServer()
{
   try
   {
      if (m_pSocketServer)
      {
         m_pSocketServer->WaitForShutdownToComplete();

         delete m_pSocketServer;
      }

      if (m_pIAddress)
      {
         m_pIAddress->Release();
      }

      if (m_pIData)
      {
         m_pIData->Release();
      }

      if (m_pIDataInit)
      {
         m_pIDataInit->Release();
      }
   }
   catch(...)
   {
   }
}

STDMETHODIMP CServer::InterfaceSupportsErrorInfo(REFIID riid)
{
	if (InlineIsEqualGUID(IID_IServer,riid))
   {
	   return S_OK;
	}

	return S_FALSE;
}

STDMETHODIMP CServer::get_LocalAddress(IAddress **ppVal)
{
   if (ppVal == 0)
   {
      return Error(L"ppVal is an invalid pointer", GUID_NULL, E_POINTER);
   }

   *ppVal = 0;

   if (!m_pIAddress)
   {
      return Error(L"Server hasn't been initialised - programming error!");
   }

   m_pIAddress->AddRef();

   *ppVal = m_pIAddress;

	return S_OK;
}

STDMETHODIMP CServer::StartListening()
{
   if (!m_pSocketServer)
   {
      return Error(L"Server hasn't been initialised - programming error!");
   }

   try
   {
      m_pSocketServer->StartAcceptingConnections();
   }
   catch(CException &e)
   {
      return ExceptionToError(e);
   }

	return S_OK;
}

STDMETHODIMP CServer::StopListening()
{
   if (!m_pSocketServer)
   {
      return Error( L"Server hasn't been initialised - programming error!");
   }

   try
   {
      m_pSocketServer->StopAcceptingConnections();
   }
   catch(CException &e)
   {
      return ExceptionToError(e);
   }

   return S_OK;
}

STDMETHODIMP CServer::Init(
   /*[in]*/ unsigned long address,
   /*[in]*/ unsigned short port)
{
   // create the socket server
   
   IAddressInit *pInit = 0;

   HRESULT hr = CAddress::CreateInstance(&pInit);

   if (SUCCEEDED(hr))
   {
      hr = pInit->Init(address, port);

      if (SUCCEEDED(hr))
      {
         hr = pInit->QueryInterface(&m_pIAddress);
      }

      pInit->Release();
   }

   if (SUCCEEDED(hr))
   {
      hr = CData::CreateInstance(&m_pIDataInit);

      if (SUCCEEDED(hr))
      {
         hr = m_pIDataInit->QueryInterface(&m_pIData);
      }
   }

   if (SUCCEEDED(hr))
   {
      try
      {
         // now create the socket server

         m_pSocketServer = new CCOMSocketServer(address, port, &m_asyncServerEventHelper);

         m_pSocketServer->Start();
      }
      catch(CException &e)
      {
         return ExceptionToError(e);
      }
   }

   return hr;
}

STDMETHODIMP CServer::OnEvent(
   long eventID)
{
   switch (eventID)
   {
      case 0 :

         return OnConnectionEstablished();

      break;

      case 1 :

         return OnDataRecieved();

      break;

      case 2 :

         return OnConnectionClosed();

      break;
   }

   return Error(L"Unexpected event in OnEvent() - programming error!");
}

HRESULT CServer::OnConnectionEstablished()
{
   CCOMSocketServer::Socket *pSocket = m_pSocketServer->GetSocket();

   // grab the socket, address and port, wrap in an ISocket, store the ISocket in the user data of the socket... 

   long address = m_pSocketServer->GetAddress();
   short port = m_pSocketServer->GetPort();

   // and then fire the event

   ISocketInit *pInit = 0;

   HRESULT hr = CSocket::CreateInstance(&pInit);

   if (SUCCEEDED(hr))
   {
      ISocket *pISocket = 0;
      
      hr = pInit->Init(address, port, pSocket);

      if (SUCCEEDED(hr))
      {
         hr = pInit->QueryInterface(&pISocket);
      }

      if (SUCCEEDED(hr))
      {
         pSocket->SetUserPtr(pISocket);
      }
   
      if (SUCCEEDED(hr))
      {
         Fire_OnConnectionEstablished(pISocket);
      }

      pInit->Release();
   }

   return hr;
}

HRESULT CServer::OnDataRecieved()
{
   if (!m_pSocketServer)
   {
      return Error(L"Server hasn't been initialised - programming error!");
   }

   if (!m_pIDataInit || !m_pIData)
   {
      return Error(L"Internal error: failed to create Data object");
   }
   
   CCOMSocketServer::Socket *pSocket = m_pSocketServer->GetSocket();

   CIOBuffer *pBuffer = m_pSocketServer->GetBuffer();

   if (!pSocket || !pBuffer)
   {
      return Error(L"Internal error: pSocket or pBuffer is 0");
   }

   ISocket *pISocket = reinterpret_cast<ISocket*>(pSocket->GetUserPtr());

   HRESULT hr = m_pIDataInit->Init(pBuffer->GetBuffer(), pBuffer->GetUsed());

   if (SUCCEEDED(hr))
   {  
      Fire_OnDataReceived(pISocket, m_pIData);   // this can stall if the handler doesnt return - 
                                                 // the ATL implementation is difficult to multi thread...
   }

   return hr;
}

HRESULT CServer::OnConnectionClosed()
{
   // retrieve socket 
   CCOMSocketServer::Socket *pSocket = m_pSocketServer->GetSocket();

   if (pSocket)
   {
      ISocket *pISocket = reinterpret_cast<ISocket *>(pSocket->GetUserPtr());

      if (pISocket)
      {
         Fire_OnConnectionClosed(pISocket);

         pISocket->Release();
      }

      pSocket->SetUserPtr(0);

      return S_OK;
   }

   return Error(L"Internal error: failed to obtain ISocket");
}

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// End of file...
///////////////////////////////////////////////////////////////////////////////
