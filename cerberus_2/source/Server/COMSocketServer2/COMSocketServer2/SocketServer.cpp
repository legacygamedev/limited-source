///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: SocketServer.cpp $
// Version        : $Revision: 3 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 18/06/02 18:45 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer2/COMSocketServer2/SocketServer.cpp $
// 
// 3     18/06/02 18:45 Len
// Removed ReuseAddress() as it's not required and it's an error to set it
// on the listening socket - you shouldn't need to and if you do it's more
// than likely a bug somewhere!
// 
// 2     7/06/02 14:32 Len
// Changes due to change in CIOBuffer. The buffer now derives from
// OVERLAPPED so the explicit conversion functions are no longer required.
// 
// 1     6/06/02 12:36 Len
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

#include "SocketServer.h"
#include "SocketServerWorkerThread.h"

#include "JetByteTools\Win32Tools\Win32Exception.h"
#include "JetByteTools\Win32Tools\Socket.h"
#include "JetByteTools\Win32Tools\Utils.h"
#include "JetByteTools\COMTools\Utils.h"
#include "JetByteTools\COMTools\UsesCom.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using JetByteTools::Win32::CIOCompletionPort;
using JetByteTools::Win32::CIOBuffer;
using JetByteTools::Win32::CCriticalSection;
using JetByteTools::Win32::CWin32Exception;
using JetByteTools::Win32::CSocket;
using JetByteTools::Win32::CIOCPWorkerThread;
using JetByteTools::Win32::Output;
using JetByteTools::Win32::ToString;
using JetByteTools::Win32::CException;

using JetByteTools::COM::CUsesCOM;
using JetByteTools::COM::WaitWithMessageLoop;
using JetByteTools::COM::MarshalInterThreadInterfaceInStream;

///////////////////////////////////////////////////////////////////////////////
// CCOMSocketServer
///////////////////////////////////////////////////////////////////////////////

CCOMSocketServer::CCOMSocketServer(
   unsigned long address,
   unsigned short port,
   IUnknown *pUnkSink)
   :  JetByteTools::Win32::CSocketServer(address, port, 10, 10, 1024, 0),
      m_eventThread(m_data, MarshalInterThreadInterfaceInStream(pUnkSink, IID__AsyncServerEvent))
{
   m_eventThread.Start();
}

CCOMSocketServer::~CCOMSocketServer()
{
   // If we want to be informed of any buffers or sockets being destroyed at destruction 
   // time then we need to release these resources now, whilst we, the derived class,
   // still exists. Once our destructor exits and the base destructor takes over we wont 
   // get any more notifications...

   try
   {
      m_eventThread.WaitForShutdownToComplete();

      ReleaseSockets();
      ReleaseBuffers();
   }
   catch(...)
   {
   }
}

SOCKET CCOMSocketServer::CreateListeningSocket(
   unsigned long address,
   unsigned short port)
{
   SOCKET s = ::WSASocket(AF_INET, SOCK_STREAM, IPPROTO_IP, NULL, 0, WSA_FLAG_OVERLAPPED); 

   if (s == INVALID_SOCKET)
   {
      throw CWin32Exception(_T("CCOMSocketServer::CreateListeningSocket()"), ::WSAGetLastError());
   }

   CSocket listeningSocket(s);

   CSocket::InternetAddress localAddress(address, port);

   listeningSocket.Bind(localAddress);

   listeningSocket.Listen(5);

   return listeningSocket.Detatch();
}

JetByteTools::Win32::CSocketServer::WorkerThread *CCOMSocketServer::CreateWorkerThread(
   CIOCompletionPort &iocp)
{
   return new CSocketServerWorkerThread(iocp, m_eventThread);
}

void CCOMSocketServer::OnConnectionEstablished(
   Socket *pSocket,
   CIOBuffer *pBuffer)
{
   pSocket->AddRef();
   pBuffer->AddRef();

   m_eventThread.Dispatch(reinterpret_cast<ULONG_PTR>(pSocket), 1, pBuffer);
}

void CCOMSocketServer::OnConnectionClosed(
   Socket *pSocket)
{
   pSocket->AddRef();

   m_eventThread.Dispatch(reinterpret_cast<ULONG_PTR>(pSocket), 2, 0);
}

JetByteTools::Win32::CSocketServer::Socket *CCOMSocketServer::GetSocket() const
{
   return m_data.pSocket;
}

JetByteTools::Win32::CIOBuffer *CCOMSocketServer::GetBuffer() const
{
   return m_data.pBuffer;
}

unsigned long CCOMSocketServer::GetAddress() const
{
   return m_data.address;
}

unsigned short CCOMSocketServer::GetPort() const
{
   return m_data.port;
}

///////////////////////////////////////////////////////////////////////////////
// CCOMSocketServer::EventData
///////////////////////////////////////////////////////////////////////////////

CCOMSocketServer::EventData::EventData()
   :  pSocket(0),
      pBuffer(0),
      port(0),
      address(0)
{

}

CCOMSocketServer::EventData::~EventData()
{
   pSocket = 0;
   pBuffer = 0;
}

///////////////////////////////////////////////////////////////////////////////
// CCOMSocketServer::COMEventThread
///////////////////////////////////////////////////////////////////////////////

CCOMSocketServer::COMEventThread::COMEventThread(
   EventData &data,
   IStream *pMarshalledInterface)
   :  CIOCPWorkerThread(m_iocp),
      m_data(data),
      m_iocp(0),
      m_pMarshalledInterface(pMarshalledInterface)
{

}

void CCOMSocketServer::COMEventThread::WaitForShutdownToComplete()
{
   InitiateShutdown();

   WaitWithMessageLoop(GetHandle());
}

int CCOMSocketServer::COMEventThread::Run()
{
   try
   {
      CUsesCOM usesCOM(COINIT_APARTMENTTHREADED);

      HRESULT hr = ::CoGetInterfaceAndReleaseStream(m_pMarshalledInterface, IID__AsyncServerEvent, (void**)&m_pEventSink);

      if (FAILED(hr))
      {
         return 0;
      }

      int result = CIOCPWorkerThread::Run();

      m_pEventSink->Release();

      m_pEventSink = 0;

      return result;
   }
   catch(const CException &e)
   {
      Output(_T("COMEventThread::Run() - Exception - ") + e.GetWhere() + _T(" - ") + e.GetMessage());
   }
   catch(...)
   {
      Output(_T("COMEventThread::Run() - Unexpected exception"));
   }

   return 0;
}

void CCOMSocketServer::COMEventThread::Process(
   ULONG_PTR completionKey,
   DWORD operation,
   OVERLAPPED *pOverlapped)
{
   Socket *pSocket = reinterpret_cast<CCOMSocketServer::Socket *>(completionKey);
   CIOBuffer *pBuffer = static_cast<CIOBuffer *>(pOverlapped);

   try
   {
      switch (operation)
      {
         case 0 :    // ReadCompleted
 
            FireReadCompletedEvent(pSocket, pBuffer);
 
         break;

         case 1 :    // OnConnectionEstablished

            FireConnectionEstablishedEvent(pSocket, pBuffer);
 
         break;

         case 2 :    // OnConnectionClosed

            FireConnectionClosedEvent(pSocket);

         break;

         default :

            Output(_T("COMEventThread::Process() - unexpected operation: ") + ToString(operation));

         break;
      }
   }
   catch(...)
   {
      pSocket->Shutdown();
   }

   pSocket->Release();

   if (pBuffer)
   {
      pBuffer->Release();
   }
}

void CCOMSocketServer::COMEventThread::FireReadCompletedEvent(
   Socket *pSocket,
   CIOBuffer *pBuffer)
{
   m_data.pSocket = pSocket;
   m_data.pBuffer = pBuffer;

   m_pEventSink->OnEvent(1);

   m_data.pSocket = 0;
   m_data.pBuffer = 0;
}

void CCOMSocketServer::COMEventThread::FireConnectionEstablishedEvent(
   Socket *pSocket,
   CIOBuffer *pAddress)
{
   const sockaddr_in *pSockAddrIn = reinterpret_cast<const sockaddr_in *>(pAddress->GetBuffer());

   if (pSockAddrIn->sin_family == AF_INET)
   {
      m_data.address = pSockAddrIn->sin_addr.S_un.S_addr;
      m_data.port = pSockAddrIn->sin_port;
   }
   
   m_data.pSocket = pSocket;
   m_data.pBuffer = 0;

   m_pEventSink->OnEvent(0);

   m_data.address = 0;
   m_data.port = 0;
   m_data.pSocket = 0;
}

void CCOMSocketServer::COMEventThread::FireConnectionClosedEvent(
   Socket *pSocket)
{
   m_data.pSocket = pSocket;
   m_data.pBuffer = 0;

   m_pEventSink->OnEvent(2);

   m_data.pSocket = 0;
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
