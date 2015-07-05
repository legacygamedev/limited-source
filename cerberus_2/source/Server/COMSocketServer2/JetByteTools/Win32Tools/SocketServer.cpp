///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: SocketServer.cpp $
// Version        : $Revision: 23 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 7/06/02 14:15 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/SimpleProtocolServer2/JetByteTools/Win32Tools/SocketServer.cpp $
// 
// 23    7/06/02 14:15 Len
// Changes due to change in CIOBuffer. The buffer now derives from
// OVERLAPPED so the explicit conversion functions are no longer required.
// 
// 22    5/06/02 19:17 Len
// Abortive socket closure is now done by an IO pool worker thread. This
// is a workaround for a problem with the COM wrapper.
// 
// 21    29/05/02 12:05 Len
// Lint issues.
// 
// 20    26/05/02 15:10 Len
// Factored out common 'user data' code into a mixin base class.
// 
// 19    24/05/02 12:13 Len
// Refactored all the linked list stuff for the sockets into a NodeList
// class.
// 
// 18    21/05/02 11:36 Len
// User data can now be stored/retrieved as either an unsigned long or a
// void *.
// A CIOBuffer containing the client's address is now passed with
// OnConnectionEstablished().
// 
// 17    21/05/02 8:33 Len
// Allow derived class to flush buffer allocator in destructor so that it
// can receive notifications about buffer release.
// 
// 16    21/05/02 8:05 Len
// SocketServer now derives from the buffer allocator.
// 
// 15    20/05/02 23:17 Len
// Updated copyright and disclaimers.
// 
// 14    20/05/02 17:26 Len
// Merged OnNewConnection() into OnConnectionEstablished(). 
// We now pass the socket to OnConnectionClosed() so that the derived
// class can dealocate any per connection user data when the connection is
// closed.
// 
// 13    20/05/02 14:45 Len
// SocketServer doesn't need to pass allocator to WorkerThread.
// 
// 12    20/05/02 14:38 Len
// WorkerThread never needs to use the allocator.
// 
// 11    20/05/02 8:09 Len
// Moved the concept of the io operation used for the io buffer into the
// socket server. The io buffer now simply presents 'user data' access
// functions. Added a similar concept of user data to the socket class so
// that users can associate their own data with a connection . Derived
// class is now notified when a connection occurs so that they can send a
// greeting or request a read, etc. 
// General code cleanup and refactoring.
// 
// 10    16/05/02 21:35 Len
// Users now signal that we're finished with a socket by calling
// Shutdown() rather than Close().
// 
// 9     15/05/02 11:07 Len
// TX and RX data logging are now wrapped in a DEBUG_ONLY() macro as the
// call to DumpData() was occurring even though the output wasnt being
// logged. This change almost doubled the throughput of the server...
// 
// 8     15/05/02 10:45 Len
// Enabled TX and RX data logging in debug build
// 
// 7     14/05/02 14:37 Len
// Expose CThread::Start() using a using declaration rather than a
// forwarding function.
// Lint cleanup.
// 
// 6     14/05/02 13:53 Len
// We now explicitly start the thread pool rather than allowing it to
// start itself in the constructor. There was a race condition over the
// completion of construction of derived classes and the first access to
// the pure virtual functions.
// Refactored some of the socket code to improve encapsulation.
// 
// 5     13/05/02 13:44 Len
// Added OnError() methods so that derived class can do something about
// obscure error situations.
// Added a 'max free sockets' concept so that the socket pool can shrink
// as well as grow. This exposed a problem in how we were handling sockets
// - knowing when we can actually delete them was complicated so they're
// now reference counted.
// 
// 4     11/05/02 11:05 Len
// Removed CreateListeningSocket() as it's now the responsibility of the
// derived class. General code cleaning.
// 
// 3     10/05/02 19:52 Len
// Bug fix. During the code cleanup we'd renamed most, but not all
// instances of 'socket'... 
// 
// 2     10/05/02 19:25 Len
// Lint options and code cleaning.
// 
// 1     9/05/02 18:47 Len
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

#include "SocketServer.h"
#include "IOCompletionPort.h"
#include "Win32Exception.h"
#include "Utils.h"
#include "SystemInfo.h"

#include <vector>

#pragma comment(lib, "ws2_32.lib")

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
// Symbol did not appear in the constructor initialiser list 
//lint -esym(1928, CThread)
//lint -esym(1928, CUsesWinsock)  
//lint -esym(1928, Node)  
//lint -esym(1928, COpaqueUserData)
//
// Symbol's default constructor implicitly called
//lint -esym(1926, CSocketServer::m_listManipulationSection)
//lint -esym(1926, CSocketServer::m_shutdownEvent)
//lint -esym(1926, CSocketServer::m_acceptConnectionsEvent)
//lint -esym(1926, CSocketServer::m_activeList)
//lint -esym(1926, CSocketServer::m_freeList)
//
// Member not defined
//lint -esym(1526, CSocketServer::CSocketServer)
//lint -esym(1526, CSocketServer::operator=)
//lint -esym(1526, Socket::Socket)
//lint -esym(1526, Socket::operator=)
//lint -esym(1526, WorkerThread::WorkerThread)
//lint -esym(1526, WorkerThread::operator=)
//
//lint -esym(534, InterlockedIncrement)   ignoring return value
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using std::vector;

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// Static helper methods
///////////////////////////////////////////////////////////////////////////////

static size_t CalculateNumberOfThreads(
   size_t numThreads);

///////////////////////////////////////////////////////////////////////////////
// Local enums
///////////////////////////////////////////////////////////////////////////////

enum IO_Operation 
{ 
   IO_Read_Request, 
   IO_Read_Completed, 
   IO_Write_Request, 
   IO_Write_Completed,
   IO_Close
};

///////////////////////////////////////////////////////////////////////////////
// CSocketServer
///////////////////////////////////////////////////////////////////////////////

CSocketServer::CSocketServer(
   unsigned long addressToListenOn,
   unsigned short portToListenOn,
   size_t maxFreeSockets,
   size_t maxFreeBuffers,
   size_t bufferSize /* = 1024 */,
   size_t numThreads /* = 0 */)
   :  CIOBuffer::Allocator(bufferSize, maxFreeBuffers),
      m_numThreads(CalculateNumberOfThreads(numThreads)),
      m_listeningSocket(INVALID_SOCKET),
      m_iocp(0),
      m_address(addressToListenOn),
      m_port(portToListenOn),
      m_maxFreeSockets(maxFreeSockets)
{
}

CSocketServer::~CSocketServer()
{
   try
   {
      ReleaseSockets();
   }
   catch(...)
   {
   }
}

void CSocketServer::ReleaseSockets()
{
   CCriticalSection::Owner lock(m_listManipulationSection);

   Socket *pSocket = m_activeList.Head();

   while (pSocket)
   {
      Socket *pNext = SocketList::Next(pSocket);

      pSocket->Close();
   
      pSocket = pNext;
   }
      
   while (m_activeList.Head())
   {
      ReleaseSocket(m_activeList.Head());
   }

   while (m_freeList.Head())
   {
      DestroySocket(m_freeList.PopNode());
   }

   if (m_freeList.Count() + m_freeList.Count() != 0)
   {
      //lint -e{1933} call to unqualified virtual function
      OnError(_T("CSocketServer::ReleaseSockets() - Leaked sockets"));
   }
}

void CSocketServer::ReleaseBuffers()
{
   Flush();
}

void CSocketServer::StartAcceptingConnections()
{
   if (m_listeningSocket == INVALID_SOCKET)
   {
      //lint -e{1933} call to unqualified virtual function
      OnStartAcceptingConnections();

      //lint -e{1933} call to unqualified virtual function
      m_listeningSocket = CreateListeningSocket(m_address, m_port);
   
      m_acceptConnectionsEvent.Set();
   }
}

void CSocketServer::StopAcceptingConnections()
{
   if (m_listeningSocket != INVALID_SOCKET)
   {
      m_acceptConnectionsEvent.Reset();

      if (0 != ::closesocket(m_listeningSocket))
      {
         //lint -e{1933} call to unqualified virtual function
         OnError(_T("CSocketServer::StopAcceptingConnections() - closesocket - ") + GetLastErrorMessage(::WSAGetLastError()));
      }

      m_listeningSocket = INVALID_SOCKET;

      //lint -e{1933} call to unqualified virtual function
      OnStopAcceptingConnections();
   }
}

void CSocketServer::InitiateShutdown()
{
   // signal that the dispatch thread should shut down all worker threads and then exit

   StopAcceptingConnections();

   m_shutdownEvent.Set();

      //lint -e{1933} call to unqualified virtual function
   OnShutdownInitiated();
}

void CSocketServer::WaitForShutdownToComplete()
{
   // if we havent already started a shut down, do so...

   InitiateShutdown();

   Wait();
}

int CSocketServer::Run()
{
   try
   {
      vector<WorkerThread *> workers;
      
      workers.reserve(m_numThreads);
      
      for (size_t i = 0; i < m_numThreads; ++i)
      {
         //lint -e{1933} call to unqualified virtual function
         WorkerThread *pThread = CreateWorkerThread(m_iocp); 

         workers.push_back(pThread);

         pThread->Start();
      }

      HANDLE handlesToWaitFor[2];

      handlesToWaitFor[0] = m_shutdownEvent.GetEvent();
      handlesToWaitFor[1] = m_acceptConnectionsEvent.GetEvent();

      while (!m_shutdownEvent.Wait(0))
      {
         DWORD waitResult = ::WaitForMultipleObjects(2, handlesToWaitFor, false, INFINITE);

         if (waitResult == WAIT_OBJECT_0)
         {
            // Time to shutdown
            break;
         }
         else if (waitResult == WAIT_OBJECT_0 + 1)
         {
            // accept connections

            while (!m_shutdownEvent.Wait(0) && m_acceptConnectionsEvent.Wait(0))
            {
               CIOBuffer *pAddress = Allocate();

               int addressSize = (int)pAddress->GetSize();

               SOCKET acceptedSocket = ::WSAAccept(
                  m_listeningSocket, 
                  reinterpret_cast<sockaddr*>(const_cast<BYTE*>(pAddress->GetBuffer())), 
                  &addressSize, 
                  0, 
                  0);

               pAddress->Use(addressSize);

               if (acceptedSocket != INVALID_SOCKET)
               {
                  Socket *pSocket = AllocateSocket(acceptedSocket);
               
                  //lint -e{1933} call to unqualified virtual function
                  OnConnectionEstablished(pSocket, pAddress);
               }
               else if (m_acceptConnectionsEvent.Wait(0))
               {
                  //lint -e{1933} call to unqualified virtual function
                  OnError(_T("CSocketServer::Run() - WSAAccept:") + GetLastErrorMessage(::WSAGetLastError()));
               }

               pAddress->Release();
            }
         }
         else
         {
            //lint -e{1933} call to unqualified virtual function
            OnError(_T("CSocketServer::Run() - WaitForMultipleObjects: ") + GetLastErrorMessage(::GetLastError()));
         }
      }

      for (i = 0; i < m_numThreads; ++i)
      {
         workers[i]->InitiateShutdown();
      }  

      for (i = 0; i < m_numThreads; ++i)
      {
         workers[i]->WaitForShutdownToComplete();

         delete workers[i];

         workers[i] = 0;
      }  
   }
   catch(const CException &e)
   {
      //lint -e{1933} call to unqualified virtual function
      OnError(_T("CSocketServer::Run() - Exception: ") + e.GetWhere() + _T(" - ") + e.GetMessage());
   }
   catch(...)
   {
      //lint -e{1933} call to unqualified virtual function
      OnError(_T("CSocketServer::Run() - Unexpected exception"));
   }

   //lint -e{1933} call to unqualified virtual function
   OnShutdownComplete();

   return 0;
}

CSocketServer::Socket *CSocketServer::AllocateSocket(
   SOCKET theSocket)
{
   CCriticalSection::Owner lock(m_listManipulationSection);

   Socket *pSocket = 0;

   if (!m_freeList.Empty())
   {
      pSocket = m_freeList.PopNode();

      pSocket->Attach(theSocket);

      pSocket->AddRef();
   }
   else
   {
      pSocket = new Socket(*this, theSocket);

      //lint -e{1933} call to unqualified virtual function
      OnConnectionCreated();
   }

   m_activeList.PushNode(pSocket);

   //lint -e{611} suspicious cast
   m_iocp.AssociateDevice(reinterpret_cast<HANDLE>(theSocket), (ULONG_PTR)pSocket);

   return pSocket;
}

void CSocketServer::ReleaseSocket(Socket *pSocket)
{
   if (!pSocket)
   {
      throw CException(_T("CSocketServer::ReleaseSocket()"), _T("pSocket is null"));
   }

   CCriticalSection::Owner lock(m_listManipulationSection);

   pSocket->RemoveFromList();

   if (m_maxFreeSockets == 0 || 
       m_freeList.Count() < m_maxFreeSockets)
   {
      m_freeList.PushNode(pSocket);
   }
   else
   {
      DestroySocket(pSocket);
   }
}

void CSocketServer::DestroySocket(
   Socket *pSocket)
{
   delete pSocket;

   //lint -e{1933} call to unqualified virtual function
   OnConnectionDestroyed();
}

void CSocketServer::PostAbortiveClose(
   Socket *pSocket)
{
   CIOBuffer *pBuffer = Allocate();

   pBuffer->SetUserData(IO_Close);

   pSocket->AddRef();

   m_iocp.PostStatus((ULONG_PTR)pSocket, 0, pBuffer);
}


void CSocketServer::Read(
   Socket *pSocket,
   CIOBuffer *pBuffer)
{
   // Post a read request to the iocp so that the actual socket read gets performed by
   // one of our IO threads...

   if (!pBuffer)
   {
      pBuffer = Allocate();
   }
   else
   {
      pBuffer->AddRef();
   }

   pBuffer->SetUserData(IO_Read_Request);

   pSocket->AddRef();

   m_iocp.PostStatus((ULONG_PTR)pSocket, 0, pBuffer);
}

void CSocketServer::Write(
   Socket *pSocket,
   const char *pData,
   size_t dataLength, 
   bool thenShutdown)
{
   // Post a write request to the iocp so that the actual socket write gets performed by
   // one of our IO threads...

   CIOBuffer *pBuffer = Allocate();

   pBuffer->AddData(pData, dataLength);

   pBuffer->SetUserData(IO_Write_Request);

   pSocket->AddRef();

   m_iocp.PostStatus((ULONG_PTR)pSocket, thenShutdown, pBuffer);
}

void CSocketServer::Write(
   Socket *pSocket,
   CIOBuffer *pBuffer, 
   bool thenShutdown)
{
   // Post a write request to the iocp so that the actual socket write gets performed by
   // one of our IO threads...

   pBuffer->AddRef();

   pBuffer->SetUserData(IO_Write_Request);

   pSocket->AddRef();

   m_iocp.PostStatus((ULONG_PTR)pSocket, thenShutdown, pBuffer);
}

void CSocketServer::OnError(
   const _tstring &message)
{
   Output(message);
}
  
///////////////////////////////////////////////////////////////////////////////
// CSocketServer::Socket
///////////////////////////////////////////////////////////////////////////////

CSocketServer::Socket::Socket(
   CSocketServer &server,                                 
   SOCKET theSocket)
   :  m_server(server),
      m_socket(theSocket),
      m_ref(1)
{
   if (INVALID_SOCKET == m_socket)
   {
      throw CException(_T("CSocketServer::Socket::Socket()"), _T("Invalid socket"));
   }
}

CSocketServer::Socket::~Socket()
{
}

void CSocketServer::Socket::Attach(
   SOCKET theSocket)
{
   if (INVALID_SOCKET != m_socket)
   {
      throw CException(_T("CSocketServer::Socket::Attach()"), _T("Socket already attached"));
   }

   m_socket = theSocket;

   SetUserData(0);
}

void CSocketServer::Socket::AddRef()
{
   ::InterlockedIncrement(&m_ref);
}

void CSocketServer::Socket::Release()
{
   if (0 == ::InterlockedDecrement(&m_ref))
   {
      m_server.ReleaseSocket(this);
   }
}

void CSocketServer::Socket::Shutdown(
   int how /* = SD_BOTH */)
{
   Output(_T("CSocketServer::Socket::Shutdown() ") + ToString(how));

   if (INVALID_SOCKET != m_socket)
   {
      if (0 != ::shutdown(m_socket, how))
      {
         m_server.OnError(_T("CSocketServer::Server::Shutdown() - ") + GetLastErrorMessage(::WSAGetLastError()));
      }

      Output(_T("shutdown initiated"));
   }
}

void CSocketServer::Socket::Close()
{
   CCriticalSection::Owner lock(m_server.m_listManipulationSection);

   if (INVALID_SOCKET != m_socket)
   {
      if (0 != ::closesocket(m_socket))
      {
         m_server.OnError(_T("CSocketServer::Socket::Close() - closesocket - ") + GetLastErrorMessage(::WSAGetLastError()));
      }

      m_socket = INVALID_SOCKET;

      m_server.OnConnectionClosed(this);

      Release();
   }
}

void CSocketServer::Socket::AbortiveClose()
{
   m_server.PostAbortiveClose(this);
}

void CSocketServer::Socket::Read(
   CIOBuffer *pBuffer /* = 0 */)
{
   m_server.Read(this, pBuffer);
}

void CSocketServer::Socket::Write(
   const char *pData, 
   size_t dataLength,
   bool thenShutdown /* = false */)
{
   m_server.Write(this, pData, dataLength, thenShutdown);
}

void CSocketServer::Socket::Write(
   CIOBuffer *pBuffer,
   bool thenShutdown /* = false */)
{
   m_server.Write(this, pBuffer, thenShutdown);
}

///////////////////////////////////////////////////////////////////////////////
// CSocketServer::WorkerThread
///////////////////////////////////////////////////////////////////////////////

CSocketServer::WorkerThread::WorkerThread(
   CIOCompletionPort &iocp)
   :  m_iocp(iocp)
{
   // All work done in initialiser list
}

int CSocketServer::WorkerThread::Run()
{
   try
   {
      //lint -e{716} while(1)
      while (true)   
      {
         // continually loop to service io completion packets

         bool closeSocket = false;

         DWORD dwIoSize = 0;
         Socket *pSocket = 0;
         CIOBuffer *pBuffer = 0;
         
         try
         {
            m_iocp.GetStatus((PDWORD_PTR)&pSocket, &dwIoSize, (OVERLAPPED**)&pBuffer);
         }
         catch (const CWin32Exception &e)
         {
            if (e.GetError() != ERROR_NETNAME_DELETED &&
                e.GetError() != WSA_OPERATION_ABORTED)
            {
               throw;
            }
            
            Output(_T("IOCP error - client connection dropped"));

            closeSocket = true;
         }

         if (!pSocket)
         {
            // A completion key of 0 is posted to the iocp to request us to shut down...

            break;
         }

         //lint -e{1933} call to unqualified virtual function
         OnBeginProcessing();

         if (pBuffer)
         {
            const IO_Operation operation = static_cast<IO_Operation>(pBuffer->GetUserData());

            switch (operation)
            {
               case IO_Read_Request :

                  Read(pSocket, pBuffer);
               
               break;
         
               case IO_Read_Completed :

                  if (0 != dwIoSize)
                  {
                     pBuffer->Use(dwIoSize);
                  
                     DEBUG_ONLY(Output(_T("RX: ") + ToString(pBuffer) + _T("\n") + DumpData(reinterpret_cast<const BYTE*>(pBuffer->GetWSABUF()->buf), dwIoSize, 40)));

                     //lint -e{1933} call to unqualified virtual function
                     ReadCompleted(pSocket, pBuffer);
                  }
                  else
                  {
                     // client connection dropped...

                     Output(_T("ReadCompleted - 0 bytes - client connection dropped"));

                     closeSocket = true;
                  }

                  pSocket->Release();
                  pBuffer->Release();

               break;

               case IO_Write_Request :

                  Write(pSocket, pBuffer);

                  if (dwIoSize != 0)
                  {
                     // final write, now shutdown send side of connection
                     pSocket->Shutdown(SD_SEND);
                  }

               break;

               case IO_Write_Completed :

                  pBuffer->Use(dwIoSize);
               
                  DEBUG_ONLY(Output(_T("TX: ") + ToString(pBuffer) + _T("\n") + DumpData(reinterpret_cast<const BYTE*>(pBuffer->GetWSABUF()->buf), dwIoSize, 40)));

                  //lint -e{1933} call to unqualified virtual function
                  WriteCompleted(pSocket, pBuffer);

                  pSocket->Release();
                  pBuffer->Release();

               break;

               case IO_Close :

                  AbortiveClose(pSocket);
            
                  pSocket->Release();
                  pBuffer->Release();

               break;

               default :
                  //lint -e{1933} call to unqualified virtual function
                  OnError(_T("CSocketServer::WorkerThread::Run() - Unexpected operation"));
               break;
            } 
         }
         else
         {
            //lint -e{1933} call to unqualified virtual function
            OnError(_T("CSocketServer::WorkerThread::Run() - Unexpected - pBuffer is 0"));
         }

         if (closeSocket)
         {
            pSocket->Close();
         }

         //lint -e{1933} call to unqualified virtual function
         OnEndProcessing();
      } 
   }
   catch(const CException &e)
   {
      //lint -e{1933} call to unqualified virtual function
      OnError(_T("CSocketServer::WorkerThread::Run() - Exception: ") + e.GetWhere() + _T(" - ") + e.GetMessage());
   }
   catch(...)
   {
      //lint -e{1933} call to unqualified virtual function
      OnError(_T("CSocketServer::WorkerThread::Run() - Unexpected exception"));
   }

   return 0;
}

void CSocketServer::WorkerThread::InitiateShutdown()
{
   m_iocp.PostStatus(0);         
}

void CSocketServer::WorkerThread::WaitForShutdownToComplete()
{
   // if we havent already started a shut down, do so...

   InitiateShutdown();

   Wait();
}

void CSocketServer::WorkerThread::Read(
   Socket *pSocket,
   CIOBuffer *pBuffer) const
{
   pBuffer->SetUserData(IO_Read_Completed);

   pBuffer->SetupRead();

   DWORD dwNumBytes = 0;
   DWORD dwFlags = 0;

   if (SOCKET_ERROR == ::WSARecv(
      pSocket->m_socket, 
      pBuffer->GetWSABUF(), 
      1, 
      &dwNumBytes,
      &dwFlags,
      pBuffer, 
      NULL))
   {
      DWORD lastError = ::WSAGetLastError();

      if (ERROR_IO_PENDING != lastError)
      {
         Output(_T("CSocketServer::Read() - WSARecv: ") + GetLastErrorMessage(lastError));

         if (lastError == WSAECONNABORTED || 
             lastError == WSAECONNRESET ||
             lastError == WSAEDISCON)
         {
            pSocket->Close();
         }

         pSocket->Release();
         pBuffer->Release();
      }
   }
}

void CSocketServer::WorkerThread::Write(
   Socket *pSocket,
   CIOBuffer *pBuffer) const
{
   pBuffer->SetUserData(IO_Write_Completed);

   pBuffer->SetupWrite();

   DWORD dwFlags = 0;
   DWORD dwSendNumBytes = 0;

   if (SOCKET_ERROR == ::WSASend(
      pSocket->m_socket,
      pBuffer->GetWSABUF(), 
      1, 
      &dwSendNumBytes,
      dwFlags,
      pBuffer, 
      NULL))
   {
      DWORD lastError = ::WSAGetLastError();

      if (ERROR_IO_PENDING != lastError)
      {
         Output(_T("CSocketServer::Write() - WSASend: ") + GetLastErrorMessage(lastError));

         if (lastError == WSAECONNABORTED || 
             lastError == WSAECONNRESET ||
             lastError == WSAEDISCON)
         {
            pSocket->Close();
         }

         pSocket->Release();
         pBuffer->Release();
      }
   }
}

void CSocketServer::WorkerThread::WriteCompleted(
   Socket * /*pSocket*/,
   CIOBuffer *pBuffer)
{
   if (pBuffer->GetUsed() != pBuffer->GetWSABUF()->len)
   {
      //lint -e{1933} call to unqualified virtual function
      OnError(_T("CSocketServer::WorkerThread::WriteCompleted - Socket write where not all data was written"));
   }

   //lint -e{818} pointer pBuffer could be declared const (but not in derived classes...)
}

void CSocketServer::WorkerThread::AbortiveClose(
   Socket *pSocket)
{
   // Force an abortive close.

   LINGER lingerStruct;

   lingerStruct.l_onoff = 1;
   lingerStruct.l_linger = 0;

   if (SOCKET_ERROR == ::setsockopt(pSocket->m_socket, SOL_SOCKET, SO_LINGER, (char *)&lingerStruct, sizeof(lingerStruct)))
   {
      //lint -e{1933} call to unqualified virtual function
      OnError(_T("CSocketServer::Socket::AbortiveClose() - setsockopt(SO_LINGER) - ")  + GetLastErrorMessage(::WSAGetLastError()));
   }

   pSocket->Close();
}

void CSocketServer::WorkerThread::OnError(
   const _tstring &message)
{
   Output(message);
}
                                          
///////////////////////////////////////////////////////////////////////////////
// Static helper methods
///////////////////////////////////////////////////////////////////////////////

static size_t CalculateNumberOfThreads(size_t numThreads)
{
   if (numThreads == 0)
   {
      CSystemInfo systemInfo;
   
      numThreads = systemInfo.dwNumberOfProcessors * 2;
   }

   return numThreads;
}

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

} // End of namespace Win32
} // End of namespace JetByteTools 

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// End of file...
///////////////////////////////////////////////////////////////////////////////
