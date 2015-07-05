#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_WIN32_SOCKET_SERVER_INCLUDED__
#define JETBYTE_TOOLS_WIN32_SOCKET_SERVER_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: SocketServer.h $
// Version        : $Revision: 20 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 7/06/02 14:13 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/SimpleProtocolServer2/JetByteTools/Win32Tools/SocketServer.h $
// 
// 20    7/06/02 14:13 Len
// Lint issues.
// 
// 19    5/06/02 19:17 Len
// Abortive socket closure is now done by an IO pool worker thread. This
// is a workaround for a problem with the COM wrapper.
// 
// 18    29/05/02 12:05 Len
// Lint issues.
// 
// 17    26/05/02 15:10 Len
// Factored out common 'user data' code into a mixin base class.
// 
// 16    24/05/02 12:13 Len
// Refactored all the linked list stuff for the sockets into a NodeList
// class.
// 
// 15    21/05/02 11:36 Len
// User data can now be stored/retrieved as either an unsigned long or a
// void *.
// A CIOBuffer containing the client's address is now passed with
// OnConnectionEstablished().
// 
// 14    21/05/02 8:33 Len
// Allow derived class to flush buffer allocator in destructor so that it
// can receive notifications about buffer release.
// 
// 13    21/05/02 8:05 Len
// SocketServer now derives from the buffer allocator.
// 
// 12    20/05/02 23:17 Len
// Updated copyright and disclaimers.
// 
// 11    20/05/02 17:26 Len
// Merged OnNewConnection() into OnConnectionEstablished(). 
// We now pass the socket to OnConnectionClosed() so that the derived
// class can dealocate any per connection user data when the connection is
// closed.
// 
// 10    20/05/02 14:45 Len
// SocketServer doesn't need to pass allocator to WorkerThread.
// 
// 9     20/05/02 14:38 Len
// WorkerThread never needs to use the allocator.
// 
// 8     20/05/02 8:09 Len
// Moved the concept of the io operation used for the io buffer into the
// socket server. The io buffer now simply presents 'user data' access
// functions. Added a similar concept of user data to the socket class so
// that users can associate their own data with a connection . Derived
// class is now notified when a connection occurs so that they can send a
// greeting or request a read, etc. 
// General code cleanup and refactoring.
// 
// 7     16/05/02 21:35 Len
// Users now signal that we're finished with a socket by calling
// Shutdown() rather than Close().
// 
// 6     14/05/02 14:37 Len
// Expose CThread::Start() using a using declaration rather than a
// forwarding function.
// Lint cleanup.
// 
// 5     14/05/02 13:53 Len
// We now explicitly start the thread pool rather than allowing it to
// start itself in the constructor. There was a race condition over the
// completion of construction of derived classes and the first access to
// the pure virtual functions.
// Refactored some of the socket code to improve encapsulation.
// 
// 4     13/05/02 13:44 Len
// Added OnError() methods so that derived class can do something about
// obscure error situations.
// Added a 'max free sockets' concept so that the socket pool can shrink
// as well as grow. This exposed a problem in how we were handling sockets
// - knowing when we can actually delete them was complicated so they're
// now reference counted.
// 
// 3     11/05/02 11:04 Len
// Made CreateListeningSocket() pure virtual as there are an infinte
// number of ways that you can create the listening socket so we'll allow
// the derived class to specify exactly how it's done.
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

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
// Class member is a reference
//lint -esym(1725, CSocketServer::m_allocator)
//lint -esym(1725, Socket::m_server)
//lint -esym(1725, WorkerThread::m_allocator)
//lint -esym(1725, WorkerThread::m_iocp)
//
// Private copy constructor
//lint -esym(1704, CSocketServer::CSocketServer)
//lint -esym(1704, Socket::Socket)
//lint -esym(1704, WorkerThread::WorkerThread)
//
// No default constructor
//lint -esym(1712, CSocketServer)
//lint -esym(1712, Socket)
//lint -esym(1712, WorkerThread)
//
// Base class destructor isnt virtual
//lint -esym(1509, CUsesWinsock)
//
// Data member hides inherited member
//lint -esym(1516, Allocator::m_activeList)
//lint -esym(1516, Allocator::m_freeList)
//
///////////////////////////////////////////////////////////////////////////////

#include "UsesWinsock.h"
#include "Thread.h"
#include "CriticalSection.h"
#include "IOCompletionPort.h"
#include "IOBuffer.h"
#include "ManualResetEvent.h"
#include "NodeList.h"
#include "OpaqueUserData.h"

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// CSocketServer
///////////////////////////////////////////////////////////////////////////////

class CSocketServer : 
   protected CThread, 
   private CUsesWinsock, 
   private CIOBuffer::Allocator
{
   public:

      class Socket;
      class WorkerThread;

      friend class Socket;

      virtual ~CSocketServer();

      using CThread::Start;

      void StartAcceptingConnections();
      void StopAcceptingConnections();

      void InitiateShutdown();

      void WaitForShutdownToComplete();

   protected :

      CSocketServer(
         unsigned long addressToListenOn,
         unsigned short portToListenOn,
         size_t maxFreeSockets,
         size_t maxFreeBuffers,
         size_t bufferSize = 1024,
         size_t numThreads = 0);

      void ReleaseSockets();

      void ReleaseBuffers();

      //lint -e{1768} Virtual function has different access specifier to base class
      virtual int Run();

   private :

      // Override this to create your worker thread

      virtual WorkerThread *CreateWorkerThread(
         CIOCompletionPort &iocp) = 0;

      // Override this to create the listening socket of your choice

      virtual SOCKET CreateListeningSocket(
         unsigned long address,
         unsigned short port) = 0;

      // Interface for derived classes to receive state change notifications...

      virtual void OnStartAcceptingConnections() {}
      virtual void OnStopAcceptingConnections() {}
      virtual void OnShutdownInitiated() {}
      virtual void OnShutdownComplete() {}

      virtual void OnConnectionCreated() {}

      virtual void OnConnectionEstablished(
         Socket *pSocket,
         CIOBuffer *pAddress) = 0;
      
      virtual void OnConnectionClosed(
         Socket * /*pSocket*/) {}

      virtual void OnConnectionDestroyed() {}

      virtual void OnError(
         const _tstring &message);

      virtual void OnBufferCreated() {}
      virtual void OnBufferAllocated() {}
      virtual void OnBufferReleased() {}
      virtual void OnBufferDestroyed() {}

      Socket *AllocateSocket(
         SOCKET theSocket);

      void ReleaseSocket(
         Socket *pSocket);

      void DestroySocket(
         Socket *pSocket);

      void PostAbortiveClose(
         Socket *pSocket);

      void Read(
         Socket *pSocket,
         CIOBuffer *pBuffer);

      void Write(
         Socket *pSocket,
         const char *pData,
         size_t dataLength, 
         bool thenShutdown);

      void Write(
         Socket *pSocket,
         CIOBuffer *pBuffer, 
         bool thenShutdown);

      const size_t m_numThreads;

      CCriticalSection m_listManipulationSection;

      typedef JetByteTools::TNodeList<Socket> SocketList;

      SocketList m_activeList;
      SocketList m_freeList;

      SOCKET m_listeningSocket;

      CIOCompletionPort m_iocp;

      CManualResetEvent m_shutdownEvent;

      CManualResetEvent m_acceptConnectionsEvent;

      const unsigned long m_address;
      const unsigned short m_port;

      const size_t m_maxFreeSockets;

      // No copies do not implement
      CSocketServer(const CSocketServer &rhs);
      CSocketServer &operator=(const CSocketServer &rhs);
};

///////////////////////////////////////////////////////////////////////////////
// CSocketServer::Socket
///////////////////////////////////////////////////////////////////////////////

class CSocketServer::Socket : public CNodeList::Node, public COpaqueUserData
{
   public :

      friend class CSocketServer;
      friend class CSocketServer::WorkerThread;

      void Read(
         CIOBuffer *pBuffer = 0);

      void Write(
         const char *pData, 
         size_t dataLength,
         bool thenShutdown = false);

      void Write(
         CIOBuffer *pBuffer,
         bool thenShutdown = false);

      void AddRef();
      void Release();

      void Shutdown(
         int how = SD_BOTH);

      void Close();

      void AbortiveClose();

   private :

      Socket(
         CSocketServer &server,                                 
         SOCKET socket);

      ~Socket();

      void Attach(
         SOCKET socket);

      CSocketServer &m_server;
      SOCKET m_socket;

      long m_ref;

      // No copies do not implement
      Socket(const Socket &rhs);
      Socket &operator=(const Socket &rhs);
};

///////////////////////////////////////////////////////////////////////////////
// CSocketServer::WorkerThread
///////////////////////////////////////////////////////////////////////////////

class CSocketServer::WorkerThread : public CThread
{
   public :

      virtual ~WorkerThread() {}

      void InitiateShutdown();

      void WaitForShutdownToComplete();

   protected :

      explicit WorkerThread(
         CIOCompletionPort &iocp);

      //lint -e{1768} Virtual function has different access specifier to base class
      virtual int Run();

   private :

      virtual void OnBeginProcessing() {}
      virtual void OnEndProcessing() {}

      virtual void OnError(
         const _tstring &message);

      void Read(
         Socket *pSocket,
         CIOBuffer *pBuffer) const;

      virtual void ReadCompleted(
         Socket *pSocket,
         CIOBuffer *pBuffer) = 0;

      void Write(
         Socket *pSocket,
         CIOBuffer *pBuffer) const;

      virtual void WriteCompleted(
         Socket *pSocket,
         CIOBuffer *pBuffer);

      void AbortiveClose(
         Socket *pSocket);

      CIOCompletionPort &m_iocp;

      // No copies do not implement
      WorkerThread(const WorkerThread &rhs);
      WorkerThread &operator=(const WorkerThread &rhs);
};

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

#endif // JETBYTE_TOOLS_WIN32_SOCKET_SERVER_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////

