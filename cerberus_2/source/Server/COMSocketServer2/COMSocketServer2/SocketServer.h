#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_COM_SOCKET_SERVER_INCLUDED__
#define JETBYTE_COM_SOCKET_SERVER_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: SocketServer.h $
// Version        : $Revision: 1 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 6/06/02 12:36 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer2/COMSocketServer2/SocketServer.h $
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

#include "JetByteTools\Win32Tools\SocketServer.h"
#include "JetByteTools\Win32Tools\tstring.h"
#include "JetByteTools\Win32Tools\IOCPWorkerThread.h"
#include "COMSocketServer.h"

///////////////////////////////////////////////////////////////////////////////
// Classes defined in other files...
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools
{
   namespace Win32
   {
      class CIOCompletionPort;
   }
}

class CSocketServerWorkerThread;

///////////////////////////////////////////////////////////////////////////////
// CCOMSocketServer
///////////////////////////////////////////////////////////////////////////////

class CCOMSocketServer : public JetByteTools::Win32::CSocketServer
{
   public :

      friend class CSocketServerWorkerThread;

      CCOMSocketServer(
         unsigned long address,
         unsigned short port,
         IUnknown *pUnkSink);

      ~CCOMSocketServer();

      Socket *GetSocket() const;

      JetByteTools::Win32::CIOBuffer *GetBuffer() const;

      unsigned long GetAddress() const;

      unsigned short GetPort() const;

   private :

      struct EventData
      {
            EventData();
            ~EventData();

            Socket *pSocket;
            JetByteTools::Win32::CIOBuffer *pBuffer;

            unsigned long address;
            unsigned short port;
      };

      class COMEventThread : public JetByteTools::Win32::CIOCPWorkerThread
      {
         public :

            COMEventThread(
               EventData &m_data,
               IStream *pMarshalledInterface);

            void WaitForShutdownToComplete();

         private :

            virtual int Run();

            virtual void Process(
               ULONG_PTR completionKey,
               DWORD dwNumBytes,
               OVERLAPPED *pOverlapped);

            void FireReadCompletedEvent(
               Socket *pSocket,
               JetByteTools::Win32::CIOBuffer *pBuffer);

            void FireConnectionEstablishedEvent(
               Socket *pSocket,
               JetByteTools::Win32::CIOBuffer *pBuffer);

            void FireConnectionClosedEvent(
               Socket *pSocket);

            EventData &m_data;
            
            _AsyncServerEvent *m_pEventSink;

            JetByteTools::Win32::CIOCompletionPort m_iocp;

            IStream *m_pMarshalledInterface;

            // no copying, do not implement
            COMEventThread(const COMEventThread &rhs);
            COMEventThread &operator=(const COMEventThread &rhs);
      };

      virtual WorkerThread *CreateWorkerThread(
         JetByteTools::Win32::CIOCompletionPort &iocp);

      virtual SOCKET CreateListeningSocket(
         unsigned long address,
         unsigned short port);

      virtual void OnConnectionEstablished(
         Socket *pSocket,
         JetByteTools::Win32::CIOBuffer *pAddress);

      virtual void OnConnectionClosed(
         Socket *pSocket);

      EventData m_data;

      COMEventThread m_eventThread;

      // no copying, do not implement
      CCOMSocketServer(const CCOMSocketServer &rhs);
      CCOMSocketServer &operator=(const CCOMSocketServer &rhs);
};

#endif // JETBYTE_COM_SOCKET_SERVER_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file...
///////////////////////////////////////////////////////////////////////////////
