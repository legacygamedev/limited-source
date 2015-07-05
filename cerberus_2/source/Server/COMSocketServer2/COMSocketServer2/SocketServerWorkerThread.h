#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_SOCKET_SERVER_WORKER_THREAD_INCLUDED__
#define JETBYTE_SOCKET_SERVER_WORKER_THREAD_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: SocketServerWorkerThread.h $
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
// $Log: /Web Articles/SocketServers/COMSocketServer2/COMSocketServer2/SocketServerWorkerThread.h $
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
// Lint options
//
//lint -save
//
// Private constructor
//lint -esym(1704, CSocketServerWorkerThread::CSocketServerWorkerThread)
//
// No default constructor
//lint -esym(1712, CSocketServerWorkerThread)
//
///////////////////////////////////////////////////////////////////////////////

#include "JetByteTools\Win32Tools\SocketServer.h"

///////////////////////////////////////////////////////////////////////////////
// Classes defined in other files...
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools
{
   namespace Win32
   {
      class CIOBuffer;
      class IWorkItemProcessor;
   }
}

///////////////////////////////////////////////////////////////////////////////
// CSocketServerWorkerThread
///////////////////////////////////////////////////////////////////////////////

class CSocketServerWorkerThread : 
   public JetByteTools::Win32::CSocketServer::WorkerThread
{
   public :

      explicit CSocketServerWorkerThread(
         JetByteTools::Win32::CIOCompletionPort &iocp,
         JetByteTools::Win32::IWorkItemProcessor &workItemProcessor);

   private :

      virtual void ReadCompleted(
         JetByteTools::Win32::CSocketServer::Socket *pSocket,
         JetByteTools::Win32::CIOBuffer *pBuffer);

      JetByteTools::Win32::IWorkItemProcessor &m_processor;

      // No copies do not implement
      CSocketServerWorkerThread(const CSocketServerWorkerThread &rhs);            
      CSocketServerWorkerThread &operator=(const CSocketServerWorkerThread &rhs); 
};

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_SOCKET_SERVER_WORKER_THREAD_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
