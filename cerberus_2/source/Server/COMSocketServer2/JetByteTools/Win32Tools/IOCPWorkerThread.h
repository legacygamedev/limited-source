#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_WIN32_IOCP_WORKER_THREAD_INCLUDED__
#define JETBYTE_TOOLS_WIN32_IOCP_WORKER_THREAD_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: IOCPWorkerThread.h $
// Version        : $Revision: 1 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 6/06/02 13:02 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer2/JetByteTools/Win32Tools/IOCPWorkerThread.h $
// 
// 1     6/06/02 13:02 Len
// 
// 1     6/06/02 12:59 Len
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

#include "Thread.h"
#include "IWorkItemProcessor.h"

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

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// CIOCPWorkerThread
///////////////////////////////////////////////////////////////////////////////

class CIOCPWorkerThread : 
   protected CThread,
   public IWorkItemProcessor
{
   public :

      using CThread::Start;

      virtual void Dispatch(
         ULONG_PTR completionKey, 
         DWORD dwNumBytes = 0, 
         OVERLAPPED *pOverlapped = 0);

      void InitiateShutdown();

      void WaitForShutdownToComplete();

   protected :
   
      CIOCPWorkerThread(
         CIOCompletionPort &iocp);
      
      virtual int Run();

   private :

      virtual bool Initialise();

      virtual void Process(
         ULONG_PTR completionKey,
         DWORD dwNumBytes,
         OVERLAPPED *pOverlapped) = 0;

      virtual void Shutdown();

      CIOCompletionPort &m_iocp;

      // No copies, do not implement
      CIOCPWorkerThread(const CIOCPWorkerThread &rhs);
      CIOCPWorkerThread &operator=(CIOCPWorkerThread &rhs);
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

#endif // JETBYTE_TOOLS_WIN32_IOCP_WORKER_THREAD_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
