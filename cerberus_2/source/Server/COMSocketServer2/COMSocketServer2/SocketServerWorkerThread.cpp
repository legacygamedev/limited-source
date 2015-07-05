///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: SocketServerWorkerThread.cpp $
// Version        : $Revision: 2 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 7/06/02 14:32 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer2/COMSocketServer2/SocketServerWorkerThread.cpp $
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

#include "SocketServerWorkerThread.h"

#include "JetByteTools\COMTools\Utils.h"
#include "JetByteTools\COMTools\Exception.h"

#include "JetByteTools\Win32Tools\Utils.h"
#include "JetByteTools\Win32Tools\IWorkItemProcessor.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
// Member not defined
//lint -esym(1526, CSocketServerWorkerThread::CSocketServerWorkerThread)
//lint -esym(1526, CSocketServerWorkerThread::operator=)
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using JetByteTools::Win32::CIOCompletionPort;
using JetByteTools::Win32::CIOBuffer;
using JetByteTools::Win32::CSocketServer;
using JetByteTools::Win32::IWorkItemProcessor;

///////////////////////////////////////////////////////////////////////////////
// CSocketServerWorkerThread
///////////////////////////////////////////////////////////////////////////////

CSocketServerWorkerThread::CSocketServerWorkerThread(
   CIOCompletionPort &iocp,
   IWorkItemProcessor &workItemProcessor)
   :  CSocketServer::WorkerThread(iocp),
      m_processor(workItemProcessor)
{
}

void CSocketServerWorkerThread::ReadCompleted(
   CSocketServer::Socket *pSocket,
   CIOBuffer *pBuffer)
{
   pSocket->AddRef();
   pBuffer->AddRef();

   m_processor.Dispatch(reinterpret_cast<ULONG_PTR>(pSocket), 0, pBuffer);
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

