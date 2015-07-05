///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: IOCPWorkerThread.cpp $
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
// $Log: /Web Articles/SocketServers/COMSocketServer2/JetByteTools/Win32Tools/IOCPWorkerThread.cpp $
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

#include "IOCPWorkerThread.h"
#include "IOCompletionPort.h"
#include "Exception.h"
#include "Utils.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// CIOCPWorkerThread
///////////////////////////////////////////////////////////////////////////////

CIOCPWorkerThread::CIOCPWorkerThread(
   CIOCompletionPort &iocp)
   : m_iocp(iocp)
{

}
    
void CIOCPWorkerThread::InitiateShutdown()
{
   m_iocp.PostStatus(0, 0, 0);
}

void CIOCPWorkerThread::WaitForShutdownToComplete()
{
   InitiateShutdown();

   Wait();
}

void CIOCPWorkerThread::Dispatch(
   ULONG_PTR completionKey, 
   DWORD dwNumBytes /*= 0*/, 
   OVERLAPPED *pOverlapped /*= 0*/) 
{
   if (completionKey == 0)
   {
      throw CException(_T("CIOCPWorkerThread::Dispatch()"), _T("0 is an invalid value for completionKey"));
   }

   m_iocp.PostStatus(completionKey, dwNumBytes, pOverlapped); 
}


bool CIOCPWorkerThread::Initialise()
{
   return true;
}

void CIOCPWorkerThread::Shutdown()
{

}

int CIOCPWorkerThread::Run()
{
   try   
   {
      //lint -e{1933} call to unqualified virtual function
      if (Initialise())
      {
         //lint -e{716} while(1)
         while (true)
         {
            ULONG_PTR completionKey;
            DWORD dwNumBytes;
            OVERLAPPED *pOverlapped;
      
            m_iocp.GetStatus(&completionKey, &dwNumBytes, &pOverlapped);

            if (completionKey == 0)
            {
               break; // request to shutdown
            }
            else
            {
               try
               {
                  //lint -e{1933} call to unqualified virtual function
                  Process(completionKey, dwNumBytes, pOverlapped);
               }
               catch(const CException &e)
               {
                  Output(_T("CIOCPWorkerThread::Run() - Exception: ") + e.GetWhere() + _T(" - ") + e.GetMessage());
               }
               catch(...)
               {
                  Output(_T("CIOCPWorkerThread::Run() - Unexpected exception"));
               }
            }
         }
         //lint -e{1933} call to unqualified virtual function
         Shutdown();
      }
   }
   catch(const CException &e)
   {
      Output(_T("CIOCPWorkerThread::Run() - Exception: ") + e.GetWhere() + _T(" - ") + e.GetMessage());
   }
   catch(...)
   {
      Output(_T("CIOCPWorkerThread::Run() - Unexpected exception"));
   }

   return 0;
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
// End of file
///////////////////////////////////////////////////////////////////////////////
