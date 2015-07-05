///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Thread.cpp $
// Version        : $Revision: 3 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 20/05/02 23:17 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Clients/PayPoint/e-Voucher/JetByteTools/Win32Tools/Thread.cpp $
// 
// 3     20/05/02 23:17 Len
// Updated copyright and disclaimers.
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

#include "Thread.h"

#include <process.h>

#include "Win32Exception.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//lint -esym(1763, CThread::GetHandle) const member indirectly modifies obj
//
// Member not defined
//lint -esym(1526, CThread::CThread)
//lint -esym(1526, CThread::operator=)
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// CThread
///////////////////////////////////////////////////////////////////////////////

CThread::CThread()
   :  m_hThread(INVALID_HANDLE_VALUE)
{

}
      
CThread::~CThread()
{
   if (m_hThread != INVALID_HANDLE_VALUE)
   {
      ::CloseHandle(m_hThread);
   }

   //lint -e{1540} m_hThread neither freed nor zeroed in destructor
}

HANDLE CThread::GetHandle() const
{
   return m_hThread;
}

void CThread::Start()
{
   if (m_hThread == INVALID_HANDLE_VALUE)
   {
      unsigned int threadID = 0;

      m_hThread = (HANDLE)::_beginthreadex(0, 0, ThreadFunction, (void*)this, 0, &threadID);

      if (m_hThread == INVALID_HANDLE_VALUE)
      {
         throw CWin32Exception(_T("CThread::Start() - _beginthreadex"), GetLastError());
      }
   }
   else
   {
      throw CException(_T("CThread::Start()"), _T("Thread already running - you can only call Start() once!"));
   }
}

void CThread::Wait() const
{
   if (!Wait(INFINITE))
   {
      throw CException(_T("CThread::Wait()"), _T("Unexpected timeout on infinite wait"));
   }
}

bool CThread::Wait(DWORD timeoutMillis) const
{
   // TODO base class? Waitable?
   bool ok;

   DWORD result = ::WaitForSingleObject(m_hThread, timeoutMillis);

   if (result == WAIT_TIMEOUT)
   {
      ok = false;
   }
   else if (result == WAIT_OBJECT_0)
   {
      ok = true;
   }
   else
   {
      throw CWin32Exception(_T("CThread::Wait() - WaitForSingleObject"), ::GetLastError());
   }
    
   return ok;
}

unsigned int __stdcall CThread::ThreadFunction(void *pV)
{
   int result = 0;

   CThread* pThis = (CThread*)pV;
   
   if (pThis)
   {
      try
      {
         result = pThis->Run();
      }
      catch(...)
      {
      }
   }

   return result;
}

void CThread::Terminate(
   DWORD exitCode /* = 0 */)
{
   if (!::TerminateThread(m_hThread, exitCode))
   {
      // TODO we could throw an exception here...
   }
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

