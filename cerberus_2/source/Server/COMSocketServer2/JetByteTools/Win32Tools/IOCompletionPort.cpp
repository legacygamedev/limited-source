///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: IOCompletionPort.cpp $
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
// $Log: /Clients/PayPoint/e-Voucher/JetByteTools/Win32Tools/IOCompletionPort.cpp $
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

#include "IOCompletionPort.h"
#include "Win32Exception.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
// Member not defined
//lint -esym(1526, CIOCompletionPort::CIOCompletionPort)
//lint -esym(1526, CIOCompletionPort::operator=)
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// CIOCompletionPort
///////////////////////////////////////////////////////////////////////////////

CIOCompletionPort::CIOCompletionPort(
   size_t maxConcurrency)
   :  m_iocp(::CreateIoCompletionPort(INVALID_HANDLE_VALUE, NULL, 0, maxConcurrency))
{
   if (m_iocp == 0)
   {
      throw CWin32Exception(_T("CIOCompletionPort::CIOCompletionPort() - CreateIoCompletionPort"), ::GetLastError());
   }
}

CIOCompletionPort::~CIOCompletionPort() 
{ 
   ::CloseHandle(m_iocp);
}

void CIOCompletionPort::AssociateDevice(
   HANDLE hDevice, 
   ULONG_PTR completionKey) 
{
   if (m_iocp != ::CreateIoCompletionPort(hDevice, m_iocp, completionKey, 0))
   {
      throw CWin32Exception(_T("CIOCompletionPort::AssociateDevice() - CreateIoCompletionPort"), ::GetLastError());
   }
}

void CIOCompletionPort::PostStatus(
   ULONG_PTR completionKey, 
   DWORD dwNumBytes /* = 0 */, 
   OVERLAPPED *pOverlapped /* = 0 */) 
{
   if (0 == ::PostQueuedCompletionStatus(m_iocp, dwNumBytes, completionKey, pOverlapped))
   {
      throw CWin32Exception(_T("CIOCompletionPort::PostStatus() - PostQueuedCompletionStatus"), ::GetLastError());
   }
}

void CIOCompletionPort::GetStatus(
   ULONG_PTR *pCompletionKey, 
   PDWORD pdwNumBytes,
   OVERLAPPED **ppOverlapped)
{
   if (0 == ::GetQueuedCompletionStatus(m_iocp, pdwNumBytes, pCompletionKey, ppOverlapped, INFINITE))
   {
      throw CWin32Exception(_T("CIOCompletionPort::GetStatus() - GetQueuedCompletionStatus"), ::GetLastError());
   }
}

bool CIOCompletionPort::GetStatus(
   ULONG_PTR *pCompletionKey, 
   PDWORD pdwNumBytes,
   OVERLAPPED **ppOverlapped, 
   DWORD dwMilliseconds)
{
   bool ok = true;

   if (0 == ::GetQueuedCompletionStatus(m_iocp, pdwNumBytes, pCompletionKey, ppOverlapped, dwMilliseconds))
   {
      DWORD lastError = ::GetLastError();

      if (lastError != WAIT_TIMEOUT)
      {
         throw CWin32Exception(_T("CIOCompletionPort::GetStatus() - GetQueuedCompletionStatus"), lastError);
      }

      ok = false;
   }

   return ok;
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
