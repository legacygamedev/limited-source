///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Event.cpp $
// Version        : $Revision: 4 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 20/05/02 23:17 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Clients/PayPoint/e-Voucher/JetByteTools/Win32Tools/Event.cpp $
// 
// 4     20/05/02 23:17 Len
// Updated copyright and disclaimers.
// 
// 3     20/05/02 10:34 Len
// Exposed Pulse functionality
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

#include "Event.h"
#include "Win32Exception.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//lint -esym(1763, CEvent::GetEvent) const member indirectly modifies obj
//
// Member not defined
//lint -esym(1526, CEvent::CEvent)
//lint -esym(1526, CEvent::operator=)
//
//lint -esym(534, CloseHandle)   ignoring return value
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// Static helper methods
///////////////////////////////////////////////////////////////////////////////

static HANDLE Create(
   LPSECURITY_ATTRIBUTES lpEventAttributes, 
   bool bManualReset, 
   bool bInitialState, 
   LPCTSTR lpName);

///////////////////////////////////////////////////////////////////////////////
// CEvent
///////////////////////////////////////////////////////////////////////////////

CEvent::CEvent(
   LPSECURITY_ATTRIBUTES lpEventAttributes, 
   bool bManualReset, 
   bool bInitialState)
   :  m_hEvent(Create(lpEventAttributes, bManualReset, bInitialState, 0))
{

}

CEvent::CEvent(
   LPSECURITY_ATTRIBUTES lpEventAttributes, 
   bool bManualReset, 
   bool bInitialState, 
   const _tstring &name)
   :  m_hEvent(Create(lpEventAttributes, bManualReset, bInitialState, name.c_str()))
{

}

CEvent::~CEvent()
{
   ::CloseHandle(m_hEvent);
}

HANDLE CEvent::GetEvent() const
{
   return m_hEvent;
}

void CEvent::Wait() const
{
   if (!Wait(INFINITE))
   {
      throw CException(_T("CEvent::Wait()"), _T("Unexpected timeout on infinite wait"));
   }
}

bool CEvent::Wait(DWORD timeoutMillis) const
{
   bool ok;

   DWORD result = ::WaitForSingleObject(m_hEvent, timeoutMillis);

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
      throw CWin32Exception(_T("CEvent::Wait() - WaitForSingleObject"), ::GetLastError());
   }
    
   return ok;
}

void CEvent::Reset()
{
   if (!::ResetEvent(m_hEvent))
   {
      throw CWin32Exception(_T("CEvent::Reset()"), ::GetLastError());
   }
}

void CEvent::Set()
{
   if (!::SetEvent(m_hEvent))
   {
      throw CWin32Exception(_T("CEvent::Set()"), ::GetLastError());
   }
}

void CEvent::Pulse()
{
   if (!::PulseEvent(m_hEvent))
   {
      throw CWin32Exception(_T("CEvent::Pulse()"), ::GetLastError());
   }
}

///////////////////////////////////////////////////////////////////////////////
// Static helper methods
///////////////////////////////////////////////////////////////////////////////

static HANDLE Create(
   LPSECURITY_ATTRIBUTES lpEventAttributes, 
   bool bManualReset, 
   bool bInitialState, 
   LPCTSTR lpName)
{
   HANDLE hEvent = ::CreateEvent(lpEventAttributes, bManualReset, bInitialState, lpName);

   if (hEvent == NULL)
   {
      throw CWin32Exception(_T("CEvent::Create()"), ::GetLastError());
   }

   return hEvent;
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
