///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: AsyncServerEventHelper.cpp $
// Version        : $Revision: 1 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 3/06/02 11:15 $
//
// Notes          : Simple "fake" COM object to get around the weak reference
//                  problem.
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer/JetByteTools/COMTools/AsyncServerEventHelper.cpp $
// 
// 1     3/06/02 11:15 Len
// 
///////////////////////////////////////////////////////////////////////////////
//
// Copyright 2002 JetByte Limited.
//
// JetByte Limited grants you ("Licensee") a non-exclusive, royalty free, 
// licence to use, modify and redistribute this software in source and binary 
// code form, provided that i) this copyright notice and licence appear on all 
// copies of the software; and ii) Licensee does not utilize the software in a 
// manner which is disparaging to JetByte Limited.
//
// This software is provided "AS IS," without a warranty of any kind. ALL
// EXPRESS OR IMPLIED CONDITIONS, REPRESENTATIONS AND WARRANTIES, INCLUDING 
// ANY IMPLIED WARRANTY OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE 
// OR NON-INFRINGEMENT, ARE HEREBY EXCLUDED. JETBYTE LIMITED AND ITS LICENSORS 
// SHALL NOT BE LIABLE FOR ANY DAMAGES SUFFERED BY LICENSEE AS A RESULT OF 
// USING, MODIFYING OR DISTRIBUTING THE SOFTWARE OR ITS DERIVATIVES. IN NO 
// EVENT WILL JETBYTE LIMITED BE LIABLE FOR ANY LOST REVENUE, PROFIT OR DATA, 
// OR FOR DIRECT, INDIRECT, SPECIAL, CONSEQUENTIAL, INCIDENTAL OR PUNITIVE 
// DAMAGES, HOWEVER CAUSED AND REGARDLESS OF THE THEORY OF LIABILITY, ARISING 
// OUT OF THE USE OF OR INABILITY TO USE SOFTWARE, EVEN IF JETBYTE LIMITED 
// HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
//
// This software is not designed or intended for use in on-line control of
// aircraft, air traffic, aircraft navigation or aircraft communications; or in
// the design, construction, operation or maintenance of any nuclear
// facility. Licensee represents and warrants that it will not use or
// redistribute the Software for such purposes.
//
///////////////////////////////////////////////////////////////////////////////

#include "AsyncServerEventHelper.h"

///////////////////////////////////////////////////////////////////////////////
// CAsyncServerEventHelper
///////////////////////////////////////////////////////////////////////////////

CAsyncServerEventHelper::CAsyncServerEventHelper(_AsyncServerEvent &theInterfce)
   :  m_interface(theInterfce)
{
}

STDMETHODIMP CAsyncServerEventHelper::OnEvent(
   long eventID)
{
   return m_interface.OnEvent(eventID);
}

ULONG STDMETHODCALLTYPE CAsyncServerEventHelper::AddRef()
{
   return 2;
}

ULONG STDMETHODCALLTYPE CAsyncServerEventHelper::Release()
{
   return 1;
}

STDMETHODIMP CAsyncServerEventHelper::QueryInterface(REFIID riid, PVOID *ppvObj)
{
   if (riid == IID_IUnknown || riid == IID__AsyncServerEvent)
   {
      *ppvObj = this;
      AddRef();
      return S_OK;
   }
   
   return E_NOINTERFACE;
}

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
