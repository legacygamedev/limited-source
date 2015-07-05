///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Address.cpp $
// Version        : $Revision: 1 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 3/06/02 11:37 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/Address.cpp $
// 
// 1     3/06/02 11:37 Len
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

#include "stdafx.h"
#include "COMSocketServer.h"
#include "Address.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// CAddress
///////////////////////////////////////////////////////////////////////////////

CAddress::CAddress()
   :  m_port(0),
      m_address(L"000.000.000.000")
{
}

STDMETHODIMP CAddress::InterfaceSupportsErrorInfo(REFIID riid)
{
   if (InlineIsEqualGUID(IID_IAddress,riid))
   {
      return S_OK;
	}
	
   return S_FALSE;
}

STDMETHODIMP CAddress::get_Address(BSTR *pVal)
{
   HRESULT hr = S_OK;

   if (pVal == 0)
   {
      return Error(L"pVal is an invalid pointer", GUID_NULL, E_POINTER);
   }

   *pVal = m_address.Copy();

	return hr;
}

STDMETHODIMP CAddress::get_Port(long *pVal)
{
   HRESULT hr = S_OK;

   if (pVal == 0)
   {
      return Error(L"pVal is an invalid pointer", GUID_NULL, E_POINTER);
   }

   *pVal = m_port;

	return hr;
}

STDMETHODIMP CAddress::Init(
   /*[in]*/ unsigned long address,
   /*[in]*/ unsigned short port)
{
   USES_CONVERSION;

   m_port = port;

   in_addr addr;

   addr.S_un.S_addr = address;

   m_address = A2OLE(::inet_ntoa(addr));

   return S_OK;
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
