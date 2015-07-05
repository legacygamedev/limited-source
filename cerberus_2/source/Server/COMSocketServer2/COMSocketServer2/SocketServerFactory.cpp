///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: SocketServerFactory.cpp $
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
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/SocketServerFactory.cpp $
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
#include "SocketServerFactory.h"
#include "JetByteTools\COMTools\Utils.h"

#include "Server.h"

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using JetByteTools::COM::GetOptionalString;
using JetByteTools::Win32::_tstring;

///////////////////////////////////////////////////////////////////////////////
// CSocketServerFactory
///////////////////////////////////////////////////////////////////////////////

CSocketServerFactory::CSocketServerFactory()
{
}

STDMETHODIMP CSocketServerFactory::InterfaceSupportsErrorInfo(REFIID riid)
{
   if (InlineIsEqualGUID(IID_ISocketServerFactory,riid))
   {
      return S_OK;
	}

	return S_FALSE;
}

STDMETHODIMP CSocketServerFactory::CreateSocketServer(
   long port, 
   VARIANT address,
   IServer **ppServer)
{
   if (ppServer == 0)
   {
      return Error(L"ppServer is an invalid pointer", GUID_NULL, E_POINTER);
   }

   *ppServer = 0;

   // todo deal with optional address in dotted ip format or long format

   // todo, check port will fit in a ushort

   std::string addressAsString;

   HRESULT hr = GetOptionalString(address, addressAsString, "000.000.000.000");

   if (SUCCEEDED(hr))
   {
      long addressAsLong = ::inet_addr(addressAsString.c_str());

      IAddressInit *pInit = 0;

      hr = CServer::CreateInstance(&pInit);

      if (SUCCEEDED(hr))
      {
         hr = pInit->Init(addressAsLong, static_cast<unsigned short>(port));

         if (SUCCEEDED(hr))
         {
            hr = pInit->QueryInterface(ppServer);
         }

         pInit->Release();
      }
   }
   else
   {
      return Error(L"address should be the IP address of an interface in this machine in dotted format", GUID_NULL, E_INVALIDARG);
   }

   return hr;
}

///////////////////////////////////////////////////////////////////////////////
// End of file...
///////////////////////////////////////////////////////////////////////////////
