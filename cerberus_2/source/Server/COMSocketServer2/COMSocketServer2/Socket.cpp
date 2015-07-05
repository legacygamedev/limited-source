///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Socket.cpp $
// Version        : $Revision: 3 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 11/06/02 15:44 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer2/COMSocketServer2/Socket.cpp $
// 
// 3     11/06/02 15:44 Len
// Allow Error() to do the T2OLE conversion as VC7 doesn't like us using
// aloca in a catch block.
// 
// 2     8/06/02 12:00 Len
// Implemented WriteString
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
#include "Socket.h"

#include "Address.h"

#include "JetByteTools\COMTools\Utils.h"
#include "JetByteTools\Win32Tools\Exception.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using JetByteTools::Win32::CException;
using JetByteTools::COM::SafeRelease;
using JetByteTools::COM::GetOptionalBool;

///////////////////////////////////////////////////////////////////////////////
// CSocket
///////////////////////////////////////////////////////////////////////////////

CSocket::CSocket()
   :  m_pSocket(0),
      m_pAddress(0)
{

}

CSocket::~CSocket()
{
   m_pSocket = 0;
   m_pAddress = SafeRelease(m_pAddress);
}

STDMETHODIMP CSocket::InterfaceSupportsErrorInfo(REFIID riid)
{
	if (InlineIsEqualGUID(IID_ISocket,riid))
   {
      return S_OK;
	}

   return S_FALSE;
}

STDMETHODIMP CSocket::WriteString(
   BSTR data,
   VARIANT sendAsUNICODE,
   VARIANT thenShutdown)
{
   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!", GUID_NULL, E_UNEXPECTED);
   }

   bool unicodeString;

   HRESULT hr = GetOptionalBool(sendAsUNICODE, unicodeString, false);

   if (FAILED(hr))
   {
      return Error(L"sendAsUNICODE should be a Boolean value", GUID_NULL, E_INVALIDARG);  
   }

   bool lastWrite;

   hr = GetOptionalBool(thenShutdown, lastWrite, false);

   if (FAILED(hr))
   {
      return Error(L"thenShutdown should be a Boolean value", GUID_NULL, E_INVALIDARG);  
   }

   try
   {
      if (unicodeString)
      {
         m_pSocket->Write((const char*)data, ::SysStringByteLen(data), lastWrite);
      }
      else
      {
         USES_CONVERSION;

         m_pSocket->Write(OLE2A(data), ::SysStringLen(data), lastWrite);
      }
   }
   catch(CException &e)
   {
      return ExceptionToError(e);
   }

	return S_OK;
}

STDMETHODIMP CSocket::Write(
   VARIANT arrayOfBytes,
   VARIANT thenShutdown)
{
   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!");
   }

   if (arrayOfBytes.vt != (VT_ARRAY | VT_UI1)) 
   {
      return Error(L"Expected an array of bytes", GUID_NULL, E_INVALIDARG);
   }

   const long dims = ::SafeArrayGetDim(arrayOfBytes.parray);

   if (dims != 1) 
   {
      return Error(L"Expected a one dimensional array", GUID_NULL, E_INVALIDARG);
   }

   bool lastWrite;

   HRESULT hr = GetOptionalBool(thenShutdown, lastWrite, false);

   if (SUCCEEDED(hr))
   {
      long upperBounds;
      long lowerBounds;

      ::SafeArrayGetLBound(arrayOfBytes.parray, 1, &lowerBounds);
      ::SafeArrayGetUBound(arrayOfBytes.parray, 1, &upperBounds);

      const size_t size = upperBounds - lowerBounds + 1;

      const char *pBuff;

      ::SafeArrayAccessData(arrayOfBytes.parray, (void**)&pBuff);

      try
      {
         m_pSocket->Write(pBuff, size, lastWrite);
      }
      catch(CException &e)
      {
         return ExceptionToError(e);
      }

      ::SafeArrayUnaccessData(arrayOfBytes.parray);
   }
   else
   {
      hr = Error(L"thenShutdown should be a Boolean value", GUID_NULL, E_INVALIDARG);  
   }

   return hr;
}

STDMETHODIMP CSocket::RequestRead()
{
   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!");
   }

   try
   {
      m_pSocket->Read();
   }
   catch(CException &e)
   {
      return ExceptionToError(e);
   }

	return S_OK;
}

STDMETHODIMP CSocket::get_RemoteAddress(IAddress **ppVal)
{
   if (ppVal == 0)
   {
      return Error(L"ppVal is an invalid pointer", GUID_NULL, E_POINTER);
   }

   *ppVal = 0;

   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!");
   }

   m_pAddress->AddRef();

   *ppVal = m_pAddress;

	return S_OK;
}

STDMETHODIMP CSocket::Init(
   unsigned long address,
   unsigned short port,
   void *pSocket)    
{   
   IAddressInit *pInit = 0;

   HRESULT hr = CAddress::CreateInstance(&pInit);

   if (SUCCEEDED(hr))
   {
      hr = pInit->Init(address, port);

      if (SUCCEEDED(hr))
      {
         hr = pInit->QueryInterface(&m_pAddress);
      }

      pInit->Release();
   }

   if (SUCCEEDED(hr))
   {
      m_pSocket = reinterpret_cast<CCOMSocketServer::Socket *>(pSocket);
   }

   return hr;
}

STDMETHODIMP CSocket::Shutdown(ShutdownMethod how)
{
   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!");
   }

   int shutdownHow = 0;

   if (how == ShutdownRead)
   {
      shutdownHow =  SD_RECEIVE;
   }
   else if (how == ShutdownWrite)
   {
      shutdownHow =  SD_SEND;
   }
   else if (how == ShutdownBoth)
   {
      shutdownHow = SD_BOTH;
   }
   else
   {
      return Error(L"Invalid value for ShutdownMethod!", GUID_NULL, E_INVALIDARG);
   }

   try
   {
      m_pSocket->Shutdown(shutdownHow);
   }
   catch(CException &e)
   {
      return Error(e.GetMessage().c_str());
   }

	return S_OK;
}

STDMETHODIMP CSocket::Close()
{
   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!");
   }

   try
   {
      m_pSocket->AbortiveClose();
   }
   catch(CException &e)
   {
      return ExceptionToError(e);
   }

	return S_OK;   
}

STDMETHODIMP CSocket::get_UserData(VARIANT *pVal)
{
   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!");
   }

   if (pVal == 0)
   {
      return Error(L"pVal is an invalid pointer", GUID_NULL, E_POINTER);
   }

   CComVariant result(m_userData);

   return result.Detach(pVal);
}

STDMETHODIMP CSocket::put_UserData(VARIANT newVal)
{
   if (!m_pSocket)
   {
      return Error(L"Socket hasn't been initialised - programming error!");
   }

   m_userData.Copy(&newVal);

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
