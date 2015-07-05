///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Data.cpp $
// Version        : $Revision: 3 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 11/06/02 10:37 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/Data.cpp $
// 
// 3     11/06/02 10:37 Len
// Removed the 'readAsUnicode' flag as reading a unicode string is
// problematic due to potential packet fragmentation. To read a unicode
// string you should read the data as a byte array and handle any
// conversions at the application level.
// 
// 2     11/06/02 9:00 Len
// Support reading of unicode/non unicode strings
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
#include "Data.h"

#include "JetByteTools\COMTools\Utils.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using JetByteTools::COM::CreateSafeArray;
using JetByteTools::COM::GetOptionalBool;

///////////////////////////////////////////////////////////////////////////////
// CData
///////////////////////////////////////////////////////////////////////////////

CData::CData()
   :  m_pData(0),
      m_length(0)
{
}

STDMETHODIMP CData::InterfaceSupportsErrorInfo(REFIID riid)
{
	if (InlineIsEqualGUID(IID_IData,riid))
   {
	   return S_OK;
	}

	return S_FALSE;
}

STDMETHODIMP CData::ReadString(
   BSTR *pResults)
{
   if (!pResults)
   {
      return Error(L"pResults is an invalid pointer", GUID_NULL, E_POINTER);
   }

   USES_CONVERSION;

   // We've unrolled A2OLE as we need to explicitly pass in the length of
   // the data to convert as the string may not be null terminated...

   LPOLESTR pOle = ((_lpa = (char*)m_pData) == NULL) ? NULL : ATLA2WHELPER((LPWSTR) alloca(m_length + 1*2), _lpa, m_length);

   *pResults = ::SysAllocStringLen(pOle, m_length);

	return S_OK;
}

STDMETHODIMP CData::Read(VARIANT *ppResults)
{
   if (!ppResults)
   {
      return Error(L"ppResults is an invalid pointer", GUID_NULL, E_POINTER);
   }

   return CreateSafeArray(m_pData, m_length, ppResults); 
}

HRESULT CData::Init(
   const unsigned char *pData, 
   size_t length)
{
   m_pData = pData;
   m_length = length;

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
