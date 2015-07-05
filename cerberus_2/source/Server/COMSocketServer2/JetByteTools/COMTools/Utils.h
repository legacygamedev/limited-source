#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_COM_UTILS_INCLUDED__
#define JETBYTE_TOOLS_COM_UTILS_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Utils.h $
// Version        : $Revision: 4 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 6/06/02 12:37 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/COMSocketServer2/JetByteTools/COMTools/Utils.h $
// 
// 4     6/06/02 12:37 Len
// Added MarshalInterThreadInterfaceInStream()
// 
// 3     3/06/02 11:18 Len
// Added "optional VARIANT" to data type extraction functions.
// Added RestartStream().
// 
// 2     30/05/02 15:45 Len
// Lint issues.
// 
// 1     22/05/02 11:05 Len
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

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

#include <wtypes.h>
#include "JetbyteTools\Win32Tools\tstring.h"

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::COM
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace COM {
  
///////////////////////////////////////////////////////////////////////////////
// Functions defined in this file...
///////////////////////////////////////////////////////////////////////////////
   
bool WaitWithMessageLoop(
   HANDLE hEvent, 
   DWORD timeout = INFINITE);

void RestartStream(
   IStream *pIStream);

IStream *MarshalInterThreadInterfaceInStream(
   IUnknown *pUnknown, 
   REFIID iid);

HRESULT CreateSafeArray(
   const BYTE *pData,
   DWORD dataLength,
   VARIANT *ppResults);

HRESULT GetOptionalDWORD(
   VARIANT &source, 
   DWORD &result, 
   const DWORD defaultValue = 0);

HRESULT GetOptionalBSTR(
   VARIANT &source, 
   BSTR &result, 
   const BSTR defaultValue = L"");

HRESULT GetOptionalString(
   VARIANT &source, 
   std::wstring &result, 
   const std::wstring &defaultValue = L"");

HRESULT GetOptionalString(
   VARIANT &source, 
   std::string &result, 
   const std::string &defaultValue = "");

HRESULT GetOptionalBool(
   VARIANT &source, 
   bool &result, 
   const bool &defaultValue = false);

///////////////////////////////////////////////////////////////////////////////
// Template functions defined in this file...
///////////////////////////////////////////////////////////////////////////////

template <class I> I *SafeRelease(I *pI) 
{
	if (pI)
	{
      //lint -e{534} Ignoring return value of function
		pI->Release();
	}

	return 0;
}

template <class I> I *SafeAddRef(I *pI) 
{
	if (pI)
	{
      //lint -e{534} Ignoring return value of function
		pI->AddRef();
	}

	return pI;
}

template <class I> IUnknown *SafeQI(I *pI, const IID &iid) 
{
	IUnknown *pUnknown = 0;

	if (pI)
	{
		HRESULT hr = pI->QueryInterface(iid, (void**)&pUnknown);

		if (FAILED(hr))
		{
			pUnknown = 0;
		}
	}

	return pUnknown;
}

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::COM
///////////////////////////////////////////////////////////////////////////////

} // End of namespace COM
} // End of namespace JetByteTools 

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_TOOLS_COM_UTILS_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
