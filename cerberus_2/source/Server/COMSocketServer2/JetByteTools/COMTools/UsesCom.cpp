///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: UsesCom.cpp $
// Version        : $Revision: 1 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 22/05/02 11:05 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /JetByteTools/COMTools/UsesCom.cpp $
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

#include "UsesCom.h"
#include "Exception.h"

#pragma warning(disable: 4711) // function selected for automatic inline expansion

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::COM
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace COM {

///////////////////////////////////////////////////////////////////////////////
// CUsesCOM
///////////////////////////////////////////////////////////////////////////////

#if (_WIN32_WINNT >= 0x0400 ) || defined(_WIN32_DCOM)
CUsesCOM::CUsesCOM(DWORD dwCoInit /* = COINIT_APARTMENTTHREADED */)
#else
CUsesCOM::CUsesCOM(DWORD /* reserved  = 0 */)
#endif

{
#if (_WIN32_WINNT >= 0x0400 ) || defined(_WIN32_DCOM)
   HRESULT hr = ::CoInitializeEx(NULL, dwCoInit);
#else
   HRESULT hr = ::CoInitialize(NULL);
#endif

   if (hr != S_OK && hr != S_FALSE)
   {
      throw Exception(_T("CUsesCOM::CUsesCOM()"), hr);
   }
}

CUsesCOM::~CUsesCOM()
{
   ::CoUninitialize();
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

///////////////////////////////////////////////////////////////////////////////
// End of file...
///////////////////////////////////////////////////////////////////////////////
