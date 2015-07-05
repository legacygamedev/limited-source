#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_COM_USES_COM_ITERATOR_INCLUDED__
#define JETBYTE_TOOLS_COM_USES_COM_ITERATOR_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: UsesCom.h $
// Version        : $Revision: 2 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 30/05/02 15:45 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /JetByteTools/COMTools/UsesCom.h $
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
//lint -esym(1704, Exception::Exception) private constructor
//lint -esym(1712, Exception) no defalt constructor
//
///////////////////////////////////////////////////////////////////////////////

#pragma warning(disable: 4201)   // nameless struct/union

#include <objbase.h>

#include "Exception.h"

#pragma warning(disable: 4201)

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::COM
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace COM {
      
///////////////////////////////////////////////////////////////////////////////
// CUsesCOM
///////////////////////////////////////////////////////////////////////////////

class CUsesCOM
{
   public :

      // Exceptions we might throw

      class Exception : public CException
      {
         private :

            friend class CUsesCOM;

            Exception(LPCTSTR pWhere, HRESULT hr)
               :  CException(pWhere, hr)
            {
            }
      };

#if (_WIN32_WINNT >= 0x0400 ) || defined(_WIN32_DCOM)
      explicit CUsesCOM(
         DWORD dwCoInit = COINIT_APARTMENTTHREADED);
#else
      explicit CUsesCOM(
         DWORD reserved = 0);
#endif
      virtual ~CUsesCOM();

};

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

#endif // JETBYTE_TOOLS_COM_USES_COM_ITERATOR_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
