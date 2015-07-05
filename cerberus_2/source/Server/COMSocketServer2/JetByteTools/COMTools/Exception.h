#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_COM_EXCEPTION_INCLUDED__
#define JETBYTE_TOOLS_COM_EXCEPTION_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Exception.h $
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
// $Log: /JetByteTools/COMTools/Exception.h $
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

#include "JetByteTools\Win32Tools\Win32Exception.h"

#ifndef _HRESULT_DEFINED			
#include <wtypes.h>					
#endif

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::COM
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace COM {
      
///////////////////////////////////////////////////////////////////////////////
// CException
///////////////////////////////////////////////////////////////////////////////

class CException : public Win32::CWin32Exception
{
   public :

      CException(
         const Win32::_tstring &where, 
         HRESULT hr);

		virtual Win32::_tstring GetMessage() const; 

/*		HRESULT SetAsComErrorInfo(
			REFIID iid, 
			LPCTSTR pSource = 0,
			LPCTSTR pHelpFile = 0,
			DWORD dwHelpContext = 0) const;
*/

		static void ThrowOnFailure(
         const Win32::_tstring &where, 
         HRESULT hr)
		{
			if (FAILED(hr))
			{
				throw CException(where, hr);
			}
		}
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

#endif // JETBYTE_TOOLS_COM_EXCEPTION_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
