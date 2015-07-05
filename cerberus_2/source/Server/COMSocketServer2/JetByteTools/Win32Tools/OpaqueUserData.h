#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_OPAQUE_USER_DATA_INCLUDED__
#define JETBYTE_TOOLS_OPAQUE_USER_DATA_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: OpaqueUserData.h $
// Version        : $Revision: 4 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 29/05/02 12:04 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/EchoServerEx/JetByteTools/Win32Tools/OpaqueUserData.h $
// 
// 4     29/05/02 12:04 Len
// More lint issues.
// 
// 3     29/05/02 11:34 Len
// Lint issues.
// 
// 2     26/05/02 21:19 Len
// Needed a const cast to compile in VC.Net
// 
// 1     26/05/02 15:08 Len
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
// Private constructor
//lint -esym(1704, COpaqueUserData::COpaqueUserData)
//
// Member not defined
//lint -esym(1526, COpaqueUserData::COpaqueUserData)
//lint -esym(1526, COpaqueUserData::operator=)
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {

///////////////////////////////////////////////////////////////////////////////
// COpaqueUserData
///////////////////////////////////////////////////////////////////////////////

class COpaqueUserData 
{
   public:

      void *GetUserPtr() const
      {
         //lint -e{50} Attempted to take the address of a non-lvalue
         return InterlockedExchangePointer(&(const_cast<void*>(m_pUserData)), m_pUserData);
      }
      
      void SetUserPtr(void *pData)
      {
         //lint -e{534} Ignoring return value of function 
         //lint -e{522} Expected void type, assignment, increment or decrement
         InterlockedExchangePointer(&m_pUserData, pData);
      }

      unsigned long GetUserData() const
      {
         return reinterpret_cast<unsigned long>(GetUserPtr());
      }

      void SetUserData(unsigned long data)
      {
         SetUserPtr(reinterpret_cast<void*>(data));
      }

   protected :
      
      COpaqueUserData()
         : m_pUserData(0)
      {
      }

      ~COpaqueUserData()
      {
         m_pUserData = 0;
      }

   private :

      void *m_pUserData;

      // No copies do not implement
      COpaqueUserData(const COpaqueUserData &rhs);
      COpaqueUserData &operator=(const COpaqueUserData &rhs);
};

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools
///////////////////////////////////////////////////////////////////////////////

} // End of namespace JetByteTools 

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_TOOLS_OPAQUE_USER_DATA_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////

