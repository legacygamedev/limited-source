#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_DATA_INCLUDED__
#define JETBYTE_DATA_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Data.h $
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
// $Log: /Web Articles/SocketServers/COMSocketServer/COMSocketServer/Data.h $
// 
// 3     11/06/02 10:37 Len
// Removed the 'readAsUnicode' flag as reading a unicode string is
// problematic due to potential packet fragmentation. To read a unicode
// string you should read the data as a byte array and handle any
// conversions at the application level.
// 
// 2     11/06/02 9:01 Len
// Support reading of unicode/non unicode strings.
// Changed type library id.
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

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

#include "resource.h"       
#include "IDataInit.h"

///////////////////////////////////////////////////////////////////////////////
// CData
///////////////////////////////////////////////////////////////////////////////

class ATL_NO_VTABLE CData : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CData, &CLSID_Data>,
	public ISupportErrorInfo,
	public IDispatchImpl<IData, &IID_IData, &LIBID_JBSOCKETSERVERLib>,
   public IDataInit
{
   public:
	
      CData();

      DECLARE_REGISTRY_RESOURCEID(IDR_DATA)
      DECLARE_NOT_AGGREGATABLE(CData)

      BEGIN_COM_MAP(CData)
	      COM_INTERFACE_ENTRY(IData)
	      COM_INTERFACE_ENTRY(IDispatch)
	      COM_INTERFACE_ENTRY(ISupportErrorInfo)
         COM_INTERFACE_ENTRY(IDataInit)
      END_COM_MAP()

      // ISupportsErrorInfo

      STDMETHOD(InterfaceSupportsErrorInfo)(
         REFIID riid);

      // IData
	   
      STDMETHOD(Read)(
         /*[out, retval]*/ VARIANT *ppResults);
	   
      STDMETHOD(ReadString)(
         /*[out, retval]*/ BSTR *pResults);

      // IDataInit

      STDMETHOD(Init)(
         /*[in]*/ const unsigned char *pData, 
         /*[in]*/ size_t length);

   private :

      const unsigned char *m_pData;
      size_t m_length;
};

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_DATA_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
