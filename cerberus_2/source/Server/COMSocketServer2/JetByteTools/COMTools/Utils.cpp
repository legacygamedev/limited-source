///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Utils.cpp $
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
// $Log: /Web Articles/SocketServers/COMSocketServer2/JetByteTools/COMTools/Utils.cpp $
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
// 1     30/05/02 12:46 Len
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

#include "Utils.h"
#include "Exception.h"

#include <atlbase.h>    // USES_CONVERSION

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
// Functions defined in this file...
///////////////////////////////////////////////////////////////////////////////
   
bool WaitWithMessageLoop(
   HANDLE hEvent, 
   DWORD timeout /* = INFINITE */)
{
   MSG msg;
   
   //lint -e{716} while(1)
   while(1)       
   {
      DWORD dwRet = MsgWaitForMultipleObjects( 
         1,              // One event to wait for
         &hEvent,        // The array of events
         FALSE,          // Wait for 1 event
         timeout,        // Timeout value
         QS_ALLINPUT);   // Any message wakes up
       
      if(dwRet == WAIT_OBJECT_0)       
      {
         // The event was signaled, return         
         return true;
      } 
      else if(dwRet == WAIT_OBJECT_0 + 1)       
      {
         // There is a window message available. Dispatch it.
         while(PeekMessage(&msg,NULL,NULL,NULL,PM_REMOVE))       
         {
            TranslateMessage(&msg);          

            //lint -e{534} Ignoring return value of function
            DispatchMessage(&msg);
         }       
      } 
      else       
      {          // Something else happened. Return.
         return false;       
      }       
   }   
} 

void RestartStream(
   IStream *pIStream)
{
   LARGE_INTEGER seek_pos;

   seek_pos.QuadPart = 0;

   // Seek to the start of the stream...

   HRESULT hr = pIStream->Seek(seek_pos, STREAM_SEEK_SET, 0);     

   CException::ThrowOnFailure(_T("RestartStream()"), hr);
}

IStream *MarshalInterThreadInterfaceInStream(
   IUnknown *pUnknown, 
   REFIID iid)
{
   if (!pUnknown)
   {
      throw CException(_T("MarshalInterThreadInterfaceInStream() - pUnknown"), E_INVALIDARG);
   }

   IUnknown *pInterface = 0;

   HRESULT hr = pUnknown->QueryInterface(iid, (void**)&pInterface);

   CException::ThrowOnFailure(_T("MarshalInterThreadInterfaceInStream() - QueryInterface"), hr);

   IStream *pIStream = 0;

   hr = ::CoMarshalInterThreadInterfaceInStream(iid, pInterface, &pIStream);
      
   pInterface->Release();

   CException::ThrowOnFailure(_T("MarshalInterThreadInterfaceInStream() - CoMarshalInterThreadInterfaceInStream"), hr);

   // TODO need a smart IStream here
   RestartStream(pIStream);

   return pIStream;
}

HRESULT CreateSafeArray(
   const BYTE *pData,
   DWORD dataLength,
   VARIANT *ppResults)
{
   SAFEARRAYBOUND sab;
   sab.cElements = dataLength;
   sab.lLbound = 0;

   SAFEARRAY *pSa = SafeArrayCreateEx(VT_UI1, 1, &sab, 0);
   
   PVOID pvData;
   
   HRESULT hr = SafeArrayAccessData(pSa, &pvData);
   
   if (SUCCEEDED(hr))
   {
      ::CopyMemory(pvData, pData, dataLength);
   
      hr = ::SafeArrayUnaccessData(pSa);
   }

   if (SUCCEEDED(hr))
   {
      ::VariantInit(ppResults);
           
      (*ppResults).parray = pSa;
      (*ppResults).vt = VT_ARRAY | VT_UI1;
   }

   return hr;
}

HRESULT GetOptionalDWORD(
   VARIANT &source, 
   DWORD &result, 
   const DWORD defaultValue /* = 0 */)
{
   result = defaultValue;

   if (source.vt != VT_EMPTY && source.vt != VT_ERROR)
   {
      VARIANT dest;
         
      ::VariantInit(&dest);

      HRESULT hr = ::VariantChangeType(&dest, &source, 0, VT_I4);

      if (FAILED(hr))
      {
         return hr;
      }
      
      result = dest.lVal;

      ::VariantClear(&dest);
   }

   return S_OK;
}

HRESULT GetOptionalBSTR(
   VARIANT &source, 
   BSTR &result, 
   const BSTR defaultValue /* = L""*/)
{ 
   result = defaultValue;

   if (source.vt != VT_EMPTY && source.vt != VT_ERROR)
   {
      VARIANT dest;
         
      ::VariantInit(&dest);

      HRESULT hr = ::VariantChangeType(&dest, &source, 0, VT_BSTR);

      if (FAILED(hr))
      {
         return hr;
      }
      
      result = dest.bstrVal;

      ::VariantClear(&dest);
   }

   return S_OK;
}

HRESULT GetOptionalString(
   VARIANT &source, 
   std::wstring &result, 
   const std::wstring &defaultValue /* = L""*/)
{
   result = defaultValue;

   if (source.vt != VT_EMPTY && source.vt != VT_ERROR)
   {
      VARIANT dest;
         
      ::VariantInit(&dest);

      HRESULT hr = ::VariantChangeType(&dest, &source, 0, VT_BSTR);

      if (FAILED(hr))
      {
         return hr;
      }
      
      result = dest.bstrVal;

      ::VariantClear(&dest);
   }

   return S_OK;
}

HRESULT GetOptionalString(
   VARIANT &source, 
   std::string &result, 
   const std::string &defaultValue /* = ""*/)
{
   result = defaultValue;

   if (source.vt != VT_EMPTY && source.vt != VT_ERROR)
   {
      VARIANT dest;
         
      ::VariantInit(&dest);

      HRESULT hr = ::VariantChangeType(&dest, &source, 0, VT_BSTR);

      if (FAILED(hr))
      {
         return hr;
      }
      
      USES_CONVERSION;

      result = OLE2A(dest.bstrVal);

      ::VariantClear(&dest);
   }

   return S_OK;
}

HRESULT GetOptionalBool(
   VARIANT &source, 
   bool &result, 
   const bool &defaultValue /*= false*/)
{
   result = defaultValue;

   if (source.vt != VT_EMPTY && source.vt != VT_ERROR)
   {
      VARIANT dest;
         
      ::VariantInit(&dest);

      HRESULT hr = ::VariantChangeType(&dest, &source, 0, VT_BOOL);

      if (FAILED(hr))
      {
         return hr;
      }
      
      result = (dest.boolVal == VARIANT_TRUE);

      ::VariantClear(&dest);
   }

   return S_OK;
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
// End of file
///////////////////////////////////////////////////////////////////////////////
