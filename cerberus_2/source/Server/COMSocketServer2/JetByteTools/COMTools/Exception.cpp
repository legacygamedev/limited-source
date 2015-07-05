///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: Exception.cpp $
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
// $Log: /JetByteTools/COMTools/Exception.cpp $
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

#include "Exception.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Using directives
///////////////////////////////////////////////////////////////////////////////

using JetByteTools::Win32::_tstring;

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::COM
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace COM {

///////////////////////////////////////////////////////////////////////////////
// CException
///////////////////////////////////////////////////////////////////////////////

CException::CException(
   const _tstring &where, 
   HRESULT hr)
   :  CWin32Exception(where, hr)
{

}

_tstring CException::GetMessage() const 
{ 
	return CWin32Exception::GetMessage();
}

#if 0
HRESULT CException::SetAsComErrorInfo(
	REFIID iid, 
	LPCTSTR pSource /*= 0*/,
	LPCTSTR pHelpFile /*= 0*/,
	DWORD dwHelpContext /*= 0*/) const
{
	ICreateErrorInfo *pICreateError = 0;

	HRESULT hr = CreateErrorInfo(&pICreateError);

	if (FAILED(hr))
	{
		throw CException(_T(""), hr);
	}

	// Fill in the stuff here

	hr = pICreateError->SetGUID(iid);

	if (FAILED(hr))
	{
		throw CException(_T(""), hr);
	}

//	need convert to bstrs
//	hr = SetDescription();




	IErrorInfo *pIErrorInfo = 0;

	hr = pICreateError->QueryInterface(IID_IErrorInfo, (void**)&pIErrorInfo);

	pICreateError->Release();

	if (FAILED(hr))
	{
		throw CException(_T(""), hr);
	}

	SetErrorInfo(0, pIErrorInfo);

	pIErrorInfo->Release();

	return m_error;
}
#endif

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
