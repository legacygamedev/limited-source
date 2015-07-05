// plugin.cpp : Implementation of Cplugin
#include "stdafx.h"
#include "AtlSample.h"
#include "plugin.h"

/////////////////////////////////////////////////////////////////////////////
// Cplugin
/*
 make sure your dlls are named *_2.dll for sleuth 1.35c+ to find and use them. 

 To tell which what the lower listbox is full of, you can access the frmMain
 Public curList As ParseTypes
 
 Refer to the help file for a complete listing of exposed objects names & Types
*/

enum ParseTYpes{
	ptLinks = 0,
    ptImages = 1,
    ptFrames = 2,
    ptCookies = 3,
    ptForms = 4,
    ptQsLinks = 5,
    ptScripts = 6,
    ptComments = 7,
    ptMetaTags = 8,
};


STDMETHODIMP Cplugin::SetHost(IDispatch ** newref)
{

	//save reference as private class variable for latter use
	frmMain = (IDispatch FAR*)*newref;
	
	HRESULT hresult;
	IDispatch FAR* pdisp = (IDispatch FAR*)*newref;
	DISPID dispid;
	OLECHAR FAR* szMember = L"RegisterPlugin";
    
	//get the dispatch ID of the RegisterPlugin function
   	hresult = pdisp->GetIDsOfNames(IID_NULL,&szMember,1, LOCALE_SYSTEM_DEFAULT, &dispid);
	
	if(hresult == S_OK){
	//	MessageBox(0,"got dispid  ok","ok",MB_OK);
	}else{
		MessageBox(0,"didnt get dispid","not ok",MB_OK);
		return S_OK;
	} 

	
	DISPPARAMS dispparams;
	VARIANTARG dispArgArray[3];
	dispparams.rgvarg = &dispArgArray[0];
	 
	//call the function with the 3 arguments vb prototype follows
	//arguments are in reverse order
	//
	//Function RegisterPlugin(intMenu As Integer, 
	//                        strMenuName As String,
	//                        intStartupArgument As Integer)

	dispparams.rgvarg[0].vt =  VT_I2;
	dispparams.rgvarg[0].lVal   = 1;   //intStartupArgument
	dispparams.rgvarg[1].vt  =  VT_BSTR;
	dispparams.rgvarg[1].bstrVal = SysAllocString(L"C++ Sample Plugin");
	dispparams.rgvarg[2].vt =  VT_I2;
	dispparams.rgvarg[2].lVal  = 0;    //intMenu 0 = main Plugins Menu, 1= list rt click menu
	dispparams.cArgs = 3;
	dispparams.cNamedArgs = 0;


	hresult = pdisp->Invoke(dispid,IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
		  &dispparams, NULL, NULL, NULL);


	return S_OK;
}

STDMETHODIMP Cplugin::StartUp(int * myArg)
{
    
	switch(*myArg){
		case 1:
			MessageBox(0,"C++ Sample Do Something for Arg=1","ok",MB_OK);
			break;
	}

	return S_OK;
}
