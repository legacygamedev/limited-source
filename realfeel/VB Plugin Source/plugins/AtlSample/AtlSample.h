/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sat Jun 28 08:04:10 2003
 */
/* Compiler settings for C:\Documents and Settings\Administrator\Desktop\plugin sample\plugins\AtlSample\AtlSample.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __AtlSample_h__
#define __AtlSample_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __Iplugin_FWD_DEFINED__
#define __Iplugin_FWD_DEFINED__
typedef interface Iplugin Iplugin;
#endif 	/* __Iplugin_FWD_DEFINED__ */


#ifndef __plugin_FWD_DEFINED__
#define __plugin_FWD_DEFINED__

#ifdef __cplusplus
typedef class plugin plugin;
#else
typedef struct plugin plugin;
#endif /* __cplusplus */

#endif 	/* __plugin_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __Iplugin_INTERFACE_DEFINED__
#define __Iplugin_INTERFACE_DEFINED__

/* interface Iplugin */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_Iplugin;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("E3FF7F96-8669-43AF-A583-878D4ED8E77C")
    Iplugin : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetHost( 
            /* [out][in] */ IDispatch __RPC_FAR *__RPC_FAR *newref) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE StartUp( 
            /* [out][in] */ int __RPC_FAR *myArg) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IpluginVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            Iplugin __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            Iplugin __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            Iplugin __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            Iplugin __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            Iplugin __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            Iplugin __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            Iplugin __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetHost )( 
            Iplugin __RPC_FAR * This,
            /* [out][in] */ IDispatch __RPC_FAR *__RPC_FAR *newref);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *StartUp )( 
            Iplugin __RPC_FAR * This,
            /* [out][in] */ int __RPC_FAR *myArg);
        
        END_INTERFACE
    } IpluginVtbl;

    interface Iplugin
    {
        CONST_VTBL struct IpluginVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define Iplugin_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define Iplugin_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define Iplugin_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define Iplugin_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define Iplugin_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define Iplugin_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define Iplugin_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define Iplugin_SetHost(This,newref)	\
    (This)->lpVtbl -> SetHost(This,newref)

#define Iplugin_StartUp(This,myArg)	\
    (This)->lpVtbl -> StartUp(This,myArg)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Iplugin_SetHost_Proxy( 
    Iplugin __RPC_FAR * This,
    /* [out][in] */ IDispatch __RPC_FAR *__RPC_FAR *newref);


void __RPC_STUB Iplugin_SetHost_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Iplugin_StartUp_Proxy( 
    Iplugin __RPC_FAR * This,
    /* [out][in] */ int __RPC_FAR *myArg);


void __RPC_STUB Iplugin_StartUp_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __Iplugin_INTERFACE_DEFINED__ */



#ifndef __ATLSAMPLELib_LIBRARY_DEFINED__
#define __ATLSAMPLELib_LIBRARY_DEFINED__

/* library ATLSAMPLELib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ATLSAMPLELib;

EXTERN_C const CLSID CLSID_plugin;

#ifdef __cplusplus

class DECLSPEC_UUID("07782213-552D-4FA7-8664-BE687A3F90BC")
plugin;
#endif
#endif /* __ATLSAMPLELib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
