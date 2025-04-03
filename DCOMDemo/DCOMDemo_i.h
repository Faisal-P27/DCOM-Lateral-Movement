

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 06:14:07 2038
 */
/* Compiler settings for DCOMDemo.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0622 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __DCOMDemo_i_h__
#define __DCOMDemo_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ICOMServer_FWD_DEFINED__
#define __ICOMServer_FWD_DEFINED__
typedef interface ICOMServer ICOMServer;

#endif 	/* __ICOMServer_FWD_DEFINED__ */


#ifndef __COMServer_FWD_DEFINED__
#define __COMServer_FWD_DEFINED__

#ifdef __cplusplus
typedef class COMServer COMServer;
#else
typedef struct COMServer COMServer;
#endif /* __cplusplus */

#endif 	/* __COMServer_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "shobjidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __ICOMServer_INTERFACE_DEFINED__
#define __ICOMServer_INTERFACE_DEFINED__

/* interface ICOMServer */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICOMServer;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("a68e05e3-4094-4c62-a050-b75146882d38")
    ICOMServer : public IDispatch
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE COMSend( 
            /* [in] */ BSTR SendBuffer,
            /* [in] */ int SendLength,
            /* [out] */ int *RicvLength,
            /* [retval][out] */ BSTR *DataRicv) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICOMServerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICOMServer * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICOMServer * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICOMServer * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICOMServer * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICOMServer * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICOMServer * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICOMServer * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        HRESULT ( STDMETHODCALLTYPE *COMSend )( 
            ICOMServer * This,
            /* [in] */ BSTR SendBuffer,
            /* [in] */ int SendLength,
            /* [out] */ int *RicvLength,
            /* [retval][out] */ BSTR *DataRicv);
        
        END_INTERFACE
    } ICOMServerVtbl;

    interface ICOMServer
    {
        CONST_VTBL struct ICOMServerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICOMServer_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICOMServer_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICOMServer_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICOMServer_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICOMServer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICOMServer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICOMServer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICOMServer_COMSend(This,SendBuffer,SendLength,RicvLength,DataRicv)	\
    ( (This)->lpVtbl -> COMSend(This,SendBuffer,SendLength,RicvLength,DataRicv) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICOMServer_INTERFACE_DEFINED__ */



#ifndef __DCOMDemoLib_LIBRARY_DEFINED__
#define __DCOMDemoLib_LIBRARY_DEFINED__

/* library DCOMDemoLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_DCOMDemoLib;

EXTERN_C const CLSID CLSID_COMServer;

#ifdef __cplusplus

class DECLSPEC_UUID("fb428c4d-a49d-436f-bdf1-05983a7b441e")
COMServer;
#endif
#endif /* __DCOMDemoLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  BSTR_UserSize64(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal64(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal64(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree64(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


