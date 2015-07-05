/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sat Jun 28 08:04:10 2003
 */
/* Compiler settings for C:\Documents and Settings\Administrator\Desktop\plugin sample\plugins\AtlSample\AtlSample.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_Iplugin = {0xE3FF7F96,0x8669,0x43AF,{0xA5,0x83,0x87,0x8D,0x4E,0xD8,0xE7,0x7C}};


const IID LIBID_ATLSAMPLELib = {0x08652EE3,0xD033,0x4CC1,{0x90,0xC3,0x40,0x58,0x1D,0xCB,0x1E,0x1C}};


const CLSID CLSID_plugin = {0x07782213,0x552D,0x4FA7,{0x86,0x64,0xBE,0x68,0x7A,0x3F,0x90,0xBC}};


#ifdef __cplusplus
}
#endif

