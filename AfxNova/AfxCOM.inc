' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxCom2.inc
' Purpose:  COM wrapper procedures
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

' ========================================================================================
' Note: Don't forget to initialize the COM library with CoInitialize or CoInitializeEx
' or OleInitialize or you will get GPFs. Call CoUninitialize or OleUninitialize when ending
' the application.
' ========================================================================================

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/unknwnbase.bi"
#include once "win/ocidl.bi"
#include once "AfxNova/AfxWin.inc"
USING AfxNova

NAMESPACE AfxNova

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ########################################################################################
' Base types for declaring ABSTRACT interface methods in other types that inherit from
' these ones. Afx_IUnknown extends the built-in OBJECT type which provides run-time type
' information for all types derived from it using Extends. Extending the built-in Object
' type allows to add an extra hidden vtable pointer field at the top of the Type. The
' vtable is used to dispatch Virtual and Abstract methods and to access information for
' run-time type identification used by Operator Is.
' ########################################################################################

#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
	DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
	DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif

#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_  EXTENDS Afx_Iunknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) as HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

' ########################################################################################

' ========================================================================================
' // Checks if the passed pointer is a BSTR.
' // Will return FALSE if it is a null pointer.
' // If it is an OLE string it must have a descriptor; otherwise, don't.
' // Get the length in bytes looking at the descriptor and divide by 2 to get the number of
' // unicode characters, that is the value returned by the FreeBASIC LEN operator.
' // If the retrieved length if the same that the returned by LEN, then it must be an OLE string.
' ========================================================================================
FUNCTION AfxIsBstr (BYVAL pv aS ANY PTR) AS BOOLEAN
   IF pv = NULL THEN RETURN FALSE
   DIM res AS DWORD = PEEK(DWORD, pv - 4) \ 2
   IF res = LEN(*cast(WSTRING PTR, pv)) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Takes a null terminated wide string as input, and returns a pointer to a new wide string
' allocated with CoTaskMemAlloc (SHStrDupW allocates memory using CoTaskMemAlloc).
' Free the returned string with CoTaskMemFree.
' Note: This is useful when we need to pass a pointer to a null terminated wide string to a
' function or method that will release it. If we pass a WSTRING it will GPF.
' If the length of the input string is 0, CoTaskMemAlloc allocates a zero-length item and
' returns a valid pointer to that item. If there is insufficient memory available,
' CoTaskMemAlloc returns NULL.
' ========================================================================================
PRIVATE FUNCTION AfxWstrAlloc (BYREF wszStr AS WSTRING) AS WSTRING PTR
   DIM ppwsz AS WSTRING PTR
   IF SHStrDupW(@wszStr, @ppwsz) = S_OK THEN RETURN ppwsz
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID)
' To be valid, the string must contain exactly 32 hexadecimal digits, delimited by hyphens
' and enclosed by curly braces. For example: {B09DE715-87C1-11D1-8BE3-0000F8754DA1}
' If pwszGuidText is omited, AfxGuid generates a new unique guid.
' Remarks: I have needed to call the UuidCreate function dynamically because, at the time
' of writing, the library for the RPCRT4.DLL seems broken and the linker fails.
' ========================================================================================
PRIVATE FUNCTION AfxGuid (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
   DIM rguid AS GUID
   IF pwszGuidText = NULL THEN
      ' // Generate a new guid
      DIM AS ANY PTR pLib = DyLibLoad("RPCRT4.DLL")
      IF pLib  THEN
         DIM pProc AS FUNCTION (BYVAL Uuid AS UUID PTR) AS RPC_STATUS
         pProc = DyLibSymbol(pLib, "UuidCreate")
         IF pProc THEN SetLastError(pProc(@rguid))
         DyLibFree(pLib)
      END IF
   ELSE
      SetLastError(CLSIDFromString(pwszGuidText, @rGuid))
   END IF
   RETURN rguid
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a 38-byte human-readable guid string from a 16-byte GUID.
' ========================================================================================
PRIVATE FUNCTION AfxGuidText OVERLOAD (BYVAL classID AS CLSID PTR) AS STRING
   DIM pwsz AS WSTRING PTR
   SetLastError(StringFromCLSID(classID, CAST(LPOLESTR PTR, @pwsz)))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGuidText OVERLOAD (BYVAL classID AS CLSID) AS STRING
   DIM pwsz AS WSTRING PTR
   SetLastError(StringFromCLSID(@classID, CAST(LPOLESTR PTR, @pwsz)))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGuidText OVERLOAD (BYVAL riid AS REFIID) AS STRING
   DIM pwsz AS WSTRING PTR
   SetLastError(StringFromIID(riid, CAST(LPOLESTR PTR, @pwsz)))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.
' Parameters:
' - wszProgID
'   The ProgID or the CLSID of the object to create.
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' wszLicKey (optional)
'   The license key as a unicode string.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom("Scripting.Dictionary")
' -or-
'   pDic = AfxNewCom(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
' CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszProgID AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "") AS ANY PTR
   DIM hr AS HRESULT, classID AS CLSID, pUnk AS ANY PTR, pIClassFactory2 AS IClassFactory2 PTR
   IF INSTR(wszProgID, "{") THEN CLSIDFromString(wszProgID, @classID) ELSE CLSIDFromProgID(wszProgID, @classID)
   IF IsEqualGuid(@classID, @IID_NULL) THEN RETURN NULL
   IF LEN(wszLicKey) = 0 THEN
      CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   ELSE
      hr = CoGetClassObject(@ClassID, CLSCTX_SERVER, NULL, @IID_IClassFactory2, @pIClassFactory2)
      IF hr = S_OK AND pIClassFactory2 <> NULL THEN
         pIClassFactory2->lpVtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @IID_IUnknown, @wszLicKey, @pUnk)
      END IF
      IF pIClassFactory2 THEN pIClassFactory2->lpVtbl->Release(pIClassFactory2)
   END IF
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.
' Parameters:
' - wszProgID = A CLSID in string format.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom(CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'   CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' and IID_IDictionary as
'   CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszClsID AS CONST WSTRING, BYREF wszIID AS CONST WSTRING) AS ANY PTR
   DIM classID AS CLSID, riid AS IID, pUnk AS ANY PTR
   CLSIDFromProgID(wszClsID, @classID)
   CLSIDFromProgID(wszIID, @riid)
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @riid, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified CLSID.
' Parameter:
' - classID = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF classID AS CONST CLSID) AS ANY PTR
   DIM pUnk AS ANY PTR
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with the specified CLSID and IID.
' Parameters:
' - classID = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom(CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' and IID_IDictionary as
'   DIM IID_IDictionary AS IID = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID) AS ANY PTR
   DIM pUnk AS ANY PTR
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @riid, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads the specified library from file and creates an instance of an object.
' Parameters:
' - wszLibName = Full path where the library is located.
' - rclsid = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' - wszLicKey = The license key.
' If it succeeds, returns a reference to the requested interface; otherwise, it returns null.
' Not every component is a suitable candidate for use under this overloaded AfxNewCom function.
'  - Only in-process servers (DLLs) are supported.
'  - Components that are system components or part of the operating system, such as XML,
'    Data Access, Internet Explorer, or DirectX, aren't supported
'  - Components that are part of an application, such Microsoft Office, aren't supported.
'  - Components intended for use as an add-in or a snap-in, such as an Office add-in or
'    a control in a Web browser, aren't supported.
'  - Components that manage a shared physical or virtual system resource aren't supported.
'  - Visual ActiveX controls aren't supported because they need to be initilized and
'    activated by the OLE container.
' Note: Do not use DyLibFree to unload the library once you have got a valid reference
' to an interface or your application will GPF. Before calling DyLibFree, all the
' interface references must be released. If you don't need to unload the library until
' the application ends, then you don't need to call DyLibFree because CoUninitialize
' closes the COM library on the current thread, unloads all DLLs loaded by the thread,
' frees any other resources that the thread maintains, and forces all RPC connections on
' the thread to close.
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "") AS ANY PTR

   DIM hr AS LONG, hLib AS HMODULE, pDisp AS ANY PTR
   DIM pIClassFactory AS IClassFactory PTR, pIClassFactory2 AS IClassFactory2 PTR

   ' // See if the library is already loaded in the address space
   hLib = GetModuleHandleW(wszLibName)
   ' // If it is not loaded, load it
   IF hLib = NULL THEN hLib = LoadLibraryW(wszLibName)
   ' // If it fails, abort
   IF hLib = NULL THEN EXIT FUNCTION

   ' // Retrieve the address of the exported function DllGetClassObject
   DIM pfnDllGetClassObject AS FUNCTION (BYVAL rclsid AS CONST IID CONST PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppv AS LPVOID PTR) AS HRESULT
   pfnDllGetClassObject = CAST(ANY PTR, GetProcAddress(hLib, "DllGetClassObject"))
   IF pfnDllGetClassObject = NULL THEN EXIT FUNCTION

   IF LEN(wszLicKey) = 0 THEN
      ' // Request a reference to the IClassFactory interface
      hr = pfnDllGetClassObject(@rclsid, @IID_IClassFactory, @pIClassFactory)
      IF hr <> S_OK THEN EXIT FUNCTION
      ' // Create an instance of the server or control
      hr = pIClassFactory->lpVtbl->CreateInstance(pIClassFactory, NULL, @riid, @pDisp)
      IF hr <> S_OK THEN
         pIClassFactory->lpVtbl->Release(pIClassFactory)
         EXIT FUNCTION
      END IF
   ELSE
      ' // Request a reference to the IClassFactory2 interface
      hr = pfnDllGetClassObject(@rclsid, @IID_IClassFactory, @pIClassFactory2)
      IF hr <> S_OK THEN EXIT FUNCTION
      ' // Create a licensed instance of the server or control
      hr = pIClassFactory2->lpVtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @riid, @wszLicKey, @pDisp)
      IF hr <> S_OK THEN
         pIClassFactory2->lpVtbl->Release(pIClassFactory2)
         EXIT FUNCTION
      END IF
   END IF

   IF pIClassFactory THEN pIClassFactory->lpVtbl->Release(pIClassFactory)
   IF pIClassFactory2 THEN pIClassFactory2->lpVtbl->Release(pIClassFactory2)
   RETURN pDisp

END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the wszClsid and wszIid parameters to GUIDs and calls the function above.
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszLibName AS CONST WSTRING, BYREF wszClsid AS CONST WSTRING, BYREF wszIid AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "") AS ANY PTR
   DIM rclsid AS CLSID, riid AS IID
   rclsid = AfxGuid(wszClsid)
   riid = AfxGuid(wszIid)
   RETURN AfxNewCom(wszLibName, rclsid, riid, wszLicKey)
END FUNCTION
' ========================================================================================

' ========================================================================================
' If the requested object is in an EXE (out-of-process server), such Office applications,
' and it is running and registered in the Running Object Table (ROT), AfxGetCom will
' return a pointer to its interface. AfxAnyCom will first try to use an existing, running
' application if available, or it will create a new instance if not.
' Be aware that AfxGetCom can fail under if Office is running but not registered in the ROT.
' When an Office application starts, it does not immediately register its running objects.
' This optimizes the application's startup process. Instead of registering at startup, an
' Office application registers its running objects in the ROT once it loses focus. Therefore,
' if you attempt to use GetObject or GetActiveObject to attach to a running instance of an
' Office application before the application has lost focus, you might receive an error.
' See: https://support.microsoft.com/en-us/help/238610/getobject-or-getactiveobject-cannot-find-a-running-office-application
' ========================================================================================
PRIVATE FUNCTION AfxGetCom OVERLOAD (BYREF wszProgID AS CONST WSTRING) AS IDispatch PTR
   DIM classID AS CLSID, pUnk AS IUnknown PTR, pDisp AS IDispatch PTR
   CLSIDFromProgID(wszProgID, @classID)
   IF IsEqualGuid(@classID, @IID_NULL) THEN RETURN NULL
   GetActiveObject(@classID, NULL, @pUnk)
   IF pUnk THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pDisp)
      pUnk->lpVtbl->Release(pUnk)
   END IF
   RETURN pDisp
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGetCom OVERLOAD (BYREF classID AS CONST CLSID) AS IDispatch PTR
   DIM pUnk AS IUnknown PTR, pDisp AS IDispatch PTR
   GetActiveObject(@classID, NULL, @pUnk)
   IF pUnk THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pDisp)
      pUnk->lpVtbl->Release(pUnk)
   END IF
   RETURN pDisp
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAnyCom OVERLOAD (BYREF wszProgID AS CONST WSTRING) AS IDispatch PTR
   DIM classID AS CLSID, pUnk AS IUnknown PTR, pDisp AS IDispatch PTR
   CLSIDFromProgID(wszProgID, @classID)
   IF IsEqualGuid(@classID, @IID_NULL) THEN RETURN NULL
   ' // Check if there is an instance already running
   IF GetActiveObject(@classID, NULL, @pUnk) = S_OK THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pDisp)
      pUnk->lpVtbl->Release(pUnk)
      RETURN pDisp
   END IF
   ' // Otherwise, create a new instance
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   IF pUnk THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pDisp)
      pUnk->lpVtbl->Release(pUnk)
   END IF
   RETURN pDisp
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAnyCom OVERLOAD (BYREF classID AS CONST CLSID) AS IDispatch PTR
   DIM pUnk AS IUnknown PTR, pDisp AS IDispatch PTR
   ' // Check if there is an instance already running
   IF GetActiveObject(@classID, NULL, @pUnk) = S_OK THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pDisp)
      pUnk->lpVtbl->Release(pUnk)
      RETURN pDisp
   END IF
   ' // Otherwise, create a new instance
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   IF pUnk THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pDisp)
      pUnk->lpVtbl->Release(pUnk)
   END IF
   RETURN pDisp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' The function returns the new reference count. This value is intended to be used only
' for test purposes.
' When the reference count on an object reaches zero, Release must cause the interface
' pointer to free itself. When the released pointer is the only existing reference to an
' object (whether the object supports single or multiple interfaces), the implementation
' must free the object.
' ========================================================================================
PRIVATE FUNCTION AfxSafeRelease (BYREF pv AS ANY PTR) AS ULONG
   IF pv = NULL THEN RETURN 0
   FUNCTION = cast(IUnknown PTR, pv)->lpvtbl->Release(pv)
   pv = NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. The method returns the new
' reference count. This value is intended to be used only for test purposes.
' This method should be called for every new copy of a pointer to an interface on an object.
' For example, if you are passing a copy of a pointer back from a method, you must call
' AddRef on that pointer. You must also call AddRef on a pointer before passing it as an
' in-out parameter to a method; the method will call IUnknown_Release before copying the
' out-value on top of it.
' Objects use a reference counting mechanism to ensure that the lifetime of the object
' includes the lifetime of references to it. You use AddRef to stabilize a copy of an
' interface pointer. It can also be called when the life of a cloned pointer must extend
' beyond the lifetime of the original pointer. The cloned pointer must be released by
' calling AfxSafeRelease.
' ========================================================================================
PRIVATE FUNCTION AfxSafeAddRef (BYVAL pv AS ANY PTR) AS ULONG
   IF pv = NULL THEN RETURN 0
   RETURN cast(IUnknown PTR, pv)->lpvtbl->AddRef(pv)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Establishes a connection between the connection point object and the client's sink.
' Parameters
' pUnk: [in] A pointer to the IUnknown of the object the client wants to connect with.
' pEvtObj: [in] A pointer to the client's IUnknown.
' riid: [in] The GUID of the connection point. Typically, this is the same as the outgoing
' interface managed by the connection point.
' pdwCookie: [out] A pointer to the cookie that uniquely identifies the connection.
' Return value: S_OK or an HRESULT value.
' ========================================================================================
PRIVATE FUNCTION AfxAdvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, BYVAL riid AS IID PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   IF pUnk = NULL OR pEvtObj = NULL OR riid = NULL OR pdwCookie = NULL THEN RETURN E_INVALIDARG
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = IUnknown_QueryInterface(cast(IUnknown PTR, pUnk), @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, riid, @pCP)
   IUnknown_Release(pCPC)
   IF hr <> S_OK OR pCP = NULL THEN RETURN hr
   ' // Establishes a connection between a connection point object and the client's sink.
   hr = pCP->lpvtbl->Advise(pCP, cast(IUnknown PTR, pEvtObj), pdwCookie)
   IUnknown_Release(pCP)
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAdvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, BYREF riid AS CONST IID, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   RETURN AfxAdvise(pUnk, pEvtObj, cast(IID PTR, @riid), pdwCookie)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAdvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, BYREF riid AS IID, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   RETURN AfxAdvise(pUnk, pEvtObj, cast(IID PTR, @riid), pdwCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the events connection identified with the cookie returned by the AfxAdvise method.
' Parameters
' pUnk [in] A pointer to the IUnknown of the object the client wants to disconnect with.
' riid: [in] The GUID of the connection point. Typically, this is the same as the outgoing
' interface managed by the connection point.
' dwCookie: [in] The cookie that uniquely identifies the connection.
' Return value: S_OK or an HRESULT value.
' ========================================================================================
PRIVATE FUNCTION AfxUnadvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL riid AS IID PTR, BYVAL dwCookie AS DWORD) AS HRESULT
   ' // Not a valid connection
   IF pUnk = NULL OR riid = NULL OR dwCookie = 0 THEN RETURN E_INVALIDARG
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = IUnknown_QueryInterface(cast(IUnknown PTR, pUnk), @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, riid, @pCP)
   IUnknown_Release(pCPC)
   IF hr <> S_OK OR pCP = NULL THEN RETURN hr
   ' // Terminates the advisory connection previously established with a call to AfxAdvise
   hr = pCP->lpvtbl->Unadvise(pCP, dwCookie)
   IUnknown_Release(pCP)
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxUnadvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYREF riid AS CONST IID, BYVAL dwCookie AS DWORD) AS HRESULT
   RETURN AfxUnadvise(pUnk, cast(IID PTR, @riid), dwCookie)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxUnadvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYREF riid AS IID, BYVAL dwCookie AS DWORD) AS HRESULT
   RETURN AfxUnadvise(pUnk, cast(IID PTR, @riid), dwCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the description of the most recent OLE error in the current logical thread and
' clears the error state for the thread. It should be called as soon as possible after
' calling a method of an Automation interface and will only succeed if the object supports
' the IErrorInfo interface.
' ========================================================================================
PRIVATE FUNCTION AfxGetOleErrorInfo () AS DWSTRING
   DIM pErrorInfo AS IErrorInfo PTR, dwsInfo AS DWSTRING
   DIM hr AS HRESULT = GetErrorInfo(0, @pErrorInfo)
   IF hr <> S_OK THEN RETURN ""
   DIM bstrDescription AS AFX_BSTR
   hr = pErrorInfo->lpvtbl->GetDescription(pErrorInfo, @bstrDescription)
   IF hr = S_OK THEN
      dwsInfo = *bstrDescription
      SysFreeString bstrDescription
      AfxSafeRelease(pErrorInfo)
   END IF
   RETURN dwsInfo
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a standard IFontDisp object.
' Parameters:
' - wszFontName: The typeface name, e.g. "Segoe UI"
' - cySize: The point size, e.g. 9.
' - fWeight: Initial weight of the font. If the weight is below 550 (the average of
'      FW_NORMAL, 400, and FW_BOLD, 700), then the Bold property is also initialized to
'      FALSE. If the weight is above 550, the Bold property is set to TRUE.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: A pointer to the object or NULL on failure.
' Remarks: The returned font must be destroyed calling the release method of the IFontDisp
' interface when no longer needed to prevent memory leaks.
' Usage examples:
' DIM pFont AS IFontDisp PTR = AfxOleCreateFontDisp("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
' DIM pFont AS IFontDisp PTR = AfxOleCreateFontDisp("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
' DIM pFont AS IFontDisp PTR = AfxOleCreateFontDisp("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION AfxOleCreateFontDisp ( _
   BYREF wszFontName AS WSTRING, _                  ' __in  Font name
   BYVAL cySize AS LONGLONG, _                      ' __in  Point size
   BYVAL fWeight AS SHORT = FW_NORMAL, _            ' __in  Weight (FW_NORMAL, etc.)
   BYVAL fItalic AS WINBOOL = FALSE, _              ' __in  Italic state
   BYVAL fUnderline AS WINBOOL = FALSE, _           ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL = FALSE, _       ' __in  Strikethrough state
   BYVAL fCharset AS SHORT = DEFAULT_CHARSET _      ' __in  Character set
   ) AS IFontDisp PTR

   DIM tf AS FONTDESC
   tf.cbSizeOfStruct = SIZEOF(FONTDESC)
   tf.lpstrName =  @wszFontName
   tf.cySize.int64 = cySize * 10000
   tf.sWeight = fWeight
   tf.sCharset = fCharset
   tf.fItalic = fItalic
   tf.fUnderline = fUnderline
   tf.fStrikethrough = fStrikethrough
   DIM pFont AS IFontDisp PTR
   OleCreateFontIndirect(@tf, @IID_IDispatch, @pFont)
   RETURN pFont

END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxOleCreateFont ( _
   BYREF wszFontName AS WSTRING, _                  ' __in  Font name
   BYVAL cySize AS LONGLONG, _                      ' __in  Point size
   BYVAL fWeight AS SHORT = FW_NORMAL, _            ' __in  Weight (FW_NORMAL, etc.)
   BYVAL fItalic AS WINBOOL = FALSE, _              ' __in  Italic state
   BYVAL fUnderline AS WINBOOL = FALSE, _           ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL = FALSE, _       ' __in  Strikethrough state
   BYVAL fCharset AS SHORT = DEFAULT_CHARSET _      ' __in  Character set
   ) AS IFont PTR

   DIM tf AS FONTDESC
   tf.cbSizeOfStruct = SIZEOF(FONTDESC)
   tf.lpstrName =  @wszFontName
   tf.cySize.int64 = cySize * 10000
   tf.sWeight = fWeight
   tf.sCharset = fCharset
   tf.fItalic = fItalic
   tf.fUnderline = fUnderline
   tf.fStrikethrough = fStrikethrough
   DIM pFont AS IFont PTR
   OleCreateFontIndirect(@tf, @IID_IUnknown, @pFont)
   RETURN pFont

END FUNCTION
' ========================================================================================

END NAMESPACE
