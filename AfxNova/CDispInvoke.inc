' ########################################################################################
' Platform: Microsoft Windows
' File: CDIspInvoke.inc
' Contents: COM Automation wrapper class
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

#pragma once
#include once "AfxNova/AfxCOM.inc"
#include once "AfxNova/DVARIANT.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' Macro for debug
' To allow debugging, define _CDISPINVOKE_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CDISPINVOKE_DEBUG_
   #define _CDISPINVOKE_DEBUG_ 0
#ENDIF
#ifndef _CDISPINVOKE_DP_
   #define _CDISPINVOKE_DP_ 1
   #MACRO CDISPINVOKE_DP(st)
      #IF (_CDISPINVOKE_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
'                               *** CDispInvoke class ***
' ########################################################################################
TYPE CDispInvoke

Public:
   m_pDisp AS IDispatch PTR
   m_Result AS HRESULT
   m_lcid AS LCID = LOCALE_USER_DEFAULT
   m_uArgErr AS UINT
   m_dispParams AS DISPPARAMS
   m_varResult AS VARIANT

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL vDisp AS VARIANT, BYVAL fAddRef AS BOOLEAN = TRUE)
   DECLARE CONSTRUCTOR (BYVAL vDisp AS VARIANT PTR, BYVAL fAddRef AS BOOLEAN = TRUE)
   DECLARE CONSTRUCTOR (BYREF cvDisp AS DVARIANT)
   DECLARE CONSTRUCTOR (BYREF wszProgID AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "")
   DECLARE CONSTRUCTOR (BYREF wszClsID AS CONST WSTRING, BYREF wszIID AS CONST WSTRING)
   DECLARE CONSTRUCTOR (BYREF classID AS CONST CLSID)
   DECLARE CONSTRUCTOR (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID)
   DECLARE CONSTRUCTOR (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "")
   DECLARE CONSTRUCTOR (BYREF wszLibName AS CONST WSTRING, BYREF wszClsid AS CONST WSTRING, BYREF wszIid AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "")
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS IDispatch PTR PTR
   DECLARE FUNCTION DispPtr () AS IDispatch PTR
   DECLARE FUNCTION DispObj () AS IDispatch PTR
   DECLARE FUNCTION vptr () AS IDispatch PTR PTR
   DECLARE SUB Clear
   DECLARE OPERATOR Let (BYVAL pdisp AS IDispatch PTR)
   DECLARE OPERATOR Let (BYVAL vDisp AS VARIANT PTR)
   DECLARE OPERATOR Let (BYVAL vDisp AS VARIANT)
   DECLARE OPERATOR Let (BYREF cvDisp AS DVARIANT)
   DECLARE FUNCTION Attach (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Attach (BYVAL vDisp AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Detach () AS IDispatch PTR
   DECLARE FUNCTION Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION GetLcid () AS LCID
   DECLARE SUB SetLcid (BYVAL _lcid AS LCID)

   DECLARE FUNCTION GetVarResult () AS DVARIANT
   DECLARE FUNCTION GetArgErr () AS UINT
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING

   DECLARE FUNCTION DispInvoke (BYVAL wFlags AS WORD, BYVAL dispID AS .DISPID, BYVAL prgArgs AS VARIANT PTR = NULL, BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
   DECLARE FUNCTION DispInvoke (BYVAL wFlags AS WORD, BYVAL pwszName AS WSTRING PTR, BYVAL prgArgs AS VARIANT PTR = NULL, BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT

   DECLARE FUNCTION Get (BYVAL dispID AS DISPID, BYVAL prgArgs AS VARIANT PTR, BYVAL cArgs AS UINT, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
   DECLARE FUNCTION Get (BYVAL pwszName AS WSTRING PTR, BYVAL prgArgs AS VARIANT PTR, BYVAL cArgs AS UINT, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
   DECLARE FUNCTION Get (BYVAL pwszName AS WSTRING PTR) AS DVARIANT
   DECLARE FUNCTION Get (BYVAL dispID AS DISPID) AS DVARIANT
   DECLARE FUNCTION Get (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Get (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Get (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Get (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT

   DECLARE PROPERTY Put (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT)
   DECLARE PROPERTY Put (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT)

   DECLARE FUNCTION Set (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Set (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Set (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Set (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS HRESULT

   DECLARE PROPERTY PutRef (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT)
   DECLARE PROPERTY PutRef (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT)
   DECLARE PROPERTY PutRef (BYVAL dispID AS DISPID, BYVAL pv AS ANY PTR)
   DECLARE PROPERTY PutRef (BYVAL pwszName AS WSTRING PTR, BYVAL pv AS ANY PTR)

   DECLARE FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT) AS HRESULT
   DECLARE FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS HRESULT
   DECLARE FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   DECLARE FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT

   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT, BYREF dv16 AS DVARIANT) AS DVARIANT

   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF cvAg7 AS DVARIANT, BYREF cvAg8 AS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT, BYREF dv16 AS DVARIANT) AS DVARIANT


END TYPE
' ########################################################################################

' ########################################################################################
'                             *** CDispInvoke methods ***
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke
   CDISPINVOKE_DP("Default")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Constructors
' - fAddRef parameter:
' If it is true, then AddRef is called. This is needed when we pass a raw pointer.
' If it is false, the object takes ownership of the interface pointer without calling AddRef.
' This is the usual case when we assign an already AddRefed pointer returned by a COM method.
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYVAL pDisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CDISPINVOKE_DP("BEGIN - pDisp: " & WSTR(pDisp) & "- fAddRef" & WSTR(fAddRef))
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Assign the dispatch pointer
   IF pdisp THEN
      m_pDisp = pDisp
      IF fAddRef THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
   END IF
   CDISPINVOKE_DP("END - pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYVAL vDisp AS VARIANT PTR, BYVAL fAddRef AS BOOLEAN = TRUE)
   CDISPINVOKE_DP("BEGIN - VARIANT PTR: " & WSTR(vDisp) & " - fAddRef" & WSTR(fAddRef))
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Assign the dispatch pointer
   IF vDisp <> NULL THEN
      IF vDisp->vt = VT_DISPATCH THEN m_pDisp = vDisp->pdispVal
      IF m_pDisp THEN
         IF fAddRef THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
      END IF
   END IF
   CDISPINVOKE_DP("END - VARIANT PTR: pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYVAL vDisp AS VARIANT, BYVAL fAddRef AS BOOLEAN = TRUE)
   CDISPINVOKE_DP("BEGIN - BYVAL VARIANT")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Assign the dispatch pointer
   IF vDisp.vt = VT_DISPATCH THEN m_pDisp = vDisp.pdispVal
   IF m_pDisp THEN
      IF fAddRef THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
   END IF
   CDISPINVOKE_DP("END - BYVAL VARIANT: pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYREF cvDisp AS DVARIANT)
   CDISPINVOKE_DP("BEGIN - BYREF VARIANT")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Assign the dispatch pointer
   IF cvDisp.vd.vt = VT_DISPATCH THEN m_pDisp = cvDisp.vd.pdispVal
   IF m_pDisp THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
   CDISPINVOKE_DP("END - BYREF VARIANT: pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
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
' Usage example:
'   DIM pDI AS CDispInvoke = "Scripting.Dictionary"
' -or-
'   pDI = CDispInvoke(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
' CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYREF wszProgID AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "")
   CDISPINVOKE_DP("BEGIN - ProgID, LicKey")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the object
   m_pDisp = AfxNewCom(wszProgID, wszLicKey)
   CDISPINVOKE_DP("END - ProgID, LicKey - pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.
' Parameters:
' - wszClsID = A CLSID in string format.
' - wszIID = A reference to the identifier of the interface to be used to communicate with the object.
' Usage example:
'   DIM pDI AS CDispInvoke = CDispInvoke(CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'   CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' and IID_IDictionary as
'   CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYREF wszClsID AS CONST WSTRING, BYREF wszIID AS CONST WSTRING)
   CDISPINVOKE_DP("BEGIN - clsID, IID")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the object
   m_pDisp = AfxNewCom(wszClsID, wszIID)
   CDISPINVOKE_DP("END - clsID, IID - pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified CLSID.
' Parameter:
' - classID = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' Usage examples:
'   DIM pDI AS CDispInvoke = CDispInvoke(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYREF classID AS CONST CLSID)
   CDISPINVOKE_DP("BEGIN - classID")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the object
   DIM pUnk AS IUnknown PTR
   pUnk = AfxNewCom(classID)
   IF pUnk THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @m_pDisp)
      IF m_pDisp THEN pUnk->lpVtbl->Release(pUnk) ELSE m_pDisp = cast(IDispatch PTR, pUnk)
   END IF
   CDISPINVOKE_DP("END - classID - pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with the specified CLSID and IID.
' Parameters:
' - classID = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' Usage examples:
'   DIM pDI AS CDispInvoke = CDispInvoke(CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' and IID_IDictionary as
'   DIM IID_IDictionary AS IID = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID)
   CDISPINVOKE_DP("BEGIN - classID, riid")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the object
   DIM pUnk AS IUnknown PTR
   pUnk = AfxNewCom(classID, riid)
   IF pUnk THEN
      pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @m_pDisp)
      IF m_pDisp THEN pUnk->lpVtbl->Release(pUnk) ELSE m_pDisp = cast(IDispatch PTR, pUnk)
   END IF
   CDISPINVOKE_DP("END - classID, riid - pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Loads the specified library from file and creates an instance of an object.
' Parameters:
' - wszLibName = Full path where the library is located.
' - rclsid = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' - wszLicKey = The license key.
' Not every component is a suitable candidate for use under this overloaded function.
'  - Only in-process servers (DLLs) are supported.
'  - Components that are system components or part of the operating system, such as XML,
'    Data Access, Internet Explorer, or DirectX, aren't supported
'  - Components that are part of an application, such Microsoft Office, aren't supported.
'  - Components intended for use as an add-in or a snap-in, such as an Office add-in or
'    a control in a Web browser, aren't supported.
'  - Components that manage a shared physical or virtual system resource aren't supported.
'  - Visual ActiveX controls aren't supported because they need to be initilized and
'    activated by the OLE container.
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "")
   CDISPINVOKE_DP("BEGIN - LibName, cldID, riid, LicKey")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the object
   m_pDisp = AfxNewCom(wszLibName, rclsid, riid, wszLicKey)
   CDISPINVOKE_DP("END - LibName, cldID, riid, LicKey - pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Converts the wszClsid and wszIid parameters to GUIDs and calls the function above.
' ========================================================================================
PRIVATE CONSTRUCTOR CDispInvoke (BYREF wszLibName AS CONST WSTRING, BYREF wszClsid AS CONST WSTRING, BYREF wszIid AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "")
   CDISPINVOKE_DP("BEGIN - LibName, clsID, LicKey")
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the object
   m_pDisp = AfxNewCom(wszLibName, wszClsid, wszIid, wszLicKey)
   CDISPINVOKE_DP("END - LibName, clsID, LicKey - pDisp: " & WSTR(m_pDisp))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CDispInvoke class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDispInvoke
   CDISPINVOKE_DP(" pDisp = " & WSTR(m_pDisp))
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the address of the dispatch pointer
' Removed to allow to use @ to get the address of the class.
' Use DisPtr or DispObj instead.
' ========================================================================================
'PRIVATE OPERATOR CDispInvoke.@ () AS IDispatch PTR PTR
'   CDISPINVOKE_DP("CDISPINVOKE OPERATOR @")
'   OPERATOR = @m_pDisp
'END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the underlying dispatch pointer (same as *).
' Don't call IUnknown_Release on it.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.DispPtr () AS IDispatch PTR
   CDISPINVOKE_DP("DispPtr" & WSTR(m_pDisp))
   RETURN m_pDisp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a counted reference of the underlying dispatch pointer.
' You must call IUnknown_Release when no longer need it.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.DispObj () AS IDispatch PTR
   CDISPINVOKE_DP("DispObj" & WSTR(m_pDisp))
   IF m_pDisp <> NULL THEN
      IUnknown_AddRef(m_pDisp)
      RETURN m_pDisp
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the underlying IDispatch pointer.
' ========================================================================================
PRIVATE OPERATOR * (BYREF cdi AS CDispInvoke) AS IDispatch PTR
   CDISPINVOKE_DP(WSTR(cdi.m_pDisp))
   OPERATOR = cdi.m_pDisp
END OPERATOR
' ========================================================================================

' ========================================================================================
' Clears the contents of the class and returns the address of the underlying IDispatch pointer.
' To pass the class to an OUT BYVAL IDispatch PTR parameter.
' If we pass a CDispInvoke to a function with an OUT IDIspatch parameter without first
' clearing the contents of the class, we may have a memory leak.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.vptr () AS IDIspatch PTR PTR
   CDISPINVOKE_DP(WSTR(@m_pDisp))
   this.Clear
   RETURN @m_pDisp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Clears the contents of the variables. Needed if we want to reuse the class to pass it to
' a method with vptr to an out Dispatch PTR PTR parameter.
' ========================================================================================
PRIVATE SUB CDispInvoke.Clear
   CDISPINVOKE_DP("")
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   m_pDisp = NULL
END SUB
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE OPERATOR CDispInvoke.Let (BYVAL pDisp AS IDispatch PTR)
   CDISPINVOKE_DP("IDispatch PTR - " & WSTR(pDisp))
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   m_pDisp = NULL
   ' // Assign the new dispatch pointer
   IF pDisp = NULL THEN EXIT OPERATOR
   m_pDisp = pDisp
   IF m_pDisp THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE OPERATOR CDispInvoke.Let (BYVAL vDisp AS VARIANT PTR)
   CDISPINVOKE_DP("VARIANT PTR " & WSTR(vDisp))
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   m_pDisp = NULL
   ' // Assign the new dispatch pointer
   IF vDisp = NULL THEN EXIT OPERATOR
   IF vDisp->vt <> VT_DISPATCH THEN EXIT OPERATOR
   IF vDisp->pdispVal = NULL THEN EXIT OPERATOR
   m_pDisp = vDisp->pdispVal
   IF m_pDisp THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE OPERATOR CDispInvoke.Let (BYVAL vDisp AS VARIANT)
   CDISPINVOKE_DP("VARIANT")
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   m_pDisp = NULL
   ' // Assign the new dispatch pointer
   IF vDisp.vt <> VT_DISPATCH THEN EXIT OPERATOR
   IF vDisp.pdispVal = NULL THEN EXIT OPERATOR
   m_pDisp = vDisp.pdispVal
   IF m_pDisp THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE OPERATOR CDispInvoke.Let (BYREF cvDisp AS DVARIANT)
   CDISPINVOKE_DP("DVARIANT")
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   m_pDisp = NULL
   ' // Assign the new dispatch pointer
   IF cvDisp.vd.vt <> VT_DISPATCH THEN EXIT OPERATOR
   IF cvDisp.vd.pdispVal = NULL THEN EXIT OPERATOR
   m_pDisp = cvDisp.vd.pdispVal
   IF m_pDisp THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Attaches a Dispatch pointer
' fAddRef: If it is true, then AddRef is called. If it is false, the object takes ownership
' of the raw interface pointer without calling AddRef.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Attach (BYVAL pDisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   CDISPINVOKE_DP("IDispatch PTR: " & WSTR(pdisp) & " - fAddRef: " & WSTR(fAddRef))
   IF pdisp = NULL THEN RETURN E_INVALIDARG
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   ' // Assign the new dispatch pointer
   m_pDisp = pDisp
   IF fAddRef THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches a Dispatch pointer
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Attach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CDISPINVOKE_DP("VARIANT PTR: " & WSTR(pvar))
   IF pvar = NULL THEN RETURN E_INVALIDARG
   IF pvar->vt <> VT_DISPATCH THEN RETURN E_INVALIDARG
   IF pvar->pdispVal = NULL THEN RETURN E_INVALIDARG
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   ' // Release the dispatch pointer
   IF m_pDisp THEN m_pDisp->lpvtbl->Release(m_pDisp)
   ' // Assign the new dispatch pointer
   m_pDisp = pvar->pdispVal
   IF m_pDisp THEN m_pDisp->lpvtbl->AddRef(m_pDisp)
   ' // Mark the source variant as VT_EMPTY instead of clearing it with VariantClear
   ' // because we aren't making a duplicate of the contents, but transfering ownership.
   pvar->vt = VT_EMPTY
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts and returns the encapsulated interface pointer, and then clears the encapsulated
' pointer storage to NULL. This removes the interface pointer from encapsulation. It is up
' to you to call Release on the returned interface pointer.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Detach () AS IDispatch PTR
   CDISPINVOKE_DP("IDispatch PTR: " & WSTR(m_pDisp))
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   DIM pTmp AS IDispatch PTR = m_pDisp
   m_pDisp = NULL
   RETURN pTmp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the Dispatch pointer and transfers ownership to a variant
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   CDISPINVOKE_DP("VARIANT PTR: " & WSTR(pvar))
   IF m_pDisp = NULL THEN RETURN E_POINTER
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   pvar->vt = VT_DISPATCH
   pvar->pdispVal = m_pDisp
   ' // Free the memory used by the result variant
   VariantClear(@m_varResult)
   m_pDisp = NULL
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" & HEX(nError) & CHR(13, 10) & dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the locale identifier.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.GetLcid () AS LCID
   CDISPINVOKE_DP("")
   RETURN m_lcid
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the locale identifier.
' ========================================================================================
PRIVATE SUB CDispInvoke.SetLcid (BYVAL _lcid AS LCID)
   CDISPINVOKE_DP("")
   m_lcid = _lcid
END SUB
' ========================================================================================

' ========================================================================================
' The index within rgvarg of the first argument that has an error. Arguments are stored in
' pDispParams->rgvarg in reverse order, so the first argument is the one with the highest
' index in the array. This parameter is returned only when the resulting return value is
' DISP_E_TYPEMISMATCH or DISP_E_PARAMNOTFOUND.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.GetArgErr () AS UINT
   CDISPINVOKE_DP("")
   RETURN m_uArgErr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the result
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.GetVarResult () AS DVARIANT
   CDISPINVOKE_DP("")
   RETURN m_varResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' Wrapper function to call the Invoke method to call a method or get property.
' Parameters:
' - dispIdMember [in]: Identifies the member. Use GetIDsOfNames or the object's documentation
'     to obtain the dispatch identifier.
' - wFlags [in]: Flags describing the context of the Invoke call.
'     DISPATCH_METHOD: The member is invoked as a method. If a property has the same name,
'     both this and the DISPATCH_PROPERTYGET flag can be set.
'     DISPATCH_PROPERTYGET: The member is retrieved as a property or data member.
'     DISPATCH_PROPERTYPUT: The member is changed as a property or data member.
'     DISPATCH_PROPERTYPUTREF: The member is changed by a reference assignment, rather than a
'     value assignment. This flag is valid only when the property accepts a reference to an object.
' - prgArgs: Pointer to an array of arguments in reversed order.
' - cArgs: Number of arguments in the array.
' - lcid [in] : The locale context in which to interpret arguments. The lcid is used by the
'     GetIDsOfNames method, and is also passed to Invoke to allow the object to interpret its
'     arguments specific to a locale. Applications that do not support multiple national
'     languages can ignore this parameter.
' Return value
'   This method can return one of these values.
'   S_OK: Success.
'   DISP_E_BADPARAMCOUNT: The number of elements provided to DISPPARAMS is different from the
'   number of arguments accepted by the method or property.
'   DISP_E_BADVARTYPE: One of the arguments in DISPPARAMS is not a valid variant type.
'   DISP_E_EXCEPTION: The application needs to raise an exception. In this case, the structure
'   passed in pexcepinfo should be filled in.
'   DISP_E_MEMBERNOTFOUND: The requested member does not exist.
'   DISP_E_NONAMEDARGS: This implementation of IDispatch does not support named arguments.
'   DISP_E_OVERFLOW: One of the arguments in DISPPARAMS could not be coerced to the specified type.
'   DISP_E_PARAMNOTFOUND: One of the parameter IDs does not correspond to a parameter on the
'   method. In this case, puArgErr is set to the first argument that contains the error.
'   DISP_E_TYPEMISMATCH: One or more of the arguments could not be coerced. The index of the
'   first parameter with the incorrect type within rgvarg is returned in puArgErr.
'   DISP_E_UNKNOWNINTERFACE: The interface identifier passed in riid is not IID_NULL.
'   DISP_E_UNKNOWNLCID: The member being invoked interprets string arguments according to the
'   LCID, and the LCID is not recognized. If the LCID is not needed to interpret arguments,
'   this error should not be returned
'   DISP_E_PARAMNOTOPTIONAL: A required parameter was omitted.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.DispInvoke (BYVAL wFlags AS WORD, BYVAL dispIdMember AS DISPID, _
BYVAL prgArgs AS VARIANT PTR = NULL, BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
   CDISPINVOKE_DP("1")
   ' // Clear the variables
   m_uArgErr = 0
   VariantClear @m_varResult
   memset @m_dispParams, 0, SIZEOF(m_dispParams)
   IF m_pDisp = NULL THEN RETURN E_POINTER
   ' // Fill the DISPPARAMS structure
   m_dispParams.cArgs = cArgs
   m_dispParams.rgvarg = prgArgs
   IF wFlags = DISPATCH_PROPERTYPUT OR wFlags = DISPATCH_PROPERTYPUTREF THEN
      m_dispParams.cNamedArgs = 1
      DIM dispidNamed AS DISPID = DISPID_PROPERTYPUT
      m_dispParams.rgdispidNamedArgs = @dispidNamed
   END IF
   ' // Call the Invoke method
   RETURN SetResult(m_pDisp->lpvtbl->Invoke(m_pDisp, dispIdMember, @IID_NULL, lcid, wFlags, @m_dispParams, @m_varResult, NULL, @m_uArgErr))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the DISPID (dispatch identifier) and calls the function above.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.DispInvoke (BYVAL wFlags AS WORD, BYVAL pwszName AS WSTRING PTR, _
BYVAL prgArgs AS VARIANT PTR = NULL, BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
   CDISPINVOKE_DP("2")
   ' // Clear the variables
   m_uArgErr = 0
   VariantClear @m_varResult
   memset @m_dispParams, 0, SIZEOF(m_dispParams)
   IF m_pDisp = NULL THEN RETURN E_POINTER
   IF pwszName = NULL THEN RETURN E_INVALIDARG
   DIM hr AS HRESULT, dispID AS LONG
   hr = m_pDisp->lpvtbl->GetIdsOfNames(m_pDisp, @IID_NULL, @pwszName, 1, lcid, @dispID)
   IF hr <> S_OK THEN RETURN hr
   RETURN SetResult(this.DispInvoke(wFlags, dispID, prgArgs, cArgs, lcid))
END FUNCTION
' ========================================================================================

' ****************************************************************************************
'                                   *** GET METHODS ***
' ****************************************************************************************

' ========================================================================================
' Wrapper function to call the Invoke method to get a property value.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL dispID AS DISPID, BYVAL prgArgs AS VARIANT PTR = NULL, BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
   CDISPINVOKE_DP("1 - DISPID")
   RETURN SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, prgArgs, cArgs, lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL pwszName AS WSTRING PTR, BYVAL prgArgs AS VARIANT PTR = NULL, BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
   CDISPINVOKE_DP("1 - Name")
   RETURN SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, prgArgs, cArgs, lcid))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL dispID AS DISPID) AS DVARIANT
   CDISPINVOKE_DP("2 - DISPID")
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, NULL, 0, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL pwszName AS WSTRING PTR) AS DVARIANT
   CDISPINVOKE_DP("Name")
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, NULL, 0, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("1 - DispId - DVARIANT")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 1, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("1 - Name - DVARIANT")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 1, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("2 - Name - DVARIANT")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 2, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Get (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("2 - Name - DVARIANT")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 2, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================

' ****************************************************************************************
'                                   *** PUT METHODS ***
' ****************************************************************************************

' ========================================================================================
' Wrapper functions to call the Invoke method to set a property value.
' ========================================================================================
PRIVATE PROPERTY CDispInvoke.Put (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT)
   CDISPINVOKE_DP("DISPID - DVARIANT")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, dispID, @vArgs(1), 1, m_lcid))
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CDispInvoke.Put (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT)
   CDISPINVOKE_DP("Name - DVARIANT")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, pwszName, @vArgs(1), 1, m_lcid))
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("dispID - DVARIANT1")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, dispID, @vArgs(1), 1, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("dispID - DVARIANT2")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, dispID, @vArgs(1), 2, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("dispID - DVARIANT 3")
   DIM vArgs(1 TO 3) AS VARIANT : vArgs(1) = dv3 : vArgs(2) = dv2 : vArgs(3) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, dispID, @vArgs(1), 3, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("dispID - DVARIANT 4")
   DIM vArgs(1 TO 4) AS VARIANT : vArgs(1) = dv4 : vArgs(2) = dv3 : vArgs(3) = dv2 : vArgs(4) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, dispID, @vArgs(1), 3, m_lcid))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("Name - DVARIANT1")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, pwszName, @vArgs(1), 1, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("Name - DVARIANT2")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, pwszName, @vArgs(1), 2, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("Name - DVARIANT 3")
   DIM vArgs(1 TO 3) AS VARIANT : vArgs(1) = dv3 : vArgs(2) = dv2 : vArgs(3) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, pwszName, @vArgs(1), 3, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Set (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("Name - DVARIANT 4")
   DIM vArgs(1 TO 4) AS VARIANT : vArgs(1) = dv4 : vArgs(2) = dv3 : vArgs(3) = dv2 : vArgs(4) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUT, pwszName, @vArgs(1), 3, m_lcid))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Wrapper functions to call the Invoke method to call a putref property.
' ========================================================================================
PRIVATE PROPERTY CDispInvoke.PutRef (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT)
   CDISPINVOKE_DP("DISPID - DVARIANT")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, dispID, @vArgs(1), 1, m_lcid))
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CDispInvoke.PutRef (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT)
   CDISPINVOKE_DP("Name - DVARIANT")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, pwszName, @vArgs(1), 1, m_lcid))
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CDispInvoke.PutRef (BYVAL dispID AS DISPID, BYVAL pv AS ANY PTR)
   CDISPINVOKE_DP("DISPID - ANY PTR")
   DIM dv AS DVARIANT = DVARIANT(CAST(IUnknown PTR, pv), TRUE)
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, dispID, @vArgs(1), 1, m_lcid))
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CDispInvoke.PutRef (BYVAL pwszName AS WSTRING PTR, BYVAL pv AS ANY PTR)
   CDISPINVOKE_DP("Name - ANY PTR")
   DIM dv AS DVARIANT = DVARIANT(CAST(IUnknown PTR, pv), TRUE)
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, pwszName, @vArgs(1), 1, m_lcid))
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CDispInvoke.SetRef (BYVAL dispID AS DISPID, BYREF dv AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("dispID - DVARIANT1")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, dispID, @vArgs(1), 1, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.SetRef (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("dispID - DVARIANT2")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, dispID, @vArgs(1), 2, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.SetRef (BYVAL dispID AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("dispID - DVARIANT 3")
   DIM vArgs(1 TO 3) AS VARIANT : vArgs(1) = dv3 : vArgs(2) = dv2 : vArgs(3) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, dispID, @vArgs(1), 3, m_lcid))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CDispInvoke.SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("Name - DVARIANT1")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, pwszName, @vArgs(1), 1, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("Name - DVARIANT2")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, pwszName, @vArgs(1), 2, m_lcid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.SetRef (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS HRESULT
   CDISPINVOKE_DP("Name - DVARIANT 3")
   DIM vArgs(1 TO 3) AS VARIANT : vArgs(1) = dv3 : vArgs(2) = dv2 : vArgs(3) = dv1
   RETURN SetResult(this.DispInvoke(DISPATCH_PROPERTYPUTREF, pwszName, @vArgs(1), 3, m_lcid))
END FUNCTION
' ========================================================================================


' ****************************************************************************************
'                                *** INVOKE METHODS ***
' ****************************************************************************************

' ========================================================================================
' Wrapper functions to call the Invoke method to call a method or get property.
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 0")
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, NULL, 0, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 1")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 1, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 2")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 2, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 3")
   DIM vArgs(1 TO 3) AS VARIANT : vArgs(1) = dv3 : vArgs(2) = dv2 : vArgs(3) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 3, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 4")
   DIM vArgs(1 TO 4) AS VARIANT : vArgs(1) = dv4 : vArgs(2) = dv3 : vArgs(3) = dv2 : vArgs(4) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 4, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 5")
   DIM vArgs(1 TO 5) AS VARIANT : vArgs(1) = dv5 : vArgs(2) = dv4 : vArgs(3) = dv3 : vArgs(4) = dv2 : vArgs(5) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 5, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 6")
   DIM vArgs(1 TO 6) AS VARIANT : vArgs(1) = dv6 : vArgs(2) = dv5 : vArgs(3) = dv4 : vArgs(4) = dv3 : vArgs(5) = dv2 : vArgs(6) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 6, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 7")
   DIM vArgs(1 TO 7) AS VARIANT : vArgs(1) = dv7 : vArgs(2) = dv6 : vArgs(3) = dv5 : vArgs(4) = dv4 : vArgs(5) = dv3 : vArgs(6) = dv2 : vArgs(7) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 7, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 8")
   DIM vArgs(1 TO 8) AS VARIANT : vArgs(1) = dv8 : vArgs(2) = dv7 : vArgs(3) = dv6 : vArgs(4) = dv5 : vArgs(5) = dv4 : vArgs(6) = dv3 : vArgs(7) = dv2 : vArgs(8) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 8, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 9")
   DIM vArgs(1 TO 9) AS VARIANT : vArgs(1) = dv9 : vArgs(2) = dv8 : vArgs(3) = dv7 : vArgs(4) = dv6 : vArgs(5) = dv5 : vArgs(6) = dv4 : vArgs(7) = dv3 : vArgs(8) = dv2 : vArgs(9) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 9, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 10")
   DIM vArgs(1 TO 10) AS VARIANT : vArgs(1) = dv10 : vArgs(2) = dv9 : vArgs(3) = dv8 : vArgs(4) = dv7 : vArgs(5) = dv6 : vArgs(6) = dv5 : vArgs(7) = dv4 : vArgs(8) = dv3 : vArgs(9) = dv2 : vArgs(10) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 10, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 11")
   DIM vArgs(1 TO 11) AS VARIANT : vArgs(1) = dv11 : vArgs(2) = dv10 : vArgs(3) = dv9 : vArgs(4) = dv8 : vArgs(5) = dv7 : vArgs(6) = dv6 : vArgs(7) = dv5 : vArgs(8) = dv4 : vArgs(9) = dv3 : vArgs(10) = dv2 : vArgs(11) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 11, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 12")
   DIM vArgs(1 TO 12) AS VARIANT : vArgs(1) = dv12 : vArgs(2) = dv11 : vArgs(3) = dv10 : vArgs(4) = dv9 : vArgs(5) = dv8 : vArgs(6) = dv7 : vArgs(7) = dv6 : vArgs(8) = dv5 : vArgs(9) = dv4 : vArgs(10) = dv3 : vArgs(11) = dv2 : vArgs(12) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 12, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 13")
   DIM vArgs(1 TO 13) AS VARIANT : vArgs(1) = dv13 : vArgs(2) = dv12 : vArgs(3) = dv11 : vArgs(4) = dv10 : vArgs(5) = dv9 : vArgs(6) = dv8 : vArgs(7) = dv7 : vArgs(8) = dv6 : vArgs(9) = dv5 : vArgs(10) = dv4 : vArgs(11) = dv3 : vArgs(12) = dv2 : vArgs(13) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 13, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 14")
   DIM vArgs(1 TO 14) AS VARIANT : vArgs(1) = dv14 : vArgs(2) = dv13 : vArgs(3) = dv12 : vArgs(4) = dv11 : vArgs(5) = dv10 : vArgs(6) = dv9 : vArgs(7) = dv8 : vArgs(8) = dv7 : vArgs(9) = dv6 : vArgs(10) = dv5 : vArgs(11) = dv4 : vArgs(12) = dv3 : vArgs(13) = dv2 : vArgs(14) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 14, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 15")
   DIM vArgs(1 TO 15) AS VARIANT : vArgs(1) = dv15 : vArgs(2) = dv14 : vArgs(3) = dv13 : vArgs(4) = dv12 : vArgs(5) = dv11 : vArgs(6) = dv10 : vArgs(7) = dv9 : vArgs(8) = dv8 : vArgs(9) = dv7 : vArgs(10) = dv6 : vArgs(11) = dv5 : vArgs(12) = dv4 : vArgs(13) = dv3 : vArgs(14) = dv2 : vArgs(15) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 15, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL dispId AS DISPID, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT, BYREF dv16 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("DISPID - DVARIANT 16")
   DIM vArgs(1 TO 16) AS VARIANT : vArgs(1) = dv16 : vArgs(2) = dv15 : vArgs(3) = dv14 : vArgs(4) = dv13 : vArgs(5) = dv12 : vArgs(6) = dv11 : vArgs(7) = dv10 : vArgs(8) = dv9 : vArgs(9) = dv8 : vArgs(10) = dv7 : vArgs(11) = dv6 : vArgs(12) = dv5 : vArgs(13) = dv4 : vArgs(14) = dv3 : vArgs(15) = dv2 : vArgs(16) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, dispID, @vArgs(1), 16, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR) AS DVARIANT
   CDISPINVOKE_DP("NAME - 0")
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, NULL, 0, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 1")
   DIM vArgs(1 TO 1) AS VARIANT : vArgs(1) = dv
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 1, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 2")
   DIM vArgs(1 TO 2) AS VARIANT : vArgs(1) = dv2 : vArgs(2) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 2, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 3")
   DIM vArgs(1 TO 3) AS VARIANT : vArgs(1) = dv3 : vArgs(2) = dv2 : vArgs(3) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 3, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 4")
   DIM vArgs(1 TO 4) AS VARIANT : vArgs(1) = dv4 : vArgs(2) = dv3 : vArgs(3) = dv2 : vArgs(4) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 4, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 5")
   DIM vArgs(1 TO 5) AS VARIANT : vArgs(1) = dv5 : vArgs(2) = dv4 : vArgs(3) = dv3 : vArgs(4) = dv2 : vArgs(5) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 5, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 6")
   DIM vArgs(1 TO 6) AS VARIANT : vArgs(1) = dv6 : vArgs(2) = dv5 : vArgs(3) = dv4 : vArgs(4) = dv3 : vArgs(5) = dv2 : vArgs(6) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 6, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 7")
   DIM vArgs(1 TO 7) AS VARIANT : vArgs(1) = dv7 : vArgs(2) = dv6 : vArgs(3) = dv5 : vArgs(4) = dv4 : vArgs(5) = dv3 : vArgs(6) = dv2 : vArgs(7) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 7, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 8")
   DIM vArgs(1 TO 8) AS VARIANT : vArgs(1) = dv8 : vArgs(2) = dv7 : vArgs(3) = dv6 : vArgs(4) = dv5 : vArgs(5) = dv4 : vArgs(6) = dv3 : vArgs(7) = dv2 : vArgs(8) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 8, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 9")
   DIM vArgs(1 TO 9) AS VARIANT : vArgs(1) = dv9 : vArgs(2) = dv8 : vArgs(3) = dv7 : vArgs(4) = dv6 : vArgs(5) = dv5 : vArgs(6) = dv4 : vArgs(7) = dv3 : vArgs(8) = dv2 : vArgs(9) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 9, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 10")
   DIM vArgs(1 TO 10) AS VARIANT : vArgs(1) = dv10 : vArgs(2) = dv9 : vArgs(3) = dv8 : vArgs(4) = dv7 : vArgs(5) = dv6 : vArgs(6) = dv5 : vArgs(7) = dv4 : vArgs(8) = dv3 : vArgs(9) = dv2 : vArgs(10) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 10, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 11")
   DIM vArgs(1 TO 11) AS VARIANT : vArgs(1) = dv11 : vArgs(2) = dv10 : vArgs(3) = dv9 : vArgs(4) = dv8 : vArgs(5) = dv7 : vArgs(6) = dv6 : vArgs(7) = dv5 : vArgs(8) = dv4 : vArgs(9) = dv3 : vArgs(10) = dv2 : vArgs(11) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 11, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 12")
   DIM vArgs(1 TO 12) AS VARIANT : vArgs(1) = dv12 : vArgs(2) = dv11 : vArgs(3) = dv10 : vArgs(4) = dv9 : vArgs(5) = dv8 : vArgs(6) = dv7 : vArgs(7) = dv6 : vArgs(8) = dv5 : vArgs(9) = dv4 : vArgs(10) = dv3 : vArgs(11) = dv2 : vArgs(12) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 12, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 13")
   DIM vArgs(1 TO 13) AS VARIANT : vArgs(1) = dv13 : vArgs(2) = dv12 : vArgs(3) = dv11 : vArgs(4) = dv10 : vArgs(5) = dv9 : vArgs(6) = dv8 : vArgs(7) = dv7 : vArgs(8) = dv6 : vArgs(9) = dv5 : vArgs(10) = dv4 : vArgs(11) = dv3 : vArgs(12) = dv2 : vArgs(13) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 13, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 14")
   DIM vArgs(1 TO 14) AS VARIANT : vArgs(1) = dv14 : vArgs(2) = dv13 : vArgs(3) = dv12 : vArgs(4) = dv11 : vArgs(5) = dv10 : vArgs(6) = dv9 : vArgs(7) = dv8 : vArgs(8) = dv7 : vArgs(9) = dv6 : vArgs(10) = dv5 : vArgs(11) = dv4 : vArgs(12) = dv3 : vArgs(13) = dv2 : vArgs(14) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 14, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 15")
   DIM vArgs(1 TO 15) AS VARIANT : vArgs(1) = dv15 : vArgs(2) = dv14 : vArgs(3) = dv13 : vArgs(4) = dv12 : vArgs(5) = dv11 : vArgs(6) = dv10 : vArgs(7) = dv9 : vArgs(8) = dv8 : vArgs(9) = dv7 : vArgs(10) = dv6 : vArgs(11) = dv5 : vArgs(12) = dv4 : vArgs(13) = dv3 : vArgs(14) = dv2 : vArgs(15) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 15, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDispInvoke.Invoke (BYVAL pwszName AS WSTRING PTR, BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT, BYREF dv3 AS DVARIANT, BYREF dv4 AS DVARIANT, BYREF dv5 AS DVARIANT, BYREF dv6 AS DVARIANT, BYREF dv7 AS DVARIANT, BYREF dv8 aS DVARIANT, BYREF dv9 AS DVARIANT, BYREF dv10 AS DVARIANT, BYREF dv11 AS DVARIANT, BYREF dv12 AS DVARIANT, BYREF dv13 AS DVARIANT, BYREF dv14 AS DVARIANT, BYREF dv15 AS DVARIANT, BYREF dv16 AS DVARIANT) AS DVARIANT
   CDISPINVOKE_DP("NAME - 16")
   DIM vArgs(1 TO 16) AS VARIANT : vArgs(1) = dv16 : vArgs(2) = dv15 : vArgs(3) = dv14 : vArgs(4) = dv13 : vArgs(5) = dv12 : vArgs(6) = dv11 : vArgs(7) = dv10 : vArgs(8) = dv9 : vArgs(9) = dv8 : vArgs(10) = dv7 : vArgs(11) = dv6 : vArgs(12) = dv5 : vArgs(13) = dv4 : vArgs(14) = dv3 : vArgs(15) = dv2 : vArgs(16) = dv1
   SetResult(this.DispInvoke(DISPATCH_METHOD OR DISPATCH_PROPERTYGET, pwszName, @vArgs(1), 16, m_lcid))
   RETURN m_varResult
END FUNCTION
' ========================================================================================

END NAMESPACE
