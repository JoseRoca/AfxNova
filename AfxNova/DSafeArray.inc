' ########################################################################################
' Platform: Microsoft Windows
' Filename: DSAFEARRAY.inc
' Purpose:  Implementation of the safe array data type.
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
#include once "AfxNova/DVARIANT.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' Macro for debug
' To allow debugging, define _DSAFEARRAY_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _DSAFEARRAY_DEBUG_
   #define _DSAFEARRAY_DEBUG_ 0
#ENDIF
#ifndef _DSAFEARRAY_DP_
   #define _DSAFEARRAY_DP_ 1
   #MACRO DSAFEARRAY_DP(st)
      #IF (_DSAFEARRAY_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
' DSafeArray - Safe array class
' ########################################################################################
TYPE DSafeArray

   m_psa AS SAFEARRAY PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   DECLARE CONSTRUCTOR (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DECLARE CONSTRUCTOR (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DECLARE CONSTRUCTOR (BYREF strType AS STRING, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   DECLARE CONSTRUCTOR (BYREF strType AS STRING, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DECLARE CONSTRUCTOR (BYREF strType AS STRING, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DECLARE CONSTRUCTOR (BYREF dsa AS DSafeArray)
   DECLARE CONSTRUCTOR (BYVAL psa AS SAFEARRAY PTR)
   DECLARE CONSTRUCTOR (BYVAL psa AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN)
   DECLARE CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR Let (BYREF dsa AS DSafeArray)
   DECLARE OPERATOR Let (BYVAL psa AS SAFEARRAY PTR)
   DECLARE OPERATOR Let (BYVAL pvar AS VARIANT PTR)
   DECLARE FUNCTION GetPtr () AS SAFEARRAY PTR
'   DECLARE OPERATOR @ () AS ANY PTR
'   DECLARE FUNCTION vptr () AS SAFEARRAY PTR
   DECLARE FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR) AS HRESULT
   DECLARE FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   DECLARE FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   DECLARE FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR, BYVAL pvExtra AS PVOID) AS HRESULT
   DECLARE FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DECLARE FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DECLARE FUNCTION CreateVector (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   DECLARE FUNCTION CreateVectorEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DECLARE FUNCTION NumDims () AS UINT
   DECLARE FUNCTION LBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION UBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION Count (BYVAL nDim AS UINT = 1) AS DWORD
   DECLARE FUNCTION ElemSize () AS UINT
   DECLARE FUNCTION IsResizable () AS BOOLEAN
   DECLARE FUNCTION GetType () AS VARTYPE
   DECLARE FUNCTION LocksCount () AS UINT
   DECLARE FUNCTION Flags () AS USHORT
   DECLARE FUNCTION Features () AS USHORT
   DECLARE FUNCTION AccessData () AS ANY PTR
   DECLARE FUNCTION UnaccessData () AS HRESULT
   DECLARE FUNCTION Redim (BYVAL pnewsabounds AS SAFEARRAYBOUND PTR) AS HRESULT
   DECLARE FUNCTION Redim (BYVAL cElements AS DWORD, BYVAL lLBound AS LONG) AS HRESULT
   DECLARE FUNCTION Redim (BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   DECLARE FUNCTION Copy () AS SAFEARRAY PTR
   DECLARE FUNCTION CopyData (BYVAL psaTarget AS SAFEARRAY PTR) AS HRESULT
   DECLARE FUNCTION CopyFrom (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   DECLARE FUNCTION PtrOfIndex (BYVAL prgIndices AS LONG PTR) AS ANY PTR
   DECLARE FUNCTION PtrOfIndex (BYVAL idx AS LONG) AS ANY PTR
   DECLARE FUNCTION PtrOfIndex (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS ANY PTR
   DECLARE FUNCTION Destroy () AS HRESULT
   DECLARE FUNCTION DestroyData () AS HRESULT
   DECLARE FUNCTION Clear () AS HRESULT
   DECLARE FUNCTION Erase () AS HRESULT
   DECLARE FUNCTION Reset () AS HRESULT
   DECLARE FUNCTION CopyFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION MoveFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION CopyToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION MoveToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Attach (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   DECLARE FUNCTION Detach () AS SAFEARRAY PTR
   DECLARE FUNCTION GetIID () AS GUID
   DECLARE FUNCTION SetIID (BYVAL pguid AS GUID PTR) AS HRESULT
   DECLARE FUNCTION GetRecordInfo () AS IRecordInfo PTR
   DECLARE FUNCTION SetRecordInfo (BYVAL prinfo AS IRecordInfo PTR) AS HRESULT
   DECLARE FUNCTION AppendElement (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION AppendElement (BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION AppendElement (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION AppendElement (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION Append (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Append (BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION Append (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Append (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION AppendStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AppendVar (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION InsertElement (BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION InsertElement (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION InsertElement (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION Insert (BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION Insert (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Insert (BYVAL vData AS VARIANT) AS HRESULT
   DECLARE FUNCTION InsertStr (BYVAL nPos AS LONG, BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION InsertStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION InsertVar (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION InsertVar (BYVAL nPos AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION DeleteElement (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION Remove (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION DeleteStringElement (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION DeleteVariantElement (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION RemoveStr (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION RemoveVar (BYVAL nPos AS LONG) AS HRESULT
   DECLARE FUNCTION FindElement (BYREF wszFind AS WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
   DECLARE FUNCTION Find (BYREF wszFind AS WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
   DECLARE FUNCTION Sort (BYVAL bAscend AS BOOLEAN = TRUE) AS HRESULT
   ' // ANY PTR
   DECLARE FUNCTION GetElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   ' // BYREF BSTRING
   DECLARE FUNCTION GetElement (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION Get (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION GetStr (BYVAL prgIndices AS LONG PTR) AS BSTRING
   DECLARE FUNCTION GetStr (BYVAL idx AS LONG) AS BSTRING
   DECLARE FUNCTION GetStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS BSTRING
   ' // BYREF DVARIANT
   DECLARE FUNCTION GetElement (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Get (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION GetVar (BYVAL prgIndices AS LONG PTR) AS DVARIANT
   DECLARE FUNCTION GetVar (BYVAL idx AS LONG) AS DVARIANT
   DECLARE FUNCTION GetVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS DVARIANT
   ' // ANY PTR
   DECLARE FUNCTION PutElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Put (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Put (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DECLARE FUNCTION Put (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   ' // BYREF BSTRING
   DECLARE FUNCTION PutElement (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION PutStr (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION PutStr (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DECLARE FUNCTION PutStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   ' // BYREF DVARIANT
   DECLARE FUNCTION PutElement (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION PutVar (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION PutVar (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION PutVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR DSafeArray
   DSAFEARRAY_DP("Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array.
' Parameters:
' - vtType: The base type of the array (the VARTYPE of each element of the array). The
'   VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the
'   VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array.
'   All other types are legal.
' - cDims: The number of dimensions in the array. The number cannot be changed after the
'   array is created.
' - lLBound: The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' - cElements: The number of elements in the array.
' ========================================================================================
' // Multidimensional array
PRIVATE CONSTRUCTOR DSafeArray (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   m_psa = SafeArrayCreate(vt, cDims, prgsabounds)
   DSafeArray_DP("multi-D - " & ..WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // One-dimensional array
PRIVATE CONSTRUCTOR DSafeArray (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   m_psa = SafeArrayCreate(vt, 1, @rgsabounds(0))
   DSafeArray_DP("1D - " & ..WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // Two-dimensional array
PRIVATE CONSTRUCTOR DSafeArray (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   m_psa = SafeArrayCreate(vt, 2, @rgsabounds(0))
   DSafeArray_DP("1D - " & ..WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // Multidimensional array
PRIVATE CONSTRUCTOR DSafeArray (BYREF strType AS STRING, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "STRING"      : vt = VT_BSTR
      CASE "BSTR"        : vt = VT_BSTR
      CASE "DWSTRING"    : vt = VT_BSTR
      CASE "BSTRING"     : vt = VT_BSTR
      CASE "BOOL"        : vt = VT_BOOL
      CASE "BOOLEAN"     : vt = VT_BOOL
      CASE "BYTE"        : vt = VT_I1
      CASE "UBYTE"       : vt = VT_UI1
      CASE "SHORT"       : vt = VT_I2
      CASE "USHORT"      : vt = VT_UI2
      CASE "INT"         : vt = VT_INT
      CASE "UINT"        : vt = VT_UINT
      CASE "LONG"        : vt = VT_I4
      CASE "ULONG"       : vt = VT_UI4
      CASE "LONGINT"     : vt = VT_I8
      CASE "ULONGINT"    : vt = VT_UI8
      CASE "SINGLE"      : vt = VT_R4
      CASE "FLOAT"       : vt = VT_R4
      CASE "DOUBLE"      : vt = VT_R8
      CASE "VARIANT"     : vt = VT_VARIANT
      CASE "DVARIANT"    : vt = VT_VARIANT
      CASE "CY"          : vt = VT_CY
      CASE "CURRENCY"    : vt = VT_CY
      CASE "DECIMAL"     : vt = VT_DECIMAL
      CASE "UNKNOWN"     : vt = VT_UNKNOWN
      CASE "DISPATCH"    : vt = VT_DISPATCH
      CASE "FILETIME"    : vt = VT_FILETIME
      CASE "DATE"        : vt = VT_DATE
      CASE "USERDEFINED" : vt = VT_USERDEFINED
      CASE "PTR"         : vt = VT_PTR
   END SELECT
   m_psa = SafeArrayCreate(vt, cDims, prgsabounds)
   DSafeArray_DP("multi-D - " & ..WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // One-dimensional array
PRIVATE CONSTRUCTOR DSafeArray (BYREF strType AS STRING, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "STRING"      : vt = VT_BSTR
      CASE "BSTR"        : vt = VT_BSTR
      CASE "DWSTRING"    : vt = VT_BSTR
      CASE "BSTRING"     : vt = VT_BSTR
      CASE "BOOL"        : vt = VT_BOOL
      CASE "BOOLEAN"     : vt = VT_BOOL
      CASE "BYTE"        : vt = VT_I1
      CASE "UBYTE"       : vt = VT_UI1
      CASE "SHORT"       : vt = VT_I2
      CASE "USHORT"      : vt = VT_UI2
      CASE "INT"         : vt = VT_INT
      CASE "UINT"        : vt = VT_UINT
      CASE "LONG"        : vt = VT_I4
      CASE "ULONG"       : vt = VT_UI4
      CASE "LONGINT"     : vt = VT_I8
      CASE "ULONGINT"    : vt = VT_UI8
      CASE "SINGLE"      : vt = VT_R4
      CASE "FLOAT"       : vt = VT_R4
      CASE "DOUBLE"      : vt = VT_R8
      CASE "VARIANT"     : vt = VT_VARIANT
      CASE "DVARIANT"    : vt = VT_VARIANT
      CASE "CY"          : vt = VT_CY
      CASE "CURRENCY"    : vt = VT_CY
      CASE "DECIMAL"     : vt = VT_DECIMAL
      CASE "UNKNOWN"     : vt = VT_UNKNOWN
      CASE "DISPATCH"    : vt = VT_DISPATCH
      CASE "FILETIME"    : vt = VT_FILETIME
      CASE "DATE"        : vt = VT_DATE
      CASE "USERDEFINED" : vt = VT_USERDEFINED
      CASE "PTR"         : vt = VT_PTR
   END SELECT
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   m_psa = SafeArrayCreate(vt, 1, @rgsabounds(0))
   DSafeArray_DP("1D - " & ..WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // Two-dimensional array
PRIVATE CONSTRUCTOR DSafeArray (BYREF strType AS STRING, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "STRING"      : vt = VT_BSTR
      CASE "BSTR"        : vt = VT_BSTR
      CASE "DWSTRING"    : vt = VT_BSTR
      CASE "BSTRING"     : vt = VT_BSTR
      CASE "BOOL"        : vt = VT_BOOL
      CASE "BOOLEAN"     : vt = VT_BOOL
      CASE "BYTE"        : vt = VT_I1
      CASE "UBYTE"       : vt = VT_UI1
      CASE "SHORT"       : vt = VT_I2
      CASE "USHORT"      : vt = VT_UI2
      CASE "INT"         : vt = VT_INT
      CASE "UINT"        : vt = VT_UINT
      CASE "LONG"        : vt = VT_I4
      CASE "ULONG"       : vt = VT_UI4
      CASE "LONGINT"     : vt = VT_I8
      CASE "ULONGINT"    : vt = VT_UI8
      CASE "SINGLE"      : vt = VT_R4
      CASE "FLOAT"       : vt = VT_R4
      CASE "DOUBLE"      : vt = VT_R8
      CASE "VARIANT"     : vt = VT_VARIANT
      CASE "DVARIANT"    : vt = VT_VARIANT
      CASE "CY"          : vt = VT_CY
      CASE "CURRENCY"    : vt = VT_CY
      CASE "DECIMAL"     : vt = VT_DECIMAL
      CASE "UNKNOWN"     : vt = VT_UNKNOWN
      CASE "DISPATCH"    : vt = VT_DISPATCH
      CASE "FILETIME"    : vt = VT_FILETIME
      CASE "DATE"        : vt = VT_DATE
      CASE "USERDEFINED" : vt = VT_USERDEFINED
      CASE "PTR"         : vt = VT_PTR
   END SELECT
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   m_psa = SafeArrayCreate(vt, 2, @rgsabounds(0))
   DSafeArray_DP("1D - " & ..WSTR(m_psa))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from another DSafeArray.
' ========================================================================================
PRIVATE CONSTRUCTOR DSafeArray (BYREF dsa AS DSafeArray)
   DSafeArray_DP("dsa.m_psa = " & ..WSTR(dsa.m_psa))
   this.CopyFrom(dsa.m_psa)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from another safe array.
' ========================================================================================
PRIVATE CONSTRUCTOR DSafeArray (BYVAL psa AS SAFEARRAY PTR)
   DSafeArray_DP("psa = " & ..WSTR(psa))
   this.CopyFrom(psa)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from another safe array.
' If fAttach = TRUE, the safe array is attached, else a copy is made.
' ========================================================================================
PRIVATE CONSTRUCTOR DSafeArray (BYVAL psa AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN)
   DSafeArray_DP("psa = " & ..WSTR(psa))
   IF fAttach THEN this.Attach(psa) ELSE this.CopyFrom(psa)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a safe array from a variant.
' ========================================================================================
PRIVATE CONSTRUCTOR DSafeArray (BYVAL pvar AS VARIANT PTR)
   DSafeArray_DP("VARIANT - " & ..WSTR(pVar))
   this.CopyFromVariant(pvar)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destroys the safe array when the class is destroyed.
' The array descriptor and all of the data in the array is destroyed. If objects are stored
' in the array, Release is called on each object in the array.
' Safe arrays of variant will have VariantClear called on each member and safe arrays of
' BSTR will have SysFreeString called on each element. IRecordInfo.RecordClear will be called
' to release object references and other values of a record without deallocating the record.
' ========================================================================================
PRIVATE DESTRUCTOR DSafeArray
   DSafeArray_DP(..WSTR(m_psa))
   IF m_psa THEN
      SafeArrayUnlock(m_psa)
      SafeArrayDestroy(m_psa)
   END IF
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assigns a DSafeArray.
' ========================================================================================
PRIVATE OPERATOR DSafeArray.Let (BYREF dsa AS DSafeArray)
   DSafeArray_DP("DSafeArray")
   this.CopyFrom(dsa.m_psa)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a safe array.
' ========================================================================================
PRIVATE OPERATOR DSafeArray.Let (BYVAL psa AS SAFEARRAY PTR)
   DSafeArray_DP("SAFEARRAY PTR")
   this.CopyFrom(psa)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a variant.
' ========================================================================================
PRIVATE OPERATOR DSafeArray.Let (BYVAL pvar AS VARIANT PTR)
   DSafeArray_DP("VARIANT PTR")
   this.CopyFromVariant(pvar)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the safe array descriptor.
' Removed to allow to use @ to get the address of the class.
' ========================================================================================
'PRIVATE OPERATOR DSafeArray.@ () AS ANY PTR
'   DSafeArray_DP("")
'   OPERATOR = @m_psa
'END OPERATOR
' ========================================================================================

' ========================================================================================
' Destroys the safe array and returns the address of the descriptor.
' To be used with BYVAL OUT safe array parameters to avoid a memory leak.
' Gotcha: The returned safe array won't be locked. Better declare a variable as a safe
' array, e.g. DIM psa AS SAFEARRAY PTR, pass it to that function or method, e.g. @psa,
' and then attach if to the class.
' ========================================================================================
'PRIVATE FUNCTION DSafeArray.vptr () AS SAFEARRAY PTR
'   DSafeArray_DP("")
'   IF m_psa THEN
'      SafeArrayUnlock(m_psa)
'      SafeArrayDestroy(m_psa)
'   END IF
'   RETURN @m_psa
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the safe array descriptor (same as *).
' ========================================================================================
PRIVATE FUNCTION DSafeArray.GetPtr () AS SAFEARRAY PTR
   DSafeArray_DP(..WSTR(m_psa))
   RETURN m_psa
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a pointer to the safe array descriptor.
' ========================================================================================
PRIVATE OPERATOR * (BYREF dsa AS DSafeArray) AS SAFEARRAY PTR
   DSafeArray_DP(..WSTR(dsa.m_psa))
   OPERATOR = dsa.m_psa
END OPERATOR
' ========================================================================================

' =====================================================================================
' Creates a safe array from the given VARTYPE, number of dimensions and bounds.
' Parameters:
' vt
'   [in] Base type of the array (the VARTYPE of each element of the array).
'   The VARTYPE is restricted to a subset of the variant types.
'   Neither VT_ARRAY nor the VT_BYREF flag can be set.
'   VT_EMPTY and VT_NULL are not valid base types for the array.
'   All other types are legal.
' cDims
'   [in] Number of dimensions in the array.
'   The number cannot be changed after the array is created.
' rgsabound
'   [in] Pointer to a vector of bounds (one for each dimension) to allocate for the array.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION DSafeArray.Create (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR) AS HRESULT
   DSafeArray_DP("multi-D")
   IF m_psa <> NULL THEN RETURN E_FAIL
   IF prgsabound = NULL THEN RETURN E_INVALIDARG
   IF cDims < 1 THEN RETURN E_INVALIDARG
   m_psa = SafeArrayCreate(vt, cDims, prgsabound)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.Create (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   DSafeArray_DP("1D")
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   RETURN this.Create(vt, 1, @rgsabounds(0))
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
PRIVATE FUNCTION DSafeArray.Create (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   DSafeArray_DP("2D")
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   RETURN this.Create(vt, 2, @rgsabounds(0))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a safe array from the given VARTYPE, number of dimensions and bounds.
' Parameters:
' vt
'   [in] The base type or the VARTYPE of each element of the array. The FADF_RECORD
'   flag can be set for a variant type VT_RECORD, The FADF_HAVEIID flag can be set
'   for VT_DISPATCH or VT_UNKNOWN, and FADF_HAVEVARTYPE can be set for all other VARTYPEs.
' cDims
'   [in] Number of dimensions in the array.
'   The number cannot be changed after the array is created.
' rgsabound
'   [in] Pointer to a vector of bounds (one for each dimension) to allocate for the array.
' pvExtra
'   Points to the type information of the user-defined type, if you are creating a
'   safe array of user-defined types. If the vt parameter is VT_RECORD, then
'   pvExtra will be a pointer to an IRecordInfo describing the record. If the vt
'   parameter is VT_DISPATCH or VT_UNKNOWN, then pvExtra will contain a pointer to
'   a GUID representing the type of interface being passed to the array.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
' Comments:
'   If the VARTYPE is VT_RECORD then SafeArraySetRecordInfo is called. If the
'   VARTYPE is VT_DISPATCH or VT_UNKNOWN then the elements of the array must contain
'   interfaces of the same type. Part of the process of marshaling this array to
'   other processes does include generating the proxy/stub code of the IID pointed
'   to by pvExtra parameter. To actually pass heterogeneous interfaces one will need
'   to specify either IID_IUnknown or IID_IDispatch in pvExtra and provide some
'   other means for the caller to identify how to query for the actual interface.
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION DSafeArray.CreateEx (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR, BYVAL pvExtra AS PVOID) AS HRESULT
   DSafeArray_DP("multi-D")
   IF m_psa <> NULL THEN RETURN E_FAIL
   IF prgsabound = NULL THEN RETURN E_INVALIDARG
   IF cDims < 1 THEN RETURN E_INVALIDARG
   m_psa = SafeArrayCreateEx(vt, cDims, prgsabound, pvExtra)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DSafeArray_DP("1D")
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   RETURN this.CreateEx(vt, 1, @rgsabounds(0), pvExtra)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
PRIVATE FUNCTION DSafeArray.CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DSafeArray_DP("2D")
   DIM rgsabounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   RETURN this.CreateEx(vt, 2, @rgsabounds(0), pvExtra)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a one-dimensional array. A safe array created with SafeArrayCreateVector
' is a fixed size, so the constant FADF_FIXEDSIZE is always set.
' Parameters:
' vt
'   [in] Base type of the array (the VARTYPE of each element of the array).
'   The VARTYPE is restricted to a subset of the variant types.
'   Neither VT_ARRAY nor the VT_BYREF flag can be set.
'   VT_EMPTY and VT_NULL are not valid base types for the array.
'   All other types are legal.
' cElements
'   The number of elements in the array.
' lLBound
'   The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
' Comments:
'  SafeArrayCreateVector allocates a single block of memory containing a SAFEARRAY
'  structure for a single-dimension array (24 bytes), immediately followed by the
'  array data. All of the existing safe array functions work correctly for safe
'  arrays that are allocated with SafeArrayCreateVector.
'  A SafeArrayCreateVector is allocated as a single block of memory. Both the
'  SafeArray descriptor and the array data block are allocated contiguously in one
'  allocation, which speeds up array allocation.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.CreateVector (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   DSafeArray_DP("")
   IF m_psa <> NULL THEN RETURN E_FAIL
   m_psa = SafeArrayCreateVector(vt, lLBound, cElements)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a one-dimensional array of the specified type and bounds.
' Parameters:
' vt
'   The base type of the array (the VARTYPE of each element of the array). The
'   FADF_RECORD flag can be set for VT_RECORD. The FADF_HAVEIID can be set for
'   VT_DISPATCH or VT_UNKNOWN and FADF_HAVEVARTYPE can be set for all other types.
' cElements
'   The number of elements in the array.
' lLBound
'   The lower bound value; that is, the index of the first element in the array.
'   Can be negative.
' pvExtra
'   Points to the type information of the user-defined type, if you are creating a
'   safe array of user-defined types. If the vt parameter is VT_RECORD, then
'   pvExtra will be a pointer to an IRecordInfo describing the record. If the vt
'   parameter is VT_DISPATCH or VT_UNKNOWN, then pvExtra will contain a pointer to
'   a GUID representing the type of interface being passed to the array.
' OBJRESULT:
'   Returns S_OK on success, or an error HRESULT on failure.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.CreateVectorEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, BYVAL pvExtra AS ANY PTR) AS HRESULT
   DSafeArray_DP("")
   IF m_psa <> NULL THEN RETURN E_FAIL
   m_psa = SafeArrayCreateVectorEx(vt, lLBound, cElements, pvExtra)
   IF m_psa = NULL THEN RETURN E_OUTOFMEMORY
   RETURN SafeArrayLock(m_psa)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Returns the number of dimensions in the array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.NumDims () AS UINT
   DSafeArray_DP("")
   IF m_psa THEN RETURN SafeArrayGetDim(m_psa)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the lower bound of the safe array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.LBound (BYVAL nDim AS UINT = 1) AS LONG
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN 0
   DIM plLBound AS LONG
   SafeArrayGetLBound(m_psa, nDim, @plLBound)
   RETURN plLBound
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the upper bound of the safe array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.UBound (BYVAL nDim AS UINT = 1) AS LONG
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN 0
   DIM plUBound AS LONG
   SafeArrayGetUBound(m_psa, nDim, @plUBound)
   RETURN plUBound
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of elements in the specified dimension of the array.
' Parameter:
' nDim: The array dimension for which to get the number of elements.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Count (BYVAL nDim AS UINT = 1) AS UINT
   DSafeArray_DP("multi-D")
   IF m_psa = NULL THEN RETURN 0
   IF nDim < 1 OR nDim > this.NumDims THEN RETURN 0
   DIM plLBound AS LONG, plUBound AS LONG
   SafeArrayGetLBound(m_psa, nDim, @plLBound)
   SafeArrayGetUBound(m_psa, nDim, @plUBound)
   RETURN plUbound - plLBound + 1
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the size (in bytes) of an element of the array.
' Does not include size of pointed-to data.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.ElemSize () AS UINT
   DSafeArray_DP("")
   IF m_psa THEN RETURN SafeArrayGetElemSize(m_psa)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Tests if the safe array can be resized.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.IsResizable () AS BOOLEAN
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN FALSE
   RETURN (m_psa->fFeatures AND FADF_FIXEDSIZE) <> FADF_FIXEDSIZE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the VARTYPE stored in the given safe array.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.GetType () AS VARTYPE
   IF m_psa = NULL THEN RETURN 0
   DIM vt AS VARTYPE
   SafeArrayGetVartype(m_psa, @vt)
   DSafeArray_DP("VarType = " & ..WSTR(vt))
   RETURN vt
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the number of Number of times the array has been locked without
' the corresponding unlock.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.LocksCount () AS UINT
   DSafeArray_DP("")
   IF m_psa THEN RETURN m_psa->cLocks
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the flags used by the safe array.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Flags () AS USHORT
   DSafeArray_DP("")
   IF m_psa THEN RETURN m_psa->fFeatures
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Features () AS USHORT
   DSafeArray_DP("")
   IF m_psa THEN RETURN m_psa->fFeatures
END FUNCTION
' =====================================================================================

' ========================================================================================
' Increments the lock count of an array, and retrieves a pointer to the array data.
' Return value: A pointer to the array data.
' Remarks: After calling AccessData, you must call the UnaccessData function to unlock the array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.AccessData () AS ANY PTR
   DSafeArray_DP("")
   DIM pvData AS ANY PTR
   IF m_psa THEN SafeArrayAccessData(m_psa, @pvData)
   RETURN pvData
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the lock count of an array, and invalidates the pointer retrieved by AccessData.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.UnaccessData () AS HRESULT
   DSafeArray_DP("")
   RETURN SafeArrayUnaccessData(m_psa)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Changes the right-most (least significant) bound of a safe array.
' Parameter:
' psaboundNew
'   Pointer to a new safe array bound structure that contains the new array boundary.
'   You can change only the least significant dimension of an array.
' Comments:
'   If you reduce the bound of an array, SafeArrayRedim deallocates the array
'   elements outside the new array boundary. If the bound of an array is increased,
'   SafeArrayRedim allocates and initializes the new array elements. The data is
'   preserved for elements that exist in both the old and new array.
' OBJRESULT:
'   S_OK Success.
'   %DISP_E_ARRAYISLOCKED The array is currently locked.
'   E_INVALIDARG psaboundNew is a null pointer.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Redim (BYVAL pnewsabounds AS SAFEARRAYBOUND PTR) AS HRESULT
   DSafeArray_DP("multi-D")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF (m_psa->fFeatures AND FADF_FIXEDSIZE) = FADF_FIXEDSIZE THEN RETURN E_FAIL
   DIM hr AS HRESULT = SafeArrayUnlock(m_psa)
   IF SUCCEEDED(hr) THEN
      hr = SafeArrayRedim(m_psa, pnewsabounds)
      SafeArrayLock(m_psa)
   END IF
   RETURN hr
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Redim (BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
   DSafeArray_DP("1D")
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Redim (BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
   DSafeArray_DP("2D")
   DIM sanewbounds(1) AS SAFEARRAYBOUND = {(cElements1, lLBound1), (cElements2, lLBound2)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a copy of the safe array.
' Return value:
' Pointer of the new array descriptor. You must free this pointer calling the API
' function SafeArrayDestroy.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Copy () AS SAFEARRAY PTR
   DSafeArray_DP("")
   DIM psaOut AS SAFEARRAY PTR
   IF m_psa THEN SafeArrayCopy(m_psa, @psaOut)
   RETURN psaOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the source array to the target array after releasing any resources in the
' target array. This is similar to SafeArrayCopy, except that the target array has
' to be set up by the caller. The target is not allocated or reallocated.
' Parameter:
' - psaTarget
'   The target safe array. On exit, the array referred to by psaTarget contains a
'   copy of the data in m_psa.
' Return value:
'   S_OK Success.
'   E_FAIL Failure.
'   E_INVALIDARG The argument psaTarget was not valid.
'   E_OUTOFMEMORY Insufficient memory to create the copy.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.CopyData (BYVAL psaTarget AS SAFEARRAY PTR) AS HRESULT
   DSafeArray_DP("")
   RETURN SafeArrayCopyData(m_psa, psaTarget)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the contents of a safe array.
' Parameter:
' - psaSrc: Pointer to an array descriptor created by SafeArrayCreate.
' Return value:
' Returns S_OK on success, or an HRESULT on failure.
' Remarks:
' SafeArrayPutElement, automatically calls SafeArrayLock and SafeArrayUnlock before and
' after assigning the element.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.CopyFrom (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   DSafeArray_DP("safearray ptr = " & ..WSTR(psaSrc))
   IF psaSrc = NULL THEN RETURN E_INVALIDARG
   ' // Destroy the safe array
   IF m_psa THEN this.Destroy
   ' // Copy the passed safe array and lock it
   DIM hr AS HRESULT = SafeArrayCopy(psaSrc, @m_psa)
   SafeArrayLock(m_psa)
   DSafeArray_DP("m_psa = " & ..WSTR(m_psa))
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to an array element.
' Parameters:
' prgIndices
'   An array of index values that identify an element of the array. All indexes for
'   the element must be specified.
' Return Value:
'   Pointer to the array element.
' ========================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION DSafeArray.PtrOfIndex (BYVAL prgIndices AS LONG PTR) AS ANY PTR
   DSafeArray_DP("multi-D")
   IF m_psa = NULL THEN RETURN NULL
   DIM pvData AS ANY PTR
   SafeArrayPtrOfIndex(m_psa, prgIndices, @pvData)
   RETURN pvData
END FUNCTION
' ========================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.PtrOfIndex (BYVAL idx AS LONG) AS ANY PTR
   DSafeArray_DP("1D")
   IF m_psa = NULL THEN RETURN NULL
   DIM pvData AS ANY PTR
   SafeArrayPtrOfIndex(m_psa, @idx, @pvData)
   RETURN pvData
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.PtrOfIndex (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS ANY PTR
   DSafeArray_DP("2D")
   IF m_psa = NULL THEN RETURN NULL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM pvData AS ANY PTR
   SafeArrayPtrOfIndex(m_psa, @rgIdx(0), @pvData)
   RETURN pvData
END FUNCTION
' =====================================================================================

' ========================================================================================
' Destroys an existing array descriptor and all of the data in the array.
' SysFreeString will be called on each element.
' Return value:
' S_OK Success.
' DISP_E_ARRAYISLOCKED The array is locked.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Destroy () AS HRESULT
   DSafeArray_DP("m_psa = " & ..WSTR(m_psa))
   DIM hr AS HRESULT
   IF m_psa THEN
      hr = SafeArrayUnlock(m_psa)
      hr = SafeArrayDestroy(m_psa)
      IF SUCCEEDED(hr) THEN m_psa = NULL
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Destroys all the data in a safe array.
' Return value:
'   Returns S_OK on success, or an error HRESULT on failure.
'   E_INVALIDARG: The argument psa is not valid.
'   DISP_E_ARRAYISLOCKED: The array is locked.
' Comments:
'   This method is typically used when freeing safe arrays that contain elements with
'   data types other than variants. If objects are stored in the array, Release is
'   called on each object in the array. Safe arrays of variant will have VariantClear
'   called on each member and safe arrays of BSTR will have SysFreeString called on
'   each element. IRecordInfo::RecordClear will be called to release object references
'   and other values of a record without deallocating the record.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.DestroyData () AS HRESULT
   DSafeArray_DP("")
   DIM hr AS HRESULT
   IF m_psa THEN
      SafeArrayUnlock(m_psa)
      hr = SafeArrayDestroyData(m_psa)
'      Don't lock the erased array
'      SafeArrayLock(m_psa)
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Clear () AS HRESULT
   RETURN this.Destroy
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Erase () AS HRESULT
   RETURN this.DestroyData
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Reset () AS HRESULT
   RETURN this.DestroyData
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the contents of a VARIANT of type VT_ARRAY, i.e. containing a safe array, to
' the DSafeArray array. The VARIANT remains unaltered.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.CopyFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DSafeArray_DP("")
   IF pvar = NULL THEN RETURN E_INVALIDARG
   IF (pvar->vt AND VT_ARRAY) <> VT_ARRAY THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   RETURN this.CopyFrom(pvar->parray)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Transfers ownership of the safe array contained in the variant parameter to this
' object. The variant is then changed to VT_EMPTY.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.MoveFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DSafeArray_DP("")
   IF pvar = NULL THEN RETURN E_INVALIDARG
   IF (pvar->vt AND VT_ARRAY) <> VT_ARRAY THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   FUNCTION = this.CopyFrom(pvar->parray)
   pvar->vt = VT_EMPTY
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the contents of the safe array to a variant.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.CopyToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   DIM vt AS VARTYPE
   SafeArrayGetVartype(m_psa, @vt)
   pvar->vt = vt OR VT_ARRAY
   RETURN SafeArrayCopy(m_psa, @pvar->parray)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Transfers ownership of the safe array to a variant and detaches it from the object.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.MoveToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF pvar = NULL THEN RETURN E_INVALIDARG
   VariantClear(pvar)
   DIM vt AS VARTYPE
   SafeArrayGetVartype(m_psa, @vt)
   pvar->vt = vt OR VT_ARRAY
   FUNCTION = SafeArrayUnlock(m_psa)
   pvar->parray = m_psa
   m_psa = NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches a SAFEARRAY descriptor to a DSafeArray array.
' Parameter:
' - psaSrc: A pointer to the SAFEARRAY descriptor.
' Returns S_OK on success, or an error HRESULT on failure.
' Remarks: Never attach the same safe array to more than one DSafeArray class because each one
' will try to destroy it when they are deleted or went out of scope.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Attach (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
   DSafeArray_DP(..WSTR(psaSrc))
   IF psaSrc = NULL THEN RETURN E_FAIL
   DIM hr AS HRESULT = this.Destroy
   IF SUCCEEDED(hr) THEN
      m_psa = psaSrc
      hr = SafeArrayLock(m_psa)
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the SAFEARRAY descriptor from the DSafeArray array.
' Return value: Returns a pointer to a SAFEARRAY descriptor.
' The caller takes ownership of it and must destroy it when no longer needed.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Detach () AS SAFEARRAY PTR
   DSafeArray_DP(..WSTR(m_psa))
   SafeArrayUnlock(m_psa)
   FUNCTION = m_psa
   m_psa = NULL
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the GUID of the interface contained within a given safe array.
' Return Value:
'   The GUID of the interface, on success, or a null guid on failure.
' Return value:
'   If the function succeeds it returns the guid; otherwise, returns a null guid.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.GetIID () AS GUID
   DSafeArray_DP("")
   DIM hr AS HRESULT, guid_ AS GUID
   IF m_psa = NULL THEN RETURN guid_
   IF (m_psa->fFeatures AND FADF_HAVEIID) <> FADF_HAVEIID THEN RETURN guid_
   hr = SafeArrayGetIID(m_psa, @guid_)
   IF hr = S_OK THEN RETURN guid_
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the GUID of the interface contained within a given safe array.
' Return value:
'   S_OK Success.
'   E_FAIL If m_psa is null.
'   E_INVALIDARG If the array descriptor does not have the FADF_HAVEIID flag set.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.SetIID (BYVAL pguid AS GUID PTR) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF (m_psa->fFeatures AND FADF_HAVEIID) <> FADF_HAVEIID THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   RETURN SafeArraySetIID(m_psa, pguid)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the IRecordInfo interface of the UDT contained in a given safe array.
' Return value:
' If the function succeeds it returns the guid; otherwise, returns a null guid.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.GetRecordInfo () AS IRecordInfo PTR
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN NULL
   IF (m_psa->fFeatures AND FADF_RECORD) <> FADF_RECORD THEN RETURN NULL
   DIM hr AS HRESULT, prinfo AS IRecordInfo PTR
   hr = SafeArrayGetRecordInfo(m_psa, @prinfo)
   RETURN prinfo
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the IRecordInfo Interface of the UDT contained in a given safe array.
' OBJRESULT:
'   S_OK Success.
'   E_FAIL If m_psa is null.
'   E_INVALIDARG If the array descriptor does not have the FADF_RECORD flag set.
' =====================================================================================
PRIVATE FUNCTION DSafeArray.SetRecordInfo (BYVAL prinfo AS IRecordInfo PTR) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF (m_psa->fFeatures AND FADF_RECORD) <> FADF_RECORD THEN RETURN HRESULT_FROM_WIN32(ERROR_INVALID_DATA)
   RETURN SafeArraySetRecordInfo(m_psa, prinfo)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves a single element of the array.
' Parameters:
' prgIndices
'   Pointer to a vector of indexes for each dimension of the array. The right-most
'   (least significant) dimension is prgIndices[0]. The left-most dimension is stored
'   at rgIndices[psa->cDims – 1].
' pData
'   Pointer to the location to place the element of the array.
' Comments:
'   This function calls SafeArrayLock and SafeArrayUnlock automatically, before and
'   after retrieving the element. The caller must provide a storage area of the
'   correct size to receive the data. If the data element is a string, object, or
'   variant, the function copies the element in the correct way.
' OBJRESULT:
'   The return value obtained from the returned HRESULT is one of the following.
'   S_OK Success.
'   E_FAIL This safe array has no elements
'   DISP_E_BADINDEX The specified index is invalid.
'   E_INVALIDARG One of the arguments is invalid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
' =====================================================================================

' =====================================================================================
' // *** BYVAL ANY PTR ***
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("multi-D - pData = " & ..WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, prgIndices, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.GetElement(prgIndices, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("1D - pData = " & ..WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, @idx, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.GetElement(idx, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("2D - pData = " & ..WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayGetElement(m_psa, @rgIdx(0), pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.GetElement(cElem, cDim, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' ** BYREF BSTRING ***
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("multi-D - BSTRING " & ..WSTR(bsData.m_bstr))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, prgIndices, bsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.GetElement(prgIndices, bsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.GetStr (BYVAL prgIndices AS LONG PTR) AS BSTRING
   DSafeArray_DP("multi-D")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN ""
   DIM bs AS BSTR
   SafeArrayGetElement(m_psa, prgIndices, @bs)
   RETURN bs
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("1D - BSTRING " & ..WSTR(bsData.m_bstr))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, @idx, bsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.GetElement(idx, bsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.GetStr (BYVAL idx AS LONG) AS BSTRING
   DSafeArray_DP("1D -" & ..WSTR(idx))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN ""
   DIM bs AS BSTR
   SafeArrayGetElement(m_psa, @idx, @bs)
   RETURN bs
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("2D - " & ..WSTR(bsData.m_bstr))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayGetElement(m_psa, @rgIdx(0), bsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.GetElement(cElem, cDim, bsData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.GetStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS BSTRING
   DSafeArray_DP("2D")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN ""
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM bs AS BSTR
   SafeArrayGetElement(m_psa, @rgIdx(0), @bs)
   RETURN bs
END FUNCTION
' =====================================================================================

' =====================================================================================
' ** BYREF DVARIANT ***
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("multi-D")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   ' // No need to clear the variant because SafeArrayGetElement uses VariantCopy to
   ' // copy the data and VariantCopy clears the destination variant.
   RETURN SafeArrayGetElement(m_psa, prgIndices, dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.GetElement(prgIndices, dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Multidimensional array
' // prgIndices: first the element, then the dimension
PRIVATE FUNCTION DSafeArray.GetVar (BYVAL prgIndices AS LONG PTR) AS DVARIANT
   DSafeArray_DP("multi-D")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM dvData AS DVARIANT
   SafeArrayGetElement(m_psa, prgIndices, dvData)
   RETURN dvData
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("1D - DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   RETURN SafeArrayGetElement(m_psa, @idx, dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.GetElement(idx, dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.GetVar (BYVAL idx AS LONG) AS DVARIANT
   DSafeArray_DP("1D - GET DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM dvData AS DVARIANT
   SafeArrayGetElement(m_psa, @idx, dvData)
   RETURN dvData
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.GetElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("2D - DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayGetElement(m_psa, @rgIdx(0), dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.GetElement(cElem, cDim, dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.GetVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS DVARIANT
   DSafeArray_DP("2D - DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM dvData AS DVARIANT
   SafeArrayGetElement(m_psa, @rgIdx(0), dvData)
   RETURN dvData
END FUNCTION
' =====================================================================================

' =====================================================================================
' Stores the data element at a given location in the array.
' Parameters
' prgIndices
'   Pointer to a vector of indexes for each dimension of the array. The right-most
'   (least significant) dimension is prgIndices[0]. The left-most dimension is stored
'   at rgIndices[psa->cDims – 1].
' pData
'   Pointer to the data to assign to the array. The variant types VT_DISPATCH,
'   VT_UNKNOWN, and VT_BSTR are pointers, and do not require another level of indirection.
' Comments:
'   This function automatically calls SafeArrayLock and SafeArrayUnlock before and
'   after assigning the element. If the data element is a string, object, or variant,
'   the function copies it correctly when the safe array is destroyed. If the
'   existing element is a string, object, or variant, it is cleared correctly. If
'   the data element is a VT_DISPATCH or VT_UNKNOWN, AddRef is called to increment
'   the object's reference count.
' Note
'   Multiple locks can be on an array. Elements can be put into an array while the
'   array is locked by other operations.
' OBJRESULT:
'   The return value obtained from the returned HRESULT is one of the following.
'   S_OK Success.
'   E_FAIL This safe array has no elements
'   DISP_E_BADINDEX The specified index was invalid.
'   E_INVALIDARG One of the arguments is invalid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
' =====================================================================================

' =====================================================================================
' *** ANY PTR ***
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("multiD - pData = " & ..WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, prgIndices, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Put (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.PutElement(prgIndices, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("1D - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @idx, pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Put (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.PutElement(idx, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("2D - pData = " & WSTR(pData))
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayPutElement(m_psa, @rgIdx(0), pData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.Put (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.PutElement(cElem, cDim, pData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' *** BYREF BSTRING ***
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("multi-D - BSTRING")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, prgIndices, bsData.m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.PutStr (BYVAL prgIndices AS LONG PTR, BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.PutElement(prgIndices, bsData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("1D -  BSTRING")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @idx, bsData.m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.PutStr (BYVAL idx AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.PutElement(idx, bsData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("2D -  BSTRING")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @rgIdx(0), bsData.m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.PutStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.PutElement(cElem, cDim, bsData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' *** BYREF DVARIANT ***
' =====================================================================================
' // Multidimensional array
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("multi-D - DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, prgIndices, dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.PutVar (BYVAL prgIndices AS LONG PTR, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.PutElement(prgIndices, dvData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // One-dimensional array
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("1D -  DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   RETURN SafeArrayPutElement(m_psa, @idx, dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.PutVar (BYVAL idx AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.PutElement(idx, dvData)
END FUNCTION
' =====================================================================================

' =====================================================================================
' // Two-dimensional array
' // First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).
PRIVATE FUNCTION DSafeArray.PutElement (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("2D -  DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM rgIdx(1) AS LONG = {cElem, cDim}
   RETURN SafeArrayPutElement(m_psa, @rgIdx(0), dvData)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DSafeArray.PutVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.PutElement(cElem, cDim, dvData)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Appends a value to the end of the one-dimensional safe array.
' If the safe array is not a one-dimensional array or it is not resizable it will return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.AppendElement (BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("ANY PTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Append (BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.AppendElement(pData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a BSTR to the end of the one-dimensional VT_BSTR safe array.
' If the safe array is not a one-dimensional VT_BSTR array or it is not resizable it will
' return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.AppendElement (BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("BSTRING")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, bsData.m_bstr)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Append (BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.AppendElement(bsData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.AppendStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   DIM bs AS AFX_BSTR = SysAllocString(pwszData)
   FUNCTION = SafeArrayPutElement(m_psa, @idx, bs)
   SysFreeString bs
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a VARIANT to the end of the one-dimensional VT_VARIANT safe array.
' If the safe array is not a one-dimensional VT_VARIANT array or it is not resizable it will
' return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.AppendElement (BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, dvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Append (BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.AppendElement(dvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.AppendVar (BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.AppendElement(dvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.AppendElement (BYVAL vData AS VARIANT) AS HRESULT
   DSafeArray_DP("DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   DIM idx AS LONG = cElements - 1 + lLBound
   RETURN SafeArrayPutElement(m_psa, @idx, @vData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Append (BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.AppendElement(vData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a value in the specified position of the one-dimensional safe array.
' If the safe array is not a one-dimensional array or it is not resizable it will return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' - nPos = Index of the array element where the new string will be inserted.
' - pData = Pointer to the data to insert.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   DSafeArray_DP("ANY PTR")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(pData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements up to nPos down
   DIM pvData AS ANY PTR, idx AS LONG
   FOR i AS LONG = cElements TO nPos STEP - 1
      idx = i - 1
      hr = SafeArrayGetElement(m_psa, @idx, @pvData)
      IF hr = S_OK THEN
         idx = i
         hr = SafeArrayPutElement(m_psa, @idx, @pvData)
      END IF
   NEXT
   ' // Set the value in the specified position
   idx = nPos
   RETURN SafeArrayPutElement(m_psa, @idx, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.InsertElement(this.LBound, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.InsertElement(nPos, pData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYVAL pData AS ANY PTR) AS HRESULT
   RETURN this.InsertElement(this.LBound, pData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYVAL nPos AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   DSafeArray_DP("BSTRING")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(bsData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      pvData[cElem] = SysAllocString(bsData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.InsertElement(this.LBound, bsData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYVAL nPos AS LONG, BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.InsertElement(nPos, bsData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYREF bsData AS BSTRING) AS HRESULT
   RETURN this.InsertElement(this.LBound, bsData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertStr (BYVAL nPos AS LONG, BYVAL pwszData AS WSTRING PTR) AS HRESULT
   DSafeArray_DP("DSafeArray InsertStr - WSTRING")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendStr(pwszData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      pvData[cElem] = SysAllocString(pwszData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
   RETURN this.InsertStr(this.LBound, pwszData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.InsertElement(this.LBound, dvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYVAL nPos AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.InsertElement(nPos, dvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.InsertElement(this.LBound, dvData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertVar (BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.InsertElement(this.LBound, dvData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYVAL nPos AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   DSafeArray_DP("DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(dvData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS VARIANT PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      ' // Mark it as empty to avoid that VariantClear,
      ' // called by VariantCopy, will free it.
      pvData[cElem].vt = VT_EMPTY
      VariantCopy(@pvData[cElem], dvData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertVar (BYVAL nPos AS LONG, BYREF dvData AS DVARIANT) AS HRESULT
   RETURN this.InsertElement(nPos, dvData)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   DSafeArray_DP("DVARIANT")
   IF m_psa = NULL THEN RETURN E_FAIL
   IF this.Count = 0 AND nPos = this.LBound THEN RETURN this.AppendElement(vData)
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN DISP_E_BADINDEX
   DIM cElements AS DWORD = this.Count(1) + 1
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements, lLBound)}
   DIM hr AS HRESULT = SafeArrayRedim(m_psa, @sanewbounds(0))
   IF hr THEN RETURN hr
   ' // Move all the elements down
   DIM cElem AS LONG = nPos - this.LBound
   DIM pvData AS VARIANT PTR = this.AccessData
   IF pvData THEN
      ' // Move all the elements from nPos down
      FOR i AS LONG = cElements - 1 TO cElem + 1 STEP - 1
         pvData[i] = pvData[i - 1]
      NEXT
      ' // Mark it as empty to avoid that VariantClear,
      ' // called by VariantCopy, will free it.
      pvData[cElem].vt = VT_EMPTY
      VariantCopy(@pvData[cElem], @vData)
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.InsertElement (BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.InsertElement(this.LBound, vData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.InsertElement(nPos, vData)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Insert (BYVAL vData AS VARIANT) AS HRESULT
   RETURN this.InsertElement(this.LBound, vData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a value to the end of the one-dimensional safe array.
' If the safe array is not a one-dimensional array or it is not resizable it will return E_FAIL.
' Return value:
'   S_OK Success.
'   DISP_E_BADINDEX The specified index is not valid.
'   E_INVALIDARG One of the arguments is not valid.
'   E_OUTOFMEMORY Memory could not be allocated for the element.
'   E_FAIL The item pointed to by m_psa is not a safe array descriptor.
'          It is a fixed-size array.
'          It is not a one-dimensional array.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.DeleteElement (BYVAL nPos AS LONG) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN E_FAIL
   DIM cElements AS DWORD = this.Count(1)
   ' // Save the element to be deleted
   DIM idx AS LONG, pTemp AS ANY PTR
   idx = nPos
   DIM hr AS HRESULT = SafeArrayGetElement(m_psa, @idx, @pTemp)
   IF hr THEN RETURN hr
   ' // Move all the elements from nPos + 1 up
   DIM pvData AS ANY PTR
   FOR i AS LONG = nPos TO cElements - 1 STEP 1
      idx = i + 1
      hr = SafeArrayGetElement(m_psa, @idx, @pvData)
      IF hr = S_OK THEN
         idx = i
         hr = SafeArrayPutElement(m_psa, @idx, @pvData)
      END IF
   NEXT
   ' // Copy the element to be deleted to the end of the array
   idx = cElements
   hr = SafeArrayPutElement(m_psa, @idx, @pTemp)
   ' // Shrink the array by one element (will free the last element)
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements - 1, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Remove (BYVAL nPos AS LONG) AS HRESULT
   RETURN this.DeleteElement(nPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION DSafeArray.DeleteStringElement (BYVAL nPos AS LONG) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN E_FAIL
   DIM cElem AS LONG = nPos - this.LBound
   DIM cElements AS DWORD = this.Count(1)
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      ' // Save the element to be deleted
      DIM pTemp AS AFX_BSTR = pvData[cElem]
      ' // Move all the elements from nPos + 1 up
      FOR i AS LONG = cElem TO cElements - 1 STEP 1
         pvData[i] = pvData[i + 1]
      NEXT
      ' // Copy the element to be deleted to the end of the array
      pvData[cElements - 1] = pTemp
      this.UnaccessData
   END IF
   ' // Shrink the array by one element (will free the last element)
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements - 1, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.RemoveStr (BYVAL nPos AS LONG) AS HRESULT
   RETURN this.DeleteStringElement(nPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION DSafeArray.DeleteVariantElement (BYVAL nPos AS LONG) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_VARIANT THEN RETURN E_FAIL
   IF this.IsResizable = FALSE THEN RETURN E_FAIL
   IF nPos < this.LBound OR nPos > this.UBound THEN RETURN E_FAIL
   DIM cElem AS LONG = nPos - this.LBound
   DIM cElements AS DWORD = this.Count(1)
   DIM pvData AS VARIANT PTR = this.AccessData
   IF pvData THEN
      ' // Save the element to be deleted
      DIM pTemp AS VARIANT = pvData[cElem]
      ' // Move all the elements from nPos + 1 up
      FOR i AS LONG = cElem TO cElements - 1 STEP 1
         pvData[i] = pvData[i + 1]
      NEXT
      ' // Copy the element to be deleted to the end of the array
      pvData[cElements - 1] = pTemp
      this.UnaccessData
   END IF
   ' // Shrink the array by one element (will free the last element)
   DIM lLBound AS LONG = this.LBound(1)
   DIM sanewbounds(0) AS SAFEARRAYBOUND = {(cElements - 1, lLBound)}
   RETURN SafeArrayRedim(m_psa, @sanewbounds(0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.RemoveVar (BYVAL nPos AS LONG) AS HRESULT
   RETURN this.DeleteVariantElement(nPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scans the array to search for the specified string.
' - wszFind = The string to search.
' - bNoCase = TRUE: Ignore case.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.FindElement (BYREF wszFind AS WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN 0
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN 0
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN 0
   DIM dwsFind AS DWSTRING = IIF(IgnoreCase = TRUE, UCASE(wszFind), wszFind)
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData THEN
      FOR i AS LONG = 0 TO this.UBound - this.LBound
         IF IIF(IgnoreCase = TRUE, UCASE(*pvData[i]), *pvData[i]) = dwsFind THEN
            FUNCTION = i + this.LBound
            EXIT FOR
         END IF
      NEXT
      this.UnaccessData
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Find (BYREF wszFind AS WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
   RETURN this.FindElement(wszFind, IgnoreCase)
END FUNCTION
' ========================================================================================

' ========================================================================================
' qsort DSafeArray comparison function
' ========================================================================================
PRIVATE FUNCTION AfxDSafeArrayCompare CDECL (BYVAL a AS BSTR PTR, BYVAL b AS BSTR PTR) AS LONG
   FUNCTION = wcscmp(*a, *b)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reverse qsort DSafeArray comparison function
' ========================================================================================
PRIVATE FUNCTION AfxDSafeArrayReverseCompare CDECL (BYVAL a AS BSTR PTR, BYVAL b AS BSTR PTR) AS LONG
   DIM r AS LONG = wcscmp(*a, *b)
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sorts a one-dimensional VT_BSTR DSafeArray calling the C qsort function.
' Parameter:
' - bAscend: TRUE for sorting in ascending order; FALSE for sorting in descending order.
' ========================================================================================
PRIVATE FUNCTION DSafeArray.Sort (BYVAL bAscend AS BOOLEAN = TRUE) AS HRESULT
   DSafeArray_DP("")
   IF m_psa = NULL THEN RETURN E_FAIL
   DIM nDims AS UINT = SafeArrayGetDim(m_psa)
   IF nDims <> 1 THEN RETURN E_FAIL
   DIM vt AS VARTYPE = this.GetType
   IF vt <> VT_BSTR THEN RETURN E_FAIL
   DIM pvData AS AFX_BSTR PTR = this.AccessData
   IF pvData = NULL THEN RETURN E_UNEXPECTED
   IF bAscend THEN
      qsort pvData, this.Count, SIZEOF(BSTR), CPTR(ANY PTR, @AfxDSafeArrayCompare)
   ELSE
      qsort pvData, this.Count, SIZEOF(BSTR), CPTR(ANY PTR, @AfxDSafeArrayReverseCompare)
   END IF
   IF pvData THEN this.UnaccessData
END FUNCTION
' ========================================================================================

' ========================================================================================
'                         *** Safe array related functions ***
' ========================================================================================

' ========================================================================================
' Splits a string into tokens, which are sequences of contiguous characters separated by
' any of the characters that are part of delimiters.
' - wszStr = The string to split.
' - wszDelimiters = The delimiter characters.
' Return value: A DSafeArray containing a token in each element.
' Usage example:
' DIM wsz AS WSTRING * 260 = "- This, a sample string."
' DIM dsa AS DSafeArray = AfxStrSplit(wsz, " ,.-")
' FOR i AS LONG = dsa.LBound TO dsa.UBound
'    DIM bsOut AS BSTRING
'    dsa.GetElement(i, bsOut)
'    print bsOut
' NEXT
' -or-
' FOR i AS LONG = dsa.LBound TO dsa.UBound
'    print dsa.GetStr(i)
' NEXT
' ========================================================================================
PRIVATE FUNCTION AfxStrSplit (BYREF wszStr AS WSTRING, BYREF wszDelimiters AS WSTRING) AS DSafeArray
   DIM dws AS DWSTRING = wszStr
   DIM dsa AS DSafeArray = DSafeArray(VT_BSTR, 0, 1)
   DIM pwsz AS WSTRING PTR = wcstok(dws, @wszDelimiters)
   WHILE pwsz <> NULL
      dsa.AppendStr(pwsz)
      pwsz = wcstok(NULL, @wszDelimiters)
   WEND
   RETURN dsa
END FUNCTION
' ========================================================================================


' ========================================================================================
' Returns a string consisting of all of the strings in an array, each separated by a delimiter.
' If the delimiter is a null (zero-length) string then no separators are inserted between
' the string sections. If the delimiter expression is the 3-byte value of "," which may be
' expressed in your source code as the string literal """,""" or as Chr(34,44,34) then a
' leading and trailing double-quote is added to each string section. This ensures that the
' returned string contains standard comma-delimited quoted fields that can be easily parsed.
' Usage example:
' DIM dsa AS DSafeArray = DSafeArray(VT_BSTR, 3, 1)
' dsa.PutElement(1, BSTRING("One"))
' dsa.PutElement(2, BSTRING("Two"))
' dsa.PutElement(3, BSTRING("Three"))
' DIM dws AS DWSTRING = AfxStrJoin(dsa, ",")
' PRINT dws   ' ouput: One,Two,Three
' ========================================================================================
PRIVATE FUNCTION AfxStrJoin (BYREF dsa AS DSafeArray, BYREF wszDelimiter AS WSTRING) AS DWSTRING
   DIM dws AS DWSTRING
   IF dsa.NumDims <> 1 THEN RETURN dws
   IF dsa.GetType <> VT_BSTR THEN RETURN dws
   ' // Use direct access for speed
   DIM nCount AS LONG = dsa.Count
   IF nCount = 0 THEN RETURN dws
   DIM pvData AS AFX_BSTR PTR = dsa.AccessData
   IF pvData = NULL THEN RETURN dws
   ' // Add a leading ""
   IF wszDelimiter = CHR(34, 44, 34) THEN dws = CHR(34)
   FOR i AS LONG = 0 TO nCount - 1
      dws += *pvData[i]
      IF i <> nCount - 1 AND wszDelimiter <> "" THEN dws += wszDelimiter
   NEXT
   dsa.UnaccessData
   ' // Add a trailing ""
   IF wszDelimiter = CHR(34, 44, 34) THEN dws += CHR(34)
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads all the lines of the specified file into a safe array.
' - wszFileName: Path of the file
' - szDelimiter: Delimiter of the line (CHR(13, 10) in Windows, CHR(10) in Linmux).
' ========================================================================================
PRIVATE FUNCTION AfxFileReadAllLinesA (BYREF wszFileName AS WSTRING, BYREF szDelimiter AS CONST ZSTRING = CHR(13, 10)) AS DSafeArray
   DIM _dsa AS DSafeArray = DSafeArray(VT_BSTR, 0, 1)
   DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   IF LEN(szDelimiter) = 0 THEN RETURN _dsa
   ' // Open the file
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN RETURN _dsa
   ' // Get the size of the file
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   DIM pBuffer AS UBYTE PTR
   pBuffer = CAllocate(1, dwFileSize)
   IF pBuffer = NULL THEN RETURN _dsa
   DIM bSuccess AS LONG = ReadFile(hFile, pBuffer, dwFileSize, @dwBytesRead, NULL)
   CloseHandle(hFile)
   IF bSuccess = FALSE THEN RETURN _dsa
   ' // Get the number of lines
   DIM nLen AS LONG = LEN(szDelimiter)
   DIM pstr AS ANY PTR = pBuffer
   DO
      pstr = strstr(pstr, szDelimiter)
      IF pstr = NULL THEN EXIT DO
      pstr += nLen
      dwCount += 1
   LOOP
' -------------------------------------------------------------------
  ' // Dimension a safe array with dwCount elements
'   DIM dsa AS DSafeArray = DSafeArray(VT_BSTR, dwCount, 1)
'   Note: strtok can't be used because it skips empty lines.
'   ' // Fill the array with the lines
'   DIM pwsz AS ZSTRING PTR = strtok(pBuffer, @szDelimiter)
'   DIM idx AS LONG = 1
'   WHILE pwsz <> NULL
'      IF idx < dwCount THEN dsa.PutElement(idx, BSTRING(*pwsz))
'      idx += 1
'      pwsz = strtok(NULL, @szDelimiter)
'   WEND
' -------------------------------------------------------------------
  ' // Dimension a safe array with dwCount elements
   DIM dsa AS DSafeArray = DSafeArray(VT_BSTR, dwCount, 1)
   DIM s AS STRING, sLen AS LONG
   DIM idx AS LONG = 1
   DIM _pstr AS ANY PTR = pBuffer
   pstr = pBuffer
   ' // Check for UTF-8 BOM
   s = "   "
   strncpy CAST(ANY PTR, STRPTR(s)), _pstr, 3
   IF s = CHR(&hEF, &hBB, &hBF) THEN _pstr += 3
   ' // Parse the buffer
   DO
      pstr = strstr(pstr, szDelimiter)
      IF pstr = NULL THEN EXIT DO
      sLen = pstr - _pstr
      s = ""
      IF sLen > nLen THEN
         s = STRING(sLen, CHR(0))
         strncpy CAST(ANY PTR, STRPTR(s)), _pstr, sLen
      END IF
      IF idx <= dwCount THEN dsa.PutElement(idx, BSTRING(s))
      idx += 1
      pstr += nLen
      _pstr = pstr
   LOOP
   DeAllocate(pBuffer)
   RETURN dsa
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads all the lines of the specified file into a safe array.
' - wszFileName: Path of the file
' - szDelimiter: Delimiter of the line (CHR(13, 10) in Windows, CHR(10) in Linmux).
' ========================================================================================
PRIVATE FUNCTION AfxFileReadAllLinesW (BYREF wszFileName AS WSTRING, BYREF wszDelimiter AS CONST WSTRING = CHR(13, 10)) AS DSafeArray
   DIM _dsa AS DSafeArray = DSafeArray(VT_BSTR, 0, 1)
   DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   IF LEN(wszDelimiter) = 0 THEN RETURN _dsa
   ' // Open the file
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN RETURN _dsa
   ' // Get the size of the file
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   DIM pBuffer AS UBYTE PTR
   pBuffer = CAllocate(1, dwFileSize)
   IF pBuffer = NULL THEN RETURN _dsa
   DIM bSuccess AS LONG = ReadFile(hFile, pBuffer, dwFileSize, @dwBytesRead, NULL)
   CloseHandle(hFile)
   IF bSuccess = FALSE THEN RETURN _dsa
   ' // Get the number of lines
   DIM nLen AS LONG = LEN(wszDelimiter) * 2
   DIM pstr AS ANY PTR = pBuffer
   DO
      pstr = wcsstr(pstr, @wszDelimiter)
      IF pstr = NULL THEN EXIT DO
      pstr += nLen
      dwCount += 1
   LOOP
  ' // Dimension a safe array with dwCount elements
   DIM dsa AS DSafeArray = DSafeArray(VT_BSTR, dwCount, 1)
   DIM dws AS DWSTRING, sLen AS LONG
   DIM idx AS LONG = 1
   DIM _pstr AS ANY PTR = pBuffer
   pstr = pBuffer
   ' / Skip the BOM
   DIM s AS STRING = "  "
   strncpy CAST(ANY PTR, STRPTR(s)), _pstr, 2
   ' // Check for UTF-16 BOM (little endian)
   IF s = CHR(&hFF, &hFE) THEN
      _pstr += 2
   ELSE
      ' // Check for UTF-8 BOM
      s = "   "
      strncpy CAST(ANY PTR, STRPTR(s)), _pstr, 3
      IF s = CHR(&hEF, &hBB, &hBF) THEN _pstr += 3
   END IF
   ' // Parse the buffer
   DO
      pstr = wcsstr(pstr, @wszDelimiter)
      IF pstr = NULL THEN EXIT DO
      sLen = (pstr - _pstr) \ 2
      dws = ""
      IF sLen > nLen THEN
         dws = STRING(sLen, CHR(0))
         wcsncpy dws, _pstr, sLen
      END IF
      IF idx <= dwCount THEN dsa.PutElement(idx, BSTRING(dws))
      idx += 1
      pstr += nLen
      _pstr = pstr
   LOOP
   DeAllocate(pBuffer)
   RETURN dsa
END FUNCTION
' ========================================================================================

END NAMESPACE
