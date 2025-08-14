' ########################################################################################
' Platform: Microsoft Windows
' Filename: DVARIANT.inc
' Purpose:  Implements a VARIANT data type.
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

' // Include files
#pragma ONCE
#include once "AfxNova/AfxVarProcs.inc"
USING AfxNova

' ========================================================================================
' Macro for debug
' To allow debugging, define _DVARIANT_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _DVARIANT_DEBUG_
   #define _DVARIANT_DEBUG_ 0
#ENDIF
#ifndef _DVARIANT_DP_
   #define _DVARIANT_DP_ 1
   #MACRO DVARIANT_DP(st)
      #IF (_DVARIANT_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE AfxNova

' ########################################################################################
'                                *** DVARIANT CLASS ***
' ########################################################################################
TYPE DVARIANT EXTENDS WSTRING

Public:
   vd AS VARIANT         ' // Variant data

Private:
   m_Result AS LONG      ' // Last result code
   m_dws AS DWSTRING     ' // To allow to return a WSTRING by reference

Public:

   ' ================================= CONSTRUCTORS ======================================
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE CONSTRUCTOR (BYREF dv AS DVARIANT)
   DECLARE CONSTRUCTOR (BYVAL v AS VARIANT)
   DECLARE CONSTRUCTOR (BYVAL pwsz AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYREF dws AS DWSTRING)
   DECLARE CONSTRUCTOR (BYREF bs AS BSTRING)
   DECLARE CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)
   DECLARE CONSTRUCTOR (BYVAL cy AS CURRENCY)
   DECLARE CONSTRUCTOR (BYVAL dec AS DECIMAL)
   DECLARE CONSTRUCTOR (BYVAL _value AS LONG)
'   DECLARE CONSTRUCTOR (BYVAL b AS BOOLEAN)
'   DECLARE CONSTRUCTOR (BYVAL _value AS DOUBLE)
   DECLARE CONSTRUCTOR (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   DECLARE CONSTRUCTOR (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   DECLARE CONSTRUCTOR (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   DECLARE CONSTRUCTOR (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   DECLARE CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD)
   DECLARE CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING)
   ' =====================================================================================

   ' ============================ ERRORS =================================================
   DECLARE FUNCTION GetLastResult () AS LONG
   DECLARE FUNCTION SetResult (BYVAL Result AS LONG) AS LONG
   DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS DWSTRING
   ' =====================================================================================

   ' =================================== CASTING =========================================
   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE OPERATOR CAST () AS ANY PTR
   DECLARE OPERATOR CAST () BYREF AS CONST WSTRING
   DECLARE FUNCTION vptr () AS VARIANT PTR
   ' =====================================================================================

   ' ================================== LET OPERATOR =====================================
   DECLARE OPERATOR LET (BYREF dv AS DVARIANT)
   DECLARE OPERATOR LET (BYVAL v AS VARIANT)
   DECLARE OPERATOR LET (BYVAL pvar AS VARIANT PTR)
   DECLARE OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
   DECLARE OPERATOR LET (BYREF dws AS DWSTRING)
   DECLARE OPERATOR LET (BYREF bs AS BSTRING)
   DECLARE OPERATOR LET (BYVAL cy AS CURRENCY)
   DECLARE OPERATOR LET (BYVAL dec AS DECIMAL)
   DECLARE OPERATOR LET (BYVAL _value AS LONG)
'   DECLARE OPERATOR LET (BYVAL b AS BOOLEAN)
   DECLARE OPERATOR LET (BYVAL pdisp AS IDispatch PTR)
   DECLARE OPERATOR LET (BYVAL punk AS IUnknown PTR)
   DECLARE OPERATOR LET (BYVAL _value AS LONGINT)
   DECLARE OPERATOR LET (BYVAL _value AS DOUBLE)
   ' =====================================================================================

   ' ============================== ASSIGNMENT METHODS ===================================
   DECLARE FUNCTION Put (BYREF dv AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Put (BYREF v AS VARIANT) AS HRESULT
   DECLARE FUNCTION Put (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE SUB Put (BYVAL pwszStr AS WSTRING PTR)
   DECLARE SUB Put (BYREF dws AS DWSTRING)
   DECLARE SUB Put (BYREF bs AS BSTRING)
   DECLARE FUNCTION Put (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Put (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE SUB Put (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   DECLARE SUB Put (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   DECLARE SUB Put (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   DECLARE SUB Put (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   DECLARE SUB PutNull
   DECLARE SUB PutBool (BYVAL _value AS BOOL)
   DECLARE SUB PutBoolean (BYVAL _value AS BOOLEAN)
   DECLARE SUB PutByte (BYVAL _value AS BYTE)
   DECLARE SUB PutUByte (BYVAL _value AS UBYTE)
   DECLARE SUB PutShort (BYVAL _value AS SHORT)
   DECLARE SUB PutUShort (BYVAL _value AS USHORT)
   DECLARE SUB PutInt (BYVAL _value AS INT_)
   DECLARE SUB PutUInt (BYVAL _value AS UINT)
   DECLARE SUB PutLong (BYVAL _value AS LONG)
   DECLARE SUB PutULong (BYVAL _value AS ULONG)
   DECLARE SUB PutLongInt (BYVAL _value AS LONGINT)
   DECLARE SUB PutULongInt (BYVAL _value AS ULONGINT)
   DECLARE SUB PutSingle (BYVAL _value AS SINGLE)
   DECLARE SUB PutFloat (BYVAL _value AS SINGLE)
   DECLARE SUB PutDouble (BYVAL _value AS DOUBLE)
   DECLARE FUNCTION PutUtf8 (BYREF strUtf8 AS STRING) AS HRESULT
   DECLARE FUNCTION PutBuffer(BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   DECLARE FUNCTION PutSafeArray (BYVAL parray AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION PutResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT) AS HRESULT
   DECLARE FUNCTION PutRecord (BYVAL pIRecordInfo AS IRecordInfo PTR, BYVAL pRec AS ANY PTR) AS HRESULT
   DECLARE FUNCTION PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DECLARE FUNCTION PutVbDate (BYVAL vbDate AS DATE_) AS HRESULT
   DECLARE FUNCTION PutSystemTime (BYVAL st AS SYSTEMTIME PTR) AS BOOLEAN
   DECLARE FUNCTION PutGuid (BYVAL guid AS IID PTR) AS HRESULT
   DECLARE FUNCTION PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
   DECLARE FUNCTION PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
   DECLARE FUNCTION PutDec (BYVAL dec AS DECIMAL) AS HRESULT
   DECLARE FUNCTION PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DECLARE FUNCTION PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
   DECLARE FUNCTION PutDecFromCy (BYVAL cyIn AS CY) AS HRESULT
   DECLARE FUNCTION PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutUShortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutLongArray (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG) AS DVARIANT
   DECLARE FUNCTION PutULongArray (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG) AS HRESULT
   DECLARE FUNCTION PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   DECLARE FUNCTION PutVariantArrayElement(BYVAL pvarIn AS VARIANT PTR, BYVAL iElem AS ULONG) AS DVARIANT
   ' // Assignments by reference
   DECLARE FUNCTION PutRef (BYVAL _value AS ANY PTR, BYVAL _vType AS WORD = VT_I8) AS HRESULT
   DECLARE FUNCTION PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
   ' =====================================================================================

   ' ================================== SAFEARRAYS ======================================
   DECLARE FUNCTION GetDim () AS ULONG
   DECLARE FUNCTION GetLBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION GetUBound (BYVAL nDim AS UINT = 1) AS LONG
   DECLARE FUNCTION GetVariantElement (BYVAL iElem AS ULONG) AS DVARIANT
   ' =====================================================================================

   ' ==================================== ARRAYS =========================================
   DECLARE FUNCTION GetElementCount () AS ULONG
   DECLARE FUNCTION GetBooleanElement (BYVAL iElem AS ULONG) AS BOOLEAN
   DECLARE FUNCTION GetShortElement (BYVAL iElem AS ULONG) AS SHORT
   DECLARE FUNCTION GetUShortElement (BYVAL iElem AS ULONG) AS USHORT
   DECLARE FUNCTION GetLongElement (BYVAL iElem AS ULONG) AS LONG
   DECLARE FUNCTION GetULongElement (BYVAL iElem AS ULONG) AS ULONG
   DECLARE FUNCTION GetLongIntElement (BYVAL iElem AS ULONG) AS LONGINT
   DECLARE FUNCTION GetULongIntElement (BYVAL iElem AS ULONG) AS ULONGINT
   DECLARE FUNCTION GetDoubleElement (BYVAL iElem AS ULONG) AS DOUBLE
   DECLARE FUNCTION GetStringElement (BYVAL iElem AS ULONG) AS DWSTRING
   ' =====================================================================================

   ' ============================== CONVERSION METHODS ===================================
   DECLARE FUNCTION ToStr () AS DWSTRING
   DECLARE FUNCTION ToBStr () AS BSTRING
   DECLARE FUNCTION ToUtf8 () AS STRING
   DECLARE FUNCTION ToBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   DECLARE FUNCTION ToBuffer () AS STRING
   DECLARE FUNCTION ToUnknown () AS ANY PTR
   DECLARE FUNCTION ToDispatch () AS ANY PTR
   DECLARE FUNCTION DecToDouble () AS DOUBLE
   DECLARE FUNCTION DecToCy () AS CY
   DECLARE FUNCTION ToVbDate () AS DATE_
   DECLARE FUNCTION ToSystemTime () AS SYSTEMTIME
   DECLARE FUNCTION ToGuid () AS GUID
   DECLARE FUNCTION ToGuidStr () AS DWSTRING
   DECLARE FUNCTION ToDosDateTime (BYVAL pwDate AS USHORT PTR, BYVAL pwTime AS USHORT PTR) AS HRESULT
   DECLARE FUNCTION ToFileTime (BYVAL stfOut AS AFX_PSTIME_FLAGS) AS FILETIME
   DECLARE FUNCTION ToStrRet () AS STRRET
   DECLARE FUNCTION ToBooleanArray (BYVAL pprg AS WINBOOL PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToBooleanArrayAlloc (BYVAL pprgf AS WINBOOL PTR PTR) AS ULONG
   DECLARE FUNCTION ToShortArray (BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToShortArrayAlloc (BYVAL pprgn AS SHORT PTR PTR) AS ULONG
   DECLARE FUNCTION ToUShortArray (BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToUShortArrayAlloc (BYVAL pprgn AS USHORT PTR PTR) AS ULONG
   DECLARE FUNCTION ToLongArray (BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToLongArrayAlloc (BYVAL pprgn AS LONG PTR PTR) AS ULONG
   DECLARE FUNCTION ToULongArray (BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToULongArrayAlloc (BYVAL pprgn AS ULONG PTR PTR) AS ULONG
   DECLARE FUNCTION ToLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToLongIntArrayAlloc (BYVAL pprgn AS LONGINT PTR PTR) AS ULONG
   DECLARE FUNCTION ToULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToULongIntArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
   DECLARE FUNCTION ToDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG) AS ULONG
   DECLARE FUNCTION ToDoubleArrayAlloc (BYVAL pprgn AS DOUBLE PTR PTR) AS ULONG
   DECLARE FUNCTION ToStringArray (BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG) AS ULONG
   DECLARE FUNCTION ToStringArrayAlloc (BYVAL pprgsz AS PWSTR PTR) AS ULONG
   ' =====================================================================================

   ' ================================== MISCELLANEOUS ====================================
   DECLARE SUB Clear
   DECLARE FUNCTION Attach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Attach (BYREF v AS VARIANT) AS HRESULT
   DECLARE FUNCTION Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE FUNCTION Detach (BYREF v AS VARIANT) AS HRESULT
   DECLARE FUNCTION vType () AS VARTYPE
   DECLARE FUNCTION ChangeType (BYVAL vtNew AS VARTYPE, BYVAL wFlags AS USHORT = 0) AS HRESULT
   DECLARE FUNCTION ChangeTypeEx (BYVAL vtNew AS VARTYPE, BYVAL lcid AS LCID = 0, BYVAL wFlags AS USHORT = 0) AS HRESULT
   DECLARE FUNCTION Round (BYREF dv AS DVARIANT, BYVAL cDecimals AS LONG) AS DVARIANT
   DECLARE FUNCTION Round (BYVAL cDecimals AS LONG) AS DVARIANT
   DECLARE FUNCTION FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
           BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   ' =====================================================================================

   ' ================================== MATH OPERATORS ===================================
   DECLARE OPERATOR += (BYREF dv AS DVARIANT)
   DECLARE OPERATOR &= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR -= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR *= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR /= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR \= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR ^= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR And= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR Or= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR Xor= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR Eqv= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR Imp= (BYREF dv AS DVARIANT)
   DECLARE OPERATOR Mod= (BYREF dv AS DVARIANT)
   ' =====================================================================================

END TYPE
' ########################################################################################

' ########################################################################################
'                                  CONSTRUCTORS
' ########################################################################################

' Notes about constructors:
' We can use the constructors to pass values to parameters in procedures without assigning
' them first to a variable, e.g.:
' ----------------------------------------------------------------------------------------
' SUB Foo (BYREF dv AS DVARIANT)
'    PRINT AfxDVarToStr(dv)
' END SUB
' Foo DVARIANT("Test string")
' Foo DVARIANT(12345)
' Foo DVARIANT(12345, "LONG")
' ----------------------------------------------------------------------------------------
' SUB Foo (BYVAL dv AS DVARIANT PTR)
'    PRINT AfxDVarToStr(dv)
' END SUB
' Foo @DVARIANT("Test string")
' Foo @DVARIANT(12345)
' Foo @DVARIANT(12345, "LONG")
' ----------------------------------------------------------------------------------------
' SUB Foo (BYVAL v AS VARIANT PTR)
'    PRINT AfxVarToStr(v)
' END SUB
' Foo DVARIANT("Test string")
' Foo DVARIANT(12345)
' Foo DVARIANT(12345, "LONG")
' ----------------------------------------------------------------------------------------

' ========================================================================================
' Default constructor
' Usage example:
' DIM dv AS DVARIANT
' dv = "Test string"
' print dv
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT
   DVARIANT_DP("Default - " & ..WSTR(@vd))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR DVARIANT
   DVARIANT_DP(..WSTR(@vd))
   m_Result = VariantClear(@vd)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes the DVARIANT from another DVARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYREF dv AS DVARIANT)
   DVARIANT_DP("DVARIANT")
   m_Result = VariantCopy(@vd, @dv.vd)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the DVARIANT from a VARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL v AS VARIANT)
   DVARIANT_DP("BYVAL VARIANT")
   m_Result = VariantCopy(@vd, @v)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the DVARIANT from a string.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL pwsz AS WSTRING PTR)
   DVARIANT_DP("WSTRING - " & ..WSTR(@vd))
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(pwsz)
   m_Result = 0
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYREF dws AS DWSTRING)
   DVARIANT_DP("DWSTRING")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(dws)
   m_Result = 0
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYREF bs AS BSTRING)
   DVARIANT_DP("BSTRING")
   vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(bs)
   m_Result = 0
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a pointer to a VARIANT.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL pvar AS VARIANT PTR)
   DVARIANT_DP("VARIANT PTR")
   m_Result = VariantCopy(@vd, pvar)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a CURRENCY structure.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL cy AS CURRENCY)
   DVARIANT_DP("CURRENCY")
   vd.vt = VT_CY : vd.cyVal = cy
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a DECIMAL structure.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL dec AS DECIMAL)
   DVARIANT_DP("DECIMAL")
   vd.vt = VT_DECIMAL : vd.decVal = dec
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a LONG value.
' For some reason, without it, when including CSAfeArray.inc, the constructor with a boolean
' throws and overflow error.
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL _value AS LONG)
   DVARIANT_DP("LONG")
   vd.vt = VT_I4 : vd.lVal = _value
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a BOOLEAN value.
' ========================================================================================
'PRIVATE CONSTRUCTOR DVARIANT (BYVAL b AS BOOLEAN)
'   DVARIANT_DP("BOOLEAN")
'   vd.vt = VT_BOOL : vd.boolVal = IIF(b = FALSE, 0, -1)
'END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a DOUBLE value.
' Conflicts with the constructor (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
' ========================================================================================
'PRIVATE CONSTRUCTOR DVARIANT (BYVAL _value AS DOUBLE)
'   DVARIANT_DP("DOUBLE")
'   vd.vt = VT_R8 : vd.dblVal = _value
'END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from an IDispatch pointer
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DVARIANT_DP("IDISPATCH PTR")
   IF fAddRef THEN IDispatch_AddRef(pdisp)
   V_VT(@vd) = VT_DISPATCH : V_DISPATCH(@vd) = pdisp
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from an IUnknown pointer
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DVARIANT_DP("IUNKNOWN PTR")
   IF fAddRef THEN IUnknown_AddRef(punk)
   V_VT(@vd) = VT_UNKNOWN : V_UNKNOWN(@vd) = punk
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from an integer value
' If _vType is wrong, default to VT_I4 (Long).
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
   DVARIANT_DP(" - value, type")
   IF (_vType AND VT_ARRAY) = VT_ARRAY THEN
      ' // NULL array : Suitable for parameters that, instead of being optional,
      ' // require a variant with a null safearray of a given type.
      vd.vt = _vType
   ELSE
      ' // If _value exceeds the limits of max long or max ulong, promote to ulongint
      IF SGN(_value) = -1 AND _value > LONG_MAX THEN
         _vType = VT_I8
      ELSEIF _value > ULONG_MAX THEN
         _vType = VT_I8
      END IF
      SELECT CASE _vType
         CASE VT_BOOL  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
         CASE VT_I1    : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
         CASE VT_UI1   : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
         CASE VT_I2    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
         CASE VT_UI2   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
         CASE VT_INT   : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
         CASE VT_UINT  : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
         CASE VT_I4    : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
         CASE VT_UI4   : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
         CASE VT_I8    : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
         CASE VT_UI8   : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
         ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
         CASE VT_NULL  : vd.vt = VT_NULL
         CASE ELSE     : vd.vt = VT_I4   : vd.llVal = CLNG(_value)
      END SELECT
   END IF
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   DVARIANT_DP(" - value, strType")
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "BOOL"     : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BOOLEAN"  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BYTE"     : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
      CASE "UBYTE"    : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
      CASE "SHORT"    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
      CASE "USHORT"   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
      CASE "INT"      : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
      CASE "UINT"     : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
      CASE "LONG"     : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
      CASE "ULONG"    : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
      CASE "LONGINT"  : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
      CASE "ULONGINT" : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
      ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
      CASE "NULL"     : vd.vt = VT_NULL
      CASE ELSE       : vd.vt = VT_I4   : vd.llVal = CLNG(_value)
   END SELECT
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a float value
' If _vType is wrong, default to VT_R8 (double).
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   DVARIANT_DP("FLOAT - value, vType")
   IF _value > FLT_MAX OR _value < FLT_MIN THEN _vType = VT_R8
   SELECT CASE _vType
      CASE VT_R4    : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE VT_R8    : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
      CASE ELSE     : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
   END SELECT
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   DVARIANT_DP("FLOAT - value, strType")
   SELECT CASE UCASE(strType)
      CASE "SINGLE" : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "FLOAT"  : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "DOUBLE" : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
      CASE ELSE     : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
   END SELECT
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Initializes a DVARIANT from a value by reference (a pointer to a variable), e.g.:
' DIM dblVal AS DOUBLE = 123456.12
' DIM dv AS DVARIANT = DVARIANT(@dblVal, "DOUBLE")
' print dv   ' // prints 123456.12
' dblVal = 345.12
' print dv   ' // prints 345.12
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD)
   DVARIANT_DP("ANY PTR - vType")
   SELECT CASE _vType
      CASE VT_BOOL      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE VT_I1        : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
      CASE VT_UI1       : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
      CASE VT_I2        : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
      CASE VT_UI2       : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
      CASE VT_INT       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
      CASE VT_UINT      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
      CASE VT_I4        : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
      CASE VT_UI4       : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
      CASE VT_I8        : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
      CASE VT_UI8       : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
      CASE VT_R4        : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE VT_R8        : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
      CASE VT_BSTR      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
      CASE VT_UNKNOWN   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
      CASE VT_DISPATCH  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
      CASE VT_DECIMAL   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
      CASE VT_CY        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE VT_DATE      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
      CASE VT_VARIANT   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
      CASE VT_SAFEARRAY : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pparray = _pvar
   END SELECT
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR DVARIANT (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING)
   DVARIANT_DP("ANY PTR - strType")
   IF _pvar <> NULL THEN
      SELECT CASE UCASE(strType)
         CASE "BOOL"      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
         CASE "BOOLEAN"   : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
         CASE "BYTE"      : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
         CASE "UBYTE"     : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
         CASE "SHORT"     : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
         CASE "USHORT"    : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
         CASE "INT"       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
         CASE "UINT"      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
         CASE "LONG"      : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
         CASE "ULONG"     : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
         CASE "LONGINT"   : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
         CASE "ULONGINT"  : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
         CASE "SINGLE"    : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
         CASE "FLOAT"     : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
         CASE "DOUBLE"    : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
         CASE "BSTR"      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
         CASE "UNKNOWN"   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
         CASE "DISPATCH"  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
         CASE "DECIMAL"   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
         CASE "CY"        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
         CASE "CURRENCY"  : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
         CASE "DATE"      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
         CASE "VARIANT"   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
         CASE "SAFEARRAY" : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pparray = _pvar
      END SELECT
   END IF
   m_Result = 0
END CONSTRUCTOR
' ========================================================================================

' ########################################################################################
'                                  CAST OPERATORS
' ########################################################################################

' The CAST operators allow to transparently pass the underlying VARIANT to a procedure.
' They aren't called directly.
'
' SUB Foo (BYREF v AS VARIANT)
'   PRINT AfxVarToStr(@v)
' END SUB
' Foo DVARIANT(12345, "LONG")
'
' SUB Foo2 (BYVAL v AS VARIANT)
'    PRINT AfxVarToStr(@v)
' END SUB
' Foo2 DVARIANT(12345, "LONG")
'
' SUB Foo3 (BYREF dv AS DVARIANT)
'    PRINT dv
' END SUB
' Foo3 DVARIANT(12345, "LONG")
'
' SUB Foo4 (BYVAL dv AS DVARIANT PTR)
'    PRINT *dv
' END SUB
' Foo4 @DVARIANT(12345, "LONG")
'
' Remarks: I haven't added a cast to return a numeric value because with procedures like
' PRINT that can use both a number or a string the compiler will fail, not knowing which
' cast it should use. If you want to convert it to a number, use VAL(DVARIANT).

' ========================================================================================
' Returns the variant data.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.CAST () AS VARIANT
   DVARIANT_DP("VARIANT")
   OPERATOR = vd
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the variant data.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.CAST () AS ANY PTR
   DVARIANT_DP("ANY PTR")
   OPERATOR = cast(ANY PTR, @vd)
END OPERATOR
' ========================================================================================

' =====================================================================================
' Extracts the contents of a VARIANT and returns them as a WSTRING.
' =====================================================================================
PRIVATE OPERATOR DVARIANT.CAST () BYREF AS CONST WSTRING
   DVARIANT_DP("BYREF AS WSTRING")
   m_dws = this.ToStr
   OPERATOR = *CAST(WSTRING PTR, m_dws.GetBuffer)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Clears the DVARIANT and returns the address of the underlying variant.
' To pass the variant to an OUT BYVAL VARIANT PTR parameter.
' If we pass a DVARIANT to a function with an OUT variant parameter without first clearing the
' contents of the CVAR, we may have a memory leak.
' Example:
' SUB Foo (BYVAL v AS VARIANT PTR)
'    v->vt = VT_I4
'    v->lVal = 12345
' END SUB
' DIM dv AS DVARIANT = "Test string"
' Foo dv.vptr
' PRINT dv
' Otherwise, you need to clear the underlying variant before passing the CVAR.
' DIM dv AS DVARIANT = "Test string"
' dv.Clear
' Foo *dv
' PRINT dv
' ========================================================================================
PRIVATE FUNCTION DVARIANT.vptr () AS VARIANT PTR
   DVARIANT_DP("")
   VariantClear @vd
   RETURN @vd
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                  ASSiGNMENT OPERATOR
' ########################################################################################

' ========================================================================================
' Assigns another DVARIANT.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYREF dv AS DVARIANT)
   DVARIANT_DP("DVARIANT")
   ' // No need to clear the variant because VariantCopy clears the destination variant
   m_Result = VariantCopy(@vd, @dv.vd)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL v AS VARIANT)
   DVARIANT_DP("VARIANT")
   m_Result = VariantCopy(@vd, @v)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a VARIANT PTR.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL pvar AS VARIANT PTR)
   DVARIANT_DP("VARIANT PTR")
   m_Result = VariantCopy(@vd, pvar)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a WSTRING PTR.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL pwszStr AS WSTRING PTR)
   DVARIANT_DP("WSTRING PTR")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(pwszStr)
   m_Result = 0
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a DWSTRING.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYREF dws AS DWSTRING)
   DVARIANT_DP("DWSTRING")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(dws)
   m_Result = 0
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a BSTRING.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYREF bs AS BSTRING)
   DVARIANT_DP("BSTRING")
   VariantClear(@vd) : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(bs)
   m_Result = 0
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a CURRENCY structure.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL cy AS CURRENCY)
   DVARIANT_DP("CURRENCY")
   VariantClear(@vd) : vd.vt = VT_CY : vd.cyVal = cy
   m_Result = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a CURRENCY structure.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL dec AS DECIMAL)
   DVARIANT_DP("DECIMAL")
   VariantClear(@vd) : vd.decVal = dec : vd.vt = VT_DECIMAL
   m_Result = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a LONG value.
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL _value AS LONG)
   DVARIANT_DP("LONG")
   VariantClear(@vd) : vd.vt = VT_I4 : vd.lVal = _value
   m_Result = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a BOOLEAN value.
' ========================================================================================
'PRIVATE OPERATOR DVARIANT.LET (BYVAL b AS BOOLEAN)
'   DVARIANT_DP("BOOLEAN")
'   VariantClear(@vd) : vd.vt = VT_BOOL : vd.boolVal = IIF(b = FALSE, 0, -1)
'END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL pdisp AS IDispatch PTR)
   DVARIANT_DP("IDispatch PTR")
   VariantClear(@vd) : V_VT(@vd) = VT_DISPATCH : V_DISPATCH(@vd) = pdisp : IDispatch_AddRef(pdisp)
   m_Result = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns an IUnknown pointer
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL punk AS IUnknown PTR)
   DVARIANT_DP("IUnknown PTR")
   VariantClear(@vd) : V_VT(@vd) = VT_UNKNOWN : V_UNKNOWN(@vd) = punk : IUnknown_AddRef(punk)
   m_Result = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns numeric values
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL _value AS LONGINT)
   DVARIANT_DP("LONGINT")
   VariantClear(@vd)
   IF CLNG(_value) = _value THEN
      vd.vt = VT_I4   : vd.lVal = CLNG(_value)
   ELSE
      vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
   END IF
   m_Result = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.LET (BYVAL _value AS DOUBLE)
   DVARIANT_DP("DOUBLE")
   VariantClear(@vd) : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
   m_Result = 0
END OPERATOR
' ========================================================================================

' ########################################################################################
'                                   ASSIGNMENT METHODS
' ########################################################################################

' ========================================================================================
' Assigns another DVARIANT.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Put (BYREF dv AS DVARIANT) AS HRESULT
   DVARIANT_DP("DVARIANT")
   m_Result = VariantCopy(@vd, @dv.vd)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Put (BYREF v AS VARIANT) AS HRESULT
   DVARIANT_DP("VARIANT")
   m_Result = VariantCopy(@vd, @v)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns a VARIANT.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Put (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DVARIANT_DP("VARIANT PTR")
   m_Result = VariantCopy(@vd, pvar)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns a WSTRING.
' ========================================================================================
PRIVATE SUB DVARIANT.Put (BYVAL pwszStr AS WSTRING PTR)
   DVARIANT_DP("WSTRING PTR")
   m_Result = 0
   VariantClear(@vd)
   vd.vt = VT_BSTR
   vd.bstrVal = SysAllocString(pwszStr)
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a DWSTRING.
' ========================================================================================
PRIVATE SUB DVARIANT.Put (BYREF dws AS DWSTRING)
   DVARIANT_DP("DWSTRING")
   m_Result = 0
   VariantClear(@vd)
   vd.vt = VT_BSTR
   vd.bstrVal = SysAllocString(dws)
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a BSTRING.
' ========================================================================================
PRIVATE SUB DVARIANT.Put (BYREF bs AS BSTRING)
   DVARIANT_DP("BSTRING")
   m_Result = 0
   VariantClear(@vd)
   vd.vt = VT_BSTR
   vd.bstrVal = SysAllocString(bs)
   IF vd.bstrVal = NULL THEN m_Result = E_FAIL
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a Dispatch pointer
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Put (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DVARIANT_DP("IDispatch PTR")
   m_Result = 0
   VariantClear(@vd)
   IF pDisp = NULL THEN
      m_Result = E_INVALIDARG
   ELSE
      V_VT(@vd) = VT_DISPATCH
      V_DISPATCH(@vd) = pdisp
      IF fAddRef THEN IDispatch_AddRef(pdisp)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns an IUnknown pointer
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Put (BYVAL punk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DVARIANT_DP("IUnknown PTR")
   m_Result = 0
   VariantClear(@vd)
   IF punk = NULL THEN
      m_Result = E_INVALIDARG
   ELSE
      V_VT(@vd) = VT_UNKNOWN
      V_UNKNOWN(@vd) = punk
      IF fAddRef THEN IUnknown_AddRef(punk)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns an integer value to a DVARIANT
' If _vType is wrong, default to VT_I8 (LongInt).
' ========================================================================================
PRIVATE SUB DVARIANT.Put (BYVAL _value AS LONGINT, BYVAL _vType AS WORD)
   DVARIANT_DP("value, vType")
   VariantClear(@vd)
   IF (_vType AND VT_ARRAY) = VT_ARRAY THEN
      ' // NULL array : Suitable for parameters that, instead of being optional,
      ' // require a variant with a null safearray of a given type.
      vd.vt = _vType
   ELSE
      ' // If _value exceeds the limits of max long or max ulong, promote to ulongint
      IF SGN(_value) = -1 AND _value > LONG_MAX THEN
         _vType = VT_I8
      ELSEIF _value > ULONG_MAX THEN
         _vType = VT_I8
      END IF
      SELECT CASE _vType
         CASE VT_BOOL  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
         CASE VT_I1    : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
         CASE VT_UI1   : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
         CASE VT_I2    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
         CASE VT_UI2   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
         CASE VT_INT   : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
         CASE VT_UINT  : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
         CASE VT_I4    : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
         CASE VT_UI4   : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
         CASE VT_I8    : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
         CASE VT_UI8   : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
         ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
         CASE VT_NULL  : vd.vt = VT_NULL
         CASE ELSE     : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
      END SELECT
   END IF
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB DVARIANT.Put (BYVAL _value AS LONGINT, BYREF strType AS STRING)
   DVARIANT_DP("value, strType")
   VariantClear(@vd)
   DIM vt AS WORD
   SELECT CASE UCASE(strType)
      CASE "BOOL"     : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BOOLEAN"  : vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
      CASE "BYTE"     : vd.vt = VT_I1   : vd.cVal = CBYTE(_value)
      CASE "UBYTE"    : vd.vt = VT_UI1  : vd.bVal = CUBYTE(_value)
      CASE "SHORT"    : vd.vt = VT_I2   : vd.iVal = CSHORT(_value)
      CASE "USHORT"   : vd.vt = VT_UI2  : vd.uiVal = CUSHORT(_value)
      CASE "INT"      : vd.vt = VT_INT  : vd.intVal = CLNG(_value)
      CASE "UINT"     : vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
      CASE "LONG"     : vd.vt = VT_I4   : vd.lVal = CLNG(_value)
      CASE "ULONG"    : vd.vt = VT_UI4  : vd.ulVal = CULNG(_value)
      CASE "LONGINT"  : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
      CASE "ULONGINT" : vd.vt = VT_UI8  : vd.ullVal = CULNGINT(_value)
      ' // If you need to assign an ULongInt value greater than a LongInt, use PutUlongInt.
      CASE "NULL"     : vd.vt = VT_NULL
      CASE ELSE       : vd.vt = VT_I8   : vd.llVal = CLNGINT(_value)
   END SELECT
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a float value to a DVARIANT
' If _vType is wrong, default to VT_R8 (double).
' ========================================================================================
PRIVATE SUB DVARIANT.Put (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
   DVARIANT_DP("value, vType")
   VariantClear(@vd)
   IF _value > FLT_MAX OR _value < FLT_MIN THEN _vType = VT_R8
   SELECT CASE _vType
      CASE VT_R4 : vd.vt = VT_R4 : vd.fltVal = CSNG(_value)
      CASE VT_R8 : vd.vt = VT_R8 : vd.dblVal = CDBL(_value)
      CASE VT_BOOL, VT_I1, VT_UI1, VT_I2, VT_UI2, VT_INT, VT_UINT, VT_I4, VT_UI4, VT_I8, VT_UI8, VT_NULL
         this.Put CLNGINT(_value), _vType
      CASE ELSE
         vd.vt = VT_R8 : vd.dblVal = CDBL(_value)
   END SELECT
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB DVARIANT.Put (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
   DVARIANT_DP("value, strType")
   VariantClear(@vd)
   SELECT CASE UCASE(strType)
      CASE "SINGLE" : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "FLOAT"  : vd.vt = VT_R4   : vd.fltVal = CSNG(_value)
      CASE "DOUBLE" : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
      CASE ELSE     : vd.vt = VT_R8   : vd.dblVal = CDBL(_value)
   END SELECT
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a value by reference (a pointer to a variable).
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutRef (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD) AS HRESULT
   DVARIANT_DP("pvar, vType")
   VariantClear(@vd)
   IF _pvar = NULL THEN RETURN E_INVALIDARG
   SELECT CASE _vType
      CASE VT_BOOL      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE VT_I1        : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
      CASE VT_UI1       : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
      CASE VT_I2        : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
      CASE VT_UI2       : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
      CASE VT_INT       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
      CASE VT_UINT      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
      CASE VT_I4        : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
      CASE VT_UI4       : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
      CASE VT_I8        : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
      CASE VT_UI8       : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
      CASE VT_R4        : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE VT_R8        : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
      CASE VT_BSTR      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
      CASE VT_UNKNOWN   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
      CASE VT_DISPATCH  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
      CASE VT_DECIMAL   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
      CASE VT_CY        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE VT_DATE      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
      CASE VT_VARIANT   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
      CASE VT_SAFEARRAY : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pvarVal = _pvar
      CASE VT_ERROR     : vd.vt = VT_ERROR OR VT_BYREF     : vd.pparray = _pvar
      CASE ELSE         : RETURN E_INVALIDARG
   END SELECT
   m_Result = 0
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
   DVARIANT_DP("pvar, strType")
   VariantClear(@vd)
   IF _pvar = NULL THEN RETURN E_INVALIDARG
   SELECT CASE UCASE(strType)
      CASE "BOOL"      : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE "BOOLEAN"   : vd.vt = VT_BOOL OR VT_BYREF      : vd.pboolVal = _pvar
      CASE "BYTE"      : vd.vt = VT_I1 OR VT_BYREF        : vd.pcVal = _pvar
      CASE "UBYTE"     : vd.vt = VT_UI1 OR VT_BYREF       : vd.pbVal = _pvar
      CASE "SHORT"     : vd.vt = VT_I2 OR VT_BYREF        : vd.piVal = _pvar
      CASE "USHORT"    : vd.vt = VT_UI2 OR VT_BYREF       : vd.puiVal = _pvar
      CASE "INT"       : vd.vt = VT_INT OR VT_BYREF       : vd.pintVal = _pvar
      CASE "UINT"      : vd.vt = VT_UINT OR VT_BYREF      : vd.puintVal = _pvar
      CASE "LONG"      : vd.vt = VT_I4 OR VT_BYREF        : vd.plVal = _pvar
      CASE "ULONG"     : vd.vt = VT_UI4 OR VT_BYREF       : vd.pulVal = _pvar
      CASE "LONGINT"   : vd.vt = VT_I8 OR VT_BYREF        : vd.pllVal = _pvar
      CASE "ULONGINT"  : vd.vt = VT_UI8 OR VT_BYREF       : vd.pullVal = _pvar
      CASE "SINGLE"    : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE "FLOAT"     : vd.vt = VT_R4 OR VT_BYREF        : vd.pfltVal = _pvar
      CASE "DOUBLE"    : vd.vt = VT_R8 OR VT_BYREF        : vd.pdblVal = _pvar
      CASE "BSTR"      : vd.vt = VT_BSTR OR VT_BYREF      : vd.pbstrVal = _pvar
      CASE "UNKNOWN"   : vd.vt = VT_UNKNOWN OR VT_BYREF   : vd.ppunkVal = _pvar
      CASE "DISPATCH"  : vd.vt = VT_DISPATCH OR VT_BYREF  : vd.ppdispVal = _pvar
      CASE "DECIMAL"   : vd.vt = VT_DECIMAL OR VT_BYREF   : vd.pdecVal = _pvar
      CASE "CY"        : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE "CURRENCY"  : vd.vt = VT_CY OR VT_BYREF        : vd.pcyVal = _pvar
      CASE "DATE"      : vd.vt = VT_DATE OR VT_BYREF      : vd.pdate = _pvar
      CASE "VARIANT"   : vd.vt = VT_VARIANT OR VT_BYREF   : vd.pvarVal = _pvar
      CASE "SAFEARRAY" : vd.vt = VT_SAFEARRAY OR VT_BYREF : vd.pvarVal = _pvar
      CASE "ERROR"     : vd.vt = VT_ERROR OR VT_BYREF     : vd.pparray = _pvar
      CASE ELSE        : RETURN E_INVALIDARG
   END SELECT
   m_Result = 0
END FUNCTION
' ========================================================================================
' ========================================================================================
' Assigns a null value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutNull
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_NULL
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a boolean value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutBool (BYVAL _value AS BOOL)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a boolean value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutBoolean (BYVAL _value AS BOOLEAN)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_BOOL : vd.boolVal = IIF(_value = 0, 0, -1)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a byte value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutByte (BYVAL _value AS BYTE)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_I1 : vd.cVal = CBYTE(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns an ubyte value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutUByte (BYVAL _value AS UBYTE)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_UI1 : vd.bVal = CUBYTE(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a short value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutShort (BYVAL _value AS SHORT)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_I2 : vd.iVal = CSHORT(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns an ushort value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutUShort (BYVAL _value AS USHORT)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_UI2 : vd.uiVal = CUSHORT(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns an INT_ value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutInt (BYVAL _value AS INT_)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_INT : vd.intVal = CLNG(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns an UINT value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutUInt (BYVAL _value AS UINT)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_UINT : vd.uintVal = CULNG(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a LONG value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutLong (BYVAL _value AS LONG)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_I4 : vd.lVal = CLNG(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns an ULONG value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutULong (BYVAL _value AS ULONG)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_UI4 : vd.ulVal = CULNG(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a LONGINT value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutLongInt (BYVAL _value AS LONGINT)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_I8 : vd.llVal = CLNGINT(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns an ULONGINT value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutULongInt (BYVAL _value AS ULONGINT)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_UI8 : vd.ullVal = _value
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a SINGLE value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutSingle (BYVAL _value AS SINGLE)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_R4 : vd.fltVal = CSNG(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a FLOAT value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutFloat (BYVAL _value AS SINGLE)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_R4 : vd.fltVal = CSNG(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Assigns a DOUBLE value.
' ========================================================================================
PRIVATE SUB DVARIANT.PutDouble (BYVAL _value AS DOUBLE)
   DVARIANT_DP("")
   VariantClear(@vd)
   vd.vt = VT_R8 : vd.dblVal = CDBL(_value)
   m_Result = 0
END SUB
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a buffer. Creates a VT_ARRAY OR VT_UI1 variant.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   DVARIANT_DP("")
   VariantClear(@vd)
   m_Result = AfxVariantFromBuffer(pv, cb, @vd)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from the contents of an UTF8 string.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutUtf8 (BYREF strUtf8 AS STRING) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   vd.vt = VT_BSTR
   DIM dws As DWSTRING
   dws.Utf8 = strUtf8
   vd.bstrVal = SysAllocString(dws)
   IF vd.bstrval = NULL THEN m_Result = E_FAIL
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT with the contents of a safe array. If fAttach = FALSE, the safe
' array is copied; if fAttach = TRUE, the safe array is attached.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutSafeArray (BYVAL parray AS SAFEARRAY PTR, BYVAL fAttach AS BOOLEAN = FALSE) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   IF parray = NULL THEN
      m_Result = E_INVALIDARG
      RETURN E_INVALIDARG
   END IF
   DIM vt AS VARTYPE
   m_Result = SafeArrayGetVartype(parray, @vt)
   IF m_Result <> S_OK THEN RETURN m_Result
   IF fAttach = FALSE THEN
      DIM pcopy AS SAFEARRAY PTR
      m_Result = SafeArrayCopy(parray, @pcopy)
      IF m_Result <> S_OK THEN RETURN m_Result
      vd.vt = vt OR VT_ARRAY
      vd.parray = pcopy
   ELSE
      vd.vt = vt OR VT_ARRAY
      vd.parray = parray
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Initializes a DVARIANT from a a string resource imbedded in an executable file.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT) AS HRESULT
   DVARIANT_DP("")
   VariantClear(@vd)
   m_Result = AfxVariantFromResource(hinst, id, @vd)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' =====================================================================================
' Initializes a DVARIANT with a reference to an UDT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutRecord (BYVAL pIRecordInfo AS IRecordInfo PTR, BYVAL pRec AS ANY PTR) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   IF pIRecordInfo = NULL OR pRec = NULL THEN
      m_Result = E_INVALIDARG
      RETURN E_INVALIDARG
   END IF
   vd.vt = VT_RECORD
   vd.pvRecord = pRec
   vd.pRecInfo = pIRecordInfo
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT as VT_DATE from a string.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   DIM dateOut AS DOUBLE
   m_Result = VarDateFromStr(pwszDate, lcid, dwFlags, @dateOut)
   IF m_Result = S_OK THEN vd.vt = VT_DATE : vd.date = dateOut
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT as VT_DATE from a DATE type.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutVbDate (BYVAL vbDate AS DATE_) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   vd.vt = VT_DATE
   vd.date = vbDate
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Converts a SYSTEMTIME to a VT_DATE variant.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutSystemTime (BYVAL st AS SYSTEMTIME PTR) AS BOOLEAN
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   DIM vbDate AS DATE_
   DIM bRes AS BOOLEAN = SystemTimeToVariantTime(st, @vbDate)
   IF bRes THEN vd.vt = VT_DATE : vd.date = vbDate
   IF bRes =FALSE THEN m_Result = E_FAIL
   RETURN bRes
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT based on a GUID. The structure is initialized as a VT_BSTR type.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutGuid (BYVAL guid AS IID PTR) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromGUIDAsString(guid, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT with the contents of a FILETIME structure.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromFileTime(pft, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' ========================================================================================
' Initializes a DVARIANT with an array of FILETIME structures.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromFileTimeArray(prgft, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT with a string stored in a STRRET structure.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromStrRet(pstrret, pidl, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT as VT_DECIMAL from a DECIMAL structure.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutDec (BYVAL dec AS DECIMAL) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   vd.decVal = dec
   vd.vt = VT_DECIMAL
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT as VT_DECIMAL from a string.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   IF pwszIn = NULL THEN
      m_Result = E_INVALIDARG
      RETURN E_INVALIDARG
   ELSE
      DIM dec AS DECIMAL
      m_Result = VarDecFromStr(pwszIn, lcid, dwFlags, @dec)
      IF m_Result = S_OK THEN vd.decVal = dec: vd.vt = VT_DECIMAL
   END IF
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT as VT_DECIMAL from a double value.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   DIM dec AS DECIMAL
   m_Result = VarDecFromR8(dbIn, @dec)
   IF m_Result = S_OK THEN vd.decVal = dec : vd.vt = VT_DECIMAL
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Converts a currency value to a DVARIANT of type VT_DECIMAL.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutDecFromCy (BYVAL cyIn AS CY) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   DIM decOut AS DECIMAL
   m_Result = VarDecFromCy(cyIn, @decOut)
   IF m_Result = S_OK THEN vd.decVal = decOut : vd.vt = VT_DECIMAL
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of Boolean values.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromBooleanArray(prgf, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of 16-bit integer values.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromInt16Array(prgf, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of unsigned 16-bit integer values.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutUShortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromUInt16Array(prgf, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of 32-bit integer values.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutLongArray (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG) AS DVARIANT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromInt32Array(prgn, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of 32-bit integer values.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutULongArray (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromUInt32Array(prgn, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of 64-bit integer values.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromInt64Array(prgn, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of unsigned 64-bit integer values.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromUInt64Array(prgn, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of values of type DOUBLE.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromDoubleArray(prgn, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Initializes a DVARIANT from an array of unicode strings.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromStringArray(prgsz, cElems, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Converts the contents of a PROPVARIANT structure to a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxPropVariantToVariant(pPropVar, @vd)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' ========================================================================================
' Initializes a DVARIANT from a specified variant element.
' Called VarGetElement (inline function) in the propvarutil.h header, which I find confusing.
' VariantFromVariantArrayElement is more descriptive.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.PutVariantArrayElement (BYVAL pvarIn AS VARIANT PTR, BYVAL iElem AS ULONG) AS DVARIANT
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
   m_Result = AfxVariantFromVariantArrayElem(pvarIn, iElem, @vd)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                     SAFEARRAYS
' ########################################################################################

' =====================================================================================
' Returns the number of dimensions for variants of type VT_ARRAY; returns 0 otherwise.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetDim () AS ULONG
   DVARIANT_DP("")
   m_Result = 0
   IF (vd.vt AND VT_ARRAY) = VT_ARRAY THEN
      IF vd.parray THEN RETURN SafeArrayGetDim(vd.parray)
   END IF
END FUNCTION
' =====================================================================================
' =====================================================================================
' Returns the lower bound for the specified dimension of the safe array for variants of
' type VT_ARRAY; returns 0 otherwise.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetLBound (BYVAL nDim AS UINT = 1) AS LONG
   DVARIANT_DP("")
   m_Result = 0
   IF (vd.vt AND VT_ARRAY) = VT_ARRAY THEN
      DIM nBound AS LONG
      IF vd.parray THEN m_Result = SafeArrayGetLBound(vd.parray, nDim, @nBound)
      IF m_Result = S_OK THEN RETURN nBound
   END IF
END FUNCTION
' =====================================================================================
' =====================================================================================
' Returns the upper bound for the specified dimension of the safe array for variants of
' type VT_ARRAY; returns -1 otherwise.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetUBound (BYVAL nDim AS UINT = 1) AS LONG
   DVARIANT_DP("")
   m_Result = 0
   IF (vd.vt AND VT_ARRAY) = VT_ARRAY THEN
      DIM nBound AS LONG
      IF vd.parray THEN m_Result = SafeArrayGetUBound(vd.parray, nDim, @nBound)
      IF m_Result = S_OK THEN RETURN nBound
   END IF
   RETURN -1
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single Variant element from a safe array of variants.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetVariantElement (BYVAL iElem AS ULONG) AS DVARIANT
   DVARIANT_DP("")
   m_Result = 0
   DIM vItem AS VARIANT
   IF vd.vt = VT_ARRAY OR VT_VARIANT THEN
      DIM idx AS LONG = iElem
      m_Result = SafeArrayGetElement(vd.parray, @idx, @vItem)
   END IF
   RETURN vItem
END FUNCTION
' =====================================================================================

' ########################################################################################
'                                      ARRAYS
' ########################################################################################

' =====================================================================================
' Returns the number of elements for variants of type VT_ARRAY; otherwise, returns 1.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetElementCount () AS ULONG
   DVARIANT_DP("")
   m_Result = 0
   RETURN AfxVariantGetElementCount(@vd)
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single Boolean element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetBooleanElement (BYVAL iElem AS ULONG) AS BOOLEAN
   DVARIANT_DP("")
   DIM pfVal AS WINBOOL
   m_Result = AfxVariantGetBooleanElem(@vd, iElem, @pfVal)
   RETURN pfVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single Int16 element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetShortElement (BYVAL iElem AS ULONG) AS SHORT
   DVARIANT_DP("")
   DIM pnVal AS SHORT
   m_Result = AfxVariantGetInt16Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single unsigned Int16 element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetUShortElement (BYVAL iElem AS ULONG) AS USHORT
   DVARIANT_DP("")
   DIM pnVal AS USHORT
   m_Result = AfxVariantGetUInt16Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single Int32 element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetLongElement (BYVAL iElem AS ULONG) AS LONG
   DVARIANT_DP("")
   DIM pnVal AS LONG
   m_Result = AfxVariantGetInt32Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single unsigned Int32 element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetULongElement (BYVAL iElem AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pnVal AS ULONG
   m_Result = AfxVariantGetUInt32Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single Int64 element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetLongIntElement (BYVAL iElem AS ULONG) AS LONGINT
   DVARIANT_DP("")
   DIM pnVal AS LONGINT
   m_Result = AfxVariantGetInt64Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a single UInt64 element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetULongIntElement (BYVAL iElem AS ULONG) AS ULONGINT
   DVARIANT_DP("")
   DIM pnVal AS ULONGINT
   m_Result = AfxVariantGetUInt64Elem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts one double element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetDoubleElement (BYVAL iElem AS ULONG) AS DOUBLE
   DVARIANT_DP("")
   DIM pnVal AS DOUBLE
   m_Result = AfxVariantGetDoubleElem(@vd, iElem, @pnVal)
   RETURN pnVal
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an string element from a DVARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.GetStringElement (BYVAL iElem AS ULONG) AS DWSTRING
   DVARIANT_DP("")
   DIM pwszVal AS WSTRING PTR, dwsVal AS DWSTRING
   m_Result = AfxVariantGetStringElem(@vd, iElem, @pwszVal)
   IF pwszVal THEN
      dwsVal = *pwszVal
      CoTaskMemFree pwszVal
   END IF
   RETURN dwsVal
END FUNCTION
' =====================================================================================

' ########################################################################################
'                                  CONVERSION METHODS
' ########################################################################################

' ========================================================================================
' Extracts the value of a variant structure to a string.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.ToStr () AS DWSTRING
   DVARIANT_DP(..WSTR(@vd))
   m_Result = 0
   RETURN AfxVarToStr(@vd)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts the value of a variant structure to a utf8 string.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.ToUtf8 () AS STRING
   DVARIANT_DP("")
   m_Result = 0
   DIM dws AS DWSTRING = AfxVarToStr(@vd)
   RETURN dws.utf8
END FUNCTION
' ========================================================================================
' =====================================================================================
' Extracts the contents of a buffer stored in a DVARIANT of type VT_ARRRAY OR VT_UI1.
' Note: To retrieve the size of the array call GetElementCount.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToBuffer (BYVAL pv AS ANY PTR, BYVAL cb AS UINT) AS HRESULT
   DVARIANT_DP("")
   m_Result = AfxVariantToBuffer(@vd, pv, cb)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts the contents of a buffer stored in a DVARIANT of type VT_ARRRAY OR VT_UI1 and
' returns it as string used as a buffer.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToBuffer () AS STRING
   DVARIANT_DP("STRING")
   m_Result = 0
   IF (vd.vt = VT_ARRAY OR VT_UI1) THEN
      DIM cb AS LONG = this.GetElementCount
      IF cb THEN
         DIM s AS STRING = SPACE(cb)
         this.ToBuffer(STRPTR(s), cb)
         RETURN s
      END IF
   END IF
   RETURN ""
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts the contents of a VARIANT and returns them as a BSTRING.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToBStr () AS BSTRING
   DVARIANT_DP("BSTRING")
   m_Result = 0
   RETURN AfxVarToStr(@vd)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks if the variant if of the type VT_UNKNOWN and returns the unknown pointer.
' It is the responsability of the called of release the returned pointer.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToUnknown () AS ANY PTR
   DVARIANT_DP("")
   m_Result = 0
   IF vd.vt <> VT_UNKNOWN AND vd.vt <> VT_DISPATCH THEN RETURN NULL
   IF vd.vt = VT_UNKNOWN THEN
      IF vd.punkVal = NULL THEN RETURN NULL
      IUnknown_AddRef(vd.punkVal)
      RETURN vd.punkVal
   END IF
   IF vd.vt = VT_DISPATCH THEN
      IF vd.pdispVal = NULL THEN RETURN NULL
      DIM punk AS IUnknown PTR
      m_Result = vd.pdispVal->lpvtbl->QueryInterface(vd.pdispVal, @IID_IUnknown, @punk)
      IF m_Result = S_OK THEN
         ' // QueryInterface calls IUnknown_AddRef in the pointer it returns
         RETURN punk
      ELSE
         IDispatch_AddRef(vd.pdispVal)
         RETURN vd.pdispVal
      END IF
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' If the variant is of the type VT_UNKNOWN or VT_DISPATCH it returns the dispatch pointer.
' It is the responsability of the called of release the returned pointer.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToDispatch () AS ANY PTR
   DVARIANT_DP("")
   m_Result = 0
   IF vd.vt <> VT_UNKNOWN AND vd.vt <> VT_DISPATCH THEN RETURN NULL
   IF vd.vt = VT_DISPATCH THEN
      IF vd.pdispVal = NULL THEN RETURN NULL
      IDispatch_AddRef(vd.pdispVal)
      RETURN vd.pdispVal
   END IF
   IF vd.vt = VT_UNKNOWN THEN
      IF vd.punkVal = NULL THEN RETURN NULL
      DIM pdisp AS IDispatch PTR
      m_Result = vd.punkVal->lpvtbl->QueryInterface(vd.punkVal, @IID_IDispatch, @pdisp)
      IF m_Result = S_OK THEN
         ' // QueryInterface calls IUnknown_AddRef in the pointer it returns
         RETURN pdisp
      ELSE
         IUnknown_AddRef(vd.punkVal)
         RETURN vd.punkVal
      END IF
   END IF
END FUNCTION
' =====================================================================================
' =====================================================================================
' Converts a variant of type decimal to a double.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.DecToDouble () AS DOUBLE
   DVARIANT_DP("")
   m_Result = 0
   DIM dblOut AS DOUBLE
   IF vd.vt = VT_DECIMAL THEN VarR8FromDec(@vd.decVal, @dblOut)
   RETURN dblOut
END FUNCTION
' =====================================================================================
' =====================================================================================
' Converts a variant of type decimal to currency.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.DecToCy () AS CY
   DVARIANT_DP("")
   m_Result = 0
   DIM cyOut AS CY
   IF vd.vt = VT_DECIMAL THEN VarCyFromDec(@vd.decVal, @cyOut)
   RETURN cyOut
END FUNCTION
' =====================================================================================
' =====================================================================================
' Converts a VT_DATE variant to a variant representation (double)
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToVbDate () AS DATE_
   DVARIANT_DP("")
   m_Result = 0
   IF vd.vt = VT_DATE THEN RETURN vd.date
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a FILETIME structure from a CVAR as a SYSTEMTIME structure.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToSystemTime () AS SYSTEMTIME
   DVARIANT_DP("")
   m_Result = 0
   DIM st AS SYSTEMTIME
   IF vd.vt = VT_DATE THEN VariantTimeToSystemTime(vd.date, @st)
   RETURN st
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a GUID property value of a VARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToGuid () AS GUID
   DVARIANT_DP("")
   m_Result = 0
   DIM guid_ AS GUID : AfxVariantToGUID(@vd, @guid_) : RETURN guid_
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a GUID property value as a DWSTRING.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToGuidStr () AS DWSTRING
   DVARIANT_DP("")
   m_Result = 0
   DIM guid_ AS GUID, dwsGuid AS DWSTRING, pwsz AS WSTRING PTR
   DIM m_Result AS HRESULT = AfxVariantToGUID(@vd, @guid_)
   IF m_Result = S_OK THEN
      StringFromCLSID(@guid_, CAST(LPOLESTR PTR, @pwsz))
      IF pwsz THEN dwsGuid = *pwsz
      CoTaskMemFree(pwsz)
   END IF
   RETURN dwsGuid
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a date and time value in Microsoft MS-DOS format from a VARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToDosDateTime (BYVAL pwDate AS USHORT PTR, BYVAL pwTime AS USHORT PTR) AS HRESULT
   DVARIANT_DP("")
   m_Result = AfxVariantToDosDateTime(@vd, pwDate, pwTime)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts a FILETIME structure from a VARIANT.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToFileTime (BYVAL stfOut AS AFX_PSTIME_FLAGS) AS FILETIME
   DVARIANT_DP("")
   DIM ft AS FILETIME
   m_Result = AfxVariantToFileTime(@vd, stfOut, @ft)
   RETURN ft
END FUNCTION
' =====================================================================================
' =====================================================================================
' If the source variant is a VT_BSTR, extracts string and places it into a STRRET structure.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToStrRet () AS STRRET
   DVARIANT_DP("")
   DIM strret_ AS STRRET
   m_Result = AfxVariantToStrRet(@vd, @strret_)
   RETURN strret_
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of Boolean values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToBooleanArray(@vd, prgf, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Allocates an array of BOOL values then extracts data from a VARIANT into that array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note: Use CoTaskMemFree to free the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToBooleanArrayAlloc (BYVAL pprgf AS WINBOOL PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToBooleanArrayAlloc(@vd, pprgf, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of Int16 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToShortArray (BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToInt16Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToShortArrayAlloc (BYVAL pprgn AS SHORT PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToInt16ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of UInt16 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToUShortArray (BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToUInt16Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToUShortArrayAlloc (BYVAL pprgn AS USHORT PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToUInt16ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of Int32 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToLongArray (BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToInt32Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int32 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToLongArrayAlloc (BYVAL pprgn AS LONG PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToInt32ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of UInt32 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToULongArray (BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToUInt32Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int32 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToULongArrayAlloc (BYVAL pprgn AS ULONG PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToUInt32ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of Int64 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToInt64Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated Int64 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToLongIntArrayAlloc (BYVAL pprgn AS LONGINT PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToInt64ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of Int64 values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToUInt64Array(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int64 array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToULongIntArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToUInt64ArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of Double values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToDoubleArray(@vd, prgn, crgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Allocates an array of DOUBLE values then extracts data from a VARIANT into that array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to release the memory allocated for the array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToDoubleArrayAlloc (BYVAL pprgn AS DOUBLE PTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToDoubleArrayAlloc(@vd, pprgn, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts an array of string values from a VARIANT structure.
' Returns the number of elements extracted from the source Variant structure.
' GetLastError: Returns S_OK if successful, or an error value otherwise, including the following:
' TYPE_E_BUFFERTOOSMALL: The source VARIANT contained more than crgn values.
' E_INVALIDARG: The VARIANT was not of the appropriate type.
' To retrieve the number of elements in the array call the GetElementCount method.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToStringArray (BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToStringArray(@vd, prgsz, crgsz, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================
' =====================================================================================
' Extracts data from a vector structure into a newly-allocated String array.
' GetLastError: Returns S_OK if successful, or an error value otherwise.
' Note Use CoTaskMemFree to free the memory used by each of the strings and the returned array.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ToStringArrayAlloc (BYVAL pprgsz AS PWSTR PTR) AS ULONG
   DVARIANT_DP("")
   DIM pcElem AS ULONG
   m_Result = AfxVariantToStringArrayAlloc(@vd, pprgsz, @pcElem)
   RETURN pcElem
END FUNCTION
' =====================================================================================

' ########################################################################################
 '                                    MISCELLANEOUS
' ########################################################################################

' ========================================================================================
' Clears the contents of the variant data and sets the vt field to VT_EMPTY.
' ========================================================================================
PRIVATE SUB DVARIANT.Clear
   DVARIANT_DP("")
   m_Result = 0
   VariantClear(@vd)
END SUB
' ========================================================================================
' ========================================================================================
' Attaches a variant to this class. The source variant is marked as empty.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Attach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DVARIANT_DP("VARIANT PTR")
   m_Result = 0
   IF pvar = NULL THEN
      m_Result = E_INVALIDARG
      RETURN E_INVALIDARG
   END IF
   VariantClear @vd
   ' // Copy the contents and give control to CVar
   DIM pdest AS ANY PTR = memcpy(@vd, pvar, SIZEOF(VARIANT))
   IF pdest = NULL THEN
      m_Result = E_FAIL
      RETURN E_FAIL
   END IF
   ' // Mark the source variant as VT_EMPTY instead of clearing it with VariantClear
   ' // because we aren't making a duplicate of the contents, but transfering ownership.
   pvar->vt = VT_EMPTY
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Attach (BYREF v AS VARIANT) AS HRESULT
   DVARIANT_DP("BYREF VARIANT")
   m_Result = 0
   VariantClear @vd
   ' // Copy the contents and give control to CVar
   DIM pdest AS ANY PTR = memcpy(@vd, @v, SIZEOF(VARIANT))
   IF pdest = NULL THEN
      m_Result = E_FAIL
      RETURN E_FAIL
   END IF
   ' // Mark the source variant as VT_EMPTY instead of clearing it with VariantClear
   ' // because we aren't making a duplicate of the contents, but transfering ownership.
   v.vt = VT_EMPTY
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the variant data from this class and returs it as a VARIANT.
' Don't clear vd with VariantClear because we are transfering ownership.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DVARIANT_DP("VARIANT PTR")
   m_Result = 0
   IF pvar = NULL THEN
      m_Result = E_FAIL
      RETURN E_INVALIDARG
   END IF
   VariantClear(pvar)
   DIM pdest AS ANY PTR = memcpy(pvar, @vd, SIZEOF(VARIANT))
   IF pdest = NULL THEN
      m_Result = E_FAIL
      RETURN E_FAIL
   END IF
   vd.vt = VT_EMPTY
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION DVARIANT.Detach (BYREF v AS VARIANT) AS HRESULT
   DVARIANT_DP("BYREF VARIANT")
   m_Result = 0
   VariantClear(@v)
   DIM pdest AS ANY PTR = memcpy(@v, @vd, SIZEOF(VARIANT))
   IF pdest = NULL THEN
      m_Result = E_FAIL
      RETURN E_FAIL
   END IF
   vd.vt = VT_EMPTY
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the VARIANT type.
' See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms221170(v=vs.85).aspx
' ========================================================================================
PRIVATE FUNCTION DVARIANT.vType () AS VARTYPE
   DVARIANT_DP("")
   m_Result = 0
   FUNCTION = vd.vt
END FUNCTION
' ========================================================================================

' =====================================================================================
' Converts the variant from one type to another.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ChangeType (BYVAL vtNew AS VARTYPE, BYVAL wFlags AS USHORT = 0) AS HRESULT
   DVARIANT_DP("")
   m_Result = VariantChangeType(@vd, @vd, wFlags, vtNew)
   RETURN m_Result
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DVARIANT.ChangeTypeEx (BYVAL vtNew AS VARTYPE, BYVAL lcid AS LCID = 0, BYVAL wFlags AS USHORT = 0) AS HRESULT
   DVARIANT_DP("")
   m_Result = VariantChangeTypeEx(@vd, @vd, lcid, wFlags, vtNew)
   RETURN m_Result
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION DVARIANT.Round (BYREF dv AS DVARIANT, BYVAL cDecimals AS LONG) AS DVARIANT
   DVARIANT_DP("")
   DIM dvRes AS DVARIANT
   m_Result = VarRound(@dv.vd, cDecimals, @dvRes.vd)
   RETURN dvRes
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION DVARIANT.Round (BYVAL cDecimals AS LONG) AS DVARIANT
   DVARIANT_DP("")
   m_Result = 0
   DIM dwRes AS DVARIANT = vd
   m_Result = VarRound(@vd, cDecimals, @dwRes.vd)
   IF m_Result = S_OK THEN vd = dwRes.vd
   RETURN dwRes
END FUNCTION
' =====================================================================================
' =====================================================================================
' Formats a variant containing numbers into a string form.
' =====================================================================================
PRIVATE FUNCTION DVARIANT.FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DVARIANT_DP("")
   DIM dwsOut AS DWSTRING, bstrOut AS BSTR
   m_Result = VarFormatNumber(@vd, iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut)
   dwsOut = *CAST(WSTRING PTR, bstrOut)
   SysFreeString bstrOut
   RETURN dwsOut
END FUNCTION
' =====================================================================================

' ########################################################################################
'                                   MATH OPERATORS
' ########################################################################################

' ========================================================================================
PRIVATE OPERATOR DVARIANT.+= (BYREF dv AS DVARIANT)
   IF vd.vt = VT_BSTR AND dv.vd.vt = VT_BSTR THEN
      ' // Both values are strings, so concatenate
      m_Result = VarCat(@vd, @dv.vd, @vd)
   ELSE
      ' // Add
      DIM vRes AS VARIANT
      m_Result = VarAdd(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.&= (BYREF dv AS DVARIANT)
   IF vd.vt = VT_BSTR AND dv.vd.vt = VT_BSTR THEN
      ' // Both values are strings, so concatenate
      m_Result = VarCat(@vd, @dv.vd, @vd)
   ELSE
      ' // Add
      DIM vRes AS VARIANT
      m_Result = VarAdd(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.-= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result= VarSub(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.*= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarMul(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT./= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarDiv(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.\= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarIDiv(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.^= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarPow(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.And= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarAnd(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.Or= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarOr(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.Xor= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarXor(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.Eqv= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarEqv(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.Imp= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarImp(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR DVARIANT.Mod= (BYREF dv AS DVARIANT)
   m_Result = 0
   IF vd.vt <> VT_BSTR AND dv.vd.vt <> VT_BSTR THEN
      DIM vRes AS VARIANT
      m_Result = VarMod(@vd, @dv.vd, @vRes)
      IF m_Result = S_OK THEN m_Result = VariantCopy(@vd, @vRes)
   END IF
END OPERATOR
' ========================================================================================

' ########################################################################################
 '                                      ERRORS
' ########################################################################################

' =====================================================================================
' Returns the last result code
' ========================================================================================
PRIVATE FUNCTION DVARIANT.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION DVARIANT.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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

END NAMESPACE

' ########################################################################################
'                               *** GLOBAL OPERATORS ***
' ########################################################################################

' // Outside a namespace because they are global
USING AfxNova

' ========================================================================================
' Returns the address of the data of the underlying variant. Same as STRPTR.
' **dv Can't be used to deference the string data because * is already overloaded to
' perform multiply. To deference the string data use *STRPTR(dv)
' ========================================================================================
PRIVATE OPERATOR * (BYREF dv AS DVARIANT) AS VARIANT PTR
   DVARIANT_DP(..WSTR(CAST(VARIANT PTR, STRPTR(dv))))
   OPERATOR = CAST(VARIANT PTR, STRPTR(dv))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR & (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   DIM dvRes AS DVARIANT
   VarCat(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = dvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Left OVERLOAD (BYREF dv AS DVARIANT, BYVAL nChars AS INTEGER) AS DWSTRING
   DIM dws AS DWSTRING = AfxVarToStr(@dv.vd)
   RETURN LEFT(dws, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Right OVERLOAD (BYREF dv AS DVARIANT, BYVAL nChars AS INTEGER) AS DWSTRING
   DIM dws AS DWSTRING = AfxVarToStr(@dv.vd)
   RETURN RIGHT(dws, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Val OVERLOAD (BYREF dv AS DVARIANT) AS DOUBLE
   DIM dws AS DWSTRING = AfxVarToStr(@dv.vd)
   RETURN .VAL(dws)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Math operators
' ========================================================================================
PRIVATE OPERATOR + (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   DIM dvRes AS DVARIANT
   IF dv1.vd.vt = VT_BSTR OR dv2.vd.vt = VT_BSTR THEN
      ' // One of the values is a string, so concatenate
      dvRes = dv1 & dv2
   ELSE
      ' // Add
      VarAdd(@dv1.vd, @dv2.vd, @dvRes.vd)
   END IF
   OPERATOR = dvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   ' // Subtract
   DIM dvRes AS DVARIANT
   VarSub(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = dvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   ' // Multiply
   DIM dvRes AS DVARIANT
   VarMul(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = dvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   ' // Divide
   DIM dvRes AS DVARIANT
   VarDiv(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = dvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR \ (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
   ' // Integer divide
   DIM dvRes AS DVARIANT
   VarIDiv(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = dvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR = (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
   ' // Equality
   RETURN (VarCmp(@dv1.vd, @dv2.vd, GetThreadLocale(), 0) = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
   ' // Inequality
   RETURN (VarCmp(@dv1.vd, @dv2.vd, GetThreadLocale(), 0) <> VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
   RETURN (VarCmp(@dv1.vd, @dv2.vd, GetThreadLocale(), 0) = VARCMP_LT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
   RETURN (VarCmp(@dv1.vd, @dv2.vd, GetThreadLocale(), 0) = VARCMP_GT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
   DIM hr AS HRESULT = VarCmp(@dv1.vd, @dv2.vd, GetThreadLocale(), 0)
   RETURN (hr = VARCMP_LT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
   DIM hr AS HRESULT = VarCmp(@dv1.vd, @dv2.vd, GetThreadLocale(), 0)
   RETURN (hr = VARCMP_GT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF dv AS DVARIANT) AS DVARIANT
   ' // Negate
   DIM dvRes AS DVARIANT
   VarNeg(@dv.vd, @dvRes.vd)
   OPERATOR = dvRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Not (BYREF dv AS DVARIANT) AS INTEGER
   ' // Bitwise negation
   DIM dvRes AS DVARIANT
   VarNot(@dv.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR And (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
   DIM dvRes AS DVARIANT
   VarAnd(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Or (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
   DIM dvRes AS DVARIANT
   VarOr(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Xor (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
   DIM dvRes AS DVARIANT
   VarXor(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Mod (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
   DIM dvRes AS DVARIANT
   VarMod(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Imp (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
   DIM dvRes AS DVARIANT
   VarImp(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Eqv (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
   DIM dvRes AS DVARIANT
   VarEqv(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR ^ (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DOUBLE
   DIM dvRes AS DVARIANT
   VarPow(@dv1.vd, @dv2.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Abs (BYREF dv AS DVARIANT) AS DOUBLE
   DIM dvRes AS DVARIANT
   VarAbs(@dv.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Fix (BYREF dv AS DVARIANT) AS DOUBLE
   DIM dvRes AS DVARIANT
   VarFix(@dv.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR Int (BYREF dv AS DVARIANT) AS DOUBLE
   DIM dvRes AS DVARIANT
   VarInt(@dv.vd, @dvRes.vd)
   OPERATOR = VAL(dvRes)
END OPERATOR
' ========================================================================================


' ########################################################################################
'                               *** HELPER FUNCTIONS ***
' ########################################################################################

' ========================================================================================
' Helper procedure to display messages
' ========================================================================================
PRIVATE FUNCTION AfxMsg OVERLOAD (BYREF cv AS DVARIANT, BYREF wszCaption AS WSTRING = "Message", BYVAL uType AS DWORD = 0) AS LONG
   FUNCTION = MessageBoxW(GetActiveWindow, CAST(WSTRING PTR, STRPTR(cv)), @wszCaption, MB_APPLMODAL OR uType)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the contents of a DVARIANT to a DWSTRING.
' ========================================================================================
PRIVATE FUNCTION AfxDVarToStr OVERLOAD (BYREF dv AS DVARIANT) AS DWSTRING
   RETURN AfxVarToStr(@dv.vd)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxDVarToStr OVERLOAD (BYVAL pcv AS DVARIANT PTR) AS DWSTRING
   RETURN AfxVarToStr(@pcv->vd)
END FUNCTION
' ========================================================================================
#define DVAR_TOSTR(cv) AfxDVarToStr(cv)
' ========================================================================================
PRIVATE FUNCTION AfxDVarOptPrm () AS DVARIANT
   DIM v AS VARIANT = TYPE(VT_ERROR, 0, 0, 0, DISP_E_PARAMNOTFOUND)
   DIM dvOpt AS DVARIANT = v
   RETURN dvOpt
END FUNCTION
' ========================================================================================
#define DVAR_OPTPRM AfxDVarOptPrm

' ========================================================================================
' Extracts the contents of a variant that contains an array of bytes.
' ========================================================================================
PRIVATE FUNCTION AfxDVarToBuffer (BYREF dvIn AS DVARIANT, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
   RETURN AfxVariantToBuffer(@dvIn.vd, pv, cb)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                      LINKED LIST
' ########################################################################################

' USage example:
' // Build the linked list
' DIM List AS DVarList PTR = NEW DVarList
' List->Add("Result 1")
' List->Add("Result 2")
' List->Add("Result 3")
' List->Insert(1, "New string")
' List->Replace(2, "Replaced string")
' // Retrieve and print the results
' FOR i AS LONG = 1 TO List->Count
'    PRINT List->Item(i)
' NEXT
' // Delete the list
' Delete List


NAMESPACE AfxNova

' // Doubly-linked list

'TYPE DVarNode
'   dv AS DVARIANT
'   pNext AS DVarNode PTR
'   pPrev AS DVarNode PTR
'   DECLARE DESTRUCTOR
'END TYPE

'DESTRUCTOR DVarNode
''   OutputDebugStringW "DVarNode destructor"
'END DESTRUCTOR

'TYPE DVarList
'   head AS DVarNode PTR
'   tail AS DVarNode PTR
'   nodes AS LONG ' number of nodes
'   DIM rgIdx(ANY) AS DVarNode PTR ' array for indexed access
'   DECLARE DESTRUCTOR
'   DECLARE FUNCTION Add (BYREF dv AS DVARIANT) AS LONG
'   DECLARE FUNCTION Count () AS LONG
'   DECLARE FUNCTION Item (BYVAL idx AS LONG) AS DVARIANT
'   DECLARE FUNCTION Remove (BYVAL idx AS LONG) AS BOOLEAN
'   DECLARE FUNCTION Insert (BYVAL idx AS LONG, BYREF dws AS DVARIANT) AS BOOLEAN
'   DECLARE FUNCTION Replace (BYVAL idx AS LONG, BYREF dws AS DVARIANT) AS BOOLEAN
'   DECLARE SUB Clear
'END TYPE

'PRIVATE DESTRUCTOR DVarList
''   OutputDebugStringW "DVarList destructor"
'   ' // Delete all the nodes
'   FOR i AS LONG = LBOUND(rgIdx) TO UBOUND(rgIdx)
'      Delete rgIdx(i)
'   NEXT
'   ' // Erase array of indexes
'   ERASE rgIdx
'END DESTRUCTOR

'PRIVATE SUB DVarList.Clear
''   OutputDebugStringW "DVarList.Clear"
'   FOR i AS LONG = LBOUND(rgIdx) TO UBOUND(rgIdx)
'      Delete rgIdx(i)
'   NEXT
'   ERASE rgIdx
'   head = NULL
'   tail = NULL
'   nodes = 0
'END SUB

'PRIVATE FUNCTION DVarList.Add (BYREF dv AS DVARIANT) AS LONG
'   DIM newNode AS DVarNode PTR = NEW DVarNode
'   newNode->dv = dv
'   newNode->pNext = NULL
'   newNode->pPrev = tail
'   IF tail <> NULL THEN
'      tail->pNext = newNode
'   ELSE
'      head = newNode
'   END IF
'   tail = newNode
'   nodes += 1
'   REDIM PRESERVE rgIdx(nodes - 1)
'   rgIdx(nodes - 1) = newNode
'   RETURN nodes
'END FUNCTION


'' Get the number of nodes
'PRIVATE FUNCTION DVarList.Count () AS LONG
'   RETURN nodes
'END FUNCTION

'PRIVATE FUNCTION DVarList.Item (BYVAL idx AS LONG) AS DVARIANT
'   idx -= 1
'   IF idx >= 0 AND idx < nodes THEN
'      RETURN rgIdx(idx)->dv
'   END IF
'   RETURN DVARIANT("") ' or empty variant
'END FUNCTION

'' Remove a node by index
'PRIVATE FUNCTION DVarList.Remove (BYVAL idx AS LONG) AS BOOLEAN
'   idx -= 1   ' make it one-based
'   IF idx < 0 OR idx >= nodes THEN RETURN FALSE ' Prevent out-of-bounds deletion
'   DIM current AS DVarNode PTR = rgIdx(idx) ' Get node from rgIdx array
'   IF current = NULL THEN RETURN FALSE ' Invalid rgIdx
'   ' Adjust pointers for linked list integrity
'   IF current->pPrev <> NULL THEN
'      current->pPrev->pNext = current->pNext
'   ELSE
'      head = current->pNext ' Update head if first node is removed
'   END IF
'   IF current->pNext <> NULL THEN
'      current->pNext->pPrev = current->pPrev
'   ELSE
'      tail = current->pPrev ' Update tail if last node is removed
'   END IF
'   ' Remove from rgIdx array
'   FOR i AS INTEGER = idx TO nodes - 2
'      rgIdx(i) = rgIdx(i + 1) ' Shift elements down
'   NEXT
'   REDIM PRESERVE rgIdx(nodes - 2) ' Resize rgIdx array
'   nodes -= 1
'   DELETE current ' Free memory
'   RETURN TRUE
'END FUNCTION

'PRIVATE FUNCTION DVarList.Replace (BYVAL idx AS LONG, BYREF dv AS DVARIANT) AS BOOLEAN
'   idx -= 1
'   IF idx >= 0 AND idx < nodes THEN
'      rgIdx(idx)->dv = dv
'      RETURN TRUE
'   END IF
'   RETURN FALSE
'END FUNCTION


'PRIVATE FUNCTION DVarList.Insert (BYVAL idx AS LONG, BYREF dv AS DVARIANT) AS BOOLEAN
'   idx -= 1
'   IF idx < 0 OR idx > nodes THEN RETURN FALSE
'   DIM newNode AS DVarNode PTR = NEW DVarNode
'   newNode->dv = dv
'   IF idx = nodes THEN
'      this.Add(dv)
'      RETURN TRUE
'   END IF
'   DIM current AS DVarNode PTR = rgIdx(idx)
'   newNode->pNext = current
'   newNode->pPrev = current->pPrev
'   IF current->pPrev <> NULL THEN
'      current->pPrev->pNext = newNode
'   ELSE
'      head = newNode
'   END IF
'   current->pPrev = newNode
'   REDIM PRESERVE rgIdx(nodes)
'   FOR i AS LONG = nodes TO idx + 1 STEP -1
'      rgIdx(i) = rgIdx(i - 1)
'   NEXT
'   rgIdx(idx) = newNode
'   nodes += 1
'   RETURN TRUE
'END FUNCTION

' // Singly-linked list

TYPE DVarNode
   dv AS DVARIANT
   pNext AS DVarNode PTR
   DECLARE DESTRUCTOR
END TYPE

DESTRUCTOR DVarNode
END DESTRUCTOR

TYPE DVarList
   head AS DVarNode PTR
   tail AS DVarNode PTR
   nodes AS LONG
   DIM rgIdx(ANY) AS DVarNode PTR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION Add (BYREF dv AS DVARIANT) AS LONG
   DECLARE FUNCTION Count () AS LONG
   DECLARE FUNCTION Item (BYVAL idx AS LONG) AS DVARIANT
   DECLARE FUNCTION Remove (BYVAL idx AS LONG) AS BOOLEAN
   DECLARE FUNCTION Insert (BYVAL idx AS LONG, BYREF dv AS DVARIANT) AS BOOLEAN
   DECLARE FUNCTION Replace (BYVAL idx AS LONG, BYREF dv AS DVARIANT) AS BOOLEAN
   DECLARE SUB Clear
END TYPE

PRIVATE DESTRUCTOR DVarList
   FOR i AS LONG = LBOUND(rgIdx) TO UBOUND(rgIdx)
      DELETE rgIdx(i)
   NEXT
   ERASE rgIdx
END DESTRUCTOR

PRIVATE SUB DVarList.Clear
   FOR i AS LONG = LBOUND(rgIdx) TO UBOUND(rgIdx)
      DELETE rgIdx(i)
   NEXT
   ERASE rgIdx
   head = NULL
   tail = NULL
   nodes = 0
END SUB

PRIVATE FUNCTION DVarList.Add (BYREF dv AS DVARIANT) AS LONG
   DIM newNode AS DVarNode PTR = NEW DVarNode
   newNode->dv = dv
   newNode->pNext = NULL
   IF tail <> NULL THEN
      tail->pNext = newNode
   ELSE
      head = newNode
   END IF
   tail = newNode
   nodes += 1
   REDIM PRESERVE rgIdx(nodes - 1)
   rgIdx(nodes - 1) = newNode
   RETURN nodes
END FUNCTION

PRIVATE FUNCTION DVarList.Count () AS LONG
   RETURN nodes
END FUNCTION

PRIVATE FUNCTION DVarList.Item (BYVAL idx AS LONG) AS DVARIANT
   idx -= 1
   IF idx >= 0 AND idx < nodes THEN
      RETURN rgIdx(idx)->dv
   END IF
   RETURN DVARIANT("")
END FUNCTION

PRIVATE FUNCTION DVarList.Replace (BYVAL idx AS LONG, BYREF dv AS DVARIANT) AS BOOLEAN
   idx -= 1
   IF idx >= 0 AND idx < nodes THEN
      rgIdx(idx)->dv = dv
      RETURN TRUE
   END IF
   RETURN FALSE
END FUNCTION

PRIVATE FUNCTION DVarList.Remove (BYVAL idx AS LONG) AS BOOLEAN
   idx -= 1
   IF idx < 0 OR idx >= nodes THEN RETURN FALSE
   DIM current AS DVarNode PTR = rgIdx(idx)
   IF head = current THEN
      head = current->pNext
      IF tail = current THEN tail = NULL
   ELSE
      rgIdx(idx - 1)->pNext = current->pNext
      IF tail = current THEN tail = rgIdx(idx - 1)
   END IF
   DELETE current
   FOR i AS LONG = idx TO nodes - 2
      rgIdx(i) = rgIdx(i + 1)
   NEXT
   REDIM PRESERVE rgIdx(nodes - 2)
   nodes -= 1
   RETURN TRUE
END FUNCTION

PRIVATE FUNCTION DVarList.Insert (BYVAL idx AS LONG, BYREF dv AS DVARIANT) AS BOOLEAN
   idx -= 1
   IF idx < 0 OR idx > nodes THEN RETURN FALSE
   DIM newNode AS DVarNode PTR = NEW DVarNode
   newNode->dv = dv
   IF idx = nodes THEN RETURN (Add(dv) > 0)
   DIM prevNode AS DVarNode PTR = IIF(idx > 0, rgIdx(idx - 1), NULL)
   DIM current AS DVarNode PTR = rgIdx(idx)
   newNode->pNext = current
   IF prevNode <> NULL THEN
      prevNode->pNext = newNode
   ELSE
      head = newNode
   END IF
   REDIM PRESERVE rgIdx(nodes)
   FOR i AS LONG = nodes TO idx + 1 STEP -1
      rgIdx(i) = rgIdx(i - 1)
   NEXT
   rgIdx(idx) = newNode
   nodes += 1
   RETURN TRUE
END FUNCTION

' ########################################################################################

END NAMESPACE
