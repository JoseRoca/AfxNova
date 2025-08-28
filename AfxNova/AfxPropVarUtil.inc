' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxPropVarUtil.inc
' Purpose:  Translation of propvarutil.h header
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
#include once "win/propidl.bi"
#include once "win/propsys.bi"
#include once "AfxNova/AfxVarProcs.inc"
USING AfxNova

extern "Windows"

ENUM tagPSTIME_FLAGS
   PSTF_UTC   = &h00000000
   PSTF_LOCAL = &h00000001
End ENUM
TYPE PSTIME_FLAGS AS LONG

' // PropVariant Helpers

' // Initialize a propvariant

DECLARE FUNCTION InitPropVariantFromResource LIB "propsys.dll" ALIAS "InitPropVariantFromResource" ( _
   BYVAL hinst AS HINSTANCE, BYVAL id AS UINT, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromBuffer LIB "propsys.dll" ALIAS "InitPropVariantFromBuffer" ( _
   BYVAL pv AS void PTR, BYVAL cb AS UINT, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromCLSID LIB "propsys.dll" ALIAS "InitPropVariantFromCLSID" ( _
   BYVAL clsid AS REFCLSID, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromGUIDAsString LIB "propsys.dll" ALIAS "InitPropVariantFromGUIDAsString" ( _
   BYVAL guid AS REFGUID, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromFileTime LIB "propsys.dll" ALIAS "InitPropVariantFromFileTime" ( _
   BYVAL pftIn AS const FILETIME PTR, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromPropVariantVectorElem LIB "propsys.dll" ALIAS "InitPropVariantFromPropVariantVectorElem" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantVectorFromPropVariant LIB "propsys.dll" ALIAS "InitPropVariantVectorFromPropVariant" ( _
   BYVAL propvarSingle AS REFPROPVARIANT, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromStrRet LIB "propsys.dll" ALIAS "InitPropVariantFromStrRet" ( _
   BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromBooleanVector LIB "propsys.dll" ALIAS "InitPropVariantFromBooleanVector" ( _
   BYVAL prgf AS const BOOL PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromInt16Vector LIB "propsys.dll" ALIAS "InitPropVariantFromInt16Vector" ( _
   BYVAL prgf AS const SHORT PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromUInt16Vector LIB "propsys.dll" ALIAS "InitPropVariantFromUInt16Vector" ( _
   BYVAL prgf AS const USHORT PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromInt32Vector LIB "propsys.dll" ALIAS "InitPropVariantFromInt32Vector" ( _
   BYVAL prgf AS const LONG PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromUInt32Vector LIB "propsys.dll" ALIAS "InitPropVariantFromUInt32Vector" ( _
   BYVAL prgf AS const ULONG PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromInt64Vector LIB "propsys.dll" ALIAS "InitPropVariantFromInt64Vector" ( _
   BYVAL prgf AS const LONGLONG PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromUInt64Vector LIB "propsys.dll" ALIAS "InitPropVariantFromUInt64Vector" ( _
   BYVAL prgf AS const ULONGLONG PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromDoubleVector LIB "propsys.dll" ALIAS "InitPropVariantFromDoubleVector" ( _
   BYVAL prgf AS const DOUBLE PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromFileTimeVector LIB "propsys.dll" ALIAS "InitPropVariantFromFileTimeVector" ( _
   BYVAL prgft AS const FILETIME PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromStringVector LIB "propsys.dll" ALIAS "InitPropVariantFromStringVector" ( _
   BYVAL prgsz AS PCWSTR PTR, BYVAL cElems AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
DECLARE FUNCTION InitPropVariantFromStringAsVector LIB "propsys.dll" ALIAS "InitPropVariantFromStringAsVector" ( _
   BYVAL psz AS PCWSTR, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT

' // Extract data from a propvariant

DECLARE FUNCTION PropVariantToBooleanWithDefault LIB "propsys.dll" ALIAS "PropVariantToBooleanWithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL fDefault AS BOOL) AS BOOL
DECLARE FUNCTION PropVariantToInt16WithDefault LIB "propsys.dll" ALIAS "PropVariantToInt16WithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL iDefault AS SHORT) AS SHORT
DECLARE FUNCTION PropVariantToUInt16WithDefault LIB "propsys.dll" ALIAS "PropVariantToUInt16WithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL uiDefault AS USHORT) AS USHORT
DECLARE FUNCTION PropVariantToInt32WithDefault LIB "propsys.dll" ALIAS "PropVariantToInt32WithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL lDefault AS LONG) AS LONG
DECLARE FUNCTION PropVariantToUInt32WithDefault LIB "propsys.dll" ALIAS "PropVariantToUInt32WithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL ulDefault AS ULONG) AS ULONG
DECLARE FUNCTION PropVariantToInt64WithDefault LIB "propsys.dll" ALIAS "PropVariantToInt64WithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL llDefault AS LONGLONG) AS LONGLONG
DECLARE FUNCTION PropVariantToUInt64WithDefault LIB "propsys.dll" ALIAS "PropVariantToUInt64WithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL ullDefault AS ULONGLONG) AS ULONGLONG
DECLARE FUNCTION PropVariantToDoubleWithDefault LIB "propsys.dll" ALIAS "PropVariantToDoubleWithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL dblDefault AS DOUBLE) AS DOUBLE
DECLARE FUNCTION PropVariantToStringWithDefault LIB "propsys.dll" ALIAS "PropVariantToStringWithDefault" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pszDefault AS LPCWSTR) AS PCWSTR
DECLARE FUNCTION PropVariantToBoolean LIB "propsys.dll" ALIAS "PropVariantToBoolean" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pfRet AS BOOL PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt16 LIB "propsys.dll" ALIAS "PropVariantToInt16" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL piRet AS SHORT PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt16 LIB "propsys.dll" ALIAS "PropVariantToUInt16" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL puiRet AS USHORT PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt32 LIB "propsys.dll" ALIAS "PropVariantToInt32" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL plRet AS LONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt32 LIB "propsys.dll" ALIAS "PropVariantToUInt32" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pulRet AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt64 LIB "propsys.dll" ALIAS "PropVariantToInt64" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pllRet AS LONGLONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt64 LIB "propsys.dll" ALIAS "PropVariantToUInt64" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pullRet AS ULONGLONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToDouble LIB "propsys.dll" ALIAS "PropVariantToDouble" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pdblRet AS DOUBLE PTR) AS HRESULT
DECLARE FUNCTION PropVariantToBuffer LIB "propsys.dll" ALIAS "PropVariantToBuffer" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pv AS void PTR, BYVAL cb AS UINT) AS HRESULT
DECLARE FUNCTION PropVariantToString LIB "propsys.dll" ALIAS "PropVariantToString" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL psz AS PWSTR, BYVAL cch AS UINT) AS HRESULT
DECLARE FUNCTION PropVariantToGUID LIB "propsys.dll" ALIAS "PropVariantToGUID" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pguid AS GUID PTR) AS HRESULT
DECLARE FUNCTION PropVariantToStringAlloc LIB "propsys.dll" ALIAS "PropVariantToStringAlloc" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL ppszOut AS PWSTR PTR) AS HRESULT
DECLARE FUNCTION PropVariantToBSTR LIB "propsys.dll" ALIAS "PropVariantToBSTR" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pbstrOut AS BSTR PTR) AS HRESULT
DECLARE FUNCTION PropVariantToStrRet LIB "propsys.dll" ALIAS "PropVariantToStrRet" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pstrret AS STRRET PTR) AS HRESULT
DECLARE FUNCTION PropVariantToFileTime LIB "propsys.dll" ALIAS "PropVariantToFileTime" ( _
   BYVAL propvarIn AS REFPROPVARIANT, BYVAL pstfOut AS PSTIME_FLAGS, BYVAL pftOut AS FILETIME PTR) AS HRESULT

' // Returns element count of a VT_VECTOR or VT_ARRAY value; or 1 otherwise

DECLARE FUNCTION PropVariantGetElementCount LIB "propsys.dll" ALIAS "PropVariantGetElementCount" (BYVAL propvar AS REFPROPVARIANT) AS ULONG

' // Extract data from a propvariant into a vector

DECLARE FUNCTION PropVariantToBooleanVector LIB "propsys.dll" ALIAS "PropVariantToBooleanVector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgf AS BOOL PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt16Vector LIB "propsys.dll" ALIAS "PropVariantToInt16Vector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgn AS SHORT PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt16Vector LIB "propsys.dll" ALIAS "PropVariantToUInt16Vector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgn AS USHORT PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt32Vector LIB "propsys.dll" ALIAS "PropVariantToInt32Vector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgn AS LONG PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt32Vector LIB "propsys.dll" ALIAS "PropVariantToUInt32Vector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgn AS ULONG PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt64Vector LIB "propsys.dll" ALIAS "PropVariantToInt64Vector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgn AS LONGLONG PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt64Vector LIB "propsys.dll" ALIAS "PropVariantToUInt64Vector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgn AS ULONGLONG PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToDoubleVector LIB "propsys.dll" ALIAS "PropVariantToDoubleVector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgn AS DOUBLE PTR, BYVAL crgf AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToFileTimeVector LIB "propsys.dll" ALIAS "PropVariantToFileTimeVector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgft AS FILETIME PTR, BYVAL crgft AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToStringVector LIB "propsys.dll" ALIAS "PropVariantToStringVector" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL prgsz AS PWSTR PTR, BYVAL crgsz AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT

' // Extract data from a propvariant and return an newly allocated vector (free with CoTaskMemFree)

DECLARE FUNCTION PropVariantToBooleanVectorAlloc LIB "propsys.dll" ALIAS "PropVariantToBooleanVectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgf AS BOOL PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt16VectorAlloc LIB "propsys.dll" ALIAS "PropVariantToInt16VectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS SHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt16VectorAlloc LIB "propsys.dll" ALIAS "PropVariantToUInt16VectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS USHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt32VectorAlloc LIB "propsys.dll" ALIAS "PropVariantToInt32VectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS LONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt32VectorAlloc LIB "propsys.dll" ALIAS "PropVariantToUInt32VectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS ULONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToInt64VectorAlloc LIB "propsys.dll" ALIAS "PropVariantToInt64VectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS LONGLONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToUInt64VectorAlloc LIB "propsys.dll" ALIAS "PropVariantToUInt64VectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS ULONGLONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToDoubleVectorAlloc LIB "propsys.dll" ALIAS "PropVariantToDoubleVectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS DOUBLE PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToFileTimeVectorAlloc LIB "propsys.dll" ALIAS "PropVariantToFileTimeVectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgn AS FILETIME PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantToStringVectorAlloc LIB "propsys.dll" ALIAS "PropVariantToFileTimeVectorAlloc" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL pprgsz AS PWSTR PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT

' // Extract a single element from a propvariant.  If it is a VT_VECTOR or VT_ARRAY, returns the element you request.
' // Otherwise iElem must equal 0 and the function will returns the value.

DECLARE FUNCTION PropVariantGetBooleanElem LIB "propsys.dll" ALIAS "PropVariantGetBooleanElem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pfVal AS BOOL PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetInt16Elem LIB "propsys.dll" ALIAS "PropVariantGetInt16Elem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS SHORT PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetUInt16Elem LIB "propsys.dll" ALIAS "PropVariantGetUInt16Elem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS USHORT PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetInt32Elem LIB "propsys.dll" ALIAS "PropVariantGetInt32Elem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS LONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetUInt32Elem LIB "propsys.dll" ALIAS "PropVariantGetUInt32Elem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetInt64Elem LIB "propsys.dll" ALIAS "PropVariantGetInt64Elem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS LONGLONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetUInt64Elem LIB "propsys.dll" ALIAS "PropVariantGetUInt64Elem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONGLONG PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetDoubleElem LIB "propsys.dll" ALIAS "PropVariantGetDoubleElem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS DOUBLE PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetFileTimeElem LIB "propsys.dll" ALIAS "PropVariantGetFileTimeElem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL pftVal AS FILETIME PTR) AS HRESULT
DECLARE FUNCTION PropVariantGetStringElem LIB "propsys.dll" ALIAS "PropVariantGetStringElem" ( _
   BYVAL propvar AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL ppszVal AS PWSTR PTR) AS HRESULT

' // Helpers

DECLARE FUNCTION ClearPropVariantArray LIB "propsys.dll" ALIAS "ClearPropVariantArray" ( _
   BYVAL rgPropVar AS PROPVARIANT PTR, BYVAL cVars AS UINT) AS HRESULT

TYPE PROPVAR_COMPARE_UNIT AS LONG
ENUM
   PVCU_DEFAULT = 0
   PVCU_SECOND  = 1
   PVCU_MINUTE  = 2
   PVCU_HOUR    = 3
   PVCU_DAY     = 4
   PVCU_MONTH   = 5
   PVCU_YEAR    = 6
END ENUM

TYPE tagPROPVAR_COMPARE_FLAGS AS LONG
ENUM
	PVCF_DEFAULT                       = &h00000000   ' // When comparing strings, use StrCmpLogical
	PVCF_TREATEMPTYASGREATERTHAN       = &h00000001   ' // Empty/null values are greater-than non-empty values
	PVCF_USESTRCMP                     = &h00000002   ' // When comparing strings, use StrCmp
	PVCF_USESTRCMPC                    = &h00000004   ' // When comparing strings, use StrCmpC
	PVCF_USESTRCMPI                    = &h00000008   ' // When comparing strings, use StrCmpI
	PVCF_USESTRCMPIC                   = &h00000010   ' // When comparing strings, use StrCmpIC
   PVCF_DIGITSASNUMBERS_CASESENSITIVE = &h00000020   ' // When comparing strings, use CompareStringEx with LOCALE_NAME_USER_DEFAULT and SORT_DIGITSASNUMBERS.  This corresponds to the linguistically correct order for UI lists.
END ENUM
TYPE PROPVAR_COMPARE_FLAGS AS LONG

ENUM tagPROPVAR_CHANGE_FLAGS
   PVCHF_DEFAULT           = &h00000000
   PVCHF_NOVALUEPROP       = &h00000001   ' // Maps to VARIANT_NOVALUEPROP for VariantChangeType
   PVCHF_ALPHABOOL         = &h00000002   ' // Maps to VARIANT_ALPHABOOL for VariantChangeType
   PVCHF_NOUSEROVERRIDE    = &h00000004   ' // Maps to VARIANT_NOUSEROVERRIDE for VariantChangeType
   PVCHF_LOCALBOOL         = &h00000008   ' // Maps to VARIANT_LOCALBOOL for VariantChangeType
   PVCHF_NOHEXSTRING       = &h00000010   ' // Don't convert a string that looks like hexadecimal (0xABCD) to the numerical equivalent
END ENUM
TYPE PROPVAR_CHANGE_FLAGS AS LONG

' // Comparisons

DECLARE FUNCTION PropVariantCompareEx LIB "propsys.dll" ALIAS "PropVariantCompareEx" ( _
   BYVAL propvar1 AS REFPROPVARIANT, BYVAL propvar2 AS REFPROPVARIANT, BYVAL unit AS PROPVAR_COMPARE_UNIT, _
   BYVAL flags AS PROPVAR_COMPARE_FLAGS) AS INT_


' // Coersions

DECLARE FUNCTION PropVariantChangeType LIB "propsys.dll" ALIAS "PropVariantChangeType" ( _
   BYVAL ppropvarDest AS PROPVARIANT PTR, BYVAL propvarSrc AS REFPROPVARIANT, _
   BYVAL flags AS PROPVAR_CHANGE_FLAGS, BYVAL vt AS VARTYPE) AS HRESULT

' // Conversions

DECLARE FUNCTION PropVariantToVariant LIB "propsys.dll" ALIAS "PropVariantToVariant" ( _
   BYVAL propvar AS PROPVARIANT PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION VariantToPropVariant LIB "propsys.dll" ALIAS "VariantToPropVariant" ( _
   BYVAL pVar AS VARIANT PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT

' // Stg functions

DECLARE FUNCTION StgSerializePropVariant LIB "propsys.dll" ALIAS "StgSerializePropVariant" ( _
   BYVAL ppropvar AS PROPVARIANT PTR, BYVAL ppProp AS SERIALIZEDPROPERTYVALUE PTR PTR, BYVAL pcb AS ULONG PTR) AS HRESULT
DECLARE FUNCTION StgDeserializePropVariant LIB "propsys.dll" ALIAS "StgDeserializePropVariant" ( _
   BYVAL pprop AS const SERIALIZEDPROPERTYVALUE PTR, BYVAL cbMax AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT

' // Initialize a VARIANT

DECLARE FUNCTION InitVariantFromResource LIB "propsys.dll" ALIAS "InitVariantFromResource" ( _
   BYVAL hinst AS HINSTANCE, BYVAL id AS UINT, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromBuffer LIB "propsys.dll" ALIAS "InitVariantFromBuffer" ( _
   BYVAL pv AS const void PTR, BYVAL cb AS UINT, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromGUIDAsString LIB "propsys.dll" ALIAS "InitVariantFromGUIDAsString" ( _
   BYVAL guid AS REFGUID, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromFileTime LIB "propsys.dll" ALIAS "InitVariantFromFileTime" ( _
   BYVAL pft AS const FILETIME PTR, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromFileTimeArray LIB "propsys.dll" ALIAS "InitVariantFromFileTimeArray" ( _
   BYVAL prgft AS const FILETIME PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromStrRet LIB "propsys.dll" ALIAS "InitVariantFromStrRet" ( _
   BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromVariantArrayElem LIB "propsys.dll" ALIAS "InitVariantFromVariantArrayElem" ( _
   BYVAL varIn AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromBooleanArray LIB "propsys.dll" ALIAS "InitVariantFromBooleanArray" ( _
   BYVAL prgf AS const BOOL PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromInt16Array LIB "propsys.dll" ALIAS "InitVariantFromInt16Array" ( _
   BYVAL prgn AS const SHORT PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromUInt16Array LIB "propsys.dll" ALIAS "InitVariantFromUInt16Array" ( _
   BYVAL prgn AS const USHORT PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromInt32Array LIB "propsys.dll" ALIAS "InitVariantFromInt32Array" ( _
   BYVAL prgn AS const LONG PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromUInt32Array LIB "propsys.dll" ALIAS "InitVariantFromUInt32Array" ( _
   BYVAL prgn AS const ULONG PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromInt64Array LIB "propsys.dll" ALIAS "InitVariantFromInt64Array" ( _
   BYVAL prgn AS const LONGLONG PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromUInt64Array LIB "propsys.dll" ALIAS "InitVariantFromUInt64Array" ( _
   BYVAL prgn AS const ULONGLONG PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromDoubleArray LIB "propsys.dll" ALIAS "InitVariantFromDoubleArray" ( _
   BYVAL prgn AS const DOUBLE PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
DECLARE FUNCTION InitVariantFromStringArray LIB "propsys.dll" ALIAS "InitVariantFromStringArray" ( _
   BYVAL prgsz AS PCWSTR PTR, BYVAL cElems AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT

' // Extract data from a VARIANT

DECLARE FUNCTION VariantToBooleanWithDefault LIB "propsys.dll" ALIAS "VariantToBooleanWithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL fDefault AS BOOL) AS BOOL
DECLARE FUNCTION VariantToInt16WithDefault LIB "propsys.dll" ALIAS "VariantToInt16WithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL iDefault AS SHORT) AS SHORT
DECLARE FUNCTION VariantToUInt16WithDefault LIB "propsys.dll" ALIAS "VariantToUInt16WithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL uiDefault AS USHORT) AS USHORT
DECLARE FUNCTION VariantToInt32WithDefault LIB "propsys.dll" ALIAS "VariantToInt32WithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL lDefault AS LONG) AS LONG
DECLARE FUNCTION VariantToUInt32WithDefault LIB "propsys.dll" ALIAS "VariantToUInt32WithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL ulDefault AS ULONG) AS ULONG
DECLARE FUNCTION VariantToInt64WithDefault LIB "propsys.dll" ALIAS "VariantToInt64WithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL llDefault AS LONGLONG) AS LONGLONG
DECLARE FUNCTION VariantToUInt64WithDefault LIB "propsys.dll" ALIAS "VariantToUInt64WithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL ullDefault AS ULONGLONG) AS ULONGLONG
DECLARE FUNCTION VariantToDoubleWithDefault LIB "propsys.dll" ALIAS "VariantToDoubleWithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL dblDefault AS DOUBLE) AS DOUBLE
DECLARE FUNCTION VariantToStringWithDefault LIB "propsys.dll" ALIAS "VariantToStringWithDefault" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pszDefault AS LPCWSTR) AS PCWSTR

DECLARE FUNCTION VariantToBoolean LIB "propsys.dll" ALIAS "VariantToBoolean" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pfRet AS BOOL PTR) AS HRESULT
DECLARE FUNCTION VariantToInt16 LIB "propsys.dll" ALIAS "VariantToInt16" ( _
   BYVAL varIn AS REFVARIANT, BYVAL piRet AS SHORT PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt16 LIB "propsys.dll" ALIAS "VariantToUInt16" ( _
   BYVAL varIn AS REFVARIANT, BYVAL puiRet AS USHORT PTR) AS HRESULT
DECLARE FUNCTION VariantToInt32 LIB "propsys.dll" ALIAS "VariantToInt32" ( _
   BYVAL varIn AS REFVARIANT, BYVAL plRet AS LONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt32 LIB "propsys.dll" ALIAS "VariantToUInt32" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pulRet AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToInt64 LIB "propsys.dll" ALIAS "VariantToInt64" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pllRet AS LONGLONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt64 LIB "propsys.dll" ALIAS "VariantToUInt64" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pullRet AS ULONGLONG PTR) AS HRESULT
DECLARE FUNCTION VariantToDouble LIB "propsys.dll" ALIAS "VariantToDouble" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pdblRet AS DOUBLE PTR) AS HRESULT
DECLARE FUNCTION VariantToBuffer LIB "propsys.dll" ALIAS "VariantToBuffer" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pv AS void PTR, BYVAL cb AS UINT) AS HRESULT
DECLARE FUNCTION VariantToGUID LIB "propsys.dll" ALIAS "VariantToGUID" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pguid AS GUID PTR) AS HRESULT
DECLARE FUNCTION VariantToString LIB "propsys.dll" ALIAS "VariantToString" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pszBuf AS PWSTR, BYVAL cchBuf AS UINT) AS HRESULT
DECLARE FUNCTION VariantToStringAlloc LIB "propsys.dll" ALIAS "VariantToStringAlloc" ( _
   BYVAL varIn AS REFVARIANT, BYVAL ppszBuf AS PWSTR PTR) AS HRESULT
DECLARE FUNCTION VariantToDosDateTime LIB "propsys.dll" ALIAS "VariantToDosDateTime" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pwDate AS WORD PTR, BYVAL pwTime AS WORD PTR) AS HRESULT
DECLARE FUNCTION VariantToStrRet LIB "propsys.dll" ALIAS "VariantToStrRet" ( _
   BYVAL varIn AS REFVARIANT, BYVAL pstrret AS STRRET PTR) AS HRESULT
DECLARE FUNCTION VariantToFileTime LIB "propsys.dll" ALIAS "VariantToFileTime" ( _
   BYVAL varIn AS REFVARIANT, BYVAL stfOut AS PSTIME_FLAGS, BYVAL pftOut AS FILETIME PTR) AS HRESULT

' // Get the element count.  Returns number of elements for values of type VT_ARRAY; returns 1 otherwise.

DECLARE FUNCTION VariantGetElementCount LIB "propsys.dll" ALIAS "VariantGetElementCount" (BYVAL varIn AS REFVARIANT) AS ULONG

' // Extract data from a VARIANT into a vector

DECLARE FUNCTION VariantToBooleanArray LIB "propsys.dll" ALIAS "VariantToBooleanArray" ( _
   BYVAL var AS REFVARIANT, BYVAL prgf AS BOOL PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToInt16Array LIB "propsys.dll" ALIAS "VariantToInt16Array" ( _
   BYVAL var AS REFVARIANT, BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt16Array LIB "propsys.dll" ALIAS "VariantToUInt16Array" ( _
   BYVAL var AS REFVARIANT, BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToInt32Array LIB "propsys.dll" ALIAS "VariantToInt32Array" ( _
   BYVAL var AS REFVARIANT, BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt32Array LIB "propsys.dll" ALIAS "VariantToUInt32Array" ( _
   BYVAL var AS REFVARIANT, BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToInt64Array LIB "propsys.dll" ALIAS "VariantToInt64Array" ( _
   BYVAL var AS REFVARIANT, BYVAL prgn AS LONGLONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt64Array LIB "propsys.dll" ALIAS "VariantToUInt64Array" ( _
   BYVAL var AS REFVARIANT, BYVAL prgn AS ULONGLONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToDoubleArray LIB "propsys.dll" ALIAS "VariantToDoubleArray" ( _
   BYVAL var AS REFVARIANT, BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToStringArray LIB "propsys.dll" ALIAS "VariantToStringArray" ( _
   BYVAL var AS REFVARIANT, BYVAL prgsz AS PWSTR PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT

' // Extract data from a VARIANT into a newly allocated vector (free with CoTaskMemFree)

DECLARE FUNCTION VariantToBooleanArrayAlloc LIB "propsys.dll" ALIAS "VariantToBooleanArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgf AS BOOL PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToInt16ArrayAlloc LIB "propsys.dll" ALIAS "VariantToInt16ArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgn AS SHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt16ArrayAlloc LIB "propsys.dll" ALIAS "VariantToUInt16ArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgn AS USHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToInt32ArrayAlloc LIB "propsys.dll" ALIAS "VariantToInt32ArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgn AS LONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt32ArrayAlloc LIB "propsys.dll" ALIAS "VariantToUInt32ArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgn AS ULONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToInt64ArrayAlloc LIB "propsys.dll" ALIAS "VariantToInt64ArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgn AS LONGLONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToUInt64ArrayAlloc LIB "propsys.dll" ALIAS "VariantToUInt64ArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgn AS ULONGLONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToDoubleArrayAlloc LIB "propsys.dll" ALIAS "VariantToDoubleArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgn AS DOUBLE PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantToStringArrayAlloc LIB "propsys.dll" ALIAS "VariantToStringArrayAlloc" ( _
   BYVAL var AS REFVARIANT, BYVAL pprgsz AS PWSTR PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT

' // Get a single element of a VARIANT.  If it is type VT_ARRAY, returns a the requested element.  Otherwise
' // iElem must equal 0 and the function returns the value.

DECLARE FUNCTION VariantGetBooleanElem LIB "propsys.dll" ALIAS "VariantGetBooleanElem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pfVal AS BOOL PTR) AS HRESULT
DECLARE FUNCTION VariantGetInt16Elem LIB "propsys.dll" ALIAS "VariantGetInt16Elem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS SHORT PTR) AS HRESULT
DECLARE FUNCTION VariantGetUInt16Elem LIB "propsys.dll" ALIAS "VariantGetUInt16Elem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS USHORT PTR) AS HRESULT
DECLARE FUNCTION VariantGetInt32Elem LIB "propsys.dll" ALIAS "VariantGetInt32Elem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS LONG PTR) AS HRESULT
DECLARE FUNCTION VariantGetUInt32Elem LIB "propsys.dll" ALIAS "VariantGetUInt32Elem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONG PTR) AS HRESULT
DECLARE FUNCTION VariantGetInt64Elem LIB "propsys.dll" ALIAS "VariantGetInt64Elem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS LONGLONG PTR) AS HRESULT
DECLARE FUNCTION VariantGetUInt64Elem LIB "propsys.dll" ALIAS "VariantGetUInt64Elem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONGLONG PTR) AS HRESULT
DECLARE FUNCTION VariantGetDoubleElem LIB "propsys.dll" ALIAS "VariantGetDoubleElem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pnVal AS DOUBLE PTR) AS HRESULT
DECLARE FUNCTION VariantGetStringElem LIB "propsys.dll" ALIAS "VariantGetStringElem" ( _
   BYVAL var AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL ppszVal AS PWSTR PTR) AS HRESULT

' // Helpers

DECLARE SUB ClearVariantArray LIB "propsys.dll" ALIAS "ClearVariantArray" ( _
   BYVAL pvars AS VARIANT PTR, BYVAL cvars AS UINT)
DECLARE FUNCTION VariantCompare LIB "propsys.dll" ALIAS "VariantCompare" ( _
   BYVAL var1 AS REFVARIANT, BYVAL var2 AS REFVARIANT) AS INT_

' // Property-specific notions

' // The progress bar property control uses a specially formatted PROPVARIANT to convey the look of the progress bar
' // propvar.vt = VT_UI4
' // propvar.caul.pElems[0] = current progress
' // propvar.caul.pElems[1] = total progress
' // propvar.caul.pElems[2] = DRAWPROGRESSFLAGS (see below);
ENUM DRAWPROGRESSFLAGS
   DPF_NONE                = &h0    ' // No progress flags.
   DPF_MARQUEE             = &h1    ' // The progress bar should draw in marquee mode.
   DPF_MARQUEE_COMPLETE    = &h2    ' // The marquee format progress bar has completed.
   DPF_ERROR               = &h4    ' // The progress bar should be drawn in the error state.
   DPF_WARNING             = &h8    ' // The progress bar should be drawn in the warning state.
   DPF_STOPPED             = &h10   ' // The progress bar is stopped.
END ENUM

' // Inline functions

PRIVATE FUNCTION InitPropVariantFromBoolean (BYVAL fVal AS BOOL, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_BOOL
   V_UNION(ppropvar, boolVal) = IIF(fVal <> 0, VARIANT_TRUE, VARIANT_FALSE)
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitPropVariantFromInt16 (BYVAL nVal AS SHORT, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_I2
   V_UNION(ppropvar, iVal) = nVal
   RETURN S_OK
END FUNcTION

PRIVATE FUNCTION InitPropVariantFromUInt16 (BYVAL uiVal AS USHORT, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_UI2
   V_UNION(ppropvar, uiVal) = uiVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitPropVariantFromInt32 (BYVAL lVal AS LONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_I4
   V_UNION(ppropvar, lVal) = lVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitPropVariantFromUInt32 (BYVAL ulVal AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_UI4
   V_UNION(ppropvar, ulVal) = ulVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitPropVariantFromInt64 (BYVAL llVal AS LONGLONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_I8
   V_UNION(ppropvar, hVal).QuadPart = llVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitPropVariantFromUInt64 (BYVAL ullVal AS ULONGLONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_UI8
   V_UNION(ppropvar, uhVal).QuadPart = ullVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitPropVariantFromDouble (BYVAL dblVal AS DOUBLE, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   V_VT(ppropvar) = VT_R8
   V_UNION(ppropvar, dblVal) = dblVal
   RETURN S_OK
END FUNCTION

'// Creates a VT_LPWSTR propvariant.
PRIVATE FUNCTION InitPropVariantFromString (BYVAL psz AS PCWSTR, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   DIM hr AS HRESULT = IIF(psz <> NULL, S_OK, E_INVALIDARG)   ' // Previous API behavior counter to the SAL requirement.
   IF SUCCEEDED(hr) THEN
      DIM byteCount AS const SIZE_T = CAST(SIZE_T, (wcslen(psz) + 1) * sizeof(*psz))
      V_UNION(ppropvar, pwszVal) = CAST(PWSTR, CoTaskMemAlloc(byteCount))
      hr = V_UNION(ppropvar, pwszVal) = IIF(psz <> NULL, S_OK, E_OUTOFMEMORY)
      IF SUCCEEDED(hr) THEN
'         memcpy_s(V_UNION(ppropvar, pwszVal), byteCount, psz, byteCount)
         memcpy(V_UNION(ppropvar, pwszVal), psz, byteCount)
         V_VT(ppropvar) = VT_LPWSTR
      END IF
   END IF
   IF FAILED(hr) THEN
      PropVariantInit(ppropvar)
   END IF
   RETURN hr
END FUNCTION

'FUNCTION InitPropVariantFromString (BYVAL psz AS PCWSTR, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
'   V_VT(ppropvar) = VT_LPWSTR
'   DIM hr AS HRESULT = SHStrDupW(psz, @ppropvar->pwszVal)
'   IF FAILED(hr) THEN PropVariantInit(ppropvar)
'   RETURN hr
'END FUNCTION

' // Creates a VT_VECTOR | VT_UI1 propvariant.

PRIVATE FUNCTION InitPropVariantFromGUIDAsBuffer (BYVAL guid AS REFGUID, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   RETURN InitPropVariantFromBuffer(CAST(void PTR, @guid), sizeof(..GUID), ppropvar)
END FUNCTION

PRIVATE FUNCTION IsPropVariantVector (BYVAL propvar AS REFPROPVARIANT) AS BOOLEAN
   RETURN (propvar->vt AND (VT_ARRAY OR VT_VECTOR))
END FUNCTION

' // If TRUE, propvar contains a unicode string.  Use PropVariantToStringWithDefault(propvar, NULL) to retrieve it.
PRIVATE FUNCTION IsPropVariantString (BYVAL propvar AS REFPROPVARIANT) AS BOOLEAN
   RETURN (PropVariantToStringWithDefault(propvar, NULL) <> NULL)
END FUNCTION

' // Handles int instead of LONG
PRIVATE FUNCTION PropVariantToInt32 (BYVAL propvarIn AS REFPROPVARIANT, BYVAL piRet AS INT_ PTR) AS HRESULT
   RETURN PropVariantToInt32(propvarIn, piRet)
END FUNCTION

' // Handles UINT instead of ULONG
PRIVATE FUNCTION PropVariantToUInt32 (BYVAL propvarIn AS REFPROPVARIANT, BYVAL puiRet AS UINT PTR) AS HRESULT
   RETURN PropVariantToUInt32(propvarIn, puiRet)
END FUNCTION

PRIVATE FUNCTION PropVariantToCLSID (BYVAL propvarIn AS REFPROPVARIANT, BYVAL pclsid AS CLSID PTR) AS HRESULT
   RETURN PropVariantToGUID(propvarIn, pclsid)
END FUNCTION

PRIVATE FUNCTION PropVariantCompare (BYVAL propvar1 AS REFPROPVARIANT, BYVAL propvar2 AS REFPROPVARIANT) AS INT_
   RETURN PropVariantCompareEx(propvar1, propvar2, PVCU_DEFAULT, PVCU_DEFAULT)
END FUNCTION

PRIVATE FUNCTION PropVariantGetElem (BYVAL propvarIn AS REFPROPVARIANT, BYVAL iElem AS ULONG, BYVAL ppropvar AS PROPVARIANT PTR) AS HRESULT
   RETURN InitPropVariantFromPropVariantVectorElem(propvarIn, iElem, ppropvar)
END FUNCTION

PRIVATE FUNCTION InitVariantFromBoolean (BYVAL fVal AS BOOL, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_BOOL
   V_UNION(pvar, boolVal) = IIF(fVal, VARIANT_TRUE, VARIANT_FALSE)
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromInt16 (BYVAL iVal AS SHORT, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_I2
   V_UNION(pvar, iVal) = iVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromUInt16 (BYVAL uiVal AS USHORT, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_UI2
   V_UNION(pvar, uiVal) = uiVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromInt32 (BYVAL lVal AS LONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_I4
   V_UNION(pvar, lVal) = lVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromUInt32 (BYVAL ulVal AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_UI4
   V_UNION(pvar, ulVal) = ulVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromInt64 (BYVAL llVal AS LONGLONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_I8
   V_UNION(pvar, llVal) = llVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromUInt64 (BYVAL ullVal AS ULONGLONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_UI8
   V_UNION(pvar, ullVal) = ullVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromDouble (BYVAL dblVal AS DOUBLE, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_R8
   V_UNION(pvar, dblVal) = dblVal
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION InitVariantFromString (BYVAL psz AS PCWSTR, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_BSTR
   V_UNION(pvar, bstrVal) = SysAllocString(psz)
   DIM hr AS HRESULT = IIF(pvar->bstrVal <> 0, S_OK, E_OUTOFMEMORY)
   IF FAILED(hr) THEN VariantInit(pvar)
   RETURN hr
END FUNCTION

PRIVATE FUNCTION InitVariantFromDispatch (BYVAL pdisp AS IDispatch PTR, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_DISPATCH
   V_UNION(pvar, pdispVal) = pdisp
   IF (V_UNION(pvar, pdispVal)) THEN
      IF pvar->pdispVal THEN pvar->pdispVal->lpvtbl->AddRef(pvar->pdispVal)
   END IF
   RETURN S_OK
END FUNCTION

' // Creates a VT_DATE variant
PRIVATE FUNCTION InitVariantFromDosDateTime (BYVAL wDate AS WORD, BYVAL wTime AS WORD, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_DATE
   RETURN IIF(DosDateTimeToVariantTime(wDate, wTime, @V_UNION(pvar, date)), S_OK, S_FALSE)
END FUNCTION

PRIVATE FUNCTION InitVariantFromUnknown (BYVAL pUnknown AS IUnknown PTR, BYVAL pvar AS VARIANT PTR) AS HRESULT
   V_VT(pvar) = VT_UNKNOWN
   V_UNION(pvar, punkVal) = pUnknown
   IF (V_UNION(pvar, punkVal)) THEN
      IF pvar->punkVal THEN pvar->punkVal->lpvtbl->AddRef(pvar->punkVal)
   END IF
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION IsVarTypeFloat (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN (vt = VT_R4 OR vt = VT_R8)
END FUNCTION

PRIVATE FUNCTION IsVariantArray (BYVAL pvar AS REFVARIANT) AS BOOLEAN
   RETURN (pvar->vt AND VT_ARRAY)
END FUNCTION

' // if TRUE, you can use VariantToStringCast to obtain the string pointer
PRIVATE FUNCTION IsVariantString (BYVAL pvar AS REFVARIANT) AS BOOLEAN
   RETURN VariantToStringWithDefault(pvar, NULL) <> NULL
END FUNCTION

PRIVATE FUNCTION IsVarTypeSignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   DIM fRet AS BOOLEAN = FALSE
   SELECT CASE vt
      CASE VT_I1, VT_I2, VT_I4, VT_I8
         fRet = TRUE
   END SELECT
   FUNCTION = fRet
END FUNCTION

PRIVATE FUNCTION IsVarTypeUnsignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   DIM fRet AS BOOLEAN = FALSE
   SELECT CASE vt
      CASE VT_UI1, VT_UI2, VT_UI4, VT_UI8
         fRet = TRUE
   END SELECT
   FUNCTION = fRet
END FUNCTION

PRIVATE FUNCTION IsVarTypeInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN IsVarTypeSignedInteger(vt) OR IsVarTypeUnsignedInteger(vt)
END FUNCTION

PRIVATE FUNCTION IsVarTypeNumber (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN IsVarTypeInteger(vt) OR IsVarTypeFloat(vt)
END FUNCTION

PRIVATE FUNCTION InitVariantFromGUIDAsBuffer (BYVAL guid AS REFGUID, BYVAL pvar AS VARIANT PTR) AS HRESULT
   RETURN InitVariantFromBuffer(CAST(void PTR, @guid), sizeof(..GUID), pvar)
END FUNCTION

PRIVATE FUNCTION VariantGetElem (BYVAL varIn AS REFVARIANT, BYVAL iElem AS ULONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
   RETURN InitVariantFromVariantArrayElem(varIn, iElem, pvar)
END FUNCTION


' // From propapi.bi
DECLARE FUNCTION StgPropertyLengthAsVariant LIB "ole32.dll" ALIAS "StgPropertyLengthAsVariant" ( _
   BYVAL pProp AS const SERIALIZEDPROPERTYVALUE PTR, BYVAL cbProp AS ULONG, BYVAL CodePage AS USHORT, BYVAL bReserved AS BYTE) AS ULONG

' // Missing in propidl.bi
'DECLAaRE FUNCTION StgConvertPropertyToVariant LIB "ole32.dll" ALIAS "StgConvertPropertyToVariant" ( _
'   BYVAL pprop AS SERIALIZEDPROPERTYVALUE PTR, BYVAL CodePage AS USHORT, BYVAL pvar AS PROPVARIANT PTR, BYVAL pma AS PMemoryAllocator) AS BOOLEAN

end extern
