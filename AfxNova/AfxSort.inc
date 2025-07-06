' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxSort.inc
' Purpose:  Sorting procedures
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
#include once "windows.bi"
#include once "AfxNova/AfxArrays.inc"
USING AfxNova

NAMESPACE AfxNowa

' ########################################################################################
'                                     DWSTRING
' ########################################################################################
#if _WIN32_WINNT >= &h0601

FUNCTION Afx_ArraySortDWStringCompare CDECL (BYVAL a AS DWSTRING PTR, BYVAL b AS DWSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_pBuffer), a->m_BufferLen, cast(WSTRING PTR, b->m_pBuffer), b->m_BufferLen, NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   RETURN r
END FUNCTION

FUNCTION Afx_ArraySortDWStringReverseCompare CDECL (BYVAL a AS DWSTRING PTR, BYVAL b AS DWSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_pBuffer), a->m_BufferLen, cast(WSTRING PTR, b->m_pBuffer), b->m_BufferLen, NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

SUB Afx_ArraySortDWString OVERLOAD (BYVAL rgdwstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgdwstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgdwstr, numElm, SIZEOF(DWSTRING), CPTR(ANY PTR, @Afx_ArraySortDWStringCompare)
   ELSE
      qsort rgdwstr, numElm, SIZEOF(DWSTRING) , CPTR(ANY PTR, @Afx_ArraySortDWStringReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortDWString OVERLOAD (rgwstr() AS DWSTRING, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgwstr) - LBOUND(rgwstr) + 1
   Afx_ArraySortDWString @rgwstr(LBOUND(rgwstr)), numElm, bAscend
END SUB

' ########################################################################################
'                                      BSTRING
' ########################################################################################

FUNCTION Afx_ArraySortBStringCompare CDECL (BYVAL a AS BSTRING PTR, BYVAL b AS BSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_bstr), SysStringLen(a->m_bstr), cast(WSTRING PTR, b->m_bstr), SysStringLen(b->m_bstr), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   RETURN r
END FUNCTION

FUNCTION Afx_ArraySortBStringReverseCompare CDECL (BYVAL a AS BSTRING PTR, BYVAL b AS BSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, cast(WSTRING PTR, a->m_bstr), SysStringLen(a->m_bstr), cast(WSTRING PTR, b->m_bstr), SysStringLen(b->m_bstr), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

SUB Afx_ArraySortBString OVERLOAD (BYVAL rgbstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgbstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgbstr, numElm, SIZEOF(BSTRING), CPTR(ANY PTR, @Afx_ArraySortBStringCompare)
   ELSE
      qsort rgbstr, numElm, SIZEOF(BSTRING) , CPTR(ANY PTR, @Afx_ArraySortBStringReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortBString OVERLOAD (rgbstr() AS BSTRING, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgbstr) - LBOUND(rgbstr) + 1
   Afx_ArraySortBString @rgbstr(LBOUND(rgbstr)), numElm, bAscend
END SUB

' ########################################################################################
'                                      STRING
' ########################################################################################

FUNCTION Afx_AraySortStringCompare CDECL (BYVAL a AS STRING PTR, BYVAL b AS STRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, STRPTR(*a), LEN(*a), STRPTR(*b), LEN(*b))
   IF r <> 0 THEN r = r - 2
   RETURN r
END FUNCTION

FUNCTION Afx_ArraySortStringReverseCompare CDECL (BYVAL a AS STRING PTR, BYVAL b AS STRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, STRPTR(*a), LEN(*a), STRPTR(*b), LEN(*b))
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

SUB Afx_ArraySortString OVERLOAD (BYVAL rgstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgstr, numElm, SIZEOF(STRING), CPTR(ANY PTR, @Afx_AraySortStringCompare)
   ELSE
      qsort rgstr, numElm, SIZEOF(STRING) , CPTR(ANY PTR, @Afx_ArraySortStringReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortString OVERLOAD (rgstr() AS STRING, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgstr) - LBOUND(rgstr) + 1
   Afx_ArraySortString @rgstr(LBOUND(rgstr)), numElm, bAscend
END SUB

' ########################################################################################
'                                      ZSTRING
' ########################################################################################

FUNCTION Afx_ArraySortZStringCompare CDECL (BYVAL a AS ZSTRING PTR, BYVAL b AS ZSTRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b))
   IF r <> 0 THEN r = r - 2
   RETURN r
END FUNCTION

FUNCTION Afx_ArraySortZStringReverseCompare CDECL (BYVAL a AS ZSTRING PTR, BYVAL b AS ZSTRING PTR) AS LONG
   STATIC Flags AS DWORD = LINGUISTIC_IGNORECASE
   DIM r AS LONG = CompareStringA(LOCALE_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b))
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

SUB Afx_ArraySortZString OVERLOAD (BYVAL rgzstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL sizeElem AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF bAscend THEN
      qsort rgzstr, numElm, sizeElem, CPTR(ANY PTR, @Afx_ArraySortZStringCompare)
   ELSE
      qsort rgzstr, numElm, sizeElem, CPTR(ANY PTR, @Afx_ArraySortZStringReverseCompare)
   END IF
END SUB

' ########################################################################################
'                                      WSTRING
' ########################################################################################

FUNCTION Afx_ArraySortWStringCompare CDECL (BYVAL a AS WSTRING PTR, BYVAL b AS WSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   RETURN r
END FUNCTION

FUNCTION Afx_ArraySortWStringArrayReverseCompare CDECL (BYVAL a AS WSTRING PTR, BYVAL b AS WSTRING PTR) AS LONG
   STATIC Flags AS DWORD = SORT_DIGITSASNUMBERS OR LINGUISTIC_IGNORECASE OR LINGUISTIC_IGNOREDIACRITIC OR NORM_IGNOREKANATYPE OR NORM_IGNOREWIDTH
   DIM r AS LONG = CompareStringEx(LOCALE_NAME_USER_DEFAULT, Flags, a, LEN(*a), b, LEN(*b), NULL, NULL, 0)
   IF r <> 0 THEN r = r -2
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION

SUB Afx_ArraySortWString OVERLOAD (BYVAL rgwstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL sizeElem AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF bAscend THEN
      qsort rgwstr, numElm, sizeElem, CPTR(ANY PTR, @Afx_ArraySortWStringCompare)
   ELSE
      qsort rgwstr, numElm, sizeElem, CPTR(ANY PTR, @Afx_ArraySortWStringArrayReverseCompare)
   END IF
END SUB

' ########################################################################################
'                                      BYTE
' ########################################################################################

FUNCTION Afx_ArraySortByteCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Byte PTR, a) > *CPTR(Byte PTR, b) THEN RETURN 1
   IF *CPTR(Byte PTR, a) < *CPTR(Byte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortByteReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Byte PTR, a) < *CPTR(Byte PTR, b) THEN RETURN 1
   IF *CPTR(Byte PTR, a) > *CPTR(Byte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

' qsort function
SUB Afx_ArraySortByte OVERLOAD (BYVAL rgbyte AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgbyte = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgbyte, numElm, SIZEOF(BYTE), CPTR(ANY PTR, @Afx_ArraySortByteCompare)
   ELSE
      qsort rgbyte, numElm, SIZEOF(BYTE), CPTR(ANY PTR, @Afx_ArraySortByteReverseCompare)
   END IF
END SUB

' qsort for numbers
SUB Afx_ArraySortByte OVERLOAD (rgbyte() AS BYTE, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgbyte) - LBOUND(rgbyte) + 1
   Afx_ArraySortByte @rgbyte(LBOUND(rgbyte)), numElm, bAscend
END SUB

' ########################################################################################
'                                      UBYTE
' ########################################################################################

FUNCTION Afx_ArraySortUByteCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UByte PTR, a) > *CPTR(UByte PTR, b) THEN RETURN 1
   IF *CPTR(UByte PTR, a) < *CPTR(UByte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortUByteReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UByte PTR, a) < *CPTR(UByte PTR, b) THEN RETURN 1
   IF *CPTR(UByte PTR, a) > *CPTR(UByte PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

' qsort function
SUB Afx_ArraySortUByte OVERLOAD (BYVAL rgubyte AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgubyte = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgubyte, numElm, SIZEOF(UBYTE), CPTR(ANY PTR, @Afx_ArraySortUByteCompare)
   ELSE
      qsort rgubyte, numElm, SIZEOF(UBYTE), CPTR(ANY PTR, @Afx_ArraySortUByteReverseCompare)
   END IF
END SUB

' qsort for numbers
SUB Afx_ArraySortUByte OVERLOAD (rgubyte() AS UBYTE, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgubyte) - LBOUND(rgubyte) + 1
   Afx_ArraySortUByte @rgubyte(LBOUND(rgubyte)), numElm, bAscend
END SUB

' ########################################################################################
'                                      SHORT
' ########################################################################################

FUNCTION Afx_ArraySortShortCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Short PTR, a) > *CPTR(Short PTR, b) THEN RETURN 1
   IF *CPTR(Short PTR, a) < *CPTR(Short PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortUShortReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Short PTR, a) < *CPTR(Short PTR, b) THEN RETURN 1
   IF *CPTR(Short PTR, a) > *CPTR(Short PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

SUB Afx_ArraySortShort OVERLOAD (BYVAL rgshort AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgshort = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgshort, numElm, SIZEOF(SHORT), CPTR(ANY PTR, @Afx_ArraySortShortCompare)
   ELSE
      qsort rgshort, numElm, SIZEOF(SHORT), CPTR(ANY PTR, @Afx_ArraySortUShortReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortShort OVERLOAD (rgshort() AS SHORT, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgshort) - LBOUND(rgshort) + 1
   Afx_ArraySortShort @rgshort(LBOUND(rgshort)), numElm, bAscend
END SUB

' ########################################################################################
'                                      USHORT
' ########################################################################################

FUNCTION Afx_ArraySortUShortCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UShort PTR, a) > *CPTR(UShort PTR, b) THEN RETURN 1
   IF *CPTR(UShort PTR, a) < *CPTR(UShort PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySorUShortReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UShort PTR, a) < *CPTR(UShort PTR, b) THEN RETURN 1
   IF *CPTR(UShort PTR, a) > *CPTR(UShort PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

SUB Afx_ArraySortUShort OVERLOAD (BYVAL rgUShort AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgUShort = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgUShort, numElm, SIZEOF(USHORT), CPTR(ANY PTR, @Afx_ArraySortUShortCompare)
   ELSE
      qsort rgUShort, numElm, SIZEOF(USHORT), CPTR(ANY PTR, @Afx_ArraySorUShortReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortUShort OVERLOAD (rgUShort() AS UShort, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgUShort) - LBOUND(rgUShort) + 1
   Afx_ArraySortUShort @rgUShort(LBOUND(rgUShort)), numElm, bAscend
END SUB

' ########################################################################################
'                                      INTEGER
' ########################################################################################

FUNCTION Afx_ArraySortIntegerCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Integer PTR, a) > *CPTR(Integer PTR, b) THEN RETURN 1
   IF *CPTR(Integer PTR, a) < *CPTR(Integer PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortIntegerReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(Integer PTR, a) < *CPTR(Integer PTR, b) THEN RETURN 1
   IF *CPTR(Integer PTR, a) > *CPTR(Integer PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

SUB Afx_ArraySortInteger OVERLOAD (BYVAL rgInteger AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgInteger = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgInteger, numElm, SIZEOF(INTEGER), CPTR(ANY PTR, @Afx_ArraySortIntegerCompare)
   ELSE
      qsort rgInteger, numElm, SIZEOF(INTEGER), CPTR(ANY PTR, @Afx_ArraySortIntegerReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortInteger OVERLOAD (rgInteger() AS Integer, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgInteger) - LBOUND(rgInteger) + 1
   Afx_ArraySortInteger @rgInteger(LBOUND(rgInteger)), numElm, bAscend
END SUB

' ########################################################################################
'                                      UINTEGER
' ########################################################################################

FUNCTION Afx_ArraySortUIntegerCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UInteger PTR, a) > *CPTR(UInteger PTR, b) THEN RETURN 1
   IF *CPTR(UInteger PTR, a) < *CPTR(UInteger PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortUIntegerReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(UInteger PTR, a) < *CPTR(UInteger PTR, b) THEN RETURN 1
   IF *CPTR(UInteger PTR, a) > *CPTR(UInteger PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

SUB Afx_ArraySortUInteger OVERLOAD (BYVAL rgUInteger AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgUInteger = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgUInteger, numElm, SIZEOF(UINTEGER), CPTR(ANY PTR, @Afx_ArraySortUIntegerCompare)
   ELSE
      qsort rgUInteger, numElm, SIZEOF(UINTEGER), CPTR(ANY PTR, @Afx_ArraySortUIntegerReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortUInteger OVERLOAD (rgUInteger() AS UInteger, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rgUInteger) - LBOUND(rgUInteger) + 1
   Afx_ArraySortUInteger @rgUInteger(LBOUND(rgUInteger)), numElm, bAscend
END SUB

' ########################################################################################
'                                      LONG
' ########################################################################################

FUNCTION Afx_ArraySortLongCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(LONG PTR, a) > *CPTR(LONG PTR, b) THEN RETURN 1
   IF *CPTR(LONG PTR, a) < *CPTR(LONG PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortLongReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
   IF *CPTR(LONG PTR, a) < *CPTR(LONG PTR, b) THEN RETURN 1
   IF *CPTR(LONG PTR, a) > *CPTR(LONG PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

SUB Afx_ArraySortLong OVERLOAD (BYVAL rglong AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rglong = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rglong, numElm, SIZEOF(LONG), CPTR(ANY PTR, @Afx_ArraySortLongCompare)
   ELSE
      qsort rglong, numElm, SIZEOF(LONG), CPTR(ANY PTR, @Afx_ArraySortLongReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortLong OVERLOAD (rglong() AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LONG = UBOUND(rglong) - LBOUND(rglong) + 1
   Afx_ArraySortLong @rglong(LBOUND(rglong)), numElm, bAscend
END SUB

' ########################################################################################
'                                      ULONG
' ########################################################################################

FUNCTION Afx_ArraySortULongCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS ULong
   IF *CPTR(ULong PTR, a) > *CPTR(ULong PTR, b) THEN RETURN 1
   IF *CPTR(ULong PTR, a) < *CPTR(ULong PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortULongReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS ULong
   IF *CPTR(ULong PTR, a) < *CPTR(ULong PTR, b) THEN RETURN 1
   IF *CPTR(ULong PTR, a) > *CPTR(ULong PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

' qsort function
SUB Afx_ArraySortULong OVERLOAD (BYVAL rgULong AS ANY PTR, BYVAL numElm AS ULong, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgULong = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgULong, numElm, SIZEOF(ULONG), CPTR(ANY PTR, @Afx_ArraySortULongCompare)
   ELSE
      qsort rgULong, numElm, SIZEOF(ULONG), CPTR(ANY PTR, @Afx_ArraySortULongReverseCompare)
   END IF
END SUB

' qsort for numbers
SUB Afx_ArraySortULong OVERLOAD (rgULong() AS ULong, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS ULong = UBOUND(rgULong) - LBOUND(rgULong) + 1
   Afx_ArraySortULong @rgULong(LBOUND(rgUlong)), numElm, bAscend
END SUB

' ########################################################################################
'                                      LONGINT
' ########################################################################################

FUNCTION Afx_ArraySortLongIntCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LongInt
   IF *CPTR(LongInt PTR, a) > *CPTR(LongInt PTR, b) THEN RETURN 1
   IF *CPTR(LongInt PTR, a) < *CPTR(LongInt PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortLongIntReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LongInt
   IF *CPTR(LongInt PTR, a) < *CPTR(LongInt PTR, b) THEN RETURN 1
   IF *CPTR(LongInt PTR, a) > *CPTR(LongInt PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

SUB Afx_ArraySortLongInt OVERLOAD (BYVAL rgLongInt AS ANY PTR, BYVAL numElm AS LongInt, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgLongInt = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgLongInt, numElm, SIZEOF(LONGINT), CPTR(ANY PTR, @Afx_ArraySortLongIntCompare)
   ELSE
      qsort rgLongInt, numElm, SIZEOF(LONGINT), CPTR(ANY PTR, @Afx_ArraySortLongIntReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortLongInt OVERLOAD (rgLongInt() AS LongInt, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS LongInt = UBOUND(rgLongInt) - LBOUND(rgLongInt) + 1
   Afx_ArraySortLongInt @rgLongInt(LBOUND(rgLongInt)), numElm, bAscend
END SUB

' ########################################################################################
'                                      ULONGINT
' ########################################################################################

FUNCTION Afx_ArraySortULongIntCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS ULongInt
   IF *CPTR(ULongInt PTR, a) > *CPTR(ULongInt PTR, b) THEN RETURN 1
   IF *CPTR(ULongInt PTR, a) < *CPTR(ULongInt PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

FUNCTION Afx_ArraySortULongIntReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS ULongInt
   IF *CPTR(ULongInt PTR, a) < *CPTR(ULongInt PTR, b) THEN RETURN 1
   IF *CPTR(ULongInt PTR, a) > *CPTR(ULongInt PTR, b) THEN RETURN -1
   RETURN 0
END FUNCTION

SUB Afx_ArraySortULongInt OVERLOAD (BYVAL rgULongInt AS ANY PTR, BYVAL numElm AS ULongInt, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgULongInt = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgULongInt, numElm, SIZEOF(ULONGINT), CPTR(ANY PTR, @Afx_ArraySortULongIntCompare)
   ELSE
      qsort rgULongInt, numElm, SIZEOF(ULONGINT), CPTR(ANY PTR, @Afx_ArraySortULongIntReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortULongInt OVERLOAD (rgULongInt() AS ULongInt, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS ULongInt = UBOUND(rgULongInt) - LBOUND(rgULongInt) + 1
   Afx_ArraySortULongInt @rgULongInt(LBOUND(rgULongInt)), numElm, bAscend
END SUB

' ########################################################################################
'                                      SINGLE
' ########################################################################################

FUNCTION Afx_ArraySortSingleCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS Integer
   STATIC EPSILON AS Single = 1.0E-6 ' Small tolerance for floating-point precision errors
   DIM valA AS SINGLE = *CPTR(SINGLE PTR, a)
   DIM valB AS SINGLE = *CPTR(SINGLE PTR, b)
   IF ABS(valA - valB) < EPSILON THEN RETURN 0 ' Treat very close values as equal
   IF valA > valB THEN RETURN 1
   RETURN -1
END FUNCTION

FUNCTION Afx_ArraySortSingleReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS Integer
   STATIC EPSILON AS Single = 1.0E-6
   DIM valA AS SINGLE = *CPTR(SINGLE PTR, a)
   DIM valB AS SINGLE = *CPTR(SINGLE PTR, b)
   IF ABS(valA - valB) < EPSILON THEN RETURN 0
   IF valA < valB THEN RETURN 1
   RETURN -1
END FUNCTION

SUB Afx_ArraySortSingle OVERLOAD (BYVAL rgSingle AS ANY PTR, BYVAL numElm AS Single, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgSingle = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgSingle, numElm, SIZEOF(SINGLE), CPTR(ANY PTR, @Afx_ArraySortSingleCompare)
   ELSE
      qsort rgSingle, numElm, SIZEOF(SINGLE), CPTR(ANY PTR, @Afx_ArraySortSingleReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortSingle OVERLOAD (rgSingle() AS Single, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS Single = UBOUND(rgSingle) - LBOUND(rgSingle) + 1
   Afx_ArraySortSingle @rgSingle(LBOUND(rgSingle)), numElm, bAscend
END SUB

' ########################################################################################
'                                      DOUBLE
' ########################################################################################

FUNCTION Afx_ArraySortDoubleCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS Integer
   STATIC EPSILON AS DOUBLE = 1.0E-12 ' Smaller tolerance for higher precision
   DIM valA AS DOUBLE = *CPTR(DOUBLE PTR, a)
   DIM valB AS DOUBLE = *CPTR(DOUBLE PTR, b)
   IF ABS(valA - valB) < EPSILON THEN RETURN 0 ' Treat very close values as equal
   IF valA > valB THEN RETURN 1
   RETURN -1
END FUNCTION

FUNCTION Afx_ArraySortDoubleReverseCompare CDECL (BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS Integer
   STATIC EPSILON AS DOUBLE = 1.0E-12
   DIM valA AS DOUBLE = *CPTR(DOUBLE PTR, a)
   DIM valB AS DOUBLE = *CPTR(DOUBLE PTR, b)
   IF ABS(valA - valB) < EPSILON THEN RETURN 0
   IF valA < valB THEN RETURN 1
   RETURN -1
END FUNCTION

SUB Afx_ArraySortDouble OVERLOAD (BYVAL rgDouble AS ANY PTR, BYVAL numElm AS Single, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgDouble = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgDouble, numElm, SIZEOF(DOUBLE), CPTR(ANY PTR, @Afx_ArraySortDoubleCompare)
   ELSE
      qsort rgDouble, numElm, SIZEOF(DOUBLE), CPTR(ANY PTR, @Afx_ArraySortDoubleReverseCompare)
   END IF
END SUB

SUB Afx_ArraySortDouble OVERLOAD (rgDouble() AS DOUBLE, BYVAL bAscend AS BOOLEAN = TRUE)
   DIM numElm AS DOUBLE = UBOUND(rgDouble) - LBOUND(rgDouble) + 1
   Afx_ArraySortDouble @rgDouble(LBOUND(rgDouble)), numElm, bAscend
END SUB

' ########################################################################################
' Macros to identify the type of aaray and call the appropriate sorting procedures.
' Works with one-dimensional arrays.
' ########################################################################################

#macro AfxSortStringArray(rg, ascending)
#if TypeOf(rg) = TypeOf(DWSTRING)
   Afx_ArraySortDWString(rg(), ascending)
'   Afx_ArraySortDWString(@rg(0), UBOUND(rg) - LBOUND(rg) + 1, ascending)
#elseif TypeOf(rg) = TypeOf(BSTRING)
   Afx_ArraySortBString(rg(), ascending)
'   Afx_ArraySortBString(@rg(0), UBOUND(rg) - LBOUND(rg) + 1, ascending)
#elseif TypeOf(rg) = TypeOf(STRING)
   Afx_ArraySortString(rg(), ascending)
'   Afx_ArraySortString(@rg(0), UBOUND(rg) - LBOUND(rg) + 1, ascending)
#elseif TypeOf(rg) = TypeOf(STRING * SIZEOF(TypeOf((rg(0)))) -1)
' The length stored in the descriptor of a fixed-length string includes the terminating null.
' We can sort them as if they were ZStrings by substratic one from the value returned by SIZEOF,
   Afx_ArraySortZString(@rg(0), UBOUND(rg) - LBOUND(rg) + 1, SIZEOF(rg(LBOUND(rg))), ascending)
#elseif TypeOf(rg) = TypeOf(ZSTRING * SIZEOF(TypeOf((rg(0)))))
   Afx_ArraySortZString(@rg(0), UBOUND(rg) - LBOUND(rg) + 1, SIZEOF(rg(LBOUND(rg))), ascending)
#elseif TypeOf(rg) = TypeOf(WSTRING * SIZEOF(TypeOf((rg(0)))) \ 2)
   Afx_ArraySortWString(@rg(0), UBOUND(rg) - LBOUND(rg) + 1, SIZEOF(rg(LBOUND(rg))), ascending)
#endif
#endmacro

' Usage example:
'#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
'DIM rgstr(ANY) AS XSTRING
'DIM rgstr2(ANY) AS XSTRING
'DIM xStr AS XSTRING = "String - "
'DIM res AS BOOLEAN
'' // Fill the array
'FOR i AS LONG = 1 TO 10
'   AppendElementToArray(rgstr, xStr & WSTR(i), res)
'NEXT
'' Fill the array to append
'xStr = "String 2 - "
'FOR i AS LONG = 1 TO 5
'   AppendElementToArray(rgstr2, xStr & WSTR(i), res)
'NEXT
'' // Append the array to the first array
'AppendArrayToArray(rgstr, rgstr2, res)
'' // Display the array
'FOR i AS LONG = LBOUND(rgstr) TO UBOUND(rgstr)
'   print rgstr(i)
'NEXT
' // Sort the array
'AfxSortStringArray(rgstr, true)
' // Display the sorted array
'FOR i AS LONG = LBOUND(rgstr) TO UBOUND(rgstr)
'   print rgstr(i)
'NEXT

#macro AfxSortNumericArray(rg, ascending)
#if TypeOf(rg) = BYTE
   Afx_ArraySortByte(rg(), ascending)
#elseif TypeOf(rg) = UBYTE
   Afx_ArraySortUByte(rg(), ascending)
#elseif TypeOf(rg) = SHORT
   Afx_ArraySortShort(rg(), ascending)
#elseif TypeOf(rg) = USHORT
   Afx_ArraySortUShort(rg(), ascending)
#elseif TypeOf(rg) = INTEGER
   Afx_ArraySortInteger(rg(), ascending)
#elseif TypeOf(rg) = UINTEGER
   Afx_ArraySortUInteger(rg(), ascending)
#elseif TypeOf(rg) = LONG
   Afx_ArraySortLong(rg(), ascending)
#elseif TypeOf(rg) = ULONG
   Afx_ArraySortULong(rg(), ascending)
#elseif TypeOf(rg) = LONGINT
   Afx_ArraySortLongInt(rg(), ascending)
#elseif TypeOf(rg) = ULONGINT
   Afx_ArraySortULongInt(rg(), ascending)
#elseif TypeOf(rg) = SINGLE
   Afx_ArraySortSingle(rg(), ascending)
#elseif TypeOf(rg) = DOUBLE
   Afx_ArraySortDouble(rg(), ascending)
#endif
#endmacro

'Usage example:
'#define XNUMBER DOUBLE ' // or SHORT, INTEGER, LONG, LONGINT, etc.
'DIM rgNum(ANY) AS XNUMBER
'DIM rgNum2(ANY) AS XNUMBER
'DIM bNumRes AS BOOLEAN
'DIM xNum AS XNUMBER = 1234567.89
'' // Fill the array
'FOR i AS LONG = 1 TO 10
'   xNum += 0.01
'   AppendElementToArray(rgNum, xNum, bNumRes)
'NEXT
'' Fill the array to insert
'xNum = 111.01
'FOR i AS LONG = 1 TO 5
'   xNum += 1
'   AppendElementToArray(rgNum, xNum, bNumRes)
'NEXT
'' // Insert the array to the first array
'InsertArrayIntoArray(rgNum, 5, rgNum2, bNumRes)
'' // Display the array
'FOR i AS LONG = LBOUND(rgNum) TO UBOUND(rgNum)
'   print rgNum(i)
'NEXT
' // Sort the numeric array
'AfxSortNumericArray(rgNum, true)
'' // Display the sorted array
'FOR i AS LONG = LBOUND(rgNum) TO UBOUND(rgNum)
'   print rgNum(i)
'NEXT

#endif

END NAMESPACE
