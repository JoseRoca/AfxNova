' ########################################################################################
' Platform: Microsoft Windows
' Filename: CInt96.inc
' Contents: Wrapper class for the DECIMAL data type.
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
#include once "win/ole2.bi"

' ========================================================================================
' Macro for debug
' To allow debugging, define _CINT96_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _CINT96_DEBUG_
   #define _CINT96_DEBUG_ 0
#ENDIF
#ifndef _CINT96_DP_
   #define _CINT96_DP_ 1
   #MACRO CINT96_DP(st)
      #IF (_CINT96_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE AfxNova

' ========================================================================================
' CInt96 class
' ========================================================================================
TYPE CInt96

Public:
   m_int96 AS DECIMAL   ' // The underlying DECIMAL structure

Public:

   ' === Constructores and destructor ===
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF cSrc AS CInt96)
   DECLARE CONSTRUCTOR (BYREF decSrc AS DECIMAL)
   DECLARE CONSTRUCTOR (BYVAL nInteger AS LONGINT)
   DECLARE CONSTRUCTOR (BYVAL nUInteger AS ULONGINT)
   DECLARE CONSTRUCTOR (BYREF wszSrc AS WSTRING)
   DECLARE DESTRUCTOR

   ' === Assigment operators ===
   DECLARE OPERATOR LET (BYREF cSrc AS CInt96)
   DECLARE OPERATOR LET (BYREF decSrc AS DECIMAL)
   DECLARE OPERATOR LET (BYVAL n AS LONGINT)
   DECLARE OPERATOR LET (BYVAL u AS ULONGINT)
   DECLARE OPERATOR LET (BYREF s AS WSTRING)

   ' === Cast operators ===
   DECLARE OPERATOR CAST () AS CURRENCY
   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE OPERATOR CAST () AS STRING

   ' === Arithmetic operators ===
   DECLARE OPERATOR += (BYREF int96 AS CInt96)
   DECLARE OPERATOR -= (BYREF int96 AS CInt96)
   DECLARE OPERATOR *= (BYREF int96 AS CInt96)
   DECLARE OPERATOR /= (BYREF cOperand AS CInt96)

   ' === Bitwise logical operators ===
   DECLARE OPERATOR MOD= (BYREF rhs AS CInt96)
   DECLARE OPERATOR AND= (BYREF rhs AS CInt96)
   DECLARE OPERATOR OR= (BYREF rhs AS CInt96)
   DECLARE OPERATOR XOR= (BYREF rhs AS CInt96)
   DECLARE OPERATOR IMP= (BYREF rhs AS CInt96)
   DECLARE OPERATOR EQV= (BYREF lhs AS CInt96)
   DECLARE OPERATOR SHL= (BYVAL nBits AS UINTEGER)
   DECLARE OPERATOR SHR= (BYVAL nBits AS UINTEGER)

   ' === Utility methods ===
   DECLARE FUNCTION Sign () AS UBYTE
   DECLARE FUNCTION IsSigned () AS BOOLEAN
   DECLARE FUNCTION IsUnsigned () AS BOOLEAN
   DECLARE FUNCTION ToVar () AS VARIANT

END TYPE
' ========================================================================================

' // Helper procedures

' ========================================================================================
' Convert DECIMAL (scale=0, sign flag +/-) into 96-bit two's complement words
' ========================================================================================
PRIVATE SUB _Bits96_ToTwos(BYREF d AS DECIMAL, BYREF w0 AS ULONG, BYREF w1 AS ULONG, BYREF w2 AS ULONG )
   CINT96_DP("")
   ' Order: w0 = low 32, w1 = mid 32, w2 = high 32
   w0 = d.Lo32 : w1 = d.Mid32 : w2 = d.Hi32
   ' If negative, build two's complement in-place
   IF d.sign <> 0 THEN
      w0 = NOT w0 : w1 = NOT w1 : w2 = NOT w2
      DIM c AS ULONG = 1, t AS ULONG
      t = w0 : w0 = t + c : c = IIF(w0 < t, 1, 0)
      t = w1 : w1 = t + c : c = IIF(w1 < t, 1, 0)
      t = w2 : w2 = t + c
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Convert 96-bit two's complement words back to DECIMAL (scale=0, sign set)
' ========================================================================================
PRIVATE SUB _Bits96_FromTwos(BYVAL w0 AS ULONG, BYVAL w1 AS ULONG, BYVAL w2 AS ULONG, BYREF d AS DECIMAL )
   CINT96_DP("")
   DIM neg AS INTEGER = IIF((w2 AND &h80000000u) <> 0, -1, 0)
   DIM m0 AS ULONG = w0, m1 AS ULONG = w1, m2 AS ULONG = w2
   IF neg THEN
      ' magnitude = two's complement
      m0 = NOT m0 : m1 = NOT m1 : m2 = NOT m2
      DIM c AS ULONG = 1, t AS ULONG
      t = m0 : m0 = t + c : c = IIF(m0 < t, 1, 0)
      t = m1 : m1 = t + c : c = IIF(m1 < t, 1, 0)
      t = m2 : m2 = t + c
   END IF
   d.Lo32 = m0 : d.Mid32 = m1 : d.Hi32 = m2
   d.scale = 0
   d.sign  = IIF(neg, &h80, 0)
END SUB
' ========================================================================================

' ========================================================================================
' CInt96 constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96
   CINT96_DP("")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYREF cSrc AS CInt96)
   CINT96_DP("CInt96")
   m_int96 = cSrc.m_int96
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYREF decSrc AS DECIMAL)
   CINT96_DP("DECIMAL")
   IF decSrc.Scale THEN SetLastError(VarDecInt(@decSrc, @m_int96)) ELSE m_int96 = decSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Make sure that nValue is greater that -9223372036854775808 or it will be truncated.
' For values smaller than -9223372036854775808 use a string.
PRIVATE CONSTRUCTOR CInt96 (BYVAL nInteger AS LONGINT)
   CINT96_DP("LONGINT")
   SetLastError(VarDecFromI8(nInteger, @m_int96))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
PRIVATE CONSTRUCTOR CInt96 (BYVAL nUInteger AS ULONGINT)
   CINT96_DP("ULONGINT")
   SetLastError(VarDecFromUI8(nuInteger, @m_int96))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Because the bigger numeric variable natively supported by Free Basic is a long/ulong integer,
' if we want to set bigger values we need to use strings, e.g.
' DIM int96 AS CInt96 = "79228162514264337593543950335"
' ========================================================================================
PRIVATE CONSTRUCTOR CInt96 (BYREF wszSrc AS WSTRING)
   CINT96_DP("WSTRING")
   IF INSTR(wszSrc, ".") THEN SetLastError(E_INVALIDARG) : EXIT CONSTRUCTOR
   SetLastError(VarDecFromStr(@wszSrc, 0, 0, @m_int96))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CInt96 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CInt96
   CINT96_DP("")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operators
' ========================================================================================
PRIVATE OPERATOR CInt96.LET (BYREF cSrc AS CInt96)
   CINT96_DP("CInt96")
   m_int96 = cSrc.m_int96
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.LET (BYREF decSrc AS DECIMAL)
   CINT96_DP("DECIMAL")
   IF decSrc.Scale THEN SetLastError(VarDecInt(@decSrc, @m_int96)) ELSE m_int96 = decSrc
END OPERATOR
' ========================================================================================
' ========================================================================================
' Make sure that nValue is greater that -9223372036854775808 or it will be truncated.
' For values smaller than -9223372036854775808 use a string.
PRIVATE OPERATOR CInt96.LET (BYVAL nInteger AS LONGINT)
   CINT96_DP("LONGINT")
   SetLastError(VarDecFromI8(nInteger, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
' Make sure that nValue does not exceed 18446744073709551615 or it will be truncated.
' For values greater than 18446744073709551615 use a string.
PRIVATE OPERATOR CInt96.LET (BYVAL nUInteger AS ULONGINT)
   CINT96_DP("ULONGINT")
   SetLastError(VarDecFromUI8(nuInteger, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.LET (BYREF wszSrc AS WSTRING)
   CINT96_DP("WSTRING")
   IF INSTR(wszSrc, ".") THEN SetLastError(E_INVALIDARG) : EXIT OPERATOR
   SetLastError(VarDecFromStr(@wszSrc, 0, 0, @m_int96))
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operators.
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS CURRENCY
   CINT96_DP("CURRENCY")
   DIM cy AS CURRENCY
   SetLastError(VarCyFromDec(@m_int96, @cy))
   OPERATOR = cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS VARIANT
   CINT96_DP("VARIANT")
   DIM v AS VARIANT
   v.vt = VT_DECIMAL
   v.decVal = m_int96
   RETURN v
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.CAST () AS STRING
   CINT96_DP("STRING")
   DIM bstrOut AS BSTR
   VarBstrFromDec(@m_int96, 0, 0, @bstrOut)
   DIM s AS STRING = *cast(WSTRING PTR, bstrOut)
   SysFreeString(bstrOut)
   OPERATOR = s
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying decimal value.
' ========================================================================================
PRIVATE OPERATOR * (BYREF int96 AS CInt96) AS DECIMAL PTR
   CINT96_DP("CInt96")
   OPERATOR = VARPTR(int96.m_int96)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators.
' ========================================================================================
PRIVATE OPERATOR = (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   CINT96_DP("CInt96")
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   CINT96_DP("CInt96")
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) <> VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   CINT96_DP("CInt96")
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) = VARCMP_LT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   CINT96_DP("CInt96")
   RETURN (VarDecCmp(@int961.m_int96, @int962.m_int96) = VARCMP_GT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   CINT96_DP("CInt96")
   DIM hr AS HRESULT = VarDecCmp(@int961.m_int96, @int962.m_int96)
   RETURN (hr = VARCMP_LT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
   CINT96_DP("CInt96")
   DIM hr AS HRESULT = VarDecCmp(@int961.m_int96, @int962.m_int96)
   RETURN (hr = VARCMP_GT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Math operators.
' ========================================================================================
PRIVATE OPERATOR + (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
   CINT96_DP("CInt96")
   DIM dec AS DECIMAL
   DIM hr AS HRESULT = VarDecAdd(@int961.m_int96, @int962.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int961
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.+= (BYREF int96Src AS CInt96)
   CINT96_DP("CInt96")
   SetLastError(VarDecAdd(@m_int96, @int96Src.m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
   CINT96_DP("CInt96")
   DIM dec AS DECIMAL
   DIM hr AS HRESULT = VarDecSub(@int961.m_int96, @int962.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int961
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.-= (BYREF int96Src AS CInt96)
   CINT96_DP("CInt96")
   SetLastError(VarDecSub(@m_int96, @int96Src.m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF int96Src AS CInt96) AS CInt96
   CINT96_DP("CInt96")
   DIM dec AS DECIMAL
   SetLastError(VarDecNeg(@int96Src.m_int96, @dec))
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
   CINT96_DP("CInt96")
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecMul(@int961.m_int96, @int962.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int961
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96.*= (BYREF int96Src AS CInt96)
   CINT96_DP("CInt96")
   SetLasterror(VarDecMul(@m_int96, @int96Src.m_int96, @m_int96))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF int96Src AS CInt96, BYREF cOperand AS CInt96) AS CInt96
   CINT96_DP("CInt96")
   DIM dec AS DECIMAL
   DIM hr AS LONG = VarDecDiv(@int96Src.m_int96, @cOperand.m_int96, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   IF dec.Scale THEN hr = VarDecInt(@dec, @dec)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN int96Src
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CInt96./= (BYREF cOperand AS CInt96)
   CINT96_DP("CInt96")
   DIM hr AS LONG = VarDecDiv(@m_int96, @cOperand.m_int96, @m_int96)
   IF hr <> S_OK THEN SetLastError(hr) : EXIT OPERATOR
   IF m_int96.Scale THEN SetLastError(VarDecInt(@m_int96, @m_int96))
END OPERATOR
' ========================================================================================

' ========================================================================================
' Divides this CInt96 value and assigns the remainder to it.
' Usage example:
' DIM int96 AS CInt96 = -17
' int96 MOD= 5
' print int96
' Result: -2
' ========================================================================================
PRIVATE OPERATOR CInt96.MOD= (BYREF rhs AS CInt96)
   CINT96_DP("CInt96")

   DIM zero AS DECIMAL : VarDecFromI4(0, @zero)
   ' Check divisor = 0
   IF VarDecCmp(@rhs.m_int96, @zero) = VARCMP_EQ THEN
      SetLastError(DISP_E_DIVBYZERO)
      EXIT OPERATOR
  END IF
   ' 0 MOD b = 0
   IF VarDecCmp(@m_int96, @zero) = VARCMP_EQ THEN
      EXIT OPERATOR
   END IF

   DIM quotient AS DECIMAL
   DIM product  AS DECIMAL
   DIM remainder AS DECIMAL
   DIM hr AS HRESULT

   ' // Divide the two decimals
   hr = VarDecDiv(@m_int96, @rhs.m_int96, @quotient)
   IF hr <> S_OK THEN
      SetLastError(hr)
      EXIT OPERATOR
   END IF
   ' // quotient = Fix(lhs / rhs)  (truncate toward zero)
   hr = VarDecFix(@quotient, @quotient)
   IF hr <> S_OK THEN
      SetLastError(hr)
      EXIT OPERATOR
   END IF
   ' // product = quotient * rhs
   hr = VarDecMul(@quotient, @rhs.m_int96, @product)
   IF hr <> S_OK THEN
      SetLastError(hr)
      EXIT OPERATOR
   END IF
   ' // remainder = lhs - product
   hr = VarDecSub(@m_int96, @product, @remainder)
   IF hr <> S_OK THEN
      SetLastError(hr)
      EXIT OPERATOR
   END IF
   ' // Store back into this instance
   m_int96 = remainder

END OPERATOR
' ========================================================================================

' ========================================================================================
' Performs a bitwise-and (conjunction) and assigns the result to a variable
' DIM int96 AS CInt96 = 5
' int96 AND= 3
' print int96
' Result = 1
' ========================================================================================
PRIVATE OPERATOR CInt96.AND= (BYREF rhs AS CInt96)
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   a0 AND= b0 : a1 AND= b1 : a2 AND= b2
   _Bits96_FromTwos(a0, a1, a2, m_int96)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Performs a bitwise-or (inclusive disjunction) and assigns the result to a variable.
' Usage example:
' DIM int96 AS CInt96 = 5
' int96 OR= 3
' PRINT int96
' Result = 7
' ========================================================================================
PRIVATE OPERATOR CInt96.OR= (BYREF rhs AS CInt96)
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   a0 OR= b0 : a1 OR= b1 : a2 OR= b2
   _Bits96_FromTwos(a0, a1, a2, m_int96)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Performs a bitwise-xor (exclusive disjunction) and assigns the result to a variable.
' Usage example: 
' DIM int96 AS CInt96 = -5
' int96 XOR= 2
' PRINT int96
' Result = -7
' ========================================================================================
PRIVATE OPERATOR CInt96.XOR= (BYREF rhs AS CInt96)
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   a0 XOR= b0 : a1 XOR= b1 : a2 XOR= b2
   _Bits96_FromTwos(a0, a1, a2, m_int96)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the bitwise-and (implication) of two numeric values.
' Usage example:
' DIM int96 AS CInt96 = 5
' int96 IMP= 3
' PRINT int96
' Result = -5
' ========================================================================================
PRIVATE OPERATOR CInt96.IMP= (BYREF rhs AS CInt96)
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   DIM r0 AS ULONG = (NOT a0) OR b0
   DIM r1 AS ULONG = (NOT a1) OR b1
   DIM r2 AS ULONG = (NOT a2) OR b2
   DIM d AS DECIMAL
   _Bits96_FromTwos(r0, r1, r2, d)
   m_int96 = d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Performs a bitwise-eqv (equivalence) and assigns the result to a variable.
' Usage example:
' DIM int96 AS CInt96 = 5
' int96 EQV= 3
' PRINT int96
' Result = -7
' ========================================================================================
PRIVATE OPERATOR CInt96.EQV= (BYREF rhs AS CInt96)
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   DIM r0 AS ULONG = NOT (a0 XOR b0)
   DIM r1 AS ULONG = NOT (a1 XOR b1)
   DIM r2 AS ULONG = NOT (a2 XOR b2)
   DIM d AS DECIMAL
   _Bits96_FromTwos(r0, r1, r2, d)
   m_int96 = d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Shifts left and assigns a value to a variable.
' Usage examples:
' DIM int96 AS CInt96 = 7
' int96 SHL= 4
' PRINT int96   ' 112
' -------
' int96 = -3
' int96 SHL= 5
' PRINT int96   ' -96
' ========================================================================================
PRIVATE OPERATOR CInt96.SHL= (BYVAL nBits AS UINTEGER)
   CINT96_DP("")
   IF nBits = 0 THEN EXIT OPERATOR
   DIM zero AS DECIMAL : VarDecFromI4(0, @zero)
   IF VarDecCmp(@m_int96, @zero) = VARCMP_EQ THEN
      EXIT OPERATOR  ' zero << n = zero
   END IF
   DIM hr AS LONG
   DIM i  AS UINTEGER
   FOR i = 1 TO nBits
      hr = VarDecAdd(@m_int96, @m_int96, @m_int96)   ' *= 2
      IF hr <> S_OK THEN
         SetLastError(hr)
         EXIT OPERATOR
      END IF
   NEXT
END OPERATOR
' ========================================================================================

' ========================================================================================
' Shifts right and assigns a value to a variable.
' Usage examples:
' DIM int96 AS CInt96 = 100
' int96 SHR= 3
' PRINT int96   ' 12
' int96 = -100
' int96 SHR= 3
' PRINT int96   ' -13
' int96 = -3
' int96 SHR= 1
' PRINT int96   ' -2
' int96 = -1
' int96 SHR= 100
' PRINT int96   ' -2
' int96 = 5
' int96 SHR= 2
' PRINT int96   ' 1
' int96 = 5
' int96 SHR= 3
' PRINT int96   ' 0
' ========================================================================================
PRIVATE OPERATOR CInt96.SHR= (BYVAL nBits AS UINTEGER)
   CINT96_DP("")
   IF nBits = 0 THEN EXIT OPERATOR
   DIM two AS DECIMAL : VarDecFromI4(2, @two)
   DIM hr AS HRESULT
   DIM i  AS UINTEGER
   FOR i = 1 TO nBits
      hr = VarDecDiv(@m_int96, @two, @m_int96)      ' / 2
      IF hr <> S_OK THEN
         SetLastError(hr)
         EXIT OPERATOR
      END IF
      IF m_int96.Scale THEN
         hr = VarDecInt(@m_int96, @m_int96)        ' floor to integer
         IF hr <> S_OK THEN
            SetLastError(hr)
            EXIT OPERATOR
         END IF
      END IF
   NEXT
END OPERATOR
' ========================================================================================

' ========================================================================================
' Converts the DECIMAL to a VT_DECIMAL variant.
' To allow to assign the decimal to a CVAR variable
' DIM dec AS CInt96 = 12345.1234
' DIM cv AS CVAR = dec.ToVar
' ========================================================================================
PRIVATE FUNCTION CInt96.ToVar () AS VARIANT
   CINT96_DP("VARIANT")
   DIM v AS VARIANT
   v.vt = VT_DECIMAL
   v.decVal = m_int96
   RETURN v
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the sign of this decimal number.
' Returns 0 if it is not signed of &h80 (128) if it is signed.
' ========================================================================================
PRIVATE FUNCTION CInt96.Sign () AS UBYTE
   CINT96_DP("")
   RETURN m_int96.sign
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns true if this number is signed or false otherwise.
' ========================================================================================
PRIVATE FUNCTION CInt96.IsSigned () AS BOOLEAN
   CINT96_DP("")
   IF m_int96.sign THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns true if this number is unsigned or false otherwise.
' ========================================================================================
PRIVATE FUNCTION CInt96.IsUnsigned () AS BOOLEAN
   CINT96_DP("")
   IF m_int96.sign = 0 THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================

END NAMESPACE

' ========================================================================================
' Outside the namespace because it is a global operator.
' ========================================================================================

USING AfxNova

' ========================================================================================
' Retrieves the absolute value of a CInt96 data type.
' ========================================================================================
PRIVATE OPERATOR Abs (BYREF c96 AS CInt96) AS CInt96
   CINT96_DP("CInt96")
   DIM dec AS DECIMAL
   SetLastError(VarDecAbs(@c96.m_int96, @dec))
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the bitwise-and (conjunction) of two numeric values
' Usage example: PRINT CInt96(5) AND CInt96(3) ' 1
' ========================================================================================
PRIVATE OPERATOR AND (BYREF lhs AS CInt96, BYREF rhs AS CInt96) AS CInt96
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(lhs.m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   DIM r0 AS ULONG = a0 AND b0
   DIM r1 AS ULONG = a1 AND b1
   DIM r2 AS ULONG = a2 AND b2
   DIM d AS DECIMAL
   _Bits96_FromTwos(r0, r1, r2, d)
   RETURN d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the bitwise-or (inclusive disjunction) of two numeric values.
' Usage example: PRINT CInt96(5) OR CInt96(3) ' 7
' ========================================================================================
PRIVATE OPERATOR OR (BYREF lhs AS CInt96, BYREF rhs AS CInt96) AS CInt96
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(lhs.m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   DIM r0 AS ULONG = a0 OR b0
   DIM r1 AS ULONG = a1 OR b1
   DIM r2 AS ULONG = a2 OR b2
   DIM d AS DECIMAL
   _Bits96_FromTwos(r0, r1, r2, d)
   RETURN d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the bitwise-xor (exclusive disjunction) of two numeric values.
' Usage example: PRINT CInt96(5) XOR CInt96(3) ' 6
' ========================================================================================
PRIVATE OPERATOR XOR (BYREF lhs AS CInt96, BYREF rhs AS CInt96) AS CInt96
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(lhs.m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   DIM r0 AS ULONG = a0 XOR b0
   DIM r1 AS ULONG = a1 XOR b1
   DIM r2 AS ULONG = a2 XOR b2
   DIM d AS DECIMAL
   _Bits96_FromTwos(r0, r1, r2, d)
   RETURN d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the bitwise-not (complement) of a numeric value
' Usage examples:
' PRINT NOT CInt96(0)    ' -1
' PRINT NOT CInt96(-1)   ' 0
' ========================================================================================
PRIVATE OPERATOR NOT (BYREF value AS CInt96) AS CInt96
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   _Bits96_ToTwos(value.m_int96, a0, a1, a2)
   DIM d AS DECIMAL
   _Bits96_FromTwos(NOT a0, NOT a1, NOT a2, d)
   RETURN d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the remainder from a division operation
' Usage example:
' DIM int96 AS CInt96 = CInt96("5")
' PRINT int96 MOD CInt96("3")
' Result: 2
' -5 MOD 3  ' -2
'  5 MOD -3 '  2
' -5 MOD -3 ' -2
'  0 MOD 7  '  0
' ========================================================================================
PRIVATE OPERATOR MOD (BYREF lhs AS CInt96, BYREF rhs AS CInt96) AS CInt96
   CINT96_DP("CInt96, CInt96")

   DIM zero AS DECIMAL : VarDecFromI4(0, @zero)
   ' Check divisor = 0
   IF VarDecCmp(@rhs.m_int96, @zero) = VARCMP_EQ THEN
      SetLastError(DISP_E_DIVBYZERO)
      RETURN lhs    ' leave unchanged on error
  END IF
   ' 0 MOD b = 0
   IF VarDecCmp(@lhs.m_int96, @zero) = VARCMP_EQ THEN
      RETURN zero
   END IF

   DIM quotient AS DECIMAL
   DIM product  AS DECIMAL
   DIM remainder AS DECIMAL
   DIM hr AS HRESULT

   ' // Divide the two decimals
   hr = VarDecDiv(@lhs.m_int96, @rhs.m_int96, @quotient)
   IF hr <> S_OK THEN
      SetLastError(hr)
      RETURN lhs
   END IF
   ' // quotient = Fix(lhs / rhs)  (truncate toward zero)
   hr = VarDecFix(@quotient, @quotient)
   IF hr <> S_OK THEN
      SetLastError(hr)
      RETURN lhs
   END IF
   ' // Multiply the quotient by the product
   ' // product = quotient * rhs
   hr = VarDecMul(@quotient, @rhs.m_int96, @product)
   IF hr <> S_OK THEN
      SetLastError(hr)
      RETURN lhs
   END IF
   ' // remainder = lhs - product
   hr = VarDecSub(@lhs.m_int96, @product, @remainder)
   IF hr <> S_OK THEN
      SetLastError(hr)
      RETURN lhs
   END IF
   RETURN remainder

END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the bitwise-and (implication) of two numeric values.
' Usage example: PRINT CInt96(5) IMP CInt96(3)  ' -5
' ========================================================================================
PRIVATE OPERATOR IMP (BYREF lhs AS CInt96, BYREF rhs AS CInt96) AS CInt96
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(lhs.m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   DIM r0 AS ULONG = (NOT a0) OR b0
   DIM r1 AS ULONG = (NOT a1) OR b1
   DIM r2 AS ULONG = (NOT a2) OR b2
   DIM d AS DECIMAL
   _Bits96_FromTwos(r0, r1, r2, d)
   RETURN d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the bitwise-and (equivalence) of two numeric values.
' Usage example: PRINT CInt96(5) EQV CInt96(3)  ' "-7"
' ========================================================================================
PRIVATE OPERATOR EQV (BYREF lhs AS CInt96, BYREF rhs AS CInt96) AS CInt96
   CINT96_DP("")
   DIM a0 AS ULONG, a1 AS ULONG, a2 AS ULONG
   DIM b0 AS ULONG, b1 AS ULONG, b2 AS ULONG
   _Bits96_ToTwos(lhs.m_int96, a0, a1, a2)
   _Bits96_ToTwos(rhs.m_int96, b0, b1, b2)
   DIM r0 AS ULONG = NOT (a0 XOR b0)
   DIM r1 AS ULONG = NOT (a1 XOR b1)
   DIM r2 AS ULONG = NOT (a2 XOR b2)
   DIM d AS DECIMAL
   _Bits96_FromTwos(r0, r1, r2, d)
   RETURN d
END OPERATOR
' ========================================================================================

' ========================================================================================
' Shifts the bits of a numeric expression to the left.
' Shift left: value SHL n  -> value * 2^n
' - Uses iterative doubling to avoid building 2^n explicitly.
' - Returns original on error and sets LastError.
' PRINT CInt96(7) SHL 4   ' 112
' PRINT CInt96(1) SHL 95  ' 39614081257132168796771975168
' PRINT CInt96(-3) SHL 5    ' -96
' ========================================================================================
PRIVATE OPERATOR SHL (BYREF value AS CInt96, BYVAL nBits AS UINTEGER) AS CInt96
   CINT96_DP("")
   IF nBits = 0 THEN RETURN value
   DIM decRes AS DECIMAL = value.m_int96
   DIM zero   AS DECIMAL : VarDecFromI4(0, @zero)
   IF VarDecCmp(@decRes, @zero) = VARCMP_EQ THEN
      RETURN value  ' zero << n = zero, return original (already zero)
   END IF
   DIM hr AS HRESULT
   DIM i  AS UINTEGER
   FOR i = 1 TO nBits
      hr = VarDecAdd(@decRes, @decRes, @decRes)   ' decRes *= 2
      IF hr <> S_OK THEN
         SetLastError(hr)
         RETURN value
       END IF
   NEXT
   RETURN decRes
END OPERATOR
' ========================================================================================

' ========================================================================================
' Shifts the bits of a numeric expression to the right.
' Shift right (arithmetic): value SHR n  -> Int(value / 2^n)
' - Divides by 2 iteratively and applies VarDecInt to keep it integer.
' - Matches arithmetic right-shift semantics for negatives (e.g., -3 >> 1 = -2).
' Usage examples:
' PRINT CInt96(100) SHR 3   ' 12
' PRINT CInt96(-100) SHR 3   ' -13
' PRINT CInt96(-3) SHR 1   ' -2
' PRINT CInt96(-1) SHR 100   ' -1
' PRINT CInt96(5) SHR 2   ' 1
' PRINT CInt96(5) SHR 3   ' 0
' ========================================================================================
PRIVATE OPERATOR SHR (BYREF value AS CInt96, BYVAL nBits AS UINTEGER) AS CInt96
   CINT96_DP("")
   IF nBits = 0 THEN RETURN value
   DIM decRes AS DECIMAL = value.m_int96
   DIM two AS DECIMAL : VarDecFromI4(2, @two)
   DIM hr AS LONG
   DIM i  AS UINTEGER
   FOR i = 1 TO nBits
      hr = VarDecDiv(@decRes, @two, @decRes)        ' / 2
      IF hr <> S_OK THEN
         SetLastError(hr)
         RETURN value
      END IF
      IF decRes.Scale THEN
         hr = VarDecInt(@decRes, @decRes)          ' floor to integer
         IF hr <> S_OK THEN
            SetLastError(hr)
            RETURN value
         END IF
      END IF
   NEXT
   RETURN decRes
END OPERATOR
' ========================================================================================
