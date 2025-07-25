' ########################################################################################
' Platform: Microsoft Windows
' Filename: CDEC.inc
' Contents: Decimal data type
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

NAMESPACE AfxNova

' ========================================================================================
' CDEC class - Decimal data type
' ========================================================================================
TYPE CDEC

Public:
   m_dec AS DECIMAL   ' // The underlying DECIMAL structure
   m_lcid AS LCID     ' // The locale id

Public:
   DECLARE CONSTRUCTOR (BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYREF cSrc AS CDEC, BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYREF decSrc AS DECIMAL, BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYVAL cySrc AS CURRENCY, BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYVAL nInteger AS LONGINT, BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYVAL nUInteger AS ULONGINT, BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYVAL dSrc AS DOUBLE, BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYREF wszSrc AS WSTRING, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0)
   DECLARE CONSTRUCTOR (BYVAL pDispSrc AS IDispatch PTR, BYVAL lcid AS LCID = 0)
   DECLARE CONSTRUCTOR (BYVAL varSrc AS VARIANT, BYVAL lcid AS LCID = 0)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR LET (BYREF cSrc AS CDEC)
   DECLARE OPERATOR += (BYREF dec AS CDEC)
   DECLARE OPERATOR -= (BYREF dec AS CDEC)
   DECLARE OPERATOR *= (BYREF dec AS CDEC)
   DECLARE OPERATOR /= (BYREF cOperand AS CDEC)
   DECLARE OPERATOR CAST () AS CURRENCY
   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE OPERATOR CAST () AS STRING
   DECLARE FUNCTION DecAbs () AS CDEC
   DECLARE FUNCTION DecFix () AS CDEC
   DECLARE FUNCTION DecInt () AS CDEC
   DECLARE FUNCTION DecRound (BYVAL nDecimals AS LONG) AS CDEC
   DECLARE FUNCTION ToVar () AS VARIANT
   DECLARE FUNCTION Scale () AS UBYTE
   DECLARE FUNCTION Sign () AS UBYTE
   DECLARE FUNCTION IsSigned () AS BOOLEAN
   DECLARE FUNCTION IsUnsigned () AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' CDec constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYVAL lcid AS LCID = 0)
   m_lcid = lcid
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYREF cSrc AS CDec, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   m_dec = cSrc.m_dec
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYREF decSrc AS DECIMAL, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   m_dec = decSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYVAL cySrc AS CURRENCY, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   SetLastError(VarDecFromCy(cySrc, @m_dec))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYVAL nInteger AS LONGINT, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   SetLastError(VarDecFromI8(nInteger, @m_dec))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYVAL nUInteger AS ULONGINT, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   SetLastError(VarDecFromUI8(nuInteger, @m_dec))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYVAL dSrc AS DOUBLE, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   SetLastError(VarDecFromR8(dSrc, @m_dec))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYREF wszSrc AS WSTRING, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0)
   m_lcid = lcid
   SetLastError(VarDecFromStr(@wszSrc, lcid, dwFlags, @m_dec))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYVAL pDispSrc AS IDispatch PTR, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   SetLastError(VarDecFromDisp(pDispSrc, lcid, @m_dec))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDEC (BYVAL varSrc AS VARIANT, BYVAL lcid AS LCID = 0)
   m_lcid = lcid
   DIM v AS VARIANT
   DIM hr AS HRESULT = VariantChangeType(@v, @varSrc, 0, VT_DECIMAL)
   IF FAILED(hr) THEN EXIT CONSTRUCTOR
   m_dec = v.decVal
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CDec destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDEC
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operators
' ========================================================================================
PRIVATE OPERATOR CDEC.LET (BYREF cSrc AS CDec)
   m_dec = cSrc.m_dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CDec.+= (BYREF decSrc AS CDEC)
   SetLastError(VarDecAdd(@m_dec, @decSrc.m_dec, @m_dec))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CDec.-= (BYREF decSrc AS CDEC)
   SetLastError(VarDecSub(@m_dec, @decSrc.m_dec, @m_dec))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CDec.*= (BYREF decSrc AS CDEC)
   SetLastError(VarDecMul(@m_dec, @decSrc.m_dec, @m_dec))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CDec./= (BYREF cOperand AS CDEC)
   SetLastError(VarDecDiv(@m_dec, @cOperand.m_dec, @m_dec))
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operators.
' ========================================================================================
PRIVATE OPERATOR CDEC.CAST () AS CURRENCY
   DIM cy AS CURRENCY
   VarCyFromDec(@m_dec, @cy)
   OPERATOR = cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CDEC.CAST () AS VARIANT
   DIM v AS VARIANT
   v.decVal = m_dec
   v.vt = VT_DECIMAL
   RETURN v
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CDEC.CAST () AS STRING
   DIM bstrOut AS BSTR
   VarBstrFromDec(@m_dec, m_lcid, 0, @bstrOut)
   DIM s AS STRING = *cast(WSTRING PTR, bstrOut)
   SysFreeString(bstrOut)
   OPERATOR = s
END OPERATOR
' ========================================================================================

' ========================================================================================
' Retrieves the absolute value of a decimal data type.
' ========================================================================================
PRIVATE FUNCTION CDEC.DecAbs () AS CDEC
   DIM dec AS DECIMAL
   SetLastError(VarDecAbs(@m_dec, @dec))
   RETURN dec
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the integer portion of a decimal data type.
' If the value is negative, then the first negative integer greater than or equal to the
' value is returned.
' ========================================================================================
PRIVATE FUNCTION CDEC.DecFix () AS CDEC
   DIM dec AS DECIMAL
   SetLastError(VarDecFix(@m_dec, @dec))
   RETURN dec
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the integer portion of a decimal data type.
' If the value is negative, then the first negative integer less than or equal to the
' value is returned.
' ========================================================================================
PRIVATE FUNCTION CDEC.DecInt () AS CDEC
   DIM dec AS DECIMAL
   SetLastError(VarDecInt(@m_dec, @dec))
   RETURN dec
END FUNCTION
' ========================================================================================

' ========================================================================================
' Rounds the decimal data type to a specified number of decimal places.
' ========================================================================================
PRIVATE FUNCTION CDEC.DecRound (BYVAL nDecimals AS LONG) AS CDEC
   DIM dec AS DECIMAL
   SetLastError(VarDecRound(@m_dec, nDecimals, @dec))
   RETURN dec
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts this decimal number to a VT_DECIMAL variant.
' To allow to assign the decimal to a DVARIANT variable
' DIM dec AS CDEC = 12345.1234
' DIM dv AS DVARIANT = dec.ToVar
' ========================================================================================
PRIVATE FUNCTION CDEC.ToVar () AS VARIANT
   DIM v AS VARIANT
   v.decVal = m_dec
   v.vt = VT_DECIMAL
   RETURN v
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the scale of this decimal number.
' ========================================================================================
PRIVATE FUNCTION CDEC.Scale () AS UBYTE
   RETURN m_dec.scale
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the sign of this decimal number.
' Returns 0 if it is not signed of &h80 (128) if it is signed.
' ========================================================================================
PRIVATE FUNCTION CDEC.Sign () AS UBYTE
   RETURN m_dec.sign
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns true if this number is signed or false otherwise.
' ========================================================================================
PRIVATE FUNCTION CDEC.IsSigned () AS BOOLEAN
   IF m_dec.sign THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns true if this number is unsigned or false otherwise.
' ========================================================================================
PRIVATE FUNCTION CDEC.IsUnsigned () AS BOOLEAN
   IF m_dec.sign = 0 THEN RETURN TRUE ELSE RETURN FALSE
END FUNCTION
' ========================================================================================

END NAMESPACE

' ========================================================================================
' Outside the namespace because they are global operators.
' ========================================================================================

USING AfxNova


' ========================================================================================
PRIVATE OPERATOR - (BYREF decSrc AS CDEC) AS CDEC
   DIM dec AS DECIMAL
   VarDecNeg(@decSrc.m_dec, @dec)
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying decimal value.
' ========================================================================================
PRIVATE OPERATOR * (BYREF dec AS CDEC) AS DECIMAL PTR
   OPERATOR = VARPTR(dec.m_dec)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR * (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS CDEC
   DIM dec AS DECIMAL
   VarDecMul(@dec1.m_dec, @dec2.m_dec, @dec)
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF decSrc AS CDEC, BYREF cOperand AS CDEC) AS CDEC
   DIM dec AS DECIMAL
   VarDecDiv(@decSrc.m_dec, @cOperand.m_dec, @dec)
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators.
' ========================================================================================
PRIVATE OPERATOR = (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
   RETURN (VarDecCmp(@dec1.m_dec, @dec2.m_dec) = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
   RETURN (VarDecCmp(@dec1.m_dec, @dec2.m_dec) <> VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
   RETURN (VarDecCmp(@dec1.m_dec, @dec2.m_dec) = VARCMP_LT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
   RETURN (VarDecCmp(@dec1.m_dec, @dec2.m_dec) = VARCMP_GT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
   DIM hr AS HRESULT = VarDecCmp(@dec1.m_dec, @dec2.m_dec)
   RETURN (hr = VARCMP_LT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
   DIM hr AS HRESULT = VarDecCmp(@dec1.m_dec, @dec2.m_dec)
   RETURN (hr = VARCMP_GT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Math operators.
' ========================================================================================
PRIVATE OPERATOR + (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS CDEC
   DIM dec AS DECIMAL
   SetLastError(VarDecAdd(@dec1.m_dec, @dec2.m_dec, @dec))
   RETURN dec
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS CDEC
   DIM dec AS DECIMAL
   SetLastError(VarDecSub(@dec1.m_dec, @dec2.m_dec, @dec))
   RETURN dec
END OPERATOR
' ========================================================================================

' ========================================================================================
' Retrieves the absolute value of a CInt96 data type.
' ========================================================================================
PRIVATE OPERATOR Abs (BYREF dec AS CDEC) AS CDEC
   DIM decOut AS DECIMAL
   SetLastError(VarDecAbs(@dec.m_dec, @decOut))
   RETURN decOut
END OPERATOR
' ========================================================================================

' ========================================================================================
' Retrieves the integer portion of a decimal data type.
' If the value is negative, then the first negative integer greater than or equal to the
' value is returned.
' ========================================================================================
PRIVATE OPERATOR Fix (BYREF dec AS CDEC) AS CDEC
   DIM decOut AS DECIMAL
   VarDecFix(@dec.m_dec, @decOut)
   RETURN decOut
END OPERATOR
' ========================================================================================
