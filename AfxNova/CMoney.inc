' ########################################################################################
' Platform: Microsoft Windows
' Filename: CMoney.inc
' Contents: Currency data type
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
#include once "AfxNova/DWSTRING.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' CMoney is a wrapper for the CURRENCY data type. CURRENCY is implemented as an 8-byte
' two's-complement integer value scaled by 10,000. This gives a fixed-point number with
' 15 digits to the left of the decimal point and 4 digits to the right. The CURRENCY data
' type is extremely useful for calculations involving money, or for any fixed-point
' calculations where aCMoneyacy is important.
' The CMoney wrapper implements arithmetic, assignment, and comparison operations for this
' fixed-point type, and provides access to the numbers on either side of the decimal point
' in the form of two components: an integer component which stores the value to the left
' of the decimal point, and a fractional component which stores the value to the right of
' the decimal point. The fractional component is stored internally as an integer value
' between -9999 (CY_MIN_FRACTION) and +9999 (CY_MAX_FRACTION). The function GetFraction
' returns a value scaled by a factor of 10000 (CY_SCALE).
' When specifying the integer and fractional components of a CMoney object, remember that
' the fractional component is a number in the range 0 to 9999. This is important when
' dealing with a currency such as the US dollar that expresses amounts using only two
' significant digits after the decimal point. Even though the last two digits are not
' displayed, they must be taken into account.
' ========================================================================================
TYPE CMoney

Public:
   DIM m_cur AS CY   ' // The underlying CURRENCY structure
   DIM CY_MIN_INTEGER AS LONGLONG = -922337203685477LL
   DIM CY_MAX_INTEGER AS LONGLONG = 922337203685477LL
   DIM CY_MIN_FRACTION AS LONG = -9999
   DIM CY_MAX_FRACTION AS LONG = 9999
   DIM CY_SCALE AS LONG = 10000

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF cSrc AS CMoney)
   DECLARE CONSTRUCTOR (BYVAL cySrc AS CURRENCY)
   DECLARE CONSTRUCTOR (BYVAL nInteger AS LONGLONG)
   DECLARE CONSTRUCTOR (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT)
   DECLARE CONSTRUCTOR (BYVAL dSrc AS DOUBLE)
   DECLARE CONSTRUCTOR (BYVAL dSrc AS DECIMAL)
   DECLARE CONSTRUCTOR (BYREF wszSrc AS WSTRING, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0)
   DECLARE CONSTRUCTOR (BYVAL varSrc AS VARIANT)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION SetInteger (BYVAL nInteger AS LONGLONG) AS HRESULT
   DECLARE FUNCTION SetValues (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT) AS HRESULT
   DECLARE FUNCTION SetFraction (BYVAL nFraction AS SHORT) AS HRESULT
   DECLARE OPERATOR LET (BYREF cSrc AS CMoney)
   DECLARE OPERATOR += (BYREF cur AS CMoney)
   DECLARE OPERATOR -= (BYREF cur AS CMoney)
   DECLARE OPERATOR *= (BYREF cur AS CMoney)
   DECLARE OPERATOR /= (BYREF cOperand AS CMoney)
   DECLARE OPERATOR CAST () AS DOUBLE
   DECLARE OPERATOR CAST () AS CURRENCY
   DECLARE OPERATOR CAST () AS STRING
   DECLARE OPERATOR CAST () AS VARIANT
   DECLARE FUNCTION Round (BYVAL nDecimals AS LONG) AS CMoney
   DECLARE FUNCTION GetInteger () AS LONGLONG
   DECLARE FUNCTION GetFraction () AS SHORT
   DECLARE FUNCTION ToVar () AS VARIANT
   DECLARE FUNCTION FormatCurrency (BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
           BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DECLARE FUNCTION FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
           BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING

END TYPE
' ========================================================================================

' ========================================================================================
' CMoney constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYREF cSrc AS CMoney)
   m_cur.int64 = cSrc.m_cur.int64
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYVAL cySrc AS CURRENCY)
   m_cur.int64 = cySrc.int64
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYVAL nInteger AS LONGLONG)
   this.SetInteger(nInteger)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT)
   this.SetInteger(nInteger)
   this.SetFraction(nFraction)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYVAL dSrc AS DOUBLE)
   VarCyFromR8(dSrc, @m_cur)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYVAL dSrc AS DECIMAL)
   VarCyFromDec(@dSrc, @m_cur)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYREF wszSrc AS WSTRING, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0)
   VarCyFromStr(@wszSrc, lcid, dwFlags, @m_cur)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMoney (BYVAL varSrc AS VARIANT)
   DIM v AS VARIANT
   DIM hr AS HRESULT = VariantChangeType(@v, @varSrc, 0, VT_CY)
   IF FAILED(hr) THEN EXIT CONSTRUCTOR
   m_cur = v.cyVal
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Sets the integer component.
' ========================================================================================
PRIVATE FUNCTION CMoney.SetInteger (BYVAL nInteger AS LONGLONG) AS HRESULT
   ' // Check if within range
   IF nInteger < this.CY_MIN_INTEGER OR nInteger > this.CY_MAX_INTEGER THEN RETURN E_INVALIDARG
   ' // Add new integer part scaled by CY_SCALE
   IF m_cur.int64 = 0 THEN m_cur.int64 += nInteger * this.CY_SCALE : RETURN S_OK
   ' // Signs must match
   IF (m_cur.int64 < 0 AND nInteger > 0) OR _
      (m_cur.int64 > 0 AND nInteger < 0) THEN RETURN TYPE_E_TYPEMISMATCH
   DIM cyTemp AS CURRENCY
   ' // Get fractional part
   cyTemp.int64 = m_cur.int64 MOD CY_SCALE
   ' // Check if within range again
   IF (nInteger = this.CY_MAX_INTEGER AND cyTemp.int64 > 5807) OR _
      (nInteger = this.CY_MIN_INTEGER AND cyTemp.int64 < -5808) THEN RETURN TYPE_E_OUTOFBOUNDS
   ' // Set to fractional part, wiping out integer part
   m_cur.int64 = cyTemp.int64
   ' // Add new integer part scaled by this.CY_SCALE
   m_cur.int64 += nInteger * this.CY_SCALE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the fractional component.
' Based on 4 digits. To set .2, pass 2000, to set .0002, pass a 2
' ========================================================================================
PRIVATE FUNCTION CMoney.SetFraction (BYVAL nFraction AS SHORT) AS HRESULT
   ' // Check if within range
   IF (nFraction < this.CY_MIN_FRACTION OR nFraction > this.CY_MAX_FRACTION) THEN RETURN E_INVALIDARG
   IF m_cur.int64 = 0 THEN m_cur.int64 += nFraction : RETURN S_OK
   ' // Signs must match
   IF (m_cur.int64 < 0 AND nFraction > 0) OR _
      (m_cur.int64 > 0 AND nFraction < 0) THEN RETURN TYPE_E_TYPEMISMATCH
   DIM cyTemp AS CURRENCY
  '  // Get integer part, wiping out fractional part
   cyTemp.int64 = m_cur.int64 / CY_SCALE
   ' // Check if within range again
   IF (cyTemp.int64 = this.CY_MAX_INTEGER AND nFraction > 5807) OR _
      (cyTemp.int64 = this.CY_MIN_INTEGER AND nFraction < -5808) THEN RETURN TYPE_E_OUTOFBOUNDS
   ' // Scale to CY_SCALE
   m_cur.int64 = cyTemp.int64 * this.CY_SCALE
   m_cur.int64 += nFraction
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the integer and fractional components.
' Based on 4 digits. To set .2, pass 2000, to set .0002, pass a 2
' ========================================================================================
PRIVATE FUNCTION CMoney.SetValues (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT) AS HRESULT
   DIM hr AS HRESULT = this.SetInteger(nInteger)
   IF hr = S_OK THEN hr = this.SetFraction(nFraction)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' CMoney destructor
' ========================================================================================
PRIVATE DESTRUCTOR CMoney
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operators
' ========================================================================================
PRIVATE OPERATOR CMoney.LET (BYREF cSrc AS CMoney)
   m_cur.int64 = cSrc.m_cur.int64
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators.
' ========================================================================================
PRIVATE OPERATOR = (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) <> VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) = VARCMP_LT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
   RETURN (VarCyCmp(cur1.m_cur, cur2.m_cur) = VARCMP_GT)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
   DIM hr AS HRESULT = VarCyCmp(cur1.m_cur, cur2.m_cur)
   RETURN (hr = VARCMP_LT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS BOOLEAN
   DIM hr AS HRESULT = VarCyCmp(cur1.m_cur, cur2.m_cur)
   RETURN (hr = VARCMP_GT) OR (hr = VARCMP_EQ)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Math operators.
' ========================================================================================
PRIVATE OPERATOR + (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS CMoney
   DIM cy AS CURRENCY
   VarCyAdd(cur1.m_cur, cur2.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CMoney.+= (BYREF cur AS CMoney)
   VarCyAdd(m_cur, cur.m_cur, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS CMoney
   DIM c AS CURRENCY
   VarCySub(cur1.m_cur, cur2.m_cur, @c)
   RETURN c
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CMoney.-= (BYREF cur AS CMoney)
   VarCySub(m_cur, cur.m_cur, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF cur AS CMoney) AS CMoney
   DIM cy AS CURRENCY
   VarCyNeg(cur.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF cur1 AS CMoney, BYREF cur2 AS CMoney) AS CMoney
   DIM cy AS CURRENCY
   VarCyMul(cur1.m_cur, cur2.m_cur, @cy)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CMoney.*= (BYREF cur AS CMoney)
   VarCyMul(m_cur, cur.m_cur, @m_cur)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF cur AS CMoney, BYVAL cOperand AS CMoney) AS CMoney
   DIM cy AS CURRENCY
   IF cOperand.m_cur.int64 = 0 THEN RETURN cur.m_cur.int64 / cur.CY_SCALE
   cy.int64 = cur.m_cur.int64 / (cOperand.m_cur.int64 / cur.CY_SCALE)
   RETURN cy
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CMoney./= (BYREF cOperand AS CMoney)
   IF cOperand.m_cur.int64 = 0 THEN EXIT OPERATOR
   m_cur.int64 /= (cOperand.m_cur.int64 / this.CY_SCALE)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operators.
' ========================================================================================
PRIVATE OPERATOR CMoney.CAST () AS CURRENCY
   OPERATOR = m_cur
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CMoney.CAST () AS DOUBLE
   OPERATOR = m_cur.int64 / this.CY_SCALE
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CMoney.CAST () AS STRING
'   DIM bstrOut AS BSTR
'   VarBstrFromCy(m_cur, 0, 0, @bstrOut)
'   DIM s AS STRING = *CAST(WSTRING PTR, bstrOut)
'   SysFreeString(bstrOut)
'   OPERATOR = s
   DIM s AS STRING = STR(m_cur.int64 / this.CY_SCALE)
   DIM p AS LONG = INSTR(s, ".")
   DIM dec AS STRING
   IF p THEN
      dec = MID(s + "0000", p + 1, 4)
      s = LEFT(s, p) & dec
   END IF
   IF s = "0" THEN s = "0.0000"
   OPERATOR = s
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the currency value as a VT_CY variant.
' DIM c AS CMoney = 12345.1234
' DIM v AS VARIANT = c
' cv = v
' ========================================================================================
PRIVATE OPERATOR CMoney.CAST () AS VARIANT
   DIM v AS VARIANT, cy AS CURRENCY
   VarCyFromR8(m_cur.int64 / this.CY_SCALE, @cy)
   v.vt = VT_CY
   v.cyVal = cy
   RETURN v
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the underlying currency value.
' ========================================================================================
PRIVATE OPERATOR * (BYREF cur AS CMoney) AS CURRENCY PTR
   OPERATOR = VARPTR(cur.m_cur)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Rounds the currency to a specified number of decimal places.
' ========================================================================================
PRIVATE FUNCTION CMoney.Round (BYVAL nDecimals AS LONG) AS CMoney
   DIM cy AS CURRENCY
   VarCyRound(m_cur, nDecimals, @cy)
   RETURN cy
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the integer component of the m_currency data member.
' ========================================================================================
PRIVATE FUNCTION CMoney.GetInteger () AS LONGLONG
   IF m_cur.int64 THEN RETURN FIX(m_cur.int64 / this.CY_SCALE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the fractional component of the m_currency data member.
' ========================================================================================
PRIVATE FUNCTION CMoney.GetFraction () AS SHORT
   IF m_cur.int64 THEN RETURN m_cur.int64 MOD this.CY_SCALE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the currency to a VT_CY variant.
' To allow to assign the currency to a VARIANT.
' DIM c AS CMoney = 12345.1234
' DIM v AS DVARIANT = c.ToVar
' ========================================================================================
PRIVATE FUNCTION CMoney.ToVar () AS VARIANT
   DIM v AS VARIANT, cy AS CURRENCY
   VarCyFromR8(m_cur.int64 / this.CY_SCALE, @cy)
   v.vt = VT_CY
   v.cyVal = cy
   RETURN v
END FUNCTION
' ========================================================================================

' =====================================================================================
' Formats a currency number into a string form.
' Parameters:
' - iNumDig: The number of digits to pad to after the decimal point.
'     Specify -1 to use the system default value.
' - iIncLead: Specifies whether to include the leading digit on numbers.
'     -2 : Use the system default.
'     -1 : Include the leading digit.
'      0 : Do not include the leading digit.
' - iUseParens: Specifies whether negative numbers should use parentheses.
'     -2 : Use the system default.
'     -1 : Use parentheses.
'      0 : Do not use parentheses.
' - iGroup: Specifies whether thousands should be grouped. For example 10,000 versus 10000.
'      Note Regular numbers and currencies have separate system defaults for all the above options.
'     -2 : Use the system default.
'     -1 : Group thousands.
'      0 : Do not group thousands.
' - dwFlags
'     VAR_CALENDAR_HIJRI is the only flag that can be set.
' Example:
' DIM c AS CMoney = 12345.1234
' PRINT c.FormatCurrency   --> 12.345,12 € (Spain)
' =====================================================================================
PRIVATE FUNCTION CMoney.FormatCurrency (BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DIM bstrOut AS BSTR
   DIM v AS VARIANT = this.ToVar
   VarFormatCurrency(@v, iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut)
   DIM dwsOut AS DWSTRING = bstrOut
   SysFreeString bstrOut
   RETURN dwsOut
END FUNCTION
' =====================================================================================

' =====================================================================================
' Formats a currency number into a string form.
' Same as FormatCurrency but without adding the currency symbol.
' Example:
'   DIM c AS CMoney = 12345.1234
'   PRINT c.FormatNumber   --> 12.345,12 (Spain)
' =====================================================================================
PRIVATE FUNCTION CMoney.FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DIM bstrOut AS BSTR
   DIM v AS VARIANT = this.ToVar
   VarFormatNumber(@v, iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut)
   DIM dwsOut AS DWSTRING = bstrOut
   SysFreeString bstrOut
   RETURN dwsOut
END FUNCTION
' =====================================================================================

END NAMESPACE
