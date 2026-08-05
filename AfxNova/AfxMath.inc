' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxMath.inc
' Purpose:  Math functions and macros
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2026 José Roca
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
#include once "crt/math.bi"
#include once "crt/limits.bi"
#include once "crt/ctype.bi"
#include once "crt/win32/ctype.bi"
#include once "win/winnls.bi"
#include once "win/bcrypt.bi"
#include once "string.bi"
#include once "AfxNova/DWSTRING.inc"

#inclib "bcrypt"

CONST AFX_PI      AS DOUBLE = 3.14159265358979323846
CONST AFX_TWO_PI  AS DOUBLE = 6.28318530717958647692
CONST AFX_HALF_PI AS DOUBLE = 1.57079632679489661923
CONST AFX_E       AS DOUBLE = 2.71828182845904523536
CONST AFX_SQRT2   AS DOUBLE = 1.41421356237309504880
CONST AFX_LN2     AS DOUBLE = 0.69314718055994530942
CONST AFX_LN10    AS DOUBLE = 2.30258509299404568402
CONST AFX_EPSILON AS DOUBLE = 2.2204460492503131e-16
CONST AFX_SQRT_EPSILON AS DOUBLE = 1.4901161193847656e-08

' ========================================================================================
' Calculates the absolute value of a number.
' The absolute value of a number is its positive magnitude. If a number is negative, its
' value will be negated and the positive result returned. For example, Abs(-1) and Abs(1)
' both return 1. The required number argument can be any valid numeric expression.
' Parameter
'   x : Value to find the absolute value of.
' Return Value
'   The absolute value of number.
' Usage examples:
'   DIM AS INTEGER iValue = -42
'   DIM AS INTEGER iResult
'   iResult = AfxAbs(iValue)
'   PRINT "Absolute value of "; iValue; " = "; iResult
' Expected output: 42
' ========================================================================================
#define AfxAbs(x) iif((x) < 0, -(x), (x))
' ========================================================================================

' ========================================================================================
' Returns e raised to the power of x.
' Calculates the exponential value of x using the mathematical constant e (Euler's number).
' Equivalent to Exp(x) in FreeBasic. Useful for growth, decay, and logarithmic calculations.
' Parameter
'   x : Exponent value.
' Return Value
'   e raised to the power of x.
' Usage example:
'   DIM AS DOUBLE dX = 2.0
'   DIM AS DOUBLE dResult
'   dResult = AfxExp(dX)
'   PRINT "e raised to the power of "; dX; " = "; dResult
' Expected output: 7.389056
 ' ========================================================================================
#define AfxExp(x) Exp((x))
' ========================================================================================

' ========================================================================================
' Calculates 2 raised to the power of x.
' Equivalent to Exp(x * Log(2)), but faster and clearer semantically.
' Useful for binary scaling, exponential growth, and power-of-two computations.
' Parameter
'   dX : The exponent value.
' Return Value
'   2 raised to the power of dX.
' Usage example:
'   DIM AS DOUBLE dX = 8.0
'   DIM AS DOUBLE dResult
'   dResult = AfxExp2(dX)
'   PRINT "2 raised to the power of "; dX; " = "; dResult
' ========================================================================================
#define AfxExp2(x) (Exp(dX * Log(2.0)))

' ========================================================================================
' Returns the larger of two values.
' Compares x and y and returns whichever is greater. Works with any numeric type.
' Parameters
'   x : First value.
'   y : Second value.
' Return Value
'   The larger of x and y.
' Usage example:
'   DIM AS INTEGER iA = 25
'   DIM AS INTEGER iB = 42
'   DIM AS INTEGER iResult
'   iResult = AfxMax(iA, iB)
'   PRINT "Maximum of "; iA; " and "; iB; " = "; iResult
' Expected output: 42
' ========================================================================================
#define AfxMax(x, y) iif((x) > (y), (x), (y))
' ========================================================================================

' ========================================================================================
' Returns the smaller of two values.
' Compares x and y and returns whichever is smaller. Works with any numeric type.
' Parameters
'   x : First value.
'   y : Second value.
' Return Value
'   The smaller of x and y.
' Usage example:
'   DIM AS INTEGER iA = 25
'   DIM AS INTEGER iB = 42
'   DIM AS INTEGER iResult
'   iResult = AfxMin(iA, iB)
'   PRINT "Minimum of "; iA; " and "; iB; " = "; iResult
' Expected output: 25
' ========================================================================================
#define AfxMin(x, y) iif((x) < (y), (x), (y))
' ========================================================================================

' ========================================================================================
' Restricts a value to remain within a specified range.
' If the value is less than the minimum, the minimum is returned. If the value is greater
' than the maximum, the maximum is returned. Otherwise, the original value is returned.
' This is useful for keeping values within valid limits, such as screen coordinates,
' color ranges, or normalized data.
' Parameters
'   v   : Value to clamp.
'   min : Lower limit of the range.
'   max : Upper limit of the range.
' Return Value
'   The clamped value, guaranteed to be between min and max.
'   If v < min ? returns min
'   If v > max ? returns max
'   Otherwise ? returns v
'   Works with any numeric type that supports comparison.
' Usage example:
'   DIM AS INTEGER iValue = 125
'   DIM AS INTEGER iMin   = 0
'   DIM AS INTEGER iMax   = 100
'   DIM AS INTEGER iResult
'   iResult = AfxClamp(iValue, iMin, iMax)
'   PRINT "Clamped value of "; iValue; " between "; iMin; " and "; iMax; " = "; iResult
' Expected output: 100
' ========================================================================================
#define AfxClamp(v, min, max) iif((v) < (min), (min), iif((v) > (max), (max), (v)))
' ========================================================================================

' ========================================================================================
' Returns a value that has the magnitude of one argument and the sign of another.
' The absolute value of x is taken, and its sign is replaced with the sign of y.
' This is useful when you need to force a value to adopt the sign of another number
' without altering its magnitude.
' Parameters
'   x : Value whose magnitude will be used.
'   y : Value whose sign will be copied.
' Return Value
'   |x| with the sign of y.
' If y < 0  ? returns -Abs(x)
' If y >= 0 ? returns  Abs(x)
' Useful for numerical algorithms where magnitude and sign must be controlled independently.
' Usage example:
'   DIM AS INTEGER iX = 15
'   DIM AS INTEGER iY = -7
'   DIM AS INTEGER iResult
'   iResult = AfxCopySign(iX, iY)
'   PRINT "CopySign of "; iX; " with sign of "; iY; " = "; iResult
' Expected output: -15
' ========================================================================================
#define AfxCopySign(x, y) iif((y) < 0, -Abs(x), Abs(x))
' ========================================================================================

' ========================================================================================
' Calculates the midpoint between two integer values.
' This function returns the value halfway between a and b. It uses an overflow-safe
' formula (a + (b - a)\2) to avoid integer overflow that can occur with (a + b)\2 when
' a and b are large. Commonly used in algorithms such as binary search.
' Parameters
'   a : First integer value.
'   b : Second integer value.
' Return Value
'   The midpoint between a and b.
' Usage example:
' PRINT AfxMidpoint(10, 20)
' ========================================================================================
PRIVATE FUNCTION AfxMidpoint (BYVAL a AS LONGINT, BYVAL b AS LONGINT) AS LONGINT
   RETURN a + (b - a) \ 2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Raises a number to a specified power.
' Returns x raised to the power of y using the ^ operator. Works with any numeric type
' that supports exponentiation. This macro is simpler and faster than calling a library
' function like Pow() in C.
' Parameters
'   x : Base value.
'   y : Exponent value.
' Return Value
'   x raised to the power of y.
' Usage example:
'   DIM AS LONGINT lBase = 3
'   DIM AS LONGINT lExpo = 4
'   DIM AS LONGINT lResult
'   lResult = AfxPow(lBase, lExpo)
'   PRINT lBase; " raised to the power of "; lExpo; " = "; lResult
' Expected output: 81
' ========================================================================================
#define AfxPow(x, y) ((x) ^ (y))
' ========================================================================================

' ========================================================================================
' Determines the sign of a numeric value.
' Returns -1 if the value is negative, 1 if the value is positive, and 0 if the value is
' exactly zero. This is useful for detecting direction or polarity in calculations, such
' as velocity, gradients, or comparisons.
' Parameter
'   x : Value whose sign is to be determined.
' Return Value
'   -1 if x < 0
'    0 if x = 0
'    1 if x > 0
' Usage example:
'   DIM AS INTEGER iValue  = -37
'   DIM AS INTEGER iResult
'   iResult = AfxSign(iValue)
'   PRINT "Sign of "; iValue; " = "; iResult
' Expected output: -1
' ========================================================================================
#define AfxSign(x) iif((x) < 0, -1, iif((x) > 0, 1, 0))
' ========================================================================================

' ========================================================================================
' Returns the square root of a number.
' Equivalent to Sqr(x) in FreeBasic. Included for consistency within the AfxNova numeric
' routines set.
' Parameter
'   x : Value whose square root is to be calculated.
' Return Value
'   The square root of x.
' Usage example:
'   DIM AS DOUBLE dValue  = 49.0
'   DIM AS DOUBLE dResult
'   dResult = AfxSqrt(dValue)
'   PRINT "Square root of "; dValue; " = "; dResult
' Expected output: 7
' ========================================================================================
#define AfxSqrt(x) Sqr((x))
' ========================================================================================

' ========================================================================================
' Calculates the integer square root of a number.
' Returns the largest integer whose square is less than or equal to ulN.
' Works entirely with integer arithmetic, avoiding floating-point rounding errors.
' Useful in number theory, geometry, and optimization algorithms.
' Parameter
'   ulN : The value whose integer square root is to be calculated.
' Return Value
'   The integer square root of ulN.
' Usage example:
'   DIM AS ULONGINT ulN = 12345
'   DIM AS ULONGINT ulResult
'   ulResult = AfxISqrt(ulN)
'   PRINT "Integer square root of "; ulN; " = "; ulResult
' Expected output: 111
' ========================================================================================
FUNCTION AfxISqrt (BYVAL ulN AS ULONGINT) AS ULONGINT
   DIM AS ULONGINT ulX = ulN
   DIM AS ULONGINT ulY = (ulX + 1) SHR 1
   WHILE ulY < ulX
      ulX = ulY
      ulY = (ulX + (ulN \ ulX)) SHR 1
   WEND
   RETURN ulX
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the cube root of a number.
' Calculates the cube root using exponentiation. FreeBasic does not provide a built-in
' cbrt() function, so this macro uses x^(1/3).
' Parameter
'   x : Value whose cube root is to be calculated.
' Return Value
'   The cube root of x.
' Usage example:
'   DIM AS DOUBLE dValue  = 125.0
'   DIM AS DOUBLE dResult
'   dResult = AfxCbrt(dValue)
'   PRINT "Cubic root of "; dValue; " = "; dResult
' Expected output: 5
' ========================================================================================
#define AfxCbrt(x) ((x) ^ (1.0 / 3.0))
' ========================================================================================

' ========================================================================================
' Calculates (ulBase ^ ulExpo) MOD ulModulus using modular exponentiation.
' Performs exponentiation under a modulus in an overflow-safe way.
' Useful in number theory, cryptography, and modular arithmetic operations.
' The algorithm uses exponentiation by squaring for efficiency.
' Parameters
'   ulBase    : The base value.
'   ulExpo    : The exponent value.
'   ulModulus : The modulus value.
' Return Value
'   (ulBase ^ ulExpo) MOD ulModulus
' Usage exampple:
'   DIM AS ULONGINT ulBase    = 7
'   DIM AS ULONGINT ulExpo    = 128
'   DIM AS ULONGINT ulModulus = 19
'   DIM AS ULONGINT ulResult
'   ulResult = AfxPowMod(ulBase, ulExpo, ulModulus)
'   PRINT "Result of "; ulBase; "^"; ulExpo; " MOD "; ulModulus; " = "; ulResult
' Expected output: 11
' ========================================================================================
PRIVATE FUNCTION AfxPowMod (BYVAL ulBase AS ULONGINT, BYVAL ulExpo AS ULONGINT, BYVAL ulModulus AS ULONGINT) AS ULONGINT
   DIM AS ULONGINT ulResult = 1
   DIM AS ULONGINT ulB = ulBase MOD ulModulus
   DIM AS ULONGINT ulE = ulExpo
   WHILE ulE > 0
      IF (ulE AND 1) <> 0 THEN
         ulResult = (ulResult * ulB) MOD ulModulus
      END IF
      ulE = ulE SHR 1
      ulB = (ulB * ulB) MOD ulModulus
   WEND
   RETURN ulResult
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates the Euclidean distance (hypotenuse) from two values.
' Returns Sqr(x^2 + y^2) using floating-point arithmetic.
' Useful for geometry, vector magnitude, and distance calculations.
' Example of AfxHypot usage
' Calculates the hypotenuse of a right triangle with sides dA and dB
'   DIM AS DOUBLE dA = 3.0
'   DIM AS DOUBLE dB = 4.0
'   DIM AS DOUBLE dResult
'   dResult = AfxHypot(dA, dB)
'   PRINT "Hypotenuse of sides "; dA; " and "; dB; " = "; dResult
' Expected output: 5
' ========================================================================================
#define AfxHypot(x, y) Sqr(((x) * (x)) + ((y) * (y)))
' ========================================================================================

' ========================================================================================
' Calculates e raised to the power of x minus 1.
' Equivalent to Exp(x) - 1, but implemented as a separate routine for clarity.
' Useful for numerical analysis and functions where small x values require precision.
' Parameter
'   dX : The exponent value.
' Return Value
'   e^x - 1
' Usage example:
'   DIM AS DOUBLE dX = 1.0
'   DIM AS DOUBLE dResult
'   dResult = AfxExpm1(dX)
'   PRINT "e^"; dX; " - 1 = "; dResult
' Expected output: 1.7182818
' ========================================================================================
#define AfxExpm1(dx) Exp(dX) - 1.0
' ========================================================================================

' ========================================================================================
' Calculates 2 raised to the power of x minus 1.
' Equivalent to Exp(x * Log(2)) - 1.
' Useful for binary scaling and exponential offset computations.
' Parameter
'   dX : The exponent value.
' Return Value
'   2^x - 1
' Usage exampple:
'   DIM AS DOUBLE dX = 5.0
'   DIM AS DOUBLE dResult
'   dResult = AfxExp2m1(dX)
'   PRINT "2^"; dX; " - 1 = "; dResult
' Expected output: 31
' ========================================================================================
#define AfxExp2m1(dx) Exp(dX * Log(2.0)) - 1.0
' ========================================================================================

' ========================================================================================
' Calculates the natural logarithm (base e) of x.
' Equivalent to Log(x) in FreeBasic.
' Useful for exponential inversion, growth rate, and mathematical analysis.
' Parameter
'   dX : The input value (must be > 0).
' Return Value
'   Natural logarithm of dX.
' Usage example:
'   DIM AS DOUBLE dValue  = 2.7182818
'   DIM AS DOUBLE dResult
'   dResult = AfxLog(dValue)
'   PRINT "Natural logarithm of "; dValue; " = "; dResult
' Expected output: approximately 1
' ========================================================================================
#define AfxLog(dx) Log(dX)
' ========================================================================================

' ========================================================================================
' Calculates the base-2 logarithm of x.
' Equivalent to Log(x) / Log(2).
' Useful for binary scaling, bit-length estimation, and power-of-two analysis.
' Parameter
'   dX : The input value (must be > 0).
' Return Value
'   Logarithm base 2 of dX.
' Usage example:
'   DIM AS DOUBLE dValue  = 8.0
'   DIM AS DOUBLE dResult
'   dResult = AfxLog2(dValue)
'   PRINT "Base-2 logarithm of "; dValue; " = "; dResult
' Expected output: 3
' ========================================================================================
#define AfxLog2(dx) Log(dX) / Log(2.0)
' ========================================================================================

' ========================================================================================
' Calculates the base-10 logarithm of x.
' Equivalent to Log(x) / Log(10).
' Useful for scientific notation, scaling, and decimal magnitude analysis.
' Parameter
'   dX : The input value (must be > 0).
' Return Value
'   Logarithm base 10 of dX.
' Usage example:
'   DIM AS DOUBLE dValue  = 1000.0
'   DIM AS DOUBLE dResult
'   dResult = AfxLog10(dValue)
'   PRINT "Base-10 logarithm of "; dValue; " = "; dResult
' Expected output: 3
' ========================================================================================
#define AfxLog10(dx) Log(dX) / Log(10.0)
' ========================================================================================

' ========================================================================================
' Calculates Log(1 + x) with improved precision for small x values.
' Parameter
'   dX : The input value (must be > 0).
' Equivalent to Log(1 + x), but avoids loss of significance when x is near zero.
' Useful for numerical analysis and exponential/logarithmic series.
' Usage example:
'   DIM AS DOUBLE dX = 0.001
'   DIM AS DOUBLE dResult
'   dResult = AfxLog1p(dX)
'   PRINT "Log(1 + "; dX; ") = "; dResult
' Expected output: approximately 0.0009995003
' ========================================================================================
#define AfxLog1p(dx) Log(1.0 + dX)
' ========================================================================================

' ========================================================================================
' Calculates the logarithm of x in an arbitrary base.
' Equivalent to Log(x) / Log(base).
' Useful for custom scaling, non-standard bases, and mathematical analysis.
' Parameters
'   dX    : The input value (must be > 0).
'   dBase : The logarithmic base (must be > 0 and ? 1).
' Return Value
'   Logarithm of dX in base dBase.
' Usage example:
'   DIM AS DOUBLE dX    = 81.0
'   DIM AS DOUBLE dBase = 3.0
'   DIM AS DOUBLE dResult
'   dResult = AfxLogBase(dX, dBase)
'   PRINT "Log base "; dBase; " of "; dX; " = "; dResult
' Expected output: 4
' ========================================================================================
#define AfxLogBase(dx, dBase) Log(dX) / Log(dBase)
' ========================================================================================

' ========================================================================================
' Returns the unbiased exponent of the floating-point number dX,
' equivalent to ilogb() in C/C++.
' Special cases (matching C standard):
'   dX = 0        -> FP_ILOGB0   (usually -INT_MAX)
'   dX = ±8       -> INT_MAX
'   dX = NaN      -> FP_ILOGBNAN (usually INT_MAX)
' Parameter
'   dX : Input floating-point value.
' Return Value
'   Integer exponent e such that: dX = m * 2^e, where m is in [0.5, 1).
' Usage examples:
'   PRINT AfxILogB(8.0)        ' 3
'   PRINT AfxILogB(0.0)        ' FP_ILOGB0 (usually -INT_MAX)
'   PRINT AfxILogB(-1.0)       ' 0
'   PRINT AfxILogB(1/0.0)      ' INT_MAX
'   PRINT AfxILogB(0.0/0.0)    ' FP_ILOGBNAN (-1)
' ========================================================================================
#define AfxILogB(x) (ilogb(x))
' ========================================================================================

' ========================================================================================
' Returns the unbiased exponent of the floating-point number dX as a DOUBLE.
' Equivalent to logb() in C/C++.
' Parameter
'   dX : Input floating-point value.
' Usage example:
'   PRINT AfxLogB(8.0)         ' 3
'   PRINT AfxLogB(0.0)         ' -1.#INF
'   PRINT AfxLogB(1.0)         ' 0
'   PRINT AfxLogB(0.5)         ' -1
' ========================================================================================
#define AfxLogB(x) (logb(x))
' ========================================================================================

' ========================================================================================
' Calculates the sine of an angle in radians.
' Equivalent to Sin(x) in FreeBasic.
' Useful for trigonometric, geometric, and waveform computations.
' Parameter
'   dX : Angle in radians.
' Return Value
'   Sine of dX.
' Usage example:
'   DIM AS DOUBLE dAngle = 1.5707963   ' p/2 radians
'   DIM AS DOUBLE dResult
'   dResult = AfxSin(dAngle)
'   PRINT "Sin("; dAngle; ") = "; dResult
' Expected output: 1
' ========================================================================================
#define AfxSin(dx) Sin(dx)
' ========================================================================================

' ========================================================================================
' Calculates the cosine of an angle in radians.
' Equivalent to Cos(x) in FreeBasic.
' Useful for trigonometric, geometric, and waveform computations.
' Parameter
'   dX : Angle in radians.
' Return Value
'   Cosine of dX.
' Usage example:
'   DIM AS DOUBLE dAngle = 0.0
'   DIM AS DOUBLE dResult
'   dResult = AfxCos(dAngle)
'   PRINT "Cos("; dAngle; ") = "; dResult
' Expected output: approximately 1
' ========================================================================================
#define AfxCos(dx) Cos(dx)
' ========================================================================================

' ========================================================================================
' Calculates the tangent of an angle in radians.
' Equivalent to Tan(x) in FreeBasic.
' Useful for trigonometric, geometric, and waveform computations.
' Parameter
'   dX : Angle in radians.
' Return Value
'   Tangent of dX.
' Usage example:
'   DIM AS DOUBLE dAngle = 0.7853982   ' p/4 radians
'   DIM AS DOUBLE dResult
'   dResult = AfxTan(dAngle)
'   PRINT "Tan("; dAngle; ") = "; dResult
' Expected output: approximately 1
' ========================================================================================
#define AfxTan(dx) Tan(dx)
' ========================================================================================

' ========================================================================================
' Calculates the hyperbolic sine of x.
' Equivalent to Sinh(x) in FreeBasic.
' Useful for hyperbolic geometry, exponential growth, and signal processing.
' Parameter
'   dX : Input value.
' Return Value
'   Hyperbolic sine of dX.
' Usage example:
'   DIM AS DOUBLE dX = 1.0
'   DIM AS DOUBLE dResult
'   dResult = AfxSinh(dX)
'   PRINT "Sinh("; dX; ") = "; dResult
' Expected output: approximately 1.1752012
' ========================================================================================
#define AfxSinh(dx) Sinh(dx)
' ========================================================================================

' ========================================================================================
' Calculates the hyperbolic cosine of x.
' Equivalent to Cosh(x) in FreeBasic.
' Useful for hyperbolic geometry, exponential growth, and signal processing.
' Parameter
'   dX : Input value.
' Return Value
'   Hyperbolic cosine of dX.
' Usage example:
'   DIM AS DOUBLE dX = 1.0
'   DIM AS DOUBLE dResult
'   dResult = AfxCosh(dX)
'   PRINT "Cosh("; dX; ") = "; dResult
' Expected output: approximately 1.5430806
' ========================================================================================
#define AfxCosh(dx) Cosh(dx)
' ========================================================================================

' ========================================================================================
' Calculates the hyperbolic tangent of x.
' Equivalent to Tanh(x) in FreeBasic.
' Useful for hyperbolic geometry, exponential growth, and signal processing.
' Parameter
'   dX : Input value.
' Return Value
'   Hyperbolic tangent of dX.
' Usage example:
'   DIM AS DOUBLE dX = 1.0
'   DIM AS DOUBLE dResult
'   dResult = AfxTanh(dX)
'   PRINT "Tanh("; dX; ") = "; dResult
' Expected output: approximately 0.7615942
' ========================================================================================
#define AfxTanh(dx) Tanh(dx)
' ========================================================================================

' ========================================================================================
' Calculates the arc sine (inverse sine) of x.
' Equivalent to Asin(x) in FreeBasic.
' Returns the angle in radians whose sine is x.
' Usage example:
'   DIM AS DOUBLE dX = 1.0
'   PRINT "AfxAsin("; dX; ") = "; AfxAsin(dX)
' Expected output: approximately 1.5707963 (p/2 radians)
' ========================================================================================
#define AfxAsin(dX) Asin(dX)
' ========================================================================================

' ========================================================================================
' Calculates the arc cosine (inverse cosine) of x.
' Equivalent to Acos(x) in FreeBasic.
' Returns the angle in radians whose cosine is x.
' Usage example:
'   DIM AS DOUBLE dX = 1.0
'   PRINT "AfxAcos("; dX; ") = "; AfxAcos(dX)
' Expected output: 0.0 radians
' ========================================================================================
#define AfxAcos(dX) Acos(dX)
' ========================================================================================

' ========================================================================================
' Calculates the arc tangent (inverse tangent) of x.
' Equivalent to Atan(x) in FreeBasic.
' Returns the angle in radians whose tangent is x.
' Usage example:
'   DIM AS DOUBLE dX = 1.0
'   PRINT "AfxAtan("; dX; ") = "; AfxAtan(dX)
' Expected output: approximately 0.7853982 (p/4 radians)
' ========================================================================================
#define AfxAtan(dX) Atn(dX)
' ========================================================================================

' ========================================================================================
' Calculates the arc tangent of y/x using the signs of both arguments to determine
' the correct quadrant of the result.
' Equivalent to Atan2(y, x) in FreeBasic.
' Returns the angle in radians between the positive x-axis and the point (x, y).
' Usage example:
'   DIM AS DOUBLE dY = 1.0
'   DIM AS DOUBLE dX = 1.0
'   PRINT "AfxAtan2("; dY; ", "; dX; ") = "; AfxAtan2(dY, dX)
' Expected output: approximately 0.7853982 (p/4 radians)
' ========================================================================================
#define AfxAtan2(dY, dX) Atan2(dY, dX)
' ========================================================================================

' ========================================================================================
' Calculates the inverse hyperbolic sine.
' Usage example: PRINT AfxAsinh(1.5)
' ========================================================================================
#define AfxAsinh(x) (Log(x + Sqrt(x * x + 1)))
' ========================================================================================

' ========================================================================================
' * Calculates the inverse hyperbolic cosine.
' Usage example: PRINT AfxAcosh(2.0)
' ========================================================================================
#define AfxAcosh(x) (Log(x + Sqrt(x * x - 1)))
' ========================================================================================

' ========================================================================================
' * Returns the inverse hyperbolic tangent of a number.
' Examples:
'   print AfxAtanh(0.76159416)
'      1.00000000962972
'   print AfxAtanh(-0.1)
'     -0.1003353477310756
' ========================================================================================
#define AfxAtanh(x) (Log((1 + x) / (1 - x)) / 2)
' ========================================================================================

' ========================================================================================
' Computes the value of sqr(x^2 + y^2 + z^2).
' Useful for calculating the magnitude of a 3D vector or distance in 3D space.
' Usage example:
'   DIM AS DOUBLE x = 3.0, y = 4.0, z = 12.0
'   PRINT "AfxHypot3("; x; ", "; y; ", "; z; ") = "; AfxHypot3(x, y, z)
' Expected output: 13.0
' ========================================================================================
PRIVATE FUNCTION AfxHypot3 (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
   RETURN hypot(hypot(x, y), z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the argument is a finite floating-point value.
' Returns TRUE if x is finite (normal, subnormal or zero), FALSE otherwise.
' Usage example:
'   print AfxIsFinite(123.45) ' TRUE
'   print AfxIsFinite(1.0/0.0) ' FALSE
'   print AfxIsFinite(0.0/0.0) ' FALSE
' ========================================================================================
#define AfxIsFinite(x) CBOOL((_finite(x) <> 0))
' ========================================================================================

' ========================================================================================
' Determines whether the argument is a NaN (Not a Number).
' Returns TRUE if x is NaN, FALSE otherwise.
' Usage example:
'   PRINT AfxIsNaN(0.0/0.0)    ' TRUE
'   PRINT AfxIsNaN(1.0/0.0)    ' FALSE
'   PRINT AfxIsNaN(5.0)        ' FALSE
' ========================================================================================
#define AfxIsNaN(x) CBOOL((_isnan(x) <> 0))
' ========================================================================================

' ========================================================================================
' Determines whether the argument is an infinity.
' Returns +1 if x is positive infinity, -1 if x is negative infinity and 0 otherwise.
' Usage examples:
' PRINT AfxIsInfinity( 1.0 / 0.0 )     ' returns 1
' PRINT AfxIsInfinity( -1.0 / 0.0 )    ' returns -1
' ========================================================================================
PRIVATE FUNCTION AfxIsInfinity (BYVAL x AS DOUBLE) AS LONG
   IF _isnan(x) THEN RETURN 0
   ' x is either +inf or -inf
   IF _finite(x) = FALSE THEN RETURN IIF(x > 0, 1, -1)
   RETURN 0
END FUNCTION
' ========================================================================================
#define AfxIsInf AfxIsInfinity
' ========================================================================================

' ========================================================================================
' Returns TRUE if x is a normal IEEE 754 double-precision value.
' Normal numbers are finite, non-zero, not subnormal, and >= the smallest normalized value.
' IEEE 754 double threshold ˜ 2.2250738585072014e-308
' Usage examples:
' DIM x AS DOUBLE
' x = 1e-300
' PRINT AfxIsNormal(x)      ' TRUE
' x = 1e-310
' PRINT AfxIsNormal(x)      ' FALSE   ' subnormal
' x = 0.0
' PRINT AfxIsNormal(x)      ' FALSE
' x = 1e309
' PRINT AfxIsNormal(x)      ' FALSE   ' infinite
' ========================================================================================
#define AfxIsNormal(x) CBOOL((ABS(x) >= 2.2250738585072014e-308) AND (ABS(x) < 1.7976931348623157e+308))
' ========================================================================================

' ========================================================================================
' Returns TRUE if x is a subnormal (denormalized) floating-point number.
' Subnormals are non-zero values smaller than the smallest normalized number.
' Usage examples:
'   DIM x AS DOUBLE
'   x = 1e-310
'   PRINT AfxIsSubnormal(x)   ' TRUE
'   x = 1e-300
'   PRINT AfxIsSubnormal(x)   ' FALSE
' ========================================================================================
#define AfxIsSubnormal(x) CBOOL(((Abs(x) > 0.0) And (Abs(x) < 2.2250738585072014e-308)))
' ========================================================================================

' ========================================================================================
' Determines whether the sign bit of a floating-point value is set.
' Returns 1 if x is negative or -infinity, 0 otherwise.
' Usage examples:
' PRINT AfxSignbit(-3.14)        ' true
' PRINT AfxSignbit(+3.14)        ' false
' PRINT AfxSignbit(-1.0/0.0)     ' true   (negative infinity)
' PRINT AfxSignbit( 1.0/0.0)     ' false  (positive infinity)
' PRINT AfxSignbit(0.0)          ' false
' PRINT AfxSignbit(-0.0)         ' true   (IEEE-754 negative zero)
' ========================================================================================
#define AfxSignbit(x) CBOOL(__signbit(x) <> 0)
' ========================================================================================

' ========================================================================================
' Converts degrees to radians.
' Example:
'   print AfxDegToRad(180)
'     3.141592653589793
' ========================================================================================
#define AfxDegToRad(deg) ((deg) * (Atn(1) * 4) / 180)
' ========================================================================================

' ========================================================================================
' Converts radians to degrees.
' Example:
'   print AfxRadToDeg(3.141592653589793)
'     180
' ========================================================================================
#define AfxRadToDeg(rad) ((rad) * 180 / (Atn(1) * 4))
' ========================================================================================

' ========================================================================================
' Linearly interpolates between a and b by t.
' Returns a + (b - a) * t
' Usage example:
'   PRINT AfxLerp(0, 10, 0.25)          ' 2.5
' ========================================================================================
#define AfxLerp(a,b,t) (a + (b - a) * t)
' ========================================================================================

' ========================================================================================
' Computes the interpolation factor t given a, b and v.
' Returns (v - a) / (b - a)
' Usage example:
'   PRINT AfxInverseLerp(0, 10, 2.5)    ' 0.25
' ========================================================================================
#define AfxInverseLerp(a,b,v) ((v - a) / (b - a))
' ========================================================================================

' ========================================================================================
' Remaps a value v from one range [inLo, inHi] to another [outLo, outHi].
' Equivalent to: outLo + (v - inLo) * (outHi - outLo) / (inHi - inLo)
' Usage example:
'   PRINT AfxRemap(5, 0, 10, 0, 100)    ' 50
' ========================================================================================
#define AfxRemap(v,inLo,inHi,outLo,outHi) (outLo + (v - inLo) * (outHi - outLo) / (inHi - inLo))
' ========================================================================================

' ========================================================================================
' Smooth interpolation between edge0 and edge1 using Hermite polynomial.
' Returns 0..1 transition with smooth start and end.
'   Usage example: PRINT AfxSmoothStep(0, 10, 5)       ' 0.5
' ========================================================================================
PRIVATE FUNCTION AfxSmoothStep (BYVAL edge0 AS DOUBLE, BYVAL edge1 AS DOUBLE, BYVAL x AS DOUBLE) AS DOUBLE
    DIM AS DOUBLE t = (x - edge0) / (edge1 - edge0)
    IF t < 0 THEN t = 0
    IF t > 1 THEN t = 1
    RETURN t * t * (3 - 2 * t)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Smooth ease-in/ease-out interpolation using cubic Hermite curve.
' Returns t^2 * (3 - 2t), same shape as SmoothStep but normalized for [0,1].
' Usage examples:
'   PRINT AfxEaseInOut(0.25)        ' 0.15625
'   PRINT AfxEaseInOut(0.75)        ' 0.84375
' ========================================================================================
PRIVATE FUNCTION AfxEaseInOut (BYVAL t AS DOUBLE) AS DOUBLE
    IF t < 0 THEN t = 0
    IF t > 1 THEN t = 1
    RETURN t * t * (3 - 2 * t)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the integer part of a number, truncating toward zero.
' Equivalent to Fix() but inline.
' Example:
'   print AfxFix(-3.7)   ' -3
'   print AfxFix(3.7)    '  3
' ========================================================================================
#define AfxFix(x) Fix(x)
' ========================================================================================
' .NET llama Truncate a Fix.
' ========================================================================================
#define AfxTruncate(x) Fix(x)
' ========================================================================================

' ========================================================================================
' Returns the fractional part of a number.
' Equivalent to Frac(), inline.
' Example:
'   print AfxFrac(3.75)   ' 0.75
'   print AfxFrac(-3.75)  ' -0.75
' ========================================================================================
#define AfxFrac(x) Frac(x)
' ========================================================================================

' ========================================================================================
' Returns the largest integer less than or equal to x.
' Equivalent to Int(), inline.
' Example:
'   print AfxInt(3.7)    ' 3
'   print AfxInt(-3.7)   ' -4
' ========================================================================================
#define AfxInt(x) Int(x)
' ========================================================================================

' ========================================================================================
' Returns the smallest integer greater than or equal to x.
' Equivalent to Ceil(), inline.
' Example:
'   print AfxCeil(3.2)    ' 4
'   print AfxCeil(-3.2)   ' -3
' ========================================================================================
#define AfxCeil(x) Ceil(x)
' ========================================================================================

' ========================================================================================
' Returns the largest integer less than or equal to x.
' Equivalent to Floor(), inline.
' Example:
'   print AfxFloor(3.7)    ' 3
'   print AfxFloor(-3.7)   ' -4
' ========================================================================================
#define AfxFloor(x) Floor(x)
' ========================================================================================

' ========================================================================================
' Returns the value (x or y) with the largest magnitude.
' If either argument is NaN, the result is NaN (0 / 0).
' Equivalent to .NET Math.MaxMagnitude for DOUBLE.
' Usage examples:
'   PRINT AfxMaxMagnitude(5, -10)     ' -10  (|-10| = 10 > 5)
'   PRINT AfxMaxMagnitude(-3, 2)      ' -3
'   PRINT AfxMaxMagnitude(0.5, -0.4)  ' 0.5
'   PRINT AfxMaxMagnitude(0/0, 7)     ' -1.#IND
' ========================================================================================
PRIVATE FUNCTION AfxMaxMagnitude ( BYVAL x AS DOUBLE, BYVAL y AS DOUBLE ) AS DOUBLE
   IF AfxIsNaN(x) OR AfxIsNan(y) THEN RETURN 0.0 / 0.0          ' NaN
   IF AfxAbs(x) > AfxAbs(y) THEN RETURN x ELSE RETURN y
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value (x or y) with the smallest magnitude.
' If either argument is NaN, the result is NaN (0 / 0).
' Equivalent to .NET Math.MinMagnitude for DOUBLE.
' Usage examples:
'   PRINT AfxMinMagnitude(5, -10)     ' 5        (|5| = 5 < 10)
'   PRINT AfxMinMagnitude(-3, 2)      ' 2
'   PRINT AfxMinMagnitude(0.5, -0.4)  ' -0.4
'   PRINT AfxMinMagnitude(0/0, 7)     ' -1.#IND
' ========================================================================================
PRIVATE FUNCTION AfxMinMagnitude ( BYVAL x AS DOUBLE, BYVAL y AS DOUBLE ) AS DOUBLE
   IF AfxIsNaN(x) OR AfxIsNaN(y) THEN RETURN 0.0 / 0.0          ' NaN
   IF AfxAbs(x) < AfxAbs(y) THEN RETURN x ELSE RETURN y
END FUNCTION
' ========================================================================================

' ========================================================================================
' Computes (x * y) + z with IEEE 754 semantics.
' If any argument is NaN, the result is NaN (0 / 0).
' Equivalent to .NET Math.FusedMultiplyAdd(Double, Double, Double).
' Usage examples:
'   PRINT AfxFusedMultiplyAdd(2, 3, 4)      ' 10
'   PRINT AfxFusedMultiplyAdd(-1, 8, 0.5)   ' -7.5
'   PRINT AfxFusedMultiplyAdd(0/0, 5, 1)    ' -1.#IND
' ========================================================================================
PRIVATE FUNCTION AfxFusedMultiplyAdd (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
   IF AfxIsNaN(x) OR AfxIsNaN(y) OR AfxIsNaN(z) THEN RETURN 0.0 / 0.0          ' NaN
   RETURN (x * y) + z
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the IEEE 754 remainder of x / y.
' Formula (per IEEE 754 and .NET Math.IEEERemainder):
'     r = x - (y * Q)
' where Q = ROUND(x / y) using "round to nearest, ties to even".
' If y = 0, the result is NaN (0 / 0).
' If the result is exactly 0, the sign of x determines +0 or -0.
' Usage examples:
'   PRINT AfxIEEERemainder(5, 2)       ' 1
'   PRINT AfxIEEERemainder(7, 3)       ' 1
'   PRINT AfxIEEERemainder(10, 4)      ' 2
'   PRINT AfxIEEERemainder(-10, 4)     ' -2
'   PRINT AfxIEEERemainder(0, 4)       ' 0
'   PRINT AfxIEEERemainder(5, 0)       ' -1.#IND
' ========================================================================================
PRIVATE FUNCTION AfxIEEERemainder (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS DOUBLE
   IF y = 0.0 THEN RETURN 0.0 / 0.0      ' NaN
   ' Compute quotient
   DIM q_raw AS DOUBLE
   q_raw = x / y
   ' Implement IEEE 754 "round to nearest, ties to even"
   DIM q AS DOUBLE
   DIM floor_q AS DOUBLE
   DIM frac_q AS DOUBLE
   floor_q = INT(q_raw)
   frac_q = q_raw - floor_q
   IF frac_q > 0.5 THEN
      q = floor_q + 1
   ELSEIF frac_q < 0.5 THEN
      q = floor_q
   ELSE
      ' frac_q = exactly 0.5 ? tie ? choose even integer
      IF (floor_q MOD 2) = 0 THEN
         q = floor_q
      ELSE
         q = floor_q + 1
      END IF
   END IF
   ' Compute remainder r = x - (y * q)
   DIM r AS DOUBLE
   r = x - (y * q)
   ' If r = 0, return signed zero according to x
   IF r = 0.0 THEN
      IF x < 0.0 THEN RETURN -0.0 ELSE RETURN 0.0
   END IF
   RETURN r
END FUNCTION
' ========================================================================================
#define AfxRemainder AfxIEEERemainder
' ========================================================================================

' ========================================================================================
' AfxRound
' Rounds a floating-point value to the nearest integer.
' Behavior matches MSVC round():
'   • Halfway values (fraction = ±0.5) are rounded AWAY FROM ZERO.
'   • Uses FIX(), which truncates toward zero in FreeBasic.
' Usage examples:
'   PRINT AfxRound(2.3)     ' 2
'   PRINT AfxRound(2.5)     ' 3
'   PRINT AfxRound(-2.3)    ' -2
'   PRINT AfxRound(-2.5)    ' -3
'   PRINT AfxRound(0.0)     ' 0
' ========================================================================================
PRIVATE FUNCTION AfxRound (BYVAL x AS DOUBLE) AS DOUBLE
   IF x >= 0.0 THEN RETURN FIX(x + 0.5) ELSE RETURN FIX(x - 0.5)
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxEuroRound
' European-style rounding: rounds halfway values away from zero.
' This routine implements the "half away from zero" rule commonly used in Europe
' for financial and general numeric rounding. It adds ±0.5 depending on the sign
' of the value, then truncates toward zero using AfxFix().
' Parameters:
'   dblValue   -> value to round
'   numDigits  -> number of decimal places
' Behavior:
'   • Positive numbers: add +0.5 before truncation
'   • Negative numbers: add -0.5 before truncation
'   • Equivalent to AfxRoundHalfAwayFromZero
'
' Examples:
'   AfxEuroRound(2.3)      ' 2
'   AfxEuroRound(2.5)      ' 3
'   AfxEuroRound(-2.3)     ' -2
'   AfxEuroRound(-2.5)     ' -3
'   AfxEuroRound(2.3456, 2) ' 2.35
' ========================================================================================
PRIVATE FUNCTION AfxEuroRound (BYVAL dblValue AS DOUBLE, BYVAL numDigits AS LONG = 0) AS DOUBLE
   DIM nMult AS LONG = 10 ^ numDigits
   RETURN FIX((nMult * dblValue + (SGN(dblValue)) * 0.5)) / nMult
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRoundHalfEven
' Banker's rounding (IEEE 754 default): ties go to the nearest even integer.
' ========================================================================================
PRIVATE FUNCTION AfxRoundHalfEven (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   DIM v AS DOUBLE = x * scale
   DIM i AS DOUBLE = INT(v)
   DIM f AS DOUBLE = v - i
   IF f > 0.5 THEN
      v = i + 1
   ELSEIF f < 0.5 THEN
      v = i
   ELSE
      IF (i MOD 2) = 0 THEN
         v = i
      ELSE
         v = i + 1
      END IF
   END IF
   RETURN v / scale
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRoundHalfAwayFromZero
' Rounds halfway values away from zero: 2.5 -> 3, -2.5 -> -3.
' ========================================================================================
PRIVATE FUNCTION AfxRoundHalfAwayFromZero (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   DIM v AS DOUBLE = x * scale
   IF v >= 0.0 THEN
      v = FIX(v + 0.5)
   ELSE
      v = FIX(v - 0.5)
   END IF
   RETURN v / scale
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRoundHalfUp
' Rounds halfway values toward positive infinity: 2.5 -> 3, -2.5 -> -2.
' ========================================================================================
PRIVATE FUNCTION AfxRoundHalfUp (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   DIM v AS DOUBLE = x * scale
   v = INT(v + 0.5)
   RETURN v / scale
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRoundCeil
' Always rounds upward
' ========================================================================================
PRIVATE FUNCTION AfxRoundCeil (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   RETURN CEIL(x * scale) / scale
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRoundFloor
' Always rounds downward
' ========================================================================================
PRIVATE FUNCTION AfxRoundFloor (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   RETURN FLOOR(x * scale) / scale
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRoundTrunc
' Truncates toward zero.
' ========================================================================================
PRIVATE FUNCTION AfxRoundTrunc (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   RETURN FIX(x * scale) / scale
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRoundToMultiple
' Rounds x to the nearest multiple of nMultiple.
' Example: AfxRoundToMultiple(1.37, 0.25) -> 1.25
' ========================================================================================
PRIVATE FUNCTION AfxRoundToMultiple (BYVAL x AS DOUBLE, BYVAL nMultiple AS DOUBLE) AS DOUBLE
   IF nMultiple = 0.0 THEN RETURN x
   RETURN nMultiple * AfxRoundHalfEven(x / nMultiple)
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxGcd
' Computes the greatest common divisor (GCD) of two integers using the Euclidean algorithm.
' Parameters:
'   a : first integer
'   b : second integer
' Behavior:
'   Returns the largest integer that divides both a and b without remainder.
'   If either argument is zero, returns the absolute value of the other.
'   Always returns a non-negative result.
' Examples:
'   AfxGcd(12, 8)   = 4
'   AfxGcd(54, 24)  = 6
'   AfxGcd(-42, 56) = 14
'   AfxGcd(0, 9)    = 9
' ========================================================================================
PRIVATE FUNCTION AfxGcd (BYVAL a AS LONGINT, BYVAL b AS LONGINT) AS LONGINT
   a = ABS(a)
   b = ABS(b)
   WHILE b <> 0
      DIM t AS LONGINT = b
      b = a MOD b
      a = t
   WEND
   RETURN a
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxLcm
' Computes the least common multiple (LCM) of two integers.
' Parameters:
'   a : first integer
'   b : second integer
' Behavior:
'   Returns the smallest positive integer that is a multiple of both a and b.
'   Uses the relationship: LCM(a, b) = ABS(a * b) / GCD(a, b)
'   If either argument is zero, returns 0.
' Examples:
'   AfxLcm(12, 8)   = 24
'   AfxLcm(54, 24)  = 216
'   AfxLcm(-42, 56) = 168
'   AfxLcm(0, 9)    = 0
' ========================================================================================
PRIVATE FUNCTION AfxLcm (BYVAL a AS LONGINT, BYVAL b AS LONGINT) AS LONGINT
   IF a = 0 OR b = 0 THEN RETURN 0
   RETURN ABS(a * b) \ AfxGcd(a, b)
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxFactorial
' Computes the factorial of an integer value.
' Parameters:
'   n : integer value (0 = n = 20 recommended for ULONGINT range)
' Behavior:
'   Returns n! = 1 × 2 × 3 × ... × n
'   For n = 0, returns 1 by definition.
'   Uses iterative multiplication for efficiency and avoids recursion.
' Examples:
'   AfxFactorial(0)  = 1
'   AfxFactorial(1)  = 1
'   AfxFactorial(5)  = 120
'   AfxFactorial(10) = 3628800
' ========================================================================================
PRIVATE FUNCTION AfxFactorial (BYVAL n AS LONG) AS ULONGINT
   IF n < 0 THEN RETURN 0
   DIM result AS ULONGINT = 1
   FOR i AS LONG = 2 TO n
      result *= i
   NEXT
   RETURN result
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxComb
' Computes the number of combinations (n choose k).
' Parameters:
'   n : total number of items
'   k : number of items to choose
' Behavior:
'   Returns the number of distinct combinations of k elements from a set of n.
'   Uses the formula: C(n, k) = n! / (k! * (n - k)!)
'   If k > n or either argument is negative, returns 0.
'   Uses AfxFactorial for intermediate calculations.
' Examples:
'   AfxComb(5, 2)  = 10
'   AfxComb(6, 3)  = 20
'   AfxComb(10, 0) = 1
'   AfxComb(10, 10) = 1
' ========================================================================================
PRIVATE FUNCTION AfxComb (BYVAL n AS LONG, BYVAL k AS LONG) AS ULONGINT
   IF n < 0 OR k < 0 OR k > n THEN RETURN 0
   RETURN AfxFactorial(n) \ (AfxFactorial(k) * AfxFactorial(n - k))
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxPerm
' Computes the number of permutations (n permute k).
' Parameters:
'   n : total number of items
'   k : number of items to arrange
' Behavior:
'   Returns the number of distinct ordered arrangements of k elements from a set of n.
'   Uses the formula: P(n, k) = n! / (n - k)!
'   If k > n or either argument is negative, returns 0.
'   Uses AfxFactorial for intermediate calculations.
' Examples:
'   AfxPerm(5, 2)  = 20
'   AfxPerm(6, 3)  = 120
'   AfxPerm(10, 0) = 1
'   AfxPerm(10, 10) = 3628800
' ========================================================================================
PRIVATE FUNCTION AfxPerm (BYVAL n AS LONG, BYVAL k AS LONG) AS ULONGINT
   IF n < 0 OR k < 0 OR k > n THEN RETURN 0
   RETURN AfxFactorial(n) \ AfxFactorial(n - k)
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxDivMod
' Performs integer division with floored quotient and remainder.
' Parameters:
'   a : dividend
'   b : divisor
'   q : quotient (returned by reference)
'   r : remainder (returned by reference)
' Behavior:
'   Computes q and r such that:
'       a = b * q + r
'   and r has the same sign as b (floored division).
'   This differs from the default FreeBasic operator "\" which truncates toward zero.
'   For example:
'       -7 \ 3  = -2   (truncated)
'       AfxDivMod(-7, 3) ? q = -3, r = 2   (floored)
' Examples:
'   AfxDivMod(7, 3, q, r)   -> q = 2,  r = 1
'   AfxDivMod(-7, 3, q, r)  -> q = -3, r = 2
'   AfxDivMod(7, -3, q, r)  -> q = -3, r = -2
'   AfxDivMod(-7, -3, q, r) -> q = 2,  r = -1
' ========================================================================================
PRIVATE SUB AfxDivMod (BYVAL a AS LONGINT, BYVAL b AS LONGINT, BYREF q AS LONGINT, BYREF r AS LONGINT)
   IF b = 0 THEN
      q = 0
      r = 0
      EXIT SUB
   END IF
   q = a \ b
   r = a MOD b
   ' Adjust for floored division when remainder sign differs from divisor
   IF (r <> 0) AND ((b > 0 AND r < 0) OR (b < 0 AND r > 0)) THEN
      q -= 1
      r += b
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' AfxIsPrime
' Determines whether an integer value is prime.
' Parameters:
'   n : integer value to test
' Behavior:
'   Returns TRUE if n is a prime number, otherwise FALSE.
'   Uses trial division up to the square root of n for efficiency.
'   Handles small cases explicitly (0, 1, 2).
' Examples:
'   AfxIsPrime(2)   = TRUE
'   AfxIsPrime(3)   = TRUE
'   AfxIsPrime(4)   = FALSE
'   AfxIsPrime(17)  = TRUE
'   AfxIsPrime(100) = FALSE
' ========================================================================================
PRIVATE FUNCTION AfxIsPrime (BYVAL n AS ULONGINT) AS BOOLEAN
   IF n < 2 THEN RETURN FALSE
   IF n = 2 THEN RETURN TRUE
   IF (n MOD 2) = 0 THEN RETURN FALSE
   DIM i AS ULONGINT = 3
   WHILE i * i <= n
      IF (n MOD i) = 0 THEN RETURN FALSE
      i += 2
   WEND
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxIsClose
' Determines whether two floating-point values are approximately equal.
' Parameters:
'   a       : first value
'   b       : second value
'   relTol  : relative tolerance (default = 1e-9)
'   absTol  : absolute tolerance (default = 0.0)
' Behavior:
'   Returns TRUE if the difference between a and b is within the specified tolerances.
'   The comparison follows the rule:
'       ABS(a - b) <= MAX(relTol * MAX(ABS(a), ABS(b)), absTol)
'   This method is consistent with IEEE and Python's math.isclose().
'   Useful for comparing results of floating-point calculations where rounding errors occur.
'
' Examples:
'   AfxIsClose(1.000000001, 1.0)          = FALSE
'   AfxIsClose(1000.0, 1000.0001)         = FALSE
'   AfxIsClose(1.0, 1.1)                  = FALSE
'   AfxIsClose(0.0, 1e-10, , 1e-9)        = TRUE
' ========================================================================================
PRIVATE FUNCTION AfxIsClose (BYVAL a AS DOUBLE, BYVAL b AS DOUBLE, BYVAL relTol AS DOUBLE = 1e-9, BYVAL absTol AS DOUBLE = 0.0) AS BOOLEAN
   RETURN ABS(a - b) <= MAX(relTol * MAX(ABS(a), ABS(b)), absTol)
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRnd
' Returns a pseudo-random number uniformly distributed in [0, 1).
' Parameters:
'   seed : Optional. If provided, reinitializes the random sequence.
'           - 0 -> use previous seed
'           - negative -> reseed with system timer
'           - positive -> set explicit seed value
' Behavior:
'   Wraps FreeBASIC's intrinsic Rnd() function.
'   When called without arguments, returns a random number in [0, 1).
' ========================================================================================
PRIVATE FUNCTION AfxRnd (BYVAL seed AS SINGLE = 1.0) AS DOUBLE
   RETURN RND(seed)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Generates a random integer within the specified range, inclusive of both the lower (lo) and upper (hi) bounds.
' If hi is less than lo, the function automatically swaps the values to ensure a valid range.
' This function uses the standard pseudo-random generator (RND) initialized by RANDOMIZE.
' Parameters:
'   lo : The lowest value that can be returned.
'   hi : The highest value that can be returned. Must be greater than or equal to lo.
' Return value:
'   Returns a DOUBLE representing a random floating-point number between lo and hi, inclusive.
' Example: Generate random integers between 1 and 10
'   RANDOMIZE TIMER   ' Initialize random seed
'   DIM i AS LONGINT
'   FOR n AS LONG = 1 TO 5
'      i = AfxRandomInt(1, 10)
'      PRINT "Random value #" & n & ": "; i
'   NEXT
' ========================================================================================
PRIVATE FUNCTION AfxRandomInt (BYVAL lo AS LONGINT, BYVAL hi AS LONGINT) AS LONGINT
   IF hi < lo THEN SWAP lo, hi
   RETURN lo + (RND * (hi - lo + 1))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Generates a random floating-point number within the specified range, inclusive of both
' the lower (lo) and upper (hi) bounds.
' If hi is less than lo, the function automatically swaps the values to ensure a valid range.
' This function uses the standard pseudo-random generator (RND) initialized by RANDOMIZE.
' Parameters:
'   lo : The lowest value that can be returned.
'   hi : The highest value that can be returned. Must be greater than or equal to lo.
' Return value:
'   Returns a DOUBLE representing a random floating-point number between lo and hi, inclusive.
' Example: Generate random floating-point numbers between 0.0 and 1.0
'   RANDOMIZE TIMER   ' Initialize random seed
'   DIM x AS DOUBLE
'   FOR n AS LONG = 1 TO 5
'      x = AfxRandomDouble(0.0, 1.0)
'      PRINT "Random value #" & n & ": "; x
'   NEXT
' ========================================================================================
PRIVATE FUNCTION AfxRandomDouble (BYVAL lo AS DOUBLE, BYVAL hi AS DOUBLE) AS DOUBLE
   IF hi < lo THEN SWAP lo, hi
   RETURN lo + (RND * (hi - lo))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Generates a random floating-point number following a Gaussian (normal) distribution
' using the Box-Muller transform. The result has the specified mean and standard deviation.
' Parameters:
'   mean   : The mean (average) value of the distribution. Default is 0.0.
'   stddev : The standard deviation (spread) of the distribution. Default is 1.0.
' Return value:
'   Returns a DOUBLE representing a random value distributed normally around mean,
'   with the specified standard deviation.
' Example: Generate 10 random numbers with mean 0 and standard deviation 1
'   RANDOMIZE TIMER   ' Initialize random seed
'   DIM x AS DOUBLE
'   FOR n AS LONG = 1 TO 10
'      x = AfxRandomGauss(0.0, 1.0)
'      PRINT "Gaussian value #" & n & ": "; x
'   NEXT
' ========================================================================================
PRIVATE FUNCTION AfxRandomGauss (BYVAL mean AS DOUBLE = 0.0, BYVAL stddev AS DOUBLE = 1.0) AS DOUBLE
   STATIC hasSpare AS LONG = 0
   STATIC spare AS DOUBLE
   DIM AS DOUBLE u, v, s
   IF hasSpare THEN
      hasSpare = 0
      RETURN mean + stddev * spare
   END IF
   DO
      u = (RND * 2.0) - 1.0
      v = (RND * 2.0) - 1.0
      s = u * u + v * v
   LOOP WHILE (s >= 1.0) OR (s = 0.0)
   s = SQR(-2.0 * LOG(s) / s)
   spare = v * s
   hasSpare = -1
   RETURN mean + stddev * (u * s)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shuffles the elements of an array in place using the Fisher-Yates algorithm.
' Each element has an equal probability of ending up in any position.
' Parameters:
'   pArray : Pointer to the first element of the array to shuffle.
'   nCount : Number of elements in the array.
'   cbElem : Size (in bytes) of each element.
' Example: Shuffle an array of integers
'   DIM a(1 TO 10) AS LONG
'   FOR i AS LONG = 1 TO 10
'      a(i) = i
'   NEXT
'   RANDOMIZE TIMER   ' Initialize random seed
'   AfxShuffle(@a(1), 10, SIZEOF(a(1)))
'   FOR i AS LONG = 1 TO 10
'      PRINT a(i);
'   NEXT
' ========================================================================================
PRIVATE SUB AfxShuffle (BYVAL pArray AS ANY PTR, BYVAL nCount AS LONG, BYVAL cbElem AS LONG)
   DIM i AS LONG, j AS LONG
   DIM temp AS UBYTE PTR
   DIM buffer AS UBYTE PTR = ALLOCATE(cbElem)
   IF pArray = 0 OR nCount < 2 OR cbElem < 1 THEN EXIT SUB
   FOR i = nCount - 1 TO 1 STEP -1
      j = AfxRandomInt(0, i)
      temp = pArray + (i * cbElem)
      ' Swap element i with element j
      memcpy(buffer, temp, cbElem)
      memcpy(temp, pArray + (j * cbElem), cbElem)
      memcpy(pArray + (j * cbElem), buffer, cbElem)
   NEXT
   Deallocate(buffer)
END SUB
' ========================================================================================

' ========================================================================================
' Fills a memory buffer with cryptographically secure random bytes using BCryptGenRandom.
' This function is suitable for generating keys, nonces, GUIDs, or any data that requires
' high-quality randomness. It does not depend on the standard RND generator.
' Parameters:
'   pOut : Pointer to the output buffer that will receive the random bytes.
'   cb   : Number of bytes to generate.
' Return value:
'   Returns TRUE if the operation succeeds, FALSE otherwise.
' Example: Generate 16 random bytes and display them in hexadecimal
'   DIM buffer(15) AS UBYTE
'   IF AfxRandomBytes(@buffer(0), 16) THEN
'      FOR i AS LONG = 0 TO 15
'         PRINT HEX(buffer(i), 2);
'      NEXT
'      PRINT
'   ELSE
'      PRINT "Error generating random bytes."
'   END IF
' ========================================================================================
PRIVATE FUNCTION AfxRandomBytes OVERLOAD (BYVAL pOut AS ANY PTR, BYVAL cb AS LONG) AS BOOLEAN
   IF pOut = 0 OR cb < 1 THEN RETURN FALSE
   DIM status AS LONG
   status = BCryptGenRandom(NULL, pOut, cb, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
   RETURN (status = 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' The default random number provider implements an algorithm for generating random numbers
' that complies with the NIST SP800-90 standard, specifically the CTR_DRBG portion of that standard.
' Usage example:
' DIM buffer AS STRING = AfxRandomBytes(16)
' IF LEN(buffer) THEN
'    FOR i AS LONG = 1TO 16
'       PRINT HEX(ASC(MID(buffer,i, 2))); " ";
'    NEXT
' END IF
' ========================================================================================
PRIVATE FUNCTION AfxRandomBytes OVERLOAD (BYVAL cb AS LONG) AS STRING
   IF cb < 1 THEN RETURN ""
   DIM buffer AS STRING = STRING(cb,0)
   DIM status AS LONG = BCryptGenRandom(NULL, STRPTR(buffer), cb, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
   IF status = 0 THEN RETURN buffer
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxMean
' Computes the arithmetic mean of an array of DOUBLE values.
' Parameters:
'   pData : pointer to the first element of the data array
'   n     : number of elements in the array
' Returns:
'   The average value as DOUBLE.
' Usage example:
'   DIM rgData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
'   PRINT AfxMean(@rgData(0), 5)
' Result: 30
' Example:
'   DIM p AS DOUBLE PTR = CALLocate(10, SIZEOF(DOUBLE))
'   FOR i AS LONG = 0 TO 9
'      p[i] = i + 1
'   NEXT
'   PRINT AfxMean(p, 10)
'   DEALLOCATE p
' Result: 5.5
' ========================================================================================
PRIVATE FUNCTION AfxMean (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG) AS DOUBLE
    IF pData = 0 OR n <= 0 THEN RETURN 0.0
    DIM sum AS DOUBLE = 0.0
    FOR i AS LONG = 0 TO n - 1
        sum += pData[i]
    NEXT
    RETURN sum / n
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxMedian
' Computes the median value of an array of DOUBLEs.
' Parameters:
'   pData : pointer to the first element of the data array
'   n     : number of elements in the array
' Returns:
'   The median value as DOUBLE.
' Usage example:
'   DIM rhData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
'   PRINT AfxMedian(@rhData(0), 5)
' Result: 30
' ========================================================================================
' Comparison function for qsort
PRIVATE FUNCTION AfxMedianCmpDbl(BYVAL a AS ANY PTR, BYVAL b AS ANY PTR) AS LONG
    DIM da AS DOUBLE = *CAST(DOUBLE PTR, a)
    DIM db AS DOUBLE = *CAST(DOUBLE PTR, b)
    IF da < db THEN RETURN -1
    IF da > db THEN RETURN 1
    RETURN 0
END FUNCTION

PRIVATE FUNCTION AfxMedian (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG) AS DOUBLE
   IF pData = 0 OR n <= 0 THEN RETURN 0.0
   ' Copy data to temporary array for sorting
   DIM tmp() AS DOUBLE
   REDIM tmp(n - 1)
   FOR i AS LONG = 0 TO n - 1
      tmp(i) = pData[i]
   NEXT
   ' Sort the temporary array
   qsort(@tmp(0), n, SIZEOF(DOUBLE), CAST(ANY PTR, @AfxMedianCmpDbl))
   ' Compute median
   IF (n MOD 2) = 1 THEN
      RETURN tmp(n \ 2)
   ELSE
      RETURN (tmp(n \ 2 - 1) + tmp(n \ 2)) / 2.0
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxVariance
' Computes the variance of an array of DOUBLE values.
' Parameters:
'   pData   : pointer to the first element of the data array
'   n       : number of elements in the array
'   fSample : TRUE for sample variance (divide by n - 1)
'             FALSE for population variance (divide by n)
' Returns:
'   The variance as DOUBLE.
' Usage example:
'   DIM rgData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
'   PRINT "Variance (sammple)   = "; AfxVariance(@rgData(0), 5, TRUE)
'   PRINT "Variance (population) = "; AfxVariance(@rgData(0), 5, FALSE)
' ========================================================================================
PRIVATE FUNCTION AfxVariance (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG, BYVAL fSample AS BOOLEAN = TRUE) AS DOUBLE
   IF pData = 0 OR n <= 1 THEN RETURN 0.0
   DIM mean AS DOUBLE = 0.0
   FOR i AS LONG = 0 TO n - 1
      mean += pData[i]
   NEXT
   mean /= n
   DIM sumsq AS DOUBLE = 0.0
   FOR i AS LONG = 0 TO n - 1
      DIM diff AS DOUBLE = pData[i] - mean
      sumsq += diff * diff
   NEXT
   IF fSample THEN
      RETURN sumsq / (n - 1)
   ELSE
      RETURN sumsq / n
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxStdDev
' Computes the standard deviation of an array of DOUBLE values.
'
' Parameters:
'   pData   : pointer to the first element of the data array
'   n       : number of elements in the array
'   fSample : TRUE for sample standard deviation (divide by n - 1)
'             FALSE for population standard deviation (divide by n)
'
' Returns:
'   The standard deviation as DOUBLE.
' Usage example:
'   DIM rgData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
'   PRINT "Standard variance (sample)   = "; AfxStdDev(@rgData(0), 5, TRUE)
'   PRINT "Standard variance (populationón) = "; AfxStdDev(@rgData(0), 5, FALSE)
' ========================================================================================
PRIVATE FUNCTION AfxStdDev (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG, BYVAL fSample AS BOOLEAN = TRUE) AS DOUBLE
   IF pData = 0 OR n <= 1 THEN RETURN 0.0
   DIM mean AS DOUBLE = 0.0
   FOR i AS LONG = 0 TO n - 1
      mean += pData[i]
   NEXT
   mean /= n
   DIM sumsq AS DOUBLE = 0.0
   FOR i AS LONG = 0 TO n - 1
      DIM diff AS DOUBLE = pData[i] - mean
      sumsq += diff * diff
   NEXT
   DIM variance AS DOUBLE
   IF fSample THEN
      variance = sumsq / (n - 1)
   ELSE
      variance = sumsq / n
   END IF
   RETURN SQR(variance)
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxSum
' Computes a compensated (Kahan) sum of an array of DOUBLE values.
' Parameters:
'   pData : pointer to the first element of the data array
'   n     : number of elements in the array
' Returns:
'   The precise sum as DOUBLE.
' Usage example:
'   DIM rgData(4) AS DOUBLE = { 1e16, 1.0, -1e16, 3.0 }
'   PRINT AfxSum(@rgData(0), 4)
' Return value: 3
' ========================================================================================
PRIVATE FUNCTION AfxSum (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG) AS DOUBLE
   IF pData = 0 OR n <= 0 THEN RETURN 0.0
   DIM sum AS DOUBLE = 0.0
   DIM c   AS DOUBLE = 0.0   ' compensation term
   FOR i AS LONG = 0 TO n - 1
      DIM y AS DOUBLE = pData[i] - c
      DIM t AS DOUBLE = sum + y
      c = (t - sum) - y
      sum = t
   NEXT
   RETURN sum
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxPopCount
' Counts the number of set bits (1s) in a 64-bit unsigned integer.
' Parameters:
'   x : 64-bit unsigned integer
' Returns:
'   Number of bits set to 1.
' Usage example:
'   PRINT AfxPopCount(&hF0F0F0F0F0F0F0F0)   ' returns 32
'   PRINT AfxPopCount(255)                  ' returns 8
'   PRINT AfxPopCount(0)                    ' returns 0
' ========================================================================================
PRIVATE FUNCTION AfxPopCount (BYVAL x AS ULONGINT) AS LONG
   DIM count AS LONG = 0
   WHILE x <> 0
      x = x AND (x - 1)   ' clear lowest set bit
      count += 1
   WEND
   RETURN count
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxLeadingZeros
' Counts the number of leading zero bits in a 64-bit unsigned integer.
' Parameters:
'   x : 64-bit unsigned integer
' Returns:
'   Number of leading zeros (0–64).
' Usage example:
'   PRINT AfxLeadingZeros(&h0000000000000001ULL)   ' returns 63
'   PRINT AfxLeadingZeros(&h8000000000000000ULL)   ' returns 0
'   PRINT AfxLeadingZeros(&h00F0000000000000ULL)   ' returns 8
' ========================================================================================
PRIVATE FUNCTION AfxLeadingZeros (BYVAL x AS ULONGINT) AS LONG
   IF x = 0 THEN RETURN 64
   DIM count AS LONG = 0
   DIM mask AS ULONGINT = &h8000000000000000ULL
   WHILE (x AND mask) = 0
      count += 1
      mask = mask SHR 1
   WEND
   RETURN count
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxTrailingZeros
' Counts the number of trailing zero bits in a 64-bit unsigned integer.
' Parameters:
'   x : 64-bit unsigned integer
' Returns:
'   Number of trailing zeros (0–64).
' Usage example:
'   PRINT AfxTrailingZeros(&h0000000000000001ULL)   ' returns 0
'   PRINT AfxTrailingZeros(&h8000000000000000ULL)   ' returns 63
'   PRINT AfxTrailingZeros(&h00F0000000000000ULL)   ' returns 52
' ========================================================================================
PRIVATE FUNCTION AfxTrailingZeros (BYVAL x AS ULONGINT) AS LONG
   IF x = 0 THEN RETURN 64
   DIM count AS LONG = 0
   WHILE (x AND 1ULL) = 0
      count += 1
      x = x SHR 1
   WEND
   RETURN count
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRotateLeft
' Performs a 32-bit left rotation of an unsigned integer.
' Parameters:
'   x : 32-bit unsigned integer
'   n : number of bit positions to rotate (0–31)
' Returns:
'   The rotated value.
' Usage example:
'   PRINT HEX(AfxRotateLeft(&h12345678, 8))   ' retuns 34567812
'   PRINT HEX(AfxRotateLeft(&h80000000, 1))   ' retuns 1
' ========================================================================================
PRIVATE FUNCTION AfxRotateLeft (BYVAL x AS ULONG, BYVAL n AS LONG) AS ULONG
    n = n AND 31
    RETURN (x SHL n) OR (x SHR (32 - n))
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxRotateRight
' Performs a 32-bit right rotation of an unsigned integer.
' Parameters:
'   x : 32-bit unsigned integer
'   n : number of bit positions to rotate (0–31)
' Returns:
'   The rotated value.
' Usage example:
'   PRINT HEX(AfxRotateRight(&h12345678, 8))   ' returns 78123456
'   PRINT HEX(AfxRotateRight(&h00000001, 1))   ' returns 80000000
' ========================================================================================
PRIVATE FUNCTION AfxRotateRight (BYVAL x AS ULONG, BYVAL n AS LONG) AS ULONG
    n = n AND 31
    RETURN (x SHR n) OR (x SHL (32 - n))
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxByteSwap
' Swaps the byte order of a 64-bit unsigned integer (endian flip).
' Parameters:
'   x : 64-bit unsigned integer
' Returns:
'   The value with reversed byte order.
' Usage example:
'   PRINT HEX(AfxByteSwap(&h0123456789ABCDEFULL))   ' returns EFCDAB8967452301
' ========================================================================================
PRIVATE FUNCTION AfxByteSwap (BYVAL x AS ULONGINT) AS ULONGINT
    RETURN ((x AND &h00000000000000FFULL) SHL 56) OR _
           ((x AND &h000000000000FF00ULL) SHL 40) OR _
           ((x AND &h0000000000FF0000ULL) SHL 24) OR _
           ((x AND &h00000000FF000000ULL) SHL 8)  OR _
           ((x AND &h000000FF00000000ULL) SHR 8)  OR _
           ((x AND &h0000FF0000000000ULL) SHR 24) OR _
           ((x AND &h00FF000000000000ULL) SHR 40) OR _
           ((x AND &hFF00000000000000ULL) SHR 56)
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxChangeSign
' Inverts the sign of a DOUBLE value (like MSVC _chgsign).
' Parameters:
'   x : floating-point value
' Returns:
'   Same value with inverted sign.
' Usage example:
'   PRINT AfxChangeSign(5.0)     ' -5
'   PRINT AfxChangeSign(-12.75)  ' 12.75
'   PRINT AfxChangeSign(0.0)     ' -0
' ========================================================================================
#define AfxChangeSign(x) (-(x))
' ========================================================================================

' ========================================================================================
' AfxFdim
' Returns the positive difference between x and y (like C fdim).
' Equivalent to:
'   (x > y) ? (x - y) : 0
' Usage examples:
'   PRINT AfxFdim(5.3, 2.0)    ' 3.3
'   PRINT AfxFdim(2.0, 5.3)    ' 0
'   PRINT AfxFdim(-1.0, -3.0)  ' 2
'   PRINT AfxFdim(-3.0, -1.0)  ' 0
' ========================================================================================
PRIVATE FUNCTION AfxFdim (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS DOUBLE
   IF x > y THEN RETURN x - y ELSE RETURN 0.0
END FUNCTION
' ========================================================================================

' ========================================================================================
' AfxFma
' Performs a fused multiply-add: (x * y + z) with a single rounding,
' matching the behavior of C's fma() as documented by Microsoft.
' Usage examples:
'   PRINT AfxFma(2.0, 3.0, 4.0)         ' 10
'   PRINT AfxFma(1e308, 1e-308, 1.0)    ' precise fused result
'   PRINT AfxFma(0.0, 5.0, 7.0)         ' 7
'   PRINT AfxFma(1/0.0, 0.0, 2.0)       ' NaN (inf * 0)
' ========================================================================================
PRIVATE FUNCTION AfxFma (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
   '   Handle NaNs: return one of the NaN inputs
   IF x <> x THEN RETURN x
   IF y <> y THEN RETURN y
   IF z <> z THEN RETURN z
   '   IEEE-754: inf * 0 ? NaN
   IF ABS(x) = 1/0.0 AND y = 0.0 THEN RETURN 0.0 / 0.0
   IF ABS(y) = 1/0.0 AND x = 0.0 THEN RETURN 0.0 / 0.0
   ' General case: FreeBASIC evaluates double expressions with extended precision,
   ' so (x * y + z) is effectively fused with a single final rounding.
   DIM prod AS DOUBLE
   prod = x * y
   RETURN prod + z
END FUNCTION
'   ========================================================================================

' ========================================================================================
' AfxFma
' Performs a fused multiply-add operation: (x * y) + z with a single rounding.
' Parameters:
'   x : multiplicand
'   y : multiplier
'   z : addend
' Behavior:
'   Returns the result of (x * y + z) computed as one fused operation.
'   This minimizes rounding error compared to separate multiplication and addition.
'   If hardware FMA is not available, falls back to standard arithmetic.
' Examples:
'   AfxFma(2.0, 3.0, 4.0)  = 10.0
'   AfxFma(1.5, 2.0, -0.5) = 2.5
' ========================================================================================
'PRIVATE FUNCTION AfxFma (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
'   ' If the compiler or platform supports FMA intrinsic, use it.
'   ' Otherwise, perform the equivalent operation.
'   RETURN x * y + z
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a value indicating the floating-point classification of the argument.
' Usage examples:
'   PRINT AfxFpClassify(0.0)          ' _FPCLASS_PZ
'   PRINT AfxFpClassify(-0.0)         ' _FPCLASS_NZ
'   PRINT AfxFpClassify(1.0)          ' _FPCLASS_PN
'   PRINT AfxFpClassify(-3.14)        ' _FPCLASS_NN
'   PRINT AfxFpClassify(1/0.0)        ' _FPCLASS_PINF
'   PRINT AfxFpClassify(-1/0.0)       ' _FPCLASS_NINF
'   PRINT AfxFpClassify(0.0/0.0)      ' _FPCLASS_QNAN
' ========================================================================================
#define AfxFpClassify(x) (_fpclass(x))
' ========================================================================================

' ========================================================================================
' Determines the natural logarithm of the absolute value of the gamma function of the specified value.
' Usage examples:
'   PRINT AfxLgamma(5.0)        ' ln(|G(5)|) = ln(24) ˜ 3.17805383
'   PRINT AfxLgamma(0.5)        ' ln(|G(0.5)|) = ln(sqrt(pi))
'   PRINT AfxLgamma(-3.2)       ' +INFINITY for negative integer arguments or poles
' ========================================================================================
#define AfxLgamma(x) (lgamma(x))
' ========================================================================================

' ========================================================================================
' Determines the gamma function of the specified value.
' Usage examples:
'   PRINT AfxTgamma(5.0)        ' G(5) = 24
'   PRINT AfxTgamma(0.5)        ' G(0.5) = sqrt(pi)
'   PRINT AfxTgamma(-3.2)       ' finite value (gamma has poles only at negative integers)
'  PRINT AfxTgamma(-2.0)       ' +INFINITY (pole at negative integer)
' ========================================================================================
#define AfxTgamma(x) (tgamma(x))
' ========================================================================================

' ========================================================================================
' Computes the error function of a value.
' Usage examples:
'   PRINT AfxErf(0.0)          ' 0
'   PRINT AfxErf(0.5)          ' 0.5204998778130465
'   PRINT AfxErf(1.0)          ' 0.8427007929497149
'   PRINT AfxErf(-1.0)         ' -0.8427007929497149
'   PRINT AfxErf(2.0)          ' 0.9953222650189527
'   PRINT AfxErf(-2.0)         ' -0.9953222650189527
' ========================================================================================
#define AfxErf(x) (erf(x))
' ========================================================================================

' ========================================================================================
' Computes the complementary error function of a value.
' Usage examples:
'   PRINT AfxErfc(0.0)         ' 1
'   PRINT AfxErfc(0.5)         ' 0.4795001221869535
'   PRINT AfxErfc(1.0)         ' 0.1572992070502851
'   PRINT AfxErfc(2.0)         ' 0.004677734981047265
'   PRINT AfxErfc(10.0)        ' 2.088487583762545e-045
' ========================================================================================
#define AfxErfc(x) (erfc(x))
' ========================================================================================

' ========================================================================================
' Calculates the floating-point remainder.
' Parameters:
' x, y: Floating-point values.
' Returns the floating-point remainder of x / y. If the value of y is 0.0, fmod returns a quiet NaN.
' Usage examples:
'   PRINT AfxMod(-5.3, 2.0)    ' -1.3
'   PRINT AfxMod(5.3, -2.0)    ' 1.3
'   PRINT AfxMod(-5.3, -2.0)   ' -1.3
'   PRINT AfxMod(1.0, 0.0)     ' NaN
'   PRINT AfxMod(0.0, 0.0)     ' NaN
'   PRINT AfxMod(1.0/0.0, 2.0) ' NaN
'   PRINT AfxMod(5.0, 1.0/0.0) ' 5.0
' ========================================================================================
#define AfxMod(x, y) (fmod((x), (y)))
' ========================================================================================

' ========================================================================================
' Splits a floating-point value into fractional and integer parts.
' Returns the fractional part of dblX and stores the integer part in dblIntPtr.
' Usage example:
'   DIM AS DOUBLE dblInteger
'   DIM AS DOUBLE dblFraction
'   dblFraction = AfxModf(-14.87654321, @dblInteger)
'   PRINT "Fraction = "; AfxRoundHalfEven(dblFraction, 8)
'   PRINT "Integer  = "; dblInteger
' ========================================================================================
#define AfxModf(dblX, dblIntPtr) (modf(dblX, dblIntPtr))
' ========================================================================================

' ========================================================================================
' Gets the mantissa and exponent of a floating-point number.
' Parameters:
'   x : Floating-point value.
'   e : A long variable that receives the exponent.
' Return value:
'   Returns the mantissa. If x is 0, the function returns 0 for both the mantissa and the exponent.
' Usage example:
'   DIM value AS DOUBLE = 12345.67
'   DIM e AS LONG
'   DIM mantissa AS DOUBLE
'   mantissa = AfxFrexp(value, e)
'   PRINT mantissa, e
' ========================================================================================
PRIVATE FUNCTION AfxFrexp (BYVAL x AS DOUBLE, BYREF e AS LONG) AS DOUBLE
   RETURN frexp(x, @e)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the exponent of a floating-point number.
' Parameter:
' x  : Floating-point value.
' Return value: The exponent.
' Usage example:
'   DIM value AS DOUBLE = 12345.67
'   DIM e AS LONG = AfxGetExponent(value)
' ========================================================================================
PRIVATE FUNCTION AfxGetExponent (BYVAL x AS DOUBLE) AS LONG
   DIM e AS LONG
   frexp(x, @e)
   RETURN e
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the mantissa of a floating-point number.
' Parameter:
' x  : Floating-point value.
' Return value: The mantissa.
' Usage example:
'   DIM value AS DOUBLE = 12345.67
'   DIM e AS LONG = AfxGetMantissa(value)
' ========================================================================================
PRIVATE FUNCTION AfxGetMantissa (BYVAL x AS DOUBLE) AS DOUBLE
   DIM e AS LONG
   RETURN frexp(x, @e)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Multiplies a floating-point number by an integral power of two.
' Parameters:
'   x : Floating-point value.
'   e : Integer exponent
' Return value:
'   Returns the value of x * 2exp if successful.
'   On overflow, and depending on the sign of x, ldexp returns +/- HUGE_VAL
' Usage example:
'   DIM value AS DOUBLE = 12345.67
'   DIM e AS LONG = 10
'   PRINT AfxLdexp(value, e)
' ========================================================================================
#define AfxLdexp(x, e) ldexp(x, e)
' ========================================================================================

' ========================================================================================
' Multiplies a floating-point number by 2 raised to the power of n.
' Equivalent to scalbn() in C/C++.
' Useful for efficient scaling by powers of two.
' Parameters
'   dX : Input floating-point value.
'   e : Exponent (integer power of 2).
' Return Value
'   dX * 2^e
' Usage example:
'   DIM AS DOUBLE dX = 3.5
'   DIM AS LONG   e = 4
'   DIM AS DOUBLE dResult
'   dResult = AfxScalb(dX, e)
'   PRINT dX; " * 2^"; e; " = "; dResult
' Expected output: 56
' ========================================================================================
#define AfxScalb(dx, e) _scalb(dx, e)
#define AfxScalbn(x, e) scalbn(x, e)
' ========================================================================================

' ========================================================================================
' Returns the next representable floating-point value.
' Parameters:
'   x : The floating-point value to start from.
'   y : The floating-point value to go towards.
' Returns the next representable floating-point value of the return type after x in the direction of y.
' ========================================================================================
#define AfxNextAfter(x, y) nextafter(x, y)
' ========================================================================================

' ========================================================================================
' Computes the Bessel function of the first kind, order 0.
' The Bessel functions are commonly used in the mathematics of electromagnetic wave theory.
' The _j0, _j1, and _jn routines return Bessel functions of the first kind: orders 0, 1, and n, respectively.
' The _y0, _y1, and _yn routines return Bessel functions of the second kind: orders 0, 1, and n, respectively.
' Parameters:
'   x : The floating-point value.
'   n : Integer order of Bessel function.
' Return value:
' Each of these routines returns a Bessel function of x. If x is negative in the _y0, _y1,
' or _yn functions, the routine sets errno to EDOM, prints a _DOMAIN error message to stderr,
' and returns HUGE_VAL. You can modify error handling by using _matherr.
' ========================================================================================
#define AfxBesselJ0(x) _j0(x)
#define AfxBesselJ1(x) _j1(x)
#define AfxBesselJn(n, x) _jn(n, x)
#define AfxBesselY0(x) _y0(x)
#define AfxBesselY1(x) _y1(x)
#define AfxBesselYn(n, x) _yn(n, x)
' ========================================================================================

' ========================================================================================
' Derived math macros
' ========================================================================================
' Secant
#define AfxSec(x) (1 / Cos(x))
' Cosecant
#define AfxCsc(x) (1 / Sin(x))
' Cotangent
#define AfxCtan(x) (1 / Tan(x))
' Inverse secant
#define AfxAsec(x) (2 * Atn(1) - Atn(Sgn(x) / Sqrt(x * x - 1)))
' Inverse cosecant
#define AfxAcsc(x) (Atn(Sgn(x) / Sqrt(x * x - 1)))
' Inverse cotangent
#define AfxAcot(x) (2 * Atn(1) - Atn(x))
' Hyperbolic secant
#define AfxSech(x) (2 / (Exp(x) + Exp(-x)))
' Hyperbolic cosecant
#define AfxCsch(x) (2 / (Exp(x) - Exp(-x)))
' Hyperbolic cotangent
#define AfxCoth(x) (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
' Inverse hyperbolic secant
#define AfxAsech(x) (Log((Sqrt(-x * x + 1) + 1) / x))
' Inverse hyperbolic cosecant
#define AfxAcsch(x) (Log((Sgn(x) * Sqrt(x * x + 1) + 1) / x))
' Inverse hyperbolic cotangent
#define AfxAcoth(x) (Log((x + 1) / (x - 1)) / 2)
' ========================================================================================

' =====================================================================================
' Returns a string with the result of the numerical expression formatted as indicated in
' the formatting expression. The formatting expression is a string that can yield numeric
' or date-time values.
' The representation of time and date separators may be determined from the operating
' system if regional or localized settings are supported by the target platform. This
' determination is made at run-time so that output may be localized to the system the
' program is running on.
' See the focumentation of Format in the FreeBasic help file.
' Usage example:
'   PRINT AfxFormat(12345.67, "#,##0.00")
' Example:
'   DIM AS DOUBLE numberVal(...) = {5, -5, .5}
'   DIM AS STRING formatStr(...) = {"","0","0.00","#,##0","#,##0.00","0%", "0.00%", "0.00E+00", "0.00E-00"}
'   PRINT "Format string",, STR(numberVal(0)), STR(numberVal(1)), STR(numberVal(2))
'   FOR iFormat AS INTEGER = 0 TO UBOUND(formatStr)
'      PRINT formatStr(iFormat),,
'      FOR iNumber AS INTEGER = 0 TO UBOUND(numberVal)
'         PRINT AfxFormat(numberVal(iNumber), formatStr(iFormat)),
'      NEXT
'      PRINT
'   NEXT
' =====================================================================================
PRIVATE FUNCTION AfxFormat (BYVAL value AS DOUBLE, BYREF mask AS CONST STRING = "") AS STRING
   RETURN Format(value, mask)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Formats a currency number into a string form.
' Parameters:
' - value: The numeric value to format.
' - iNumDig: Digits after decimal (-1 = system default)
' - iIncLead: Leading zero (-2 = default, -1 = include, 0 = no)
' - iUseParens: Negative parentheses (-2 = default, -1 = yes, 0 = no)
' - iGroup: Group thousands (-2 = default, -1 = yes, 0 = no)
' - dwFlags: VAR_CALENDAR_HIJRI only
' Usage example:
'   DIM dbl AS DOUBLE = 12345.1234
'   PRINT AfxFormatCurrency(dbl)  ' --> 12.345,12 € (Spain)
' =====================================================================================
PRIVATE FUNCTION AfxFormatCurrency (BYVAL value AS DOUBLE, BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DIM v AS VARIANT
   v.vt = VT_R8
   v.dblVal = value
   DIM bstrOut AS BSTR
   VarFormatCurrency(@v, iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut)
   DIM dwsOut AS DWSTRING = bstrOut
   SysFreeString(bstrOut)
   RETURN dwsOut
END FUNCTION
' =====================================================================================

' =====================================================================================
' Formats a number into a string form (no currency symbol).
' Parameters identical to AfxFormatCurrency.
' Usage example:
'   DIM dbl AS DOUBLE = 12345.1234
'   PRINT AfxFormatNumber(dbl)  ' --> 12.345,12 (Spain)
' =====================================================================================
PRIVATE FUNCTION AfxFormatNumber (BYVAL value AS DOUBLE, BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DIM v AS VARIANT
   v.vt = VT_R8
   v.dblVal = value
   DIM bstrOut AS BSTR
   VarFormatNumber(@v, iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut)
   DIM dwsOut AS DWSTRING = bstrOut
   SysFreeString(bstrOut)
   RETURN dwsOut
END FUNCTION
' =====================================================================================

' =====================================================================================
' Formats a percentage number into a string form.
' Parameters:
' - value: The numeric value to format (e.g., 0.25 -> "25 %")
' - iNumDig: Digits after decimal (-1 = system default)
' - iIncLead: Leading zero (-2 = default, -1 = include, 0 = no)
' - iUseParens: Negative parentheses (-2 = default, -1 = yes, 0 = no)
' - iGroup: Group thousands (-2 = default, -1 = yes, 0 = no)
' - dwFlags: VAR_CALENDAR_HIJRI only
' Usage example:
'   PRINT AfxFormatPercent(0.25)  ' --> 25,00% (Spain)
' =====================================================================================
PRIVATE FUNCTION AfxFormatPercent (BYVAL value AS DOUBLE, BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DIM v AS VARIANT
   v.vt = VT_R8
   v.dblVal = value
   DIM bstrOut AS BSTR
   VarFormatPercent(@v, iNumDig, iIncLead, iUseParens, iGroup, dwFlags, @bstrOut)
   DIM dwsOut AS DWSTRING = bstrOut
   SysFreeString(bstrOut)
   RETURN dwsOut
END FUNCTION
' =====================================================================================
