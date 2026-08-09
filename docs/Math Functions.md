#### Mathematical functions

The `AfxMath` module provides a comprehensive collection of mathematical routines designed to offer consistent, CRT‑compatible behavior across all supported Windows platforms. These functions wrap the standard C runtime (CRT) math library and expose a unified, FreeBasic‑friendly interface that mirrors the naming and semantics of the original C functions while integrating cleanly into the `AfxNova` framework.

The goal of this module is not to replace the CRT, but to standardize access, improve discoverability, and ensure predictable behavior regardless of compiler settings or platform variations. Every function in `AfxMath` is a thin, reliable wrapper around the underlying CRT implementation, preserving IEEE‑754 semantics and matching the behavior of the corresponding C function as closely as possible.

It also extends the CRT with additional routines that are not available in the Microsoft C runtime, including IEEE‑754 helpers, secure random number generation, and several utility functions designed specifically for FreeBasic and `AfxNova`. These additions fill long‑standing gaps in the CRT and provide functionality commonly expected in modern numerical libraries.

Although most applications will only use a small subset of these routines, having the full set available ensures that `AfxNova` can support numerical code ranging from simple calculations to advanced scientific and engineering tasks. The module is lightweight, dependency‑free, and suitable for both high‑level and low‑level numerical work.

## CRT error codes

These are the possible error values returned by **errno**. In C and C++, **errno** is a macro-based global or thread-local variable that stores the error code number for the most recently failed C Runtime Library (CRT) or system function call.

| Error | Value | Description |
| ------| ----- | ----------- |
| EPERM | 1 | Operation not permitted |
| ENOFILE | 2 | no such file or directory |
| ENOFILE | 2 | Error no entity |
| ESRCH | 3 | No such process |
| EINTR | 4 | Interrupted function call |
| EIO | 5 | Input/output error |
| ENXIO | 6 | No such device or address |
| E2BIG | 7 | Argument list too long |
| ENOEXEC | 8 | Exec format error |
| EBADF | 9 | Bad file descriptor |
| ECHILD | 10 | No child processes |
| EAGAIN | 11 | Resource temporarily unavailablem try again |
| ENOMEM | 12 | Out of memory |
| EACCES | 13 | Permission denied |
| EFAULT | 14 | Bad address |
| EBUSY | 16 | Device or resource busy |
| EEXIST | 17 | File exists |
| EXDEV | 18 | Cross-device link error |
| ENODEV | 19 | No such device |
| ENOTDIR | 20 | Not a directory |
| EISDIR | 21 | Illegal operation on a directory |
| EINVAL | 22 | Invalid argument |
| ENFILE | 23 | File table overflow |
| EMFILE | 24 | Too many open files |
| ENOTTY | 25 | Inappropriate ioctl for device |
| EFBIG | 27 | File too large |
| ENOSPC | 28 | No space left on device |
| ESPIPE | 29 | Invalid seek |
| EROFS | 30 | Read-only file system |
| EMLINK | 31 | Too many links |
| EPIPE | 32 | Broken pipe |
| EDOM | 33 | Domain error |
| ERANGE | 34 | Out of range |
| EDEADLOCK | 36 | Resource deadlock would occur |
| EDEADLK | 36 | Resource deadlock would occur |
| ENAMETOOLONG | 38 | Filename too long |
| ENOLCK | 39 | No locks available |
| ENOSYS | 40 | Function not implemented |
| ENOTEMPTY | 41 | Directory not empty |
| EILSEQ | 42 | Illegal byte sequence |

---

| Function   | Description |
| ---------- | ----------- |
| [AfxAbs](#afxabs) | Calculates the absolute value of a number. |
| [AfxAcos](#afxacos) | Calculates the arccosine in radians. |
| [AfxAcosh](#afxacosh) | Calculates the inverse hyperbolic cosine. |
| [AfxAcot](#afxacot) | Calculates the inverse cotangent im radians. |
| [AfxAcoth](#afxacoth) | Calculates the inverse hyperbolic cotangent |
| [AfxAcsc](#afxacsc) | Calculates the inverse cosecant in radians. |
| [AfxACsch](#afxacsch) | Calculates the inverse hyperbolic cosecant |
| [AfxAsec](#afxasec) | Calculates the inverse secant in radians. |
| [AfxAsech](#afxasech) | Calculates the Inverse hyperbolic secant in radians. |
| [AfxAsin](#afxasin) | Calculates the arc sine (inverse sine) of a number. |
| [AfxAsinh](#afxasinh) | Calculates the inverse hyperbolic sine. |
| [AfxAtan](#afxatan) | Calculates the arc tangent (inverse tangent) of a number. |
| [AfxAtan2](#afxatan2) | Calculates the arc tangent of y/x using the signs of both arguments to determine the correct quadrant of the result. |
| [AfxAtanh](#afxatanh) | Returns the inverse hyperbolic tangent of a number. |
| [AfxBessel](#afxbessel) | Computes the Bessel function. |
| [AfxByteSwap](#afxbyteswap) | Swaps the byte order of a 64-bit unsigned integer (endian flip). |
| [AfxCbrt](#afxcbrt) | Returns the cube root of a number. |
| [AfxChangeSign](#afxchangesign) | Inverts the sign of a double value. |
| [AfxCeil](#afxceil) | Returns the smallest integer greater than or equal to x. |
| [AfxClamp](#afxclamp) | Restricts a value to remain within a specified range. |
| [AfxComb](#afxcomb) | Computes the number of combinations (n choose k). |
| [AfxCopySign](#afxcopysign) | Returns a value that has the magnitude of one argument and the sign of another. |
| [AfxCos](#afxcos) | Calculates the cosine of an angle in radians. |
| [AfxCosh](#afxcosh) | Calculates the hyperbolic cosine of a number. |
| [AfxCoth](#afxcoth) | Calculates the hyperbolic cotagent. |
| [AfxCsc](#afxcsc) | Cosecant. Calculates the reciprocal of the sine function. |
| [AfxCsch](#afxcsch) | Calculates the hyperbolic cosecant |
| [AfxCtan](#afxctan) | Cotangent. Calculates the reciprocal of the tangent function. |
| [AfxDegToRad](#afxdegtorad) | Converts degrees to radians. |
| [AfxDivMod](#afxdivmod) | Performs integer division with floored quotient and remainder. |
| [AfxEaseInOut](#afxeaseinout) | Smooth ease-in/ease-out interpolation using cubic Hermite curve. |
| [AfxErf](#afxerf) | Computes the error function of a value. |
| [AfxErfc](#afxerfc) | Computes the complementary error function of a value. |
| [AfxEuroRound](#afxeuroround) | European-style rounding: rounds halfway values away from zero. |
| [AfxExp](#afxexp) | Calculates the exponential value of a number. |
| [AfxExp2](#afxexp2) | Calculates 2 raised to the specified power. |
| [AfxExpm1](#afxexpm1) | Calculates e raised to the power of x minus 1. |
| [AfxExp2m1](#afxexp2m1) | Calculates 2 raised to the power of x minus 1. |
| [AfxFactorial](#afxfactorial) | Computes the factorial of an integer value. |
| [AfxFdim](#afxfdim) | Returns the positive difference between x and y (like C fdim). |
| [AfxFix](#afxfix) | Returns the integer part of a number, truncating toward zero. |
| [AfxFloor](#afxfloor) | Returns the largest integer less than or equal to x. |
| [AfxFma](#afxfma) | Performs a fused multiply-add: (x * y + z) with a single rounding, |
| [AfxFormat](#afxformat) | Returns a string with the result of the numerical expression formatted as indicated in the formatting expression. |
| [AfxFormatCurrency](#afxformatcurrency) | Formats a currency number into a string form. |
| [AfxFormatNumber](#afxformatnumber) | Formats a number into a string form (no currency symbol). |
| [AfxFormatPercent](#afxformatpercent) | Formats a percentage number into a string form. |
| [AfxFpClassify](#afxfpclassify) | Returns a value indicating the floating-point classification of the argument. |
| [AfxFrac](#afxfrac) | Returns the fractional part of a number. |
| [AfxFrexp](#afxfrexp) | Gets the mantissa and exponent of a floating-point number. |
| [AfxFusedMultiplyAdd](#afxfusedmultiplyadd) | Computes (x * y) + z with IEEE 754 semantics. |
| [AfxGcd](#afxgcd) | Computes the greatest common divisor (GCD) of two integers using the Euclidean algorithm. |
| [AfxGetExponent](#afxgetexponent) | Gets the exponent of a floating-point number. |
| [AfxGetMantissa](#afxgetmantissa) | Gets the mantissa of a floating-point number. |
| [AfxHypot](#afxhypot) | Calculates the Euclidean distance (hypotenuse) from two values. |
| [AfxHypot3](#afxhypot3) | Computes the value of sqr(x^2 + y^2 + z^2). |
| [AfxIEEERemainder](#afxremainder) | Returns the IEEE 754 remainder of x / y. |
| [AfxILogB](#afxilogb) | Returns the unbiased exponent of the floating-point number x. |
| [AfxInt](#afxint) | Returns the largest integer less than or equal to x. |
| [AfxInverseLerp](#afxinverselerp) | Computes the interpolation factor t given a, b and v. |
| [AfxIsClose](#afxisclose) | Determines whether two floating-point values are approximately equal. |
| [AfxIsFinite](#afxisfinite) | Determines whether the argument is a finite floating-point value. |
| [AfxIsInfinity](#afxisinfinity) | Determines whether the argument is an infinity. |
| [AfxIsNaN](#afxisnan) | Determines whether the argument is a NaN (Not a Number). |
| [AfxIsNormal](#afxisnormal) | Returns TRUE if x is a normal IEEE 754 double-precision value. |
| [AfxIsPrime](#afxisprime) | Determines whether an integer value is prime. |
| [AfxISqrt](#afxisqrt) | Calculates the integer square root of a number. |
| [AfxIsSubnormal](#afxissubnormal) | Returns TRUE if x is a subnormal (denormalized) floating-point number. |
| [AfxLcm](#afxlcm) | Computes the least common multiple (LCM) of two integers. |
| [AfxLdexp](#afxldexp) | Multiplies a floating-point number by an integral power of two. |
| [AfxLeadingZeros](#afxleadingzeros) | Counts the number of leading zero bits in a 64-bit unsigned integer. |
| [AfxLerp](#afxlerp) | Linearly interpolates between a and b by t. |
| [AfxLgamma](#afxlgamma) | Determines the natural logarithm of the absolute value of the gamma function of the specified value. |
| [AfxLog](#afxlog) | Calculates the natural logarithm (base e) of x. |
| [AfxLog2](#afxlog2) | Calculates the base-2 logarithm of x. |
| [AfxLog10](#afxlog10) | Calculates the base-2 logarithm of x. |
| [AfxLog1p](#afxlog1p) | Calculates Log(1 + x) with improved precision for small x values. |
| [AfxLogB](#afxlogb) | Returns the unbiased exponent of a floating-point number as a double. |
| [AfxLogBase](#afxlogbase) | Calculates the logarithm of x in an arbitrary base. |
| [AfxMax](#afxmax) | Returns the greatest of two values. |
| [AfxMaxMagnitude](#afxmaxmagnitude) | Returns the value (x or y) with the largest magnitude. |
| [AfxMean](#afxmean) | Computes the arithmetic mean of an array of DOUBLE values. |
| [AfxMedian](#afxmedian) | Computes the median value of an array of DOUBLEs. |
| [AfxMidPoint](#afxmidpoint) | Calculates the midpoint between two integer values. |
| [AfxMin](#afxmin) | Returns the smaller of two values. |
| [AfxMinMagnitude](#afxminmagnitude) | Returns the value (x or y) with the smallest magnitude. |
| [AfxMod](#afxmod) | Calculates the floating-point remainder. |
| [AfxModf](#afxmodf) | Splits a floating-point value into fractional and integer parts. |
| [AfxNextAfter](#afxnextafter) | Returns the next representable floating-point value. |
| [AfxPerm](#afxperm) | Computes the number of permutations (n permute k). |
| [AfxPopCount](#afxpopcount) | Counts the number of set bits (1s) in a 64-bit unsigned integer. |
| [AfxPow](#afxpow) | Raises a number to a specified power. |
| [AfxPowMod](#afxpowmod) | Performs exponentiation under a modulus in an overflow-safe way. |
| [AfxRadToDeg](#afxradtodeg) | Converts radians to degrees. |
| [AfxRandomBytes](#afxrandombytes) | Fills a memory buffer with cryptographically secure random bytes using BCryptGenRandom. |
| [AfxRandomDouble](#afxrandomdouble) | Generates a random integer within the specified range. |
| [AfxRandomGauss](#afxrandomgauss) | Generates a random floating-point number following a Gaussian (normal) distribution using the Box-Muller transform. |
| [AfxRandomInt](#afxrandomint) | Generates a random integer within the specified range. |
| [AfxRemainder](#afxremainder) | Returns the remainder of x / y. |
| [AfxRemap](#afxremap) | Remaps a value from one range to another. |
| [AfxRnd](#afxrnd) | Returns a pseudo-random number uniformly distributed in [0, 1). |
| [AfxRotateLeft](#afxrotateleft) | Performs a 32-bit left rotation of an unsigned integer. |
| [AfxRotateRight](#afxrotateright) | Performs a 32-bit right rotation of an unsigned integer. |
| [AfxRound](#afxround) | Rounds a floating-point value to the nearest integer. |
| [AfxRoundCeil](#afxroundceil) | Rounds a floating-point value downward. |
| [AfxRoundFloor](#afxroundfloor) | Rounds a floating-point value upward. |
| [AfxRoundHalfAwayFromZero](#afxroundhalfawayfromzero) | Rounds halfway values away from zero |
| [AfxRoundHalfEven](#afxroundhalfeven) | Banker's rounding (IEEE 754 default): ties go to the nearest even integer. |
| [AfxRoundHalfUp](#afxroundhalfup) | Rounds halfway values toward positive infinity. |
| [AfxRoundToMultiple](#afxroundtomultiple) | Rounds x to the nearest multiple of nMultiple. |
| [AfxRoundTrunc](#afxroundtrunc) | Rounds a floating-point value truncating toward zero. |
| [AfxScalb](#afxscalb) | Multiplies a floating-point number by 2 raised to the power of n. |
| [AfxScalbn](#afxscalb) | Multiplies a floating-point number by 2 raised to the power of n. |
| [AfxSec](#afxsec) | Secant. Calculates the reciprocal of the cosine function. |
| [AfxSech](#afxsech) | Calculates the hyperbolic secant. |
| [AfxShuffle](#afxshuffle) | Shuffles the elements of an array in place using the Fisher-Yates algorithm. |
| [AfxSign](#afxsign) | Determines the sign of a numeric value. |
| [AfxSignbit](#afxsignbit) | Determines whether the sign bit of a floating-point value is set. |
| [AfxSin](#afxsin) | Calculates the sine of an angle in radians. |
| [AfxSinh](#afxsinh) | Calculates the hyperbolic sine of x. |
| [AfxSmoothStep](#afxsmoothstep) | Smooth interpolation between edge0 and edge1 using Hermite polynomial. |
| [AfxSqrt](#afxsqrt) | Returns the square root of a number. |
| [AfxStdDev](#afxstddev) | Computes the standard deviation of an array of double values. |
| [AfxSum](#afxsum) | Computes a compensated (Kahan) sum of an array of double values. |
| [AfxTan](#afxtan) | Calculates the tangent of an angle in radians. |
| [AfxTanh](#afxtanh) | Calculates the hyperbolic tangent of x. |
| [AfxTgamma](#afxtgamma) | Determines the gamma function of the specified value. |
| [AfxTrailingZeros](#afxtrailingzeros) | Counts the number of trailing zero bits in a 64-bit unsigned integer. |
| [AfxTruncate](#afxfix) | Returns the integer part of a number, truncating toward zero. |
| [AfxVariance](#afxvariance) | Computes the variance of an array of double values. |

---

## AfxAbs

Calculates the absolute value of a number.

```
#define AfxAbs(x) iif((x) < 0, -(x), (x))
```

The absolute value of a number is its positive magnitude. If a number is negative, its value will be negated and the positive result returned. For example, Abs(-1) and Abs(1) both return 1. The required number argument can be any valid numeric expression.

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Value to find the absolute value of. |

#### Return Value

The absolute value of the number.

#### Usage example
```
PRINT AfxAbs(5)
PRINT AfxAbs(-5)
```
---

## AfxAcos

Calculates the arccosine in radians.

```
#define AfxAcos(x) Acos(x)
```

#### Return value

Returns the arccosine of x in the range 0 to π radians.

By default, if *x* is less than -1 or greater than 1, it returns an indefinite.

### Usage example:
```
PRINT AfxAcos(1.0)
```
---

## AfxAcosh

Calculates the inverse hyperbolic cosine.

```
#define AfxAcosh(x) (Log(x + Sqrt(x * x - 1)))
```

#### Return value

Returns the inverse hyperbolic cosine (arc hyperbolic cosine) of x. Values are valid over the domain x ≥ 1. If *x* is less than 1, *errno* is set to EDOM, and the result is a quiet NaN. If *x* is a quiet NaN, indefinite, or infinity, the same value is returned.

#### Usage example
```
PRINT AfxAcosh(2.0)
```
---

## AfxAcot

Calculates the inverse cotangent.

```
#define AfxAcot(x) (2 * Atn(1) - Atn(x))
```

Returns the principal value of the arccotangent, or inverse cotangent, of its argument. The returned angle is given in radians in the range 0 (zero) to π.

#### Usage example
```
PRINT AfxAcot(2.6)
```
---

## AfxAcsc

Calculates the inverse cosecant in radians.

```
#define AfxAcsc(x) (Atn(Sgn(x) / Sqrt(x * x - 1)))
```

Returns the principal value of the inverse cosecant of its argument. The returned angle is given in radians in the interval [-π/2, π/2].

#### Usage example
```
PRINT AfxAcsc(3)
```
---

## AfxAcsch

Calculates the inverse hyperbolic cosecant in radians.

```
#define AfxAcsch(x) (Log((Sgn(x) * Sqrt(x * x + 1) + 1) / x))
```

Returns the principal value of the inverse cosecant of its argument. The returned value is given in radians.

Domain: All real numbers except x = 0.

---

## AfxAcoth

Calculates the inverse hyperbolic cotangent.

```
#define AfxAcoth(x) (Log((x + 1) / (x - 1)) / 2)
```

Returns the inverse hyperbolic cotangent of a number. The number must be less than -1 or greater than 1 (absolute value must be greater than 1). It returns the value in radians.

#### Usage example
```
PRINT AfxAcoth(2)
```
---

## AfxAsec

Calculates the inverse secant in radians.

```
#define AfxAsec(x) (2 * Atn(1) - Atn(Sgn(x) / Sqrt(x * x - 1)))
```

Returns the principal value of the inverse secant of its argument. The returned angle is given in radians in the range [0,𝜋] excluding 𝜋/2.

#### Usage example
```
PRINT AfxAsec(-2.8)
```
---

## AfxAsech

Calculates the inverse hyperbolic secant.

```
#define AfxAsech(x) (Log((Sqrt(-x * x + 1) + 1) / x))
```

Returns the principal value of the inverse hyperbolic secant of its argument. The returned angle is given in radians in the range [0,+∞].

#### Usage example
```
PRINT AfxAsech(0.5)
```
---

## AfxAsin

Calculates the arc sine (inverse sine) in radians.

```
#define AfxAsin(x) Asin(x)
```

Returns the arcsine, or inverse sine, of its argument. The arcsine is the angle whose sine is the argument. The returned angle is given in radians in the range -π/2 to π/2.

#### Usage example
```
PRINT AfxAsin(1)
```
---

## AfxAsinh

Calculates the inverse hyperbolic sine.

```
#define AfxAsinh(x) (Log(x + Sqrt(x * x + 1)))
```

Returns the inverse hyperbolic sine of a number. The inverse hyperbolic sine is the value whose hyperbolic sine is x, so AfxAsinh(AfxSinh(x)) equals x.

#### Usage example
```
PRINT AfxAsinh(2)
```
---

## AfxAtan

Calculates the arc tangent (inverse tangent). |

```
#define AfxAtan(x) Atn(x)
```

Returns the arctangent, or inverse tangent, of its argument. The arctangent is the angle whose tangent is the argument. The returned angle is given in radians in the range -π/2 to π/2.

#### Usage example
```
PRINT AfxAtan(1.0)
```
---

## AfxAtan2

Calculates the arc tangent of y/x using the signs of both arguments to determine the correct quadrant of the result.

```
#define AfxAtan2(y, x) Atan2(y, x)
```

Returns the arctangent, or inverse tangent, of the specified x and y coordinates as arguments. The arctangent is the angle from the x-axis to a line that contains the origin (0, 0) and a point with coordinates (x, y). The angle is given in radians between -π and π, excluding -π. A positive result represents a counterclockwise angle from the x-axis; a negative result represents a clockwise angle. Atan2(a, b) equals Atan(b/a), except that a can equal 0 (zero) with the **AfxAtan2** function.

#### Usage example
```
PRINT AfxAtan2(1.0, 1.0)
```
---

## AfxAtanh

Returns the inverse hyperbolic tangent of a number.

```
#define AfxAtanh(x) (Log((1 + x) / (1 - x)) / 2)
```

Returns the inverse hyperbolic tangent of a number. The number must be between -1 and 1 (excluding -1 and 1). The inverse hyperbolic tangent is the value whose hyperbolic tangent is number, so that AfxATanh(AfxTamh(number)) equals number.

#### Usage example
```
PRINT AfxAtanh(0.76159416)
```
---

## AfxBessel

Computes the Bessel function. he Bessel functions are commonly used in the mathematics of electromagnetic wave theory. ' The _j0, _j1, and _jn routines return Bessel functions of the first kind: orders 0, 1, and n, respectively. The _y0, _y1, and _yn routines return Bessel functions of the second kind: orders 0, 1, and n, respectively.

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The floating-point value. |
| *n* | Integer order of Bessel function. |

```
#define AfxBesselJ0(x) _j0(x)
#define AfxBesselJ1(x) _j1(x)
#define AfxBesselJn(n, x) _jn(n, x)
#define AfxBesselY0(x) _y0(x)
#define AfxBesselY1(x) _y1(x)
#define AfxBesselYn(n, x) _yn(n, x)
```

#### Return value:

Each of these routines returns a Bessel function of x. If x is negative in the _y0, _y1, or _yn functions, the routine sets errno to EDOM, prints a _DOMAIN error message to stderr, and returns HUGE_VAL. You can modify error handling by using _matherr.

---

## AfxByteSwap

Swaps the byte order of a 64-bit unsigned integer (endian flip).

```
FUNCTION AfxByteSwap (BYVAL x AS ULONGINT) AS ULONGINT
    RETURN ((x AND &h00000000000000FFULL) SHL 56) OR _
           ((x AND &h000000000000FF00ULL) SHL 40) OR _
           ((x AND &h0000000000FF0000ULL) SHL 24) OR _
           ((x AND &h00000000FF000000ULL) SHL 8)  OR _
           ((x AND &h000000FF00000000ULL) SHR 8)  OR _
           ((x AND &h0000FF0000000000ULL) SHR 24) OR _
           ((x AND &h00FF000000000000ULL) SHR 40) OR _
           ((x AND &hFF00000000000000ULL) SHR 56)
END FUNCTION
```

#### Return value

The value with reversed byte order.

#### Usage example
```
PRINT HEX(AfxByteSwap(&h0123456789ABCDEFULL))   ' returns EFCDAB8967452301
```
---

## AfxCbrt

Returns the cube root of a number.

```
#define AfxCbrt(x) cbrt(x)
```

#### Usage example
```
PRINT AfxCbrt(125)
```
---

## AfxChangeSign

Inverts the sign of a double value.

```
#define AfxChangeSign(x) (-(x))
```

#### Return value

Returns the same value with inverted sign.

#### Usage examples:
```
PRINT AfxChangeSign(5.0)     ' -5
PRINT AfxChangeSign(-12.75)  ' 12.75
PRINT AfxChangeSign(0.0)     ' -0
```
---

## AfxCeil

Calculates the ceiling of a value.

```
#define AfxCeil(x) Ceil(x)
```

#### Usage examples
```
PRINT AfxCeil(3.2)    ' 4
PRINT AfxCeil(-3.2)   ' -3
```
---

## AfxClamp

Restricts a value to remain within a specified range.

```
#define AfxClamp(v, min, max) iif((v) < (min), (min), iif((v) > (max), (max), (v)))
```

| Parameter  | Description |
| ---------- | ----------- |
| *v* | Value to clamp. |
| *min* | Lower limit of the range. |
| *max* | Upper limit of the range. |

If the value is less than the minimum, the minimum is returned. If the value is greater than the maximum, the maximum is returned. Otherwise, the original value is returned. This is useful for keeping values within valid limits, such as screen coordinates, color ranges, or normalized data.

#### Usage example
```
PRINT AfxClamp(125, 0, 100)
```
---

## AfxComb

Computes the number of distinct combinations of k elements from a set of n.

```
FUNCTION AfxComb (BYVAL n AS LONG, BYVAL k AS LONG) AS ULONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *v* | Value to clamp. |

#### Remarks

Uses the formula: C(n, k) = n! / (k! * (n - k)!)

If k > n or either argument is negative, returns 0.

#### Usage examples
```
PRINT AfxComb(5, 2)  ' 10
PRINT AfxComb(6, 3)  ' 20
PRINT AfxComb(10, 0) ' 1
PRINT AfxComb(10, 10) ' 1
```
---

## AfxCopySign

Returns a value that has the magnitude of one argument and the sign of another.

```
#define AfxCopySign(x, y) copysign(x, y)
```

#### Usage xample
```
PRINT AfxCopySign(15, -7)
```
---

## AfxCos

Calculates the cosine of an angle in radians.

```
#define AfxCos(x) Cos(x)
```

#### Usage example
```
PRINT AfxCos(45)
```
---

## AfxCosh

Calculates the hyperbolic cosine of a number.

```
#define AfxCosh(x) Cosh(x)
```

#### Usage example
```
PRINT AfxCosh(1)
```
---

## AfxCoth

Calculates the hyperbolic cotangent of a hyperbolic angle in radians.

```
#define AfxCoth(x) (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
```

The hyperbolic cotangent is an analog of the ordinary (circular) cotangent.

The absolute value of number must be less than and cannot be 0.

#### Usage example

```
PRINT AfxCoth(1.0)
```
---

## AfxCsc

Cosecant. Calculates the reciprocal of the sine function.

Values greater than or equal to 1, or less than or equal to -1.

```
#define AfxCsc(x) (1 / Sin(x))
```

#### Usage example
```
PRINT AfxCsc(1.0)
```
---

## AfxCsch

Calculates the hyperbolic cosecant.

```
#define AfxCsch(x) (2 / (Exp(x) - Exp(-x)))
```

#### Remarks

It is equal to the reciprocal of the hyperbolic sine function

#### Usage example
```
PRINT AfxCsch(1.0)
```
---

## AfxCtan

Cotangent. Calculates the reciprocal of the tangent function.

```
#define AfxCtan(x) (1 / Tan(x))
```

#### Usage example
```
PRINT AfxCtan(0.01)
```
---

## AfxDegToRad

Converts degrees to radians.

```
#define AfxDegToRad(deg) ((deg) * (Atn(1) * 4) / 180)
```

#### Usage example
```
PRINT AfxDegToRad(180)
```
---

## AfxDivMod

Performs integer division with floored quotient and remainder.

```
SUB AfxDivMod (BYVAL a AS LONGINT, BYVAL b AS LONGINT, BYREF q AS LONGINT, BYREF r AS LONGINT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *a* | Dividend |
| *b* | Divisor |
| *q* | Quotient (returned by reference) |
| *r* | Remainder (returned by reference) |

#### Remarks

Computes q and r such that: a = b * q + r and r has the same sign as b (floored division). This differs from the default FreeBasic operator "\" which truncates toward zero.

For example:
```
-7 \ 3  = -2   (truncated)
AfxDivMod(-7, 3) -> q = -3, r = 2   (floored)
```
#### Usage examples
```
DIM AS LONGINT 1, r
AfxDivMod(7, 3, q, r)   ' q = 2,  r = 1
AfxDivMod(-7, 3, q, r)  ' q = -3, r = 2
AfxDivMod(7, -3, q, r)  ' q = -3, r = -2
AfxDivMod(-7, -3, q, r) ' q = 2,  r = -1
```
---

## AfxEaseInOut

Smooth ease-in/ease-out interpolation using cubic Hermite curve.

```
FUNCTION AfxEaseInOut (BYVAL t AS DOUBLE) AS DOUBLE
```

#### Return value

Returns t^2 * (3 - 2t), same shape as **SmoothStep** but normalized for 0,1.

#### Usage examples:
```
PRINT AfxEaseInOut(0.25)        ' 0.15625
PRINT AfxEaseInOut(0.75)        ' 0.84375
```
---

## AfxErf

Computes the error function of a value.

```
#define AfxErf(x) (erf(x))
```

Returns the Gauss error function of x.

#### Usage example
```
PRINT AfxErf(0.0)          ' 0
PRINT AfxErf(0.5)          ' 0.5204998778130465
PRINT AfxErf(1.0)          ' 0.8427007929497149
PRINT AfxErf(-1.0)         ' -0.8427007929497149
PRINT AfxErf(2.0)          ' 0.9953222650189527
PRINT AfxErf(-2.0)         ' -0.9953222650189527
```
---

## AfxErfc

Computes the complementary error function of a value.

```
#define AfxErfc(x) (erfc(x))
```

#### Usage examples
```
PRINT AfxErfc(0.0)         ' 1
PRINT AfxErfc(0.5)         ' 0.4795001221869535
PRINT AfxErfc(1.0)         ' 0.1572992070502851
PRINT AfxErfc(2.0)         ' 0.004677734981047265
PRINT AfxErfc(10.0)        ' 2.088487583762545e-045
```
---

## AfxEuroRound

European-style rounding: rounds halfway values away from zero. This routine implements the "half away from zero" rule commonly used in Europe for financial and general numeric rounding. It adds ±0.5 depending on the sign of the value, then truncates toward zero using **Fix**. Eauivalent to **AfxRoundHalfAwayFromZero**.

```
FUNCTION AfxEuroRound (BYVAL dblValue AS DOUBLE, BYVAL numDigits AS LONG = 0) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *dblValue* | Value to round. |
| *numDigits* | Number of decimal places |

#### Usage examples
```
PRINT AfxEuroRound(2.3)      ' 2
PRINT AfxEuroRound(2.5)      ' 3
PRINT AfxEuroRound(-2.3)     ' -2
PRINT AfxEuroRound(-2.5)     ' -3
PRINT AfxEuroRound(2.3456, 2) ' 2.35
```
---

## AfxExp

Computes the exponential value of a number.

```
#define AfxExp(x) Exp((x))
```

#### Usage example
```
PRINT AfxExp(2.0)
```
---


## AfxExp2

Computes 2 raised to the specified value.

```
#define AfxExp2(x) exp2(x)
```

#### Usage example
```
PRINT AfxExp2(2.0)
```
---

## AfxExpm1

Calculates e raised to the power of x minus 1.

```
#define AfxExpm1(x) Exp(x) - 1.0
```

#### Usage example
```
PRINT AfxExpm1(1.0)
```
---

## AfxExp2m1

Calculates 2 raised to the power of x minus 1. Useful for binary scaling and exponential offset computations.

```
#define AfxExp2m1(x) Exp(x * Log(2.0)) - 1.0
```

#### Usage example
```
PRINT AfxExp2m1(5.0)
```
---

## AfxFactorial

Computes the factorial of an integer value.

```
FUNCTION AfxFactorial (BYVAL n AS LONG) AS ULONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *n* | Integer value. |

#### Usage examples
```
PRINT AfxFactorial(0)  ' 1
PRINT AfxFactorial(1)  ' 1
PRINT AfxFactorial(5)  ' 120
PRINT AfxFactorial(10) ' 3628800
```
---

## AfxFdim

Determines the positive difference between the first and second values.

```
#define AfxFdim(x, y) fdim(x, y)
```

#### Usage examples
```
PRINT AfxFdim(5.3, 2.0)    ' 3.3
PRINT AfxFdim(2.0, 5.3)    ' 0
PRINT AfxFdim(-1.0, -3.0)  ' 2
PRINT AfxFdim(-3.0, -1.0)  ' 0
```
---

## AfxFix
## AfxTruncate

Returns the integer part of a number, truncating toward zero.

.NET calls it **Truncate**.

```
#define AfxFix(x) Fix(x)
#define AfxTruncate(x) Fix(x)
```
#### Usage examples
```
PRINT AfxFix(-3.7)   ' -3
PRINT AfxFix(3.7)    '  3
```

## AfxFloor

Returns the largest integer less than or equal to x.

```
#define AfxFloor(x) Floor(x)
```
#### Usage examples
```
PRINT AfxFloor(3.7)    ' 3
PRINT AfxFloor(-3.7)   ' -4
```
---

## AfxFma

Multiplies two values together, adds a third value, and then rounds the result, while only losing a small amount of precision due to intermediary rounding.

```
#define AfxFma(x, y, z) fma(x, y, z)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The first value to multiply. |
| *y* | The second value to multiply. |
| *z* | The value to add. |

#### Usage examples
```
PRINT AfxFma(2.0, 3.0, 4.0)         ' 10
PRINT AfxFma(1e308, 1e-308, 1.0)    ' 2
PRINT AfxFma(0.0, 5.0, 7.0)         ' 7
PRINT AfxFma(1/0.0, 0.0, 2.0)       ' NaN (inf * 0)
```
---

## AfxFormat

 Returns a string with the result of the numerical expression formatted as indicated in the formatting expression. The formatting expression is a string that can yield numeric or date-time values. The representation of time and date separators may be determined from the operating system if regional or localized settings are supported by the target platform. This determination is made at run-time so that output may be localized to the system the program is running on. See the focumentation of Format in the FreeBasic help file.

```
FUNCTION AfxFormat (BYVAL value AS DOUBLE, BYREF mask AS CONST STRING = "") AS STRING
   RETURN Format(value, mask)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | Number to format. |
| *mask* | Formatting pattern. |

#### Usage example
```
PRINT AfxFormat(12345.67, "#,##0.00")
```
#### Example
```
DIM AS DOUBLE numberVal(...) = {5, -5, .5}
DIM AS STRING formatStr(...) = {"","0","0.00","#,##0","#,##0.00","0%", "0.00%", "0.00E+00", "0.00E-00"}
PRINT "Format string",, STR(numberVal(0)), STR(numberVal(1)), STR(numberVal(2))
FOR iFormat AS INTEGER = 0 TO UBOUND(formatStr)
   PRINT formatStr(iFormat),,
   FOR iNumber AS INTEGER = 0 TO UBOUND(numberVal)
      PRINT AfxFormat(numberVal(iNumber), formatStr(iFormat)),
   NEXT
  PRINT
NEXT
```
---

## AfxFormatCurrency

Formats a currency number into a string form.

```
FUNCTION AfxFormatCurrency (BYVAL value AS DOUBLE, BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | The numeric value to format. |
| *iNumDig* | Digits after decimal (-1 = system default) |
| *iIncLead* | Leading zero (-2 = default, -1 = include, 0 = no) |
| *iUseParens* | Negative parentheses (-2 = default, -1 = yes, 0 = no) |
| *iGroup* | Group thousands (-2 = default, -1 = yes, 0 = no) |
| *dwFlags* | VAR_CALENDAR_HIJRI only |

#### Usage example:
```
DIM dbl AS DOUBLE = 12345.1234
PRINT AfxFormatCurrency(dbl)  ' --> 12.345,12 € (Spain)
```
---

## AfxFormatNumber

Formats a number into a string form (no currency symbol).

```
FUNCTION AfxFormatNumber (BYVAL value AS DOUBLE, BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | The numeric value to format. |
| *iNumDig* | Digits after decimal (-1 = system default) |
| *iIncLead* | Leading zero (-2 = default, -1 = include, 0 = no) |
| *iUseParens* | Negative parentheses (-2 = default, -1 = yes, 0 = no) |
| *iGroup* | Group thousands (-2 = default, -1 = yes, 0 = no) |
| *dwFlags* | VAR_CALENDAR_HIJRI only |

#### Usage example:
```
DIM dbl AS DOUBLE = 12345.1234
PRINT AfxFormatNumber(dbl)  ' --> 12.345,12 (Spain)
```
---

## AfxFormatPercent

Formats a percentage number into a string form.

```
FUNCTION AfxFormatPercent (BYVAL value AS DOUBLE, BYVAL iNumDig AS LONG = -1, BYVAL iIncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | The numeric value to format. |
| *iNumDig* | Digits after decimal (-1 = system default) |
| *iIncLead* | Leading zero (-2 = default, -1 = include, 0 = no) |
| *iUseParens* | Negative parentheses (-2 = default, -1 = yes, 0 = no) |
| *iGroup* | Group thousands (-2 = default, -1 = yes, 0 = no) |
| *dwFlags* | VAR_CALENDAR_HIJRI only |

#### Usage example:
```
PRINT AfxFormatPercent(0.25)  ' --> 25,00% (Spain)
```
---

## AfxFpClassify

Returns a value indicating the floating-point classification of the argument.

```
#define AfxFpClassify(x) (_fpclass(x))
```

#### Usage examples:
```
PRINT AfxFpClassify(0.0)          ' _FPCLASS_PZ
PRINT AfxFpClassify(-0.0)         ' _FPCLASS_NZ
PRINT AfxFpClassify(1.0)          ' _FPCLASS_PN
PRINT AfxFpClassify(-3.14)        ' _FPCLASS_NN
PRINT AfxFpClassify(1/0.0)        ' _FPCLASS_PINF
PRINT AfxFpClassify(-1/0.0)       ' _FPCLASS_NINF
PRINT AfxFpClassify(0.0/0.0)      ' _FPCLASS_QNAN
```
---

## AfxFrac

Returns the fractional part of a number.

```
#define AfxFrac(x) Frac(x)
```

#### Usage examples
```
PRINT AfxFrac(3.75)   ' 0.75
PRINT AfxFrac(-3.75)  ' -0.75
```
---

## AfxFrexp

Gets the mantissa and exponent of a floating-point number.

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Floating-point value. |
| *e* | A long variable that receives the exponent. |

#### Return value

Returns the mantissa. If x is 0, the function returns 0 for both the mantissa and the exponent.

```
FUNCTION AfxFrexp (BYVAL x AS DOUBLE, BYREF e AS LONG) AS DOUBLE
   RETURN frexp(x, @e)
END FUNCTION
```
#### Usage example
```
DIM value AS DOUBLE = 12345.67
DIM e AS LONG
DIM mantissa AS DOUBLE
mantissa = AfxFrexp(value, e)
PRINT mantissa, e
```
---

## AfxFusedMultiplyAdd

Computes (x * y) + z with IEEE 754 semantics. If any argument is NaN, the result is NaN (0 / 0).
```
FUNCTION AfxFusedMultiplyAdd (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
   IF AfxIsNaN(x) OR AfxIsNaN(y) OR AfxIsNaN(z) THEN RETURN 0.0 / 0.0          ' NaN
   RETURN (x * y) + z
END FUNCTION
```
#### Usage examples:
```
PRINT AfxFusedMultiplyAdd(2, 3, 4)      ' 10
PRINT AfxFusedMultiplyAdd(-1, 8, 0.5)   ' -7.5
PRINT AfxFusedMultiplyAdd(0/0, 5, 1)    ' -1.#IND
```
---

## AfxGcd

Computes the greatest common divisor (GCD) of two integers using the Euclidean algorithm.

| Parameter  | Description |
| ---------- | ----------- |
| *a* | First integer. |
| *b* | Second integer. |

#### Return value

Returns the largest integer that divides both *a* and *b* without remainder. If either argument is zero, returns the absolute value of the other. Always returns a non-negative result.

#### Usage examples
```
PRINT AfxGcd(12, 8)   ' 4
PRINT AfxGcd(54, 24)  ' 6
PRINT AfxGcd(-42, 56) ' 14
PRINT AfxGcd(0, 9)    ' 9
```
---

## AfxGetExponent

Gets the exponent of a floating-point number.

```
FUNCTION AfxGetExponent (BYVAL x AS DOUBLE) AS LONG
```

#### Usage example:
```
DIM value AS DOUBLE = 12345.67
DIM e AS LONG = AfxGetExponent(value)
```
---

## AfxGetMantissa

Gets the mantissa of a floating-point number.

```
FUNCTION AfxGetMantissa (BYVAL x AS DOUBLE) AS DOUBLE
```

#### Usage example:
```
DIM value AS DOUBLE = 12345.67
DIM e AS LONG = AfxGetMantissa(value)
```
---

## AfxHypot

Calculates the Euclidean distance (hypotenuse) from two values.

```
#define AfxHypot(x, y) Sqr(((x) * (x)) + ((y) * (y)))
```

#### Usage example
```
DIM AS DOUBLE dA = 3.0
DIM AS DOUBLE dB = 4.0
DIM AS DOUBLE dResult
dResult = AfxHypot(dA, dB)
PRINT "Hypotenuse of sides "; dA; " and "; dB; " = "; dResult
```
---

## AfxHypot3

Computes the value of sqr(x^2 + y^2 + z^2).

Useful for calculating the magnitude of a 3D vector or distance in 3D space.

```
FUNCTION AfxHypot3 (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
   RETURN hypot(hypot(x, y), z)
END FUNCTION
```

#### Usage example:
```
DIM AS DOUBLE x = 3.0, y = 4.0, z = 12.0
PRINT "AfxHypot3("; x; ", "; y; ", "; z; ") = "; AfxHypot3(x, y, z)
```
---

## AfxILogB

Returns the unbiased exponent of the floating-point number x,

```
#define AfxILogB(x) (ilogb(x))
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The input value (must be > 0). |

#### Return Value

Integer exponent e such that: x = m * 2^e, where m is in [0.5, 1).

#### Usage examples:
```
PRINT AfxILogB(8.0)        ' 3
PRINT AfxILogB(0.0)        ' FP_ILOGB0 (usually -INT_MAX)
PRINT AfxILogB(-1.0)       ' 0
PRINT AfxILogB(1/0.0)      ' INT_MAX
PRINT AfxILogB(0.0/0.0)    ' FP_ILOGBNAN (-1)
```
---

## AfxInt

Returns the largest integer less than or equal to x.
```
#define AfxInt(x) Int(x)
```
#### Usage example
```
PRINT AfxInt(3.7)    ' 3
PRINT AfxInt(-3.7)   ' -4
```
---

## AfxInverseLerp

Takes a start value, an end value, and a target value. It calculates where the target value sits between the start and end, returning a 0-to-1 ratio representing that position

```
#define AfxInverseLerp(a,b,v) ((v - a) / (b - a))
```

#### Usage example:
```
PRINT AfxInverseLerp(0, 10, 2.5)    ' 0.25
```
---

## AfxIsClose

Determines whether two floating-point values are approximately equal.

```
FUNCTION AfxIsClose (BYVAL a AS DOUBLE, BYVAL b AS DOUBLE, _
   BYVAL relTol AS DOUBLE = 1e-9, BYVAL absTol AS DOUBLE = 0.0) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *a* | The first value. |
| *b* | The second value. |
| *relTol* | Relative tolerance (default = 1e-9) |
| *absTol* | Absolute tolerance (default = 0.0) |

#### Return value

Returns TRUE if the difference between a and b is within the specified tolerances.

```
FUNCTION AfxIsClose (BYVAL a AS DOUBLE, BYVAL b AS DOUBLE, _
   BYVAL relTol AS DOUBLE = 1e-9, BYVAL absTol AS DOUBLE = 0.0) AS BOOLEAN
```

The comparison follows the rule:
```
ABS(a - b) <= MAX(relTol * MAX(ABS(a), ABS(b)), absTol)
```
This method is consistent with IEEE and Python's math.isclose().

Useful for comparing results of floating-point calculations where rounding errors occur.

#### Usage examples
```
PRINT AfxIsClose(1.000000001, 1.0)          ' FALSE
PRINT AfxIsClose(1000.0, 1000.0001)         ' FALSE
PRINT AfxIsClose(1.0, 1.1)                  ' FALSE
PRINT AfxIsClose(0.0, 1e-10, , 1e-9)        ' TRUE
```
---

## AfxIsFinite

Determines whether the argument is a finite floating-point value.

```
#define AfxIsFinite(x) CBOOL((_finite(x) <> 0))
```

#### Return value

Returns TRUE if x is finite (normal, subnormal or zero), FALSE otherwise.

#### Usage example:
```
PRINT AfxIsFinite(123.45) ' TRUE
PRINT AfxIsFinite(1.0/0.0) ' FALSE
PRINT AfxIsFinite(0.0/0.0) ' FALSE
```
---

## AfxIsInfinity

Determines whether the argument is an infinity.

```
FUNCTION AfxIsInfinity (BYVAL x AS DOUBLE) AS LONG
   IF _isnan(x) THEN RETURN 0
   ' x is either +inf or -inf
   IF _finite(x) = FALSE THEN RETURN IIF(x > 0, 1, -1)
   RETURN 0
END FUNCTION
```
```
#define AfxIsInf AfxIsInfinity
```

#### Return value

Returns +1 if x is positive infinity, -1 if x is negative infinity and 0 otherwise.

#### Usage examples
```
PRINT AfxIsInfinity( 1.0 / 0.0 )     ' returns 1
PRINT AfxIsInfinity( -1.0 / 0.0 )    ' returns -1
```
---

## AfxIsNaN

Determines whether the argument is a NaN (Not a Number).

#### Return value

Returns TRUE if x is NaN, FALSE otherwise.

```
#define AfxIsNaN(x) CBOOL((_isnan(x) <> 0))
```

#### Usage example
```
PRINT AfxIsNaN(0.0/0.0)    ' TRUE
PRINT AfxIsNaN(1.0/0.0)    ' FALSE
PRINT AfxIsNaN(5.0)        ' FALSE
```
---

## AfxIsNormal

Returns TRUE if x is a normal IEEE 754 double-precision value.

Normal numbers are finite, non-zero, not subnormal, and >= the smallest normalized value IEEE 754 double threshold ˜ 2.2250738585072014e-308.

```
#define AfxIsNormal(x) CBOOL((ABS(x) >= 2.2250738585072014e-308) AND (ABS(x) < 1.7976931348623157e+308))
```

#### Usage examples
```
DIM x AS DOUBLE
x = 1e-300
PRINT AfxIsNormal(x)      ' TRUE
x = 1e-310
PRINT AfxIsNormal(x)      ' FALSE   ' subnormal
x = 0.0
PRINT AfxIsNormal(x)      ' FALSE
x = 1e309
PRINT AfxIsNormal(x)      ' FALSE   ' infinite
```
---

## AfxIsPrime

Determines whether an integer value is prime.

```
FUNCTION AfxIsPrime (BYVAL n AS ULONGINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *n* | Integer value to test. |

#### Return value

Returns TRUE if n is a prime number, otherwise FALSE.

#### Usage examples
```
PRINT AfxIsPrime(2)   ' TRUE
PRINT AfxIsPrime(3)   ' TRUE
PRINT AfxIsPrime(4)   ' FALSE
PRINT AfxIsPrime(17)  ' TRUE
PRINT AfxIsPrime(100) ' FALSE
```
---

## AfxISqrt

Calculates the integer square root of a number.

```
FUCTION AfxISqrt (BYVAL ulN AS ULONGINT) AS ULONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ulN* | The value whose integer square root is to be calculated. |

#### Return value

Returns the largest integer whose square is less than or equal to *ulN*.

Works entirely with integer arithmetic, avoiding floating-point rounding errors.

Useful in number theory, geometry, and optimization algorithms.

#### Return Value

The integer square root of *ulN*.

#### Usage examples
```
DIM AS ULONGINT ulN = 12345
DIM AS ULONGINT ulResult
ulResult = AfxISqrt(ulN)
PRINT "Integer square root of "; ulN; " = "; ulResult
```
---

## AfxIsSubnormal

Returns TRUE if x is a subnormal (denormalized) floating-point number.

Subnormals are non-zero values smaller than the smallest normalized number.

```
#define AfxIsSubnormal(x) CBOOL(((Abs(x) > 0.0) And (Abs(x) < 2.2250738585072014e-308)))
```

#### Usage examples
```
DIM x AS DOUBLE
x = 1e-310
PRINT AfxIsSubnormal(x)   ' TRUE
x = 1e-300
PRINT AfxIsSubnormal(x)   ' FALSE
```
---

## AfxLcm

Computes the least common multiple (LCM) of two integers.

```
FUNCTION AfxLcm (BYVAL a AS LONGINT, BYVAL b AS LONGINT) AS LONGINT
   IF a = 0 OR b = 0 THEN RETURN 0
   RETURN ABS(a * b) \ AfxGcd(a, b)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *a* | The first integer. |
| *b* | The second integer. |

#### Return value

Returns the smallest positive integer that is a multiple of both a and b.

Uses the relationship: LCM(a, b) = ABS(a * b) / GCD(a, b). If either argument is zero, returns 0.

#### Usage examples
```
PRINT AfxLcm(12, 8)   ' 24
PRINT AfxLcm(54, 24)  ' 216
PRINT AfxLcm(-42, 56) ' 168
PRINT AfxLcm(0, 9)    ' 0
```
---

## AfxLdexp

Multiplies a floating-point number by an integral power of two.

```
#define AfxLdexp(x, e) ldexp(x, e)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The Floating-point value. |
| *e* | Integer exponent. |

#### Return value

Returns the value of x * 2exp if successful.

On overflow, and depending on the sign of x, ldexp returns +/- HUGE_VAL

#### Usage example
```
DIM value AS DOUBLE = 12345.67
DIM e AS LONG = 10
PRINT AfxLdexp(value, e)
```
---

## AfxLeadingZeros

Counts the number of leading zero bits in a 64-bit unsigned integer.

```
FUNCTION AfxLeadingZeros (BYVAL x AS ULONGINT) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | 64-bit unsigned integer. |

#### Returns value

Number of leading zeros (0–64).

#### Usage example
```
PRINT AfxLeadingZeros(&h0000000000000001ULL)   ' returns 63
PRINT AfxLeadingZeros(&h8000000000000000ULL)   ' returns 0
PRINT AfxLeadingZeros(&h00F0000000000000ULL)   ' returns 8
```
---

## AfxLerp

Linearly interpolates between *a* and *b* by *t*.

```
#define AfxLerp(a,b,t) (a + (b - a) * t)
```

#### Return value

Returns a + (b - a) * t

#### Usage example
```
PRINT AfxLerp(0, 10, 0.25)          ' 2.5
```
---

## AfxLgamma

Determines the natural logarithm of the absolute value of the gamma function of the specified value.

```
#define AfxLgamma(x) (lgamma(x))
```

#### Usage examples
```
PRINT AfxLgamma(5.0)        ' ln(|G(5)|) = ln(24) ˜ 3.17805383
PRINT AfxLgamma(0.5)        ' ln(|G(0.5)|) = ln(sqrt(pi))
PRINT AfxLgamma(-3.2)       ' +INFINITY for negative integer arguments or poles
```
---

## AfxLog

Calculates the natural logarithm (base e) of x.

```
#define AfxLog(dx) Log(dX)
```

#### Usage example
```
DIM AS DOUBLE dValue  = 2.7182818
DIM AS DOUBLE dResult
dResult = AfxLog(dValue)
PRINT "Natural logarithm of "; dValue; " = "; dResult
```
---

## AfxLog2

Calculates the base-2 logarithm of x.

```
#define AfxLog2(x) Log(x) / Log(2.0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The input value (must be > 0). |

#### Usage example
```
DIM AS DOUBLE dValue  = 8.0
DIM AS DOUBLE dResult
dResult = AfxLog2(dValue)
PRINT "Base-2 logarithm of "; dValue; " = "; dResult
```
---

## AfxLog10

Calculates the base-10 logarithm of x.

```
#define AfxLog10(x) Log(x) / Log(10.0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The input value (must be > 0). |

#### Usage example
```
DIM AS DOUBLE dValue  = 1000.0
DIM AS DOUBLE dResult
dResult = AfxLog10(dValue)
PRINT "Base-10 logarithm of "; dValue; " = "; dResult
```
---

## AfxLog1p

Calculates Log(1 + x) with improved precision for small x values.

```
#define AfxLog1p(x) Log(1.0 + x)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The input value (must be > 0). |

#### Remarks

Equivalent to Log(1 + x), but avoids loss of significance when x is near zero.

#### Usage example:
```
DIM AS DOUBLE x = 0.001
DIM AS DOUBLE dResult
dResult = AfxLog1p(x)
PRINT "Log(1 + "; x; ") = "; dResult
```
---

## AfxLogB

Returns the unbiased exponent of the floating-point number x as a double.

```
#define AfxLogB(x) (logb(x))
```
#### Usage examples
```
PRINT AfxLogB(8.0)         ' 3
PRINT AfxLogB(0.0)         ' -1.#INF
PRINT AfxLogB(1.0)         ' 0
PRINT AfxLogB(0.5)         ' -1
```
---

## AfxLogBase

Calculates the logarithm of x in an arbitrary base.

```
#define AfxLogBase(x, dBase) Log(x) / Log(dBase)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The input value (must be > 0). |
| *dBase* | The input value (must be > 0). |

#### Usage example
```
DIM AS DOUBLE x    = 81.0
DIM AS DOUBLE dBase = 3.0
DIM AS DOUBLE dResult
dResult = AfxLogBase(x, dBase)
PRINT "Log base "; dBase; " of "; x; " = "; dResult
```
---

## AfxMax

Returns the larger of two values.

```
#define AfxMax(x, y) iif((x) > (y), (x), (y))
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The first value. |
| *y* | The second value. |

#### Usage example
```
DIM AS INTEGER iA = 25
DIM AS INTEGER iB = 42
DIM AS INTEGER iResult
iResult = AfxMax(iA, iB)
PRINT "Maximum of "; iA; " and "; iB; " = "; iResult
```
---

## AfxMaxMagnitude

Returns the value (x or y) with the largest magnitude.

If either argument is NaN, the result is NaN (0 / 0).

```
FUNCTION AfxMaxMagnitude (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS DOUBLE
   IF AfxIsNaN(x) OR AfxIsNan(y) THEN RETURN 0.0 / 0.0          ' NaN
   IF AfxAbs(x) > AfxAbs(y) THEN RETURN x ELSE RETURN y
END FUNCTION
```

#### Usage examples:
```
PRINT AfxMaxMagnitude(5, -10)     ' -10  (|-10| = 10 > 5)
PRINT AfxMaxMagnitude(-3, 2)      ' -3
PRINT AfxMaxMagnitude(0.5, -0.4)  ' 0.5
PRINT AfxMaxMagnitude(0/0, 7)     ' -1.#IND
```
---

## AfxMean

Computes the arithmetic mean of an array of double values.

```
FUNCTION AfxMean (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pData* | Pointer to the first element of the data array. |
| *n* | Number of elements in the array. |

## Return value

The average value.

#### Usage example:
```
DIM rgData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
PRINT AfxMean(@rgData(0), 5)
```
---

## AfxMedian

Computes the median value of an array of doubles.

```
FUNCTION AfxMedian (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pData* | Pointer to the first element of the data array. |
| *n* | Number of elements in the array. |

#### Usage example
```
DIM rgData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
PRINT AfxMedian(@rgData(0), 5)
```
---

## AfxMidpoint

Calculates the midpoint between two integer values.

This function returns the value halfway between a and b. It uses an overflow-safe formula (a + (b - a)\2) to avoid integer overflow that can occur with (a + b)\2 when a and b are large. Commonly used in algorithms such as binary search.

```
FUNCTION AfxMidpoint (BYVAL a AS LONGINT, BYVAL b AS LONGINT) AS LONGINT
   RETURN a + (b - a) \ 2
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *a* | The first integer value. |
| *b* | The second integer value. |

#### Return Value

The midpoint between a and b.

#### Usage example:
```
PRINT AfxMidpoint(10, 20)
```
---

## AfxMin

Returns the smaller of two values.

```
#define AfxMin(x, y) iif((x) < (y), (x), (y))
```
| Parameter  | Description |
| ---------- | ----------- |
| *x* | The first alue. |
| *y* | The second value. |

#### Usage example:
```
DIM AS INTEGER iA = 25
DIM AS INTEGER iB = 42
DIM AS INTEGER iResult
iResult = AfxMin(iA, iB)
PRINT "Minimum of "; iA; " and "; iB; " = "; iResult
```
---

## AfxMinMagnitude

Returns the value (x or y) with the smallest magnitude.

If either argument is NaN, the result is NaN (0 / 0).

```
FUNCTION AfxMinMagnitude ( BYVAL x AS DOUBLE, BYVAL y AS DOUBLE ) AS DOUBLE
   IF AfxIsNaN(x) OR AfxIsNaN(y) THEN RETURN 0.0 / 0.0          ' NaN
   IF AfxAbs(x) < AfxAbs(y) THEN RETURN x ELSE RETURN y
END FUNCTION
```
#### Usage examples
```
PRINT AfxMinMagnitude(5, -10)     ' 5
PRINT AfxMinMagnitude(-3, 2)      ' 2
PRINT AfxMinMagnitude(0.5, -0.4)  ' -0.4
PRINT AfxMinMagnitude(0/0, 7)     ' -1.#IND
```
---

## AfxMod

Calculates the floating-point remainder.

```
#define AfxMod(x, y) (fmod((x), (y)))
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The first alue. |
| *y* | The second value. |

#### Return value

Returns the floating-point remainder of x / y. If the value of y is 0.0, fmod returns a quiet NaN.

#### Usage examples
```
PRINT AfxMod(-5.3, 2.0)    ' -1.3
PRINT AfxMod(5.3, -2.0)    ' 1.3
PRINT AfxMod(-5.3, -2.0)   ' -1.3
PRINT AfxMod(1.0, 0.0)     ' NaN
PRINT AfxMod(0.0, 0.0)     ' NaN
PRINT AfxMod(1.0/0.0, 2.0) ' NaN
PRINT AfxMod(5.0, 1.0/0.0) ' 5.0
```
---

## AfxModf

Splits a floating-point value into fractional and integer parts
```
#define AfxModf(dblX, dblIntPtr) (modf(dblX, dblIntPtr))
```
#### Return value

Returns the fractional part of x and stores the integer part in *intptr*.

#### Usage example:
```
DIM AS DOUBLE fraction, intptr
fraction = AfxModf(-14.87654321, @intptr)
PRINT "Fraction = "; fraction
PRINT "Integer  = "; intptr
```
---

## AfxNextAfter

Returns the next representable floating-point value of the return type after x in the direction of y.

```
#define AfxNextAfter(x, y) nextafter(x, y)
```
---

## AfxPerm

Computes the number of permutations (n permute k).

```
FUNCTION AfxPerm (BYVAL n AS LONG, BYVAL k AS LONG) AS ULONGINT
   IF n < 0 OR k < 0 OR k > n THEN RETURN 0
   RETURN AfxFactorial(n) \ AfxFactorial(n - k)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *n* | The total number of items. |
| *k* | The number of items to arrange. |

#### Return value

Returns the number of distinct ordered arrangements of k elements from a set of n.

Uses the formula: P(n, k) = n! / (n - k)!

If k > n or either argument is negative, returns 0.

Uses AfxFactorial for intermediate calculations.

#### Usage examples
```
PRINT AfxPerm(5, 2)  = 20
PRINT AfxPerm(6, 3)  = 120
PRINT AfxPerm(10, 0) = 1
PRINT AfxPerm(10, 10) = 3628800
```
---

## AfxPopCount

Counts the number of set bits (1s) in a 64-bit unsigned integer.

```
FUNCTION AfxPopCount (BYVAL x AS ULONGINT) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | 64-bit unsigned integer. |

#### Returns value
Returns the number of bits set to 1.

#### Usage example
```
PRINT AfxPopCount(&hF0F0F0F0F0F0F0F0)   ' returns 32
PRINT AfxPopCount(255)                  ' returns 8
PRINT AfxPopCount(0)                    ' returns 0
```
---

## AfxPow

Raises a number to a specified power.

```
#define AfxPow(x, y) ((x) ^ (y))
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Base value. |
| *y* | Exponent value. |

#### Return value

Returns x raised to the power of y using the ^ operator.

#### Usage example
```
PRINT AfxPow(3, 4)
```
---

## AfxPowMod

Performs exponentiation under a modulus in an overflow-safe way.

```
FUNCTION AfxPowMod (BYVAL ulBase AS ULONGINT, BYVAL ulExpo AS ULONGINT, _
   BYVAL ulModulus AS ULONGINT) AS ULONGINT
```

#### Usage example
```
PRINT AfxPowMod(7, 128, 19)
```
---

## AfxRadToDeg

Converts radians to degrees.

```
#define AfxRadToDeg(rad) ((rad) * 180 / (Atn(1) * 4))
```
#### Usage example
```
PRINT AfxRadToDeg(3.141592653589793)   ' returns 180
```
---

## AfxRandomBytes

Fills a memory buffer with cryptographically secure random bytes using BCryptGenRandom. This function is suitable for generating keys, nonces, GUIDs, or any data that requires high-quality randomness. It does not depend on the standard RND generator.

```
FUNCTION AfxRandomBytes OVERLOAD (BYVAL pOut AS ANY PTR, BYVAL cb AS LONG) AS BOOLEAN
FUNCTION AfxRandomBytes OVERLAD (BYVAL cb AS LONG) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *pOut* | Pointer to the output buffer that will receive the random bytes. |
| *cb* | Number of bytes to generate. |

#### Remarks

The default random number provider implements an algorithm for generating random numbers that complies with the NIST SP800-90 standard, specifically the CTR_DRBG portion of that standard.

#### Return value

Returns TRUE if the operation succeeds, FALSE otherwise.

#### Example (first overloaded function)

Generate 16 random bytes and display them in hexadecimal
```
DIM buffer(15) AS UBYTE
IF AfxRandomBytes(@buffer(0), 16) THEN
   FOR i AS LONG = 0 TO 15
      PRINT HEX(buffer(i), 2);
   NEXT
   PRINT
ELSE
   PRINT "Error generating random bytes."
END IF
```
#### Example (second overloaded function)
```
 DIM buffer AS STRING = AfxRandomBytes(16)
 IF LEN(buffer THEN
    FOR i AS LONG = 1 TO 16
       PRINT HEX(ASC(MID(buffer2, i, 2))); " ";
    NEXT
 END IF
```
---

## AfxRandomDouble

Generates a random floating-point number within the specified range, inclusive of both the lower (lo) and upper (hi) bounds.

```
FUNCTION AfxRandomDouble (BYVAL lo AS DOUBLE, BYVAL hi AS DOUBLE) AS DOUBLE
   IF hi < lo THEN SWAP lo, hi
   RETURN lo + (RND * (hi - lo))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *lo* | The lowest value that can be returned. |
| *h* | The highest value that can be returned. Must be greater than or equal to lo. |

#### Return value:

Returns a DOUBLE representing a random floating-point number between lo and hi, inclusive.

#### Remarks

If hi is less than lo, the function automatically swaps the values to ensure a valid range.

This function uses the standard pseudo-random generator (RND) initialized by RANDOMIZE.

#### Example

Generate random floating-point numbers between 0.0 and 1.0
```
RANDOMIZE TIMER   ' Initialize random seed
DIM x AS DOUBLE
FOR n AS LONG = 1 TO 5
   x = AfxRandomDouble(0.0, 1.0)
   PRINT "Random value #" & n & ": "; x
NEXT
```
---

## AfxRandomGauss

Generates a random floating-point number following a Gaussian (normal) distribution using the Box-Muller transform.

```
FUNCTION AfxRandomGauss (BYVAL mean AS DOUBLE = 0.0, BYVAL stddev AS DOUBLE = 1.0) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *mean* | The mean (average) value of the distribution. Default is 0.0. |
| *stddev* | The standard deviation (spread) of the distribution. Default is 1.0. |

#### Return value:

Returns a DOUBLE representing a random value distributed normally around mean, with the specified standard deviation.

#### Example

Generate 10 random numbers with mean 0 and standard deviation 1
```
RANDOMIZE TIMER   ' Initialize random seed
DIM x AS DOUBLE
FOR n AS LONG = 1 TO 10
   x = AfxRandomGauss(0.0, 1.0)
   PRINT "Gaussian value #" & n & ": "; x
NEXT
```
---

## AfxRandomInt

Generates a random integer within the specified range, inclusive of both the lower (lo) and upper (hi) bounds.

```
FUNCTION AfxRandomInt (BYVAL lo AS LONGINT, BYVAL hi AS LONGINT) AS LONGINT
   IF hi < lo THEN SWAP lo, hi
   RETURN lo + (RND * (hi - lo + 1))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *lo* | The lowest value that can be returned. |
| *hi* | The highest value that can be returned. Must be greater than or equal to *lo*. |

#### Remarks

If hi is less than lo, the function automatically swaps the values to ensure a valid range.

This function uses the standard pseudo-random generator (RND) initialized by RANDOMIZE.

#### Return value

Returns a DOUBLE representing a random floating-point number between lo and hi, inclusive.

#### Example

Generate random integers between 1 and 10
```
RANDOMIZE TIMER   ' Initialize random seed
DIM i AS LONGINT
FOR n AS LONG = 1 TO 5
   i = AfxRandomInt(1, 10)
   PRINT "Random value #" & n & ": "; i
NEXT
```
---

## AfxRemainder

Returns the IEEE 754 remainder of x / y.

```
FUNCTION AfxRemainder (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS DOUBLE
FUNCTION AfxIEEERemainder (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS DOUBLE
```

## Remarks

Formula (per IEEE 754 and .NET Math.IEEERemainder):

   r = x - (y * Q)

where Q = ROUND(x / y) using "round to nearest, ties to even".

If y = 0, the result is NaN (0 / 0).

If the result is exactly 0, the sign of x determines +0 or -0.

#### Usage examples
```
PRINT AfxRemainder(5, 2)       ' 1
PRINT AfxRemainder(7, 3)       ' 1
PRINT AfxRemainder(10, 4)      ' 2
PRINT AfxRemainder(-10, 4)     ' -2
PRINT AfxRemainder(0, 4)       ' 0
PRINT AfxRemainder(5, 0)       ' -1.#IND
```
---

## AfxRemap

Remaps a value from one range to another.

```
#define AfxRemap(v,inLo,inHi,outLo,outHi) (outLo + (v - inLo) * (outHi - outLo) / (inHi - inLo))
```

| Parameter  | Description |
| ---------- | ----------- |
| *v* | The value to be remapped. |
| *inLo* | The lower bound of the input range. |
| *inHi* | The upper bound of the input range. |
| *outLo* | The lower bound of the output range. |
| *outHi* | The upper bound of the output range. |

#### Remarks

If *v* lies within (*inLo*, *inHi*), the result lies within (*outLo*, *outHi*).

If *v* lies outside the input range, the result is extrapolated proportionally.

#### Usage example
```
PRINT AfxRemap(5, 0, 10, 0, 100)    ' 50
```
---

## AfxRnd

Returns a pseudo-random number uniformly distributed in [0, 1).

```
FUNCTION AfxRnd (BYVAL seed AS SINGLE = 1.0) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *seed* | Optional. If provided, reinitializes the random sequence.<br>0 -> use previous seed.<br>negative -> reseed with system timer.<br>positive -> set explicit seed value. |

#### Usage examples
```
PRINT AfxRnd(0.5)
PRINT AfxRnd(2.0)
PRINT AfxRnd(10.0)
```
---

## AfxRotateLeft

Performs a 32-bit left rotation of an unsigned integer.

```
FUNCTION AfxRotateLeft (BYVAL x AS ULONG, BYVAL n AS LONG) AS ULONG
    n = n AND 31
    RETURN (x SHL n) OR (x SHR (32 - n))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | 32-bit unsigned integer |
| *n* | Number of bit positions to rotate (0–31) |

#### Return value

The rotated value.

#### Usage examples
```
PRINT HEX(AfxRotateLeft(&h12345678, 8))   ' retuns 34567812
PRINT HEX(AfxRotateLeft(&h80000000, 1))   ' retuns 1
```
---

## AfxRotateRight

Performs a 32-bit right rotation of an unsigned integer.

```
FUNCTION AfxRotateRight (BYVAL x AS ULONG, BYVAL n AS LONG) AS ULONG
    n = n AND 31
    RETURN (x SHR n) OR (x SHL (32 - n))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | 32-bit unsigned integer |
| *n* | Number of bit positions to rotate (0–31) |

#### Return value

The rotated value.

## Usage examples
```
PRINT HEX(AfxRotateRight(&h12345678, 8))   ' returns 78123456
PRINT HEX(AfxRotateRight(&h00000001, 1))   ' returns 80000000
```
---

## AfxRound

Rounds a floating-point value to the nearest integer.

```
FUNCTION AfxRound (BYVAL x AS DOUBLE) AS DOUBLE
   IF x >= 0.0 THEN RETURN FIX(x + 0.5) ELSE RETURN FIX(x - 0.5)
END FUNCTION
```
#### Remaks

Halfway values (fraction = ±0.5) are rounded away from zero.

#### Usage examples
```
PRINT AfxRound(2.3)     ' 2
PRINT AfxRound(2.5)     ' 3
PRINT AfxRound(-2.3)    ' -2
PRINT AfxRound(-2.5)    ' -3
PRINT AfxRound(0.0)     ' 0
```
---

## AfxRoundCeil

Rounds a floating-point value upward.

```
FUNCTION AfxRoundCeil (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   RETURN CEIL(x * scale) / scale
END FUNCTION
```

#### Usage examples:
```
PRINT AfxRoundCeil(2.3)     ' 3
PRINT AfxRoundCeil(2.5)     ' 3
PRINT AfxRoundCeil(-2.3)    ' -2
PRINT AfxRoundCeil(-2.5)    ' -2
PRINT AfxRoundCeil(0.0)     ' 0
```
---

## AfxRoundFloor

Rounds a floating-point value downward.

```
FUNCTION AfxRoundFloor (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   RETURN FLOOR(x * scale) / scale
END FUNCTION
```

#### Usage examples
```
PRINT AfxRoundFloor(2.3)     ' 2
PRINT AfxRoundFloor(2.5)     ' 2
PRINT AfxRoundFloor(-2.3)    ' -3
PRINT AfxRoundFloor(-2.5)    ' -3
PRINT AfxRoundFloor(0.0)     ' 0
```
---

## AfxRoundHalfAwayFromZero

Rounds halfway values away from zero.

```
FUNCTION AfxRoundHalfAwayFromZero (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
```
#### Usage examples
```
PRINT AfxRoundHalfAwayFromZero(2.3)     ' 2
PRINT AfxRoundHalfAwayFromZero(2.5)     ' 3
PRINT AfxRoundHalfAwayFromZero(-2.3)    ' -2
PRINT AfxRoundHalfAwayFromZero(-2.5)    ' -3
PRINT AfxRoundHalfAwayFromZero(0.0)     ' 0
```
---

## AfxRoundHalfEven

Banker's rounding (IEEE 754 default): ties go to the nearest even integer.

```
FUNCTION AfxRoundHalfEven (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
```

#### Usage examples
```
PRINT AfxRoundHalfEven(2.3)     ' 2
PRINT AfxRoundHalfEven(2.5)     ' 2
PRINT AfxRoundHalfEven(-2.3)    ' -2
PRINT AfxRoundHalfEven(-2.5)    ' -2
PRINT AfxRoundHalfEven(0.0)     ' 0
```
---

## AfxRoundHalfUp

Rounds halfway values toward positive infinity.

```
FUNCTION AfxRoundHalfUp (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
```

#### Usage examples
```
PRINT AfxRoundHalfUp(2.3)     ' 2
PRINT AfxRoundHalfUp(2.5)     ' 3
PRINT AfxRoundHalfUp(-2.3)    ' -2
PRINT AfxRoundHalfUp(-2.5)    ' -2
PRINT AfxRoundHalfUp(0.0)     ' 0
```
---

## AfxRoundToMultiple

Rounds x to the nearest multiple of *nMultiple*.

```
FUNCTION AfxRoundToMultiple (BYVAL x AS DOUBLE, BYVAL nMultiple AS DOUBLE) AS DOUBLE
   IF nMultiple = 0.0 THEN RETURN x
   RETURN nMultiple * AfxRoundHalfEven(x / nMultiple)
END FUNCTION
```

#### Usage example
```
' Round 37 to the nearest multiple of 10
PRINT AfxRoundToMultiple(37, 10)   ' -> 40
```
---

## AfxRoundTrunc

Rounds a floating-point value truncating toward zero.

```
FUNCTION AfxRoundTrunc (BYVAL x AS DOUBLE, BYVAL nDecimals AS LONG = 0) AS DOUBLE
   DIM scale AS DOUBLE = 10.0 ^ nDecimals
   RETURN FIX(x * scale) / scale
END FUNCTION
```
#### Usage examples
```
PRINT AfxRoundTrunc(2.3)     ' 2
PRINT AfxRoundTrunc(2.5)     ' 2
PRINT AfxRoundTrunc(-2.3)    ' -2
PRINT AfxRoundTrunc(-2.5)    ' -2
PRINT AfxRoundTrunc(0.0)     ' 0
```
---

## AfxScalb
## AfxScalbn

Multiplies a floating-point number by 2 raised to the power of n.

```
#define AfxScalb(dx, e) _scalb(dx, e)
#define AfxScalbn(x, e) scalbn(x, e)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Input floating-point value. |
| *e* | Exponent (integer power of 2). |

#### Usage examples
```
PRINT AfxScalb(3.5, 4)
PRINT AfxScalbn(3.5, 4)
```
---

## AfxSec

Secant. Calculates the reciprocal of the cosine function.

```
#define AfxSec(x) (1 / Cos(x))
```

#### Usage examples
```
' Secant of 0 radians
PRINT AfxSec(0)   ' 1
' Secant of p/4 (45 degrees)
PRINT AfxSec(3.141592653589793 / 4)   ' 1.414213562373095
' Secant of p/3 (60 degrees)
PRINT AfxSec(3.141592653589793 / 3)   ' 2
```
---

## AfxSech

Calculates the hyperbolic secant.

```
#define AfxSech(x) (2 / (Exp(x) + Exp(-x)))
```

#### Usage examples
```
' Hyperbolic secant of 1
PRINT AfxSech(1)   ' 0.6480542736638855
' Hyperbolic secant of 2
PRINT AfxSech(2)   ' 0.2658022288340797
```
---

## AfxShuffle

Shuffles the elements of an array in place using the Fisher-Yates algorithm.

Each element has an equal probability of ending up in any position.

```
SUB AfxShuffle (BYVAL pArray AS ANY PTR, BYVAL nCount AS LONG, BYVAL cbElem AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pArray* | Pointer to the first element of the array to shuffle. |
| *nCount* | Number of elements in the array. |
| *cbElem* | Size (in bytes) of each element. |

#### Example

Shuffle an array of integers
```
DIM a(1 TO 10) AS LONG
FOR i AS LONG = 1 TO 10
   a(i) = i
NEXT
RANDOMIZE TIMER   ' Initialize random seed
AfxShuffle(@a(1), 10, SIZEOF(a(1)))
FOR i AS LONG = 1 TO 10
   PRINT a(i);
NEXT
```
---

## AfxSign

Determines the sign of a numeric value.

```
#define AfxSign(x) iif((x) < 0, -1, iif((x) > 0, 1, 0))
```

#### Return value

Returns -1 if the value is negative, 1 if the value is positive, and 0 if the value is exactly zero. This is useful for detecting direction or polarity in calculations, such as velocity, gradients, or comparisons.

#### Usage example
```
PRINT AfxSign(23)    ' 1
PRINT AfxSign(-37)   ' -1
PRINT AfxSign(0)     ' 0
```
---

## AfxSignbit

Determines whether the sign bit of a floating-point value is set.

```
#define AfxSignbit(x) CBOOL(__signbit(x) <> 0)
```

#### Return value

Returns 1 if x is negative or -infinity, false otherwise.

#### Usage examples
```
PRINT AfxSignbit(-3.14)        ' true
PRINT AfxSignbit(+3.14)        ' false
PRINT AfxSignbit(-1.0/0.0)     ' true   (negative infinity)
PRINT AfxSignbit( 1.0/0.0)     ' false  (positive infinity)
PRINT AfxSignbit(0.0)          ' false
PRINT AfxSignbit(-0.0)         ' true   (IEEE-754 negative zero)
```
---

## AfxSin

Calculates the sine of an angle in radians.

```
#define AfxSin(dx) Sin(dx)
```

#### Usage example

```
DIM AS DOUBLE dAngle = 1.5707963   ' p/2 radians
PRINT AfxSin(dAngle)
```
---

## AfxSinh

Calculates the hyperbolic sine of x.

```
#define AfxSinh(dx) Sinh(dx)
```

#### Usage example
```
PRINT AfxSinh(1.0)
```
---

## AfxSmoothStep

Smooth interpolation between *edge0* and *edge1* using Hermite polynomial.

```
FUNCTION AfxSmoothStep (BYVAL edge0 AS DOUBLE, BYVAL edge1 AS DOUBLE, BYVAL x AS DOUBLE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *edge0* | Lower edge of the transition interval. Values of x less than or equal to edge0 produce an output of 0. |
| *edge1* | Upper edge of the transition interval. Values of x greater than or equal to edge1 produce an output of 1. edge1 must be greater than edge0 for the interpolation to work correctly. |
| *x* | Input value to be evaluated. The function maps x smoothly from 0 to 1 as it moves from edge0 to edge1. |

#### Return value

0 when x ≤ edge0

1 when x ≥ edge1

A smooth, S‑shaped transition between 0 and 1 when x is between edge0 and edge1

#### Usage example
```
PRINT AfxSmoothStep(0, 10, 5)       ' 0.5 
```
---

## AfxSqrt

Returns the square root of a number.

```
#define AfxSqrt(x) Sqr((x))
```

#### Usahe example
```
PRINT AfxSqrt(49.0)   ' 7
```
---

## AfxStdDev

Computes the standard deviation of an array of double values.

```
FUNCTION AfxStdDev (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG, _
   BYVAL fSample AS BOOLEAN = TRUE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pData* | Pointer to the first element of the array to shuffle. |
| *n* | Number of elements in the array. |
| *fSample* | TRUE for sample standard deviation (divide by *n - 1*); FALSE for population standard deviation (divide by *n*) |

#### Return value

Returns the standard deviation.

#### Usage example
```
DIM rgData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
PRINT "Standard variance (sample)   = "; AfxStdDev(@rgData(0), 5, TRUE)
PRINT "Standard variance (population) = "; AfxStdDev(@rgData(0), 5, FALSE)
```
---

## AfxSum

Computes a compensated (Kahan) sum of an array of double values.

```
FUNCTION AfxSum (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pData* | Pointer to the first element of the data array. |
| *n* | Number of elements in the array. |

#### Return value

The precise sum as double.

#### Usage example
```
DIM rgData(4) AS DOUBLE = { 1e16, 1.0, -1e16, 3.0 }
PRINT AfxSum(@rgData(0), 4)
```
---

## AfxTan

Calculates the tangent of an angle in radians.

```
#define AfxTan(x) Tan(x)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Angle in radians. |

#### Usage example:
```
DIM AS DOUBLE dAngle = 0.7853982   ' p/4 radians
PRINT AfxTan(dAngle)
```
---

## AfxTanh

Calculates the hyperbolic tangent of x.

```
#define AfxTanh(x) Tanh(x)
```
#### Usage example:
```
PRINT AfxTanh(1.0)
```
---

## AfxTgamma

Determines the gamma function of the specified value.

```
#define AfxTgamma(x) (tgamma(x))
```

#### Usage examples
```
PRINT AfxTgamma(5.0)        ' 24
PRINT AfxTgamma(0.5)        ' 1.772453850905516
PRINT AfxTgamma(-3.2)       ' 0.6890564120059789
PRINT AfxTgamma(-2.0)       ' 1.#QNAN
```
---

## AfxTrailingZeros

Counts the number of trailing zero bits in a 64-bit unsigned integer.

```
FUNCTION AfxTrailingZeros (BYVAL x AS ULONGINT) AS LONG
```

#### Return value

Number of trailing zeros (0–64).

#### Usage example
```
PRINT AfxTrailingZeros(&h0000000000000001ULL)   ' returns 0
PRINT AfxTrailingZeros(&h8000000000000000ULL)   ' returns 63
PRINT AfxTrailingZeros(&h00F0000000000000ULL)   ' returns 52
```
---

## AfxVariance

Computes the variance of an array of double values.

Variance means how much the values ​​in a set are dispersed around its mean. It measures the average squared distance between each data point and the mean of the set.

```
FUNCTION AfxVariance (BYVAL pData AS DOUBLE PTR, BYVAL n AS LONG, _
   BYVAL fSample AS BOOLEAN = TRUE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pData* | Pointer to the first element of the data array. |
| *n* | Number of elements in the array. |
| *fSample* | TRUE for sample variance (divide by *n - 1*); FALSE for population variance (divide by *n*) |

#### Return value

The variance as double.

#### Usage example:
```
DIM rgData(4) AS DOUBLE = { 10.0, 20.0, 30.0, 40.0, 50.0 }
PRINT "Variance (sammple)   = "; AfxVariance(@rgData(0), 5, TRUE)
PRINT "Variance (population) = "; AfxVariance(@rgData(0), 5, FALSE)
```
---

