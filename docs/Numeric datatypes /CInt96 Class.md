# CInt96 Class

`CInt96` is a wrapper class for the DECIMAL data type. Holds signed 128-bit (16-byte) values representing 96-bit (12-byte) integer numbers. It supports up to 29 significant digits and can represent values in excess of 7.9228 x 10^28. The largest possible value is +/-79,228,162,514,264,337,593,543,950,335.

---

### Constructors

Creates a new instance of the `CInt96` class and assigns the values passed to it.

```
CONSTRUCTOR CInt96
CONSTRUCTOR CInt96 (BYREF cSrc AS CInt96)
CONSTRUCTOR CInt96 (BYREF decSrc AS DECIMAL)
CONSTRUCTOR CInt96 (BYVAL nInteger AS LONGINT)
CONSTRUCTOR CInt96 (BYVAL nUInteger AS ULONGINT)
CONSTRUCTOR CInt96 (BYREF wszSrc AS WSTRING)
```

#### Usage example

```
DIM int96 AS CInt96 = 1234567890
```

#### Remarks

Because the bigger numeric variable natively supported by Free Basic is a long/ulong integer, if we want to set bigger values we need to use strings, e.g.

```
DIM int96 AS CInt96 = "79228162514264337593543950335"
```
---

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator LET](#operator1) | Assigns a value to a `CInt96`variable. |
| [CAST operators](#operator2) | Converts a `CInt96` into another data type. |
| [Operator \*](#operator3) | Returns the address of the underlying `DECIMAL` structure. |
| [Comparison operators](#operator4) | Compares `CInt96` numbers. |
| [Math operators](#operator5) | Add, subtract, multiply or divide currency numbers. |

---

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Abs_](#abs_) | Returns the absolute value of a decimal data type. |
| [IsSigned](#issigned) | Returns true if this number is signed or false otherwise. |
| [IsUnsigned](#isunsigned) | Returns true if this number is unsigned or false otherwise. |
| [Sign](#sign) | Returns 0 if it is not signed of &h80 (128) if it is signed. |
| [ToVar](#tovar) | Returns the currency as a VT_CY variant. |

---

# Error Handling

The class uses the API function `SetLastError` to set error information. After calling a method of the class, you can check its success or failure calling the API function `GetLastError`. To get a localized description of the error, you can call a function such `AfxGetWinErrMsg`, passing the error code returned by `GetLastError`:

```
PRIVATE FUNCTION AfxGetWinErrMsg (BYVAL dwError AS DWORD) AS CWSTR
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, cwsMsg AS CWSTR
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, dwError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      cwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   RETURN cwsMsg
END FUNCTION
```
---

## <a name="operator1"></a>Operator LET (=)

Assigns a value to a `CInt96` variable.

Because the bigger numeric variable natively supported by Free Basic is a long/ulong integer, if we want to set bigger values we need to use strings.

#### Examples

```
DIM int96 AS CInt96
int96 = 1234567890
-or-
DIM int96 AS CInt96 = 1234567890
-or-
DIM int96 AS CInt96 = "1234567890"
```
---

## <a name="operator2"></a>CAST Operators

Converts a `CInt16` into another data type.

```
OPERATOR CAST () AS DOUBLE
OPERATOR CAST () AS CURRENCY
OPERATOR CAST () AS VARIANT
OPERATOR CAST () AS STRING
```

#### Remarks

These operators aren't called directly, they perform the conversion when the target is not a `CInt96` variable.

---

## <a name="operator3"></a>Operator *

Returns the address of the underlying `DECIMAL` structure.

```
OPERATOR * (BYREF int96 AS CInt96) AS DECIMAL PTR
```
---

## <a name="operator4"></a>Comparison operators

```
OPERATOR = (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR <> (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR < (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR > (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR <= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR >= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
```
---

## <a name="operator5"></a>Math operators

```
OPERATOR + (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
OPERATOR += (BYREF int96 AS CInt96)
OPERATOR - (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
OPERATOR -= (BYREF int96 AS CInt96)
OPERATOR * (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
OPERATOR *= (BYREF int96 AS CInt96)
OPERATOR / (BYREF int96 AS CInt96, BYVAL cOperand AS CInt96) AS CInt96
OPERATOR /= (BYREF cOperand AS CInt96)
```

#### Examples

```
DIM int96 AS CInt96 = 1234567890
int96 = int96 + 111
int96 = int96 - 111
int96 = int96 * 2
int96 = int96 / 2
int96 += 123
int96 -= 123
int96 *= 2
int96 /= 2
```

#### Remarks

The version OPERATOR - (BYREF int96 AS CInt96) AS CInt96 changes the sign of a decimal number.

```
DIM int96 AS CInt96 = 123
int96 = -int96
```
---

## <a name="abs_"></a>Abs_

Returns the absolute value of a decimal data type.

```
FUNCTION Abs_ () AS CInt96
```
---

## <a name="issigned"></a>IsSigned

Returns true if this number is signed or false otherwise.

```
FUNCTION IsSigned () AS BOOLEAN
```
---

## <a name="isunsigned"></a>IsUnsigned

Returns true if this number is unsigned or false otherwise.

```
FUNCTION IsUnsigned () AS BOOLEAN
```
---

## <a name="sign"></a>Sign

Returns 0 if it is not signed or &h80 (128) if it is signed.

```
FUNCTION Sign () AS UBYTE
```
---

## <a name="tovar"></a>ToVar

Returns the DECIMAL as a VT_DECIMAL variant.

```
FUNCTION ToVar () AS VARIANT
```

#### Remarks

Can be used to assign a currency directly to a VT_CY `VARIANT`

```
DIM int96 AS CInt96 = 1234567890
DIM v AS VARIANT = int96.ToVar
```
or a `DVARIANT`
```
DIM int96 AS CInt96 = 1234567890
DIM dv AS DVARIANT = int96.ToVar
```
---
