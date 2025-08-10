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
| [Arithmetic operators](#operator5) | Add, subtract, multiply or divide currency numbers. |
| [Bitwise logical operators](#operator6) | AND, EQV, IMP, MOD, OR, SHL, SHR, XOR. |

---

### Methods

| Name       | Description |
| ---------- | ----------- |
| [IsSigned](#issigned) | Returns true if this number is signed or false otherwise. |
| [IsUnsigned](#isunsigned) | Returns true if this number is unsigned or false otherwise. |
| [Sign](#sign) | Returns 0 if it is not signed of &h80 (128) if it is signed. |
| [ToVar](#tovar) | Returns the currency as a VT_CY variant. |

---

# Error Handling

The class uses the API function `SetLastError` to set error information. After calling a method of the class, you can check its success or failure calling the API function `GetLastError`. To get a localized description of the error, you can call a function such `AfxGetWinErrMsg`, passing the error code returned by `GetLastError`:

```
PRIVATE FUNCTION AfxGetWinErrMsg (BYVAL dwError AS DWORD) AS DWSTRING
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, dwError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   RETURN dwsMsg
END FUNCTION
```
---

## <a name="operator1"></a>Operator LET (=)

Assigns a value to a `CInt96` variable.
```
OPERATOR LET (BYREF cSrc AS CInt96)
OPERATOR LET (BYREF decSrc AS DECIMAL)
OPERATOR LET (BYVAL n AS LONGINT)
OPERATOR LET (BYVAL u AS ULONGINT)
OPERATOR LET (BYREF s AS WSTRING)
```
Because the bigger numeric variable natively supported by Free Basic is a long/ulong integer, if we want to set bigger values we need to use strings.

#### Usage examples

```
DIM int96 AS CInt96
int96 = "123456789001234567890"
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

## <a name="operator5"></a>Arithmetic operators

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

## <a name="operator6"></a>Bitwise logical operators

```
OPERATOR AND (BYREF rhs AS CInt96)
OPERATOR AND= (BYREF rhs AS CInt96)
OPERATOR EQV (BYREF lhs AS CInt96)
OPERATOR EQV= (BYREF lhs AS CInt96)
OPERATOR IMP (BYREF rhs AS CInt96)
OPERATOR IMP= (BYREF rhs AS CInt96)
OPERATOR MOD (BYREF rhs AS CInt96)
OPERATOR MOD= (BYREF rhs AS CInt96)
OPERATOR NOT (BYREF rhs AS CInt96)
OPERATOR OR (BYREF rhs AS CInt96)
OPERATOR OR= (BYREF rhs AS CInt96)
OPERATOR SHL (BYVAL nBits AS UINTEGER)
OPERATOR SHL= (BYVAL nBits AS UINTEGER)
OPERATOR SHR (BYVAL nBits AS UINTEGER)
OPERATOR SHR= (BYVAL nBits AS UINTEGER)
OPERATOR XOR (BYREF rhs AS CInt96)
OPERATOR XOR= (BYREF rhs AS CInt96)
```

| Name       | Description |
| ---------- | ----------- |
| **AND** | Returns the bitwise-and (conjunction) of two numeric values. |
| **EQV** | Returns the bitwise-and (equivalence) of two numeric values. |
| **IMP** | Returns the bitwise-and (implication) of two numeric values. |
| **MOD** | Returns the remainder from a division operation. |
| **NOT** | Returns the bitwise-not (complement) of a numeric value. |
| **OR** | Returns the bitwise-or (inclusive disjunction) of two numeric values. |
| **SHL** | Shifts the bits of a numeric expression to the left. |
| **SHR** | Shifts the bits of a numeric expression to the right. |
| **XOR** | Returns the bitwise-xor (exclusive disjunction) of two numeric values. |


| Name       | Description |
| ---------- | ----------- |
| **AND=** | Performs a bitwise-and (conjunction) and assigns the result to a variable, |
| **EQV=** | Performs a bitwise-eqv (equivalence) and assigns the result to a variable. |
| **IMP=** | Returns the bitwise-and (implication) of two numeric values. |
| **MOD=** | Divides this CInt96 value and assigns the remainder to it. |
| **OR=** | Performs a bitwise-or (inclusive disjunction) and assigns the result to a variable. |
| **SHL=** | Shifts left and assigns a value to a variable. |
| **SHR=** | Shifts right and assigns a value to a variable. |
| **XOR=** | Performs a bitwise-xor (exclusive disjunction) and assigns the result to a variable. |

#### Usaage examples

```
PRINT CInt96(5) AND CInt96(3) ' 1
```
```
DIM int96 AS CInt96 = 5
int96 AND= 3
print int96   ' 1
```
```
PRINT CInt96(5) EQV CInt96(3)  ' "-7"
```
```
DIM int96 AS CInt96 = 5
int96 EQV= 3
PRINT int96   ' 7
```
```
PRINT CInt96(5) IMP CInt96(3)  ' -5
```
```
DIM int96 AS CInt96 = 5
int96 IMP= 3
PRINT int96   ' -5
```
```
PRINT CInt96("5") MOD CInt96("3")   ' 2
```
```
DIM int96 AS CInt96 = -17
int96 MOD= 5
print int96   ' -2
```
```
PRINT CInt96(5) OR CInt96(3) ' 7
```
```
DIM int96 AS CInt96 = 5
int96 OR= 3
PRINT int96   ' 7
```
```
PRINT CInt96(7) SHL 4   ' 112
PRINT CInt96(1) SHL 95  ' 39614081257132168796771975168
PRINT CInt96(-3) SHL 5    ' -96
```
```
DIM int96 AS CInt96 = 7
int96 SHL= 4
PRINT int96   ' 112
```
```
PRINT CInt96(100) SHR 3   ' 12
PRINT CInt96(-100) SHR 3   ' -13
PRINT CInt96(-3) SHR 1   ' -2
PRINT CInt96(-1) SHR 100   ' -1
PRINT CInt96(5) SHR 2   ' 1
PRINT CInt96(5) SHR 3   ' 0
```
```
DIM int96 AS CInt96 = 100
int96 SHR= 3
PRINT int96   ' 12
```
```
PRINT CInt96(5) XOR CInt96(3) ' 6
```
```
DIM int96 AS CInt96 = -5
int96 XOR= 2
PRINT int96   ' 7
```
```
PRINT NOT CInt96(0)    ' -1
PRINT NOT CInt96(-1)   ' 0
```

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
