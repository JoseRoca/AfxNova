# DVARIANT Class

The `DVARIANT` class implements a `VARIANT` data type. The variant data type is a tagged union that can be used to represent any other data type. While lacking in efficiency, they are heavily used in COM Automation for its flexibility. The main purpose of the `DVARIANT` class is to make its use as easy as possible when you need to use them to work with COM Automation objects.

**Include file**: DVARIANT.INC.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructors) | Initialize the class with the specified value. |
| [Operators](#operators) | Procedures that perform a certain function with their operands. |
| [vType](#vtype) | Returns the VARIANT type. |
| [Attach](#attach) | Attaches a variant to the class. |
| [Detach](#detach) | Detaches the variant data from this class and transfers ownership to the passed variant. |
| [ChangeType](#changetype) | Converts the variant from one type to another. |
| [ChangeTypeEx](#changetypeex) | Converts the variant from one type to another. |
| [GetDim](#getdim) | Gets the number of dimensions in the array. |
| [GetLBound](#getlbound) | Gets the lower bound for the specified dimension of the safe array. |
| [GetUBound](#getubound) | Gets the upper bound for the specified dimension of the safe array. |
| [GetElementCount](#getelementcount) | Gets the number of elements in the array. |
| [DecToCY](#dectocy) | Converts a DVARIANT of type decimal to a CY structure. |
| [DecToDouble](#dectodouble) | Converts a DVARIANT of type decimal to a double. |
| [Round](#round) | Rounds a variant to the specified number of decimal places. |
| [FormatNumber](#formatnumber) | Formats a DVARIANT containing numbers into a string form. |
| [GetBooleanElement](#getbooleanrlem) | Extracts a single boolean element from a safe array of booleans. |
| [GetDoubleElement](#getdoublerlem) | Extracts a single DOUBLE element from a safe array of doubles. |
| [GetLongElement](#getlongelem) | Extracts a single LONG element from a safe array of longs. |
| [GetLongIntElement](#getlongintrlem) | Extracts a single LONGINT element from a safe array of long integers. |
| [GetShortElement](#getshortrlem) | Extracts a single SHORT element from a safe array of shorts. |
| [GetStringElement](#getstringelem) | Extracts a single BSTR element from a safe array of unicode strings. |
| [GetULongElement](#getulongelem) | Extracts a single ULONG element from a safe array of unsigned longs. |
| [GetULongIntElement](#getulongintelem) | Extracts a single ULONGINT element from a safe array of unsigned long integers. |
| [GetUShortElement](#getushortelem) | Extracts a single USHORT element from a safe array of unsigned shorts. |
| [GetVariantElement](#getvariantelem) | Extracts a single Variant element from a safe array of variants. |
| [Put](#put) | Assigns values to a DVARIANT. |
| [PutNull](#putnull) | Assigns a null value. |
| [PutBool](#putbool) | Assigns a boolean value. |
| [PutBoolean](#putboolean) | Assigns a boolean value. |
| [PutByte](#putbyte) | Assigns a byte value. |
| [PutUByte](#putubyte) | Assigns an ubyte value. |
| [PutShort](#putshort) | Assigns a short value. |
| [PutUShort](#putushort) | Assigns an ushort value. |
| [PutInt](#putint) | Assigns an int_ value. |
| [PutUInt](#putuint) | Assigns an uint value. |
| [PutLong](#putlong) | Assigns a long value. |
| [PutULong](#putulong) | Assigns an ulong value. |
| [PutLongInt](#putlongint) | Assigns a longint value. |
| [PutULongInt](#putulongint) | Assigns an ulongint value. |
| [PutSingle](#putsingle) | Assigns a single value. |
| [PutFloat](#putfloat) | Assigns a single value. |
| [PutDouble](#putdouble) | Assigns a double value. |
| [PutBooleanArray](#putbooleanarray) | Initializes DVARIANT from an array of Boolean values. |
| [PutShortArray](#putshortarray) | Initializes DVARIANT from an array of signed 16-bit integer values. |
| [PutUShortArray](#putushortarray) | Initializes DVARIANT from an array of unsigned 16-bit integer values. |
| [PutLongArray](#putlongarray) | Initializes DVARIANT from an array of signed 32-bit integer values. |
| [PutULongArray](#putulongarray) | Initializes DVARIANT from an array of 32-bit unsigned integer values. |
| [PutLongIntArray](#putlongintarray) | Initializes DVARIANT from an array of signed 64-bit integer values. |
| [PutULongIntArray](#putulongintarray) | Initializes DVARIANT from an array of unsigned 64-bit integer values. |
| [PutDoubleArray](#putdoublearray) | Initializes DVARIANT from an array of unsigned 64-bit integer values. |
| [PutStringArray](#putstringarray) | Initializes DVARIANT from an array of unsigned 64-bit integer values. |
| [PutBuffer](#putbuffer) | Initializes DVARIANT with the contents of a buffer. |
| [PutDateString](#putdatestring) | Initializes DVARIANT VT_DATE from a string. |
| [PutDec](#putdec) | Initializes DVARIANT with the contents of a DECIMAL structure. |
| [PutDecFromCY](#putdecfromcy) | Converts a currency value to a variant of type VT_DECIMAL. |
| [PutDecFromDouble](#putdecfromdouble) | Converts a double value to a variant of type VT_DECIMAL. |
| [PutDecFromStr](#putdecfromstr) | Initializes DVARIANT as VT_DECIMAL from a string. |
| [PutFileTime](#putfiletime) | Initializes DVARIANT with the contents of a FILETIME structure. |
| [PutFileTimeArray](#putfiletimearray) | Initializes DVARIANT with an array of FILETIME structures. |
| [PutGuid](#putguid) | Initializes DVARIANT from a GUID. |
| [PutPropVariant](#putpropvariant) | Initializes DVARIANT from the contents of a PROPVARIANT structure. |
| [PutRecord](#putrecord) | Initializes DVARIANT with a reference to an UDT. |
| [PutRef](#putref) | Assigns a value by reference (a pointer to a variable). |
| [PutResource](#putresource) | Initializes the DVARIANT based on a string resource imbedded in an executable file. |
| [PutSafeArray](#putsafearray) | Initializes DVARIANT from a safe array. |
| [PutStrRet](#putstrret) | Initializes DVARIANT with string stored in a STRRET structure. |
| [PutSystemTime](#putsystemtime) | Initializes DVARIANT with the contents of a SYSTEMTIME structure. |
| [PutUtf8](#pututf8) | Initializes DVARIANT with the contents of an UTF-8 string. |
| [PutVariantArrayElement](#putvariantarrayelem) | Initializes DVARIANT with a value stored in another VARIANT structure. |
| [PutVbDate](#putvbdate) | Initializes DVARIANT with the contents of a DATE value. |
| [ToBooleanArray](#tobooleanarray) | Extracts an array of boolean values from DVARIANT. |
| [ToBooleanArrayAlloc](#tobooleanarrayalloc) | Extracts an array of boolean values from DVARIANT. |
| [ToBstr](#tobstr) | Extracts the content of the underlying variant and returns it as a BSTRING. |
| [ToBuffer](#tobuffer) | Extracts the contents of a DVARIANT of type VT_ARRRAY OR VT_UI1 to a buffer. |
| [ToBuffer (STRING)](#tobufferstring) | Extracts the contents of a DVARIANT of type VT_ARRRAY OR VT_UI1 to a string used as a buffer. |
| [ToDosDateTime](#todosdatetime) | Extracts a date and time value in Microsoft MS-DOS format from a DVARIANT of type VT_DATE. |
| [ToDoubleArray](#todoublearray) | Extracts an array of DOUBLE values from DVARIANT. |
| [ToDoubleArrayAlloc](#todoublearrayalloc) | Extracts an array of DOUBLE values from DVARIANT. |
| [ToFileTime](#tofiletime) | Returns the contents of a DVARIANT of type VT_DATE as a FILETIME structure. |
| [ToGuid](#ToGuid) | Returns the contents of a DVARIANT containing a GUID string as a GUID structure. |
| [ToGuidBStr](#toguidbdtr) | Returns the contents of a DVARIANT containing a GUID string as an unicode GUID string. |
| [ToGuidStr](#toguidstr) | Returns the contents of a DVARIANT containing a GUID string as an unicode GUID string. |
| [ToGuidWStr](#toguidwstr) | Returns the contents of a DVARIANT containing a GUID string as an unicode GUID string. |
| [ToLongArray](#tolongarray) | Extracts an array of LONG values from DVARIANT. |
| [ToLongArrayAlloc](#tolongarrayalloc) | Extracts an array of LONG values from DVARIANT. |
| [ToLongIntArray](#tolongintarray) | Extracts an array of LONGINT values from DVARIANT. |
| [ToLongIntArrayAlloc](#tolongintarrayalloc) | Extracts an array of LONGINT values from DVARIANT. |
| [ToShortArray](#toshortarray) | Extracts an array of Int16 values from DVARIANT. |
| [ToShortArrayAlloc](#toshortarrayalloc) | Extracts an array of SHORT values from DVARIANT. |
| [ToStr](#tostr) | Extracts the content of the underlying variant and returns it as a DWSTRING. |
| [ToStringArray](#tostringarray) | Extracts data from a vector structure into a PWSTR array. |
| [ToStringArrayAlloc](#tostringarrayalloc) | Extracts an array of PWSTR values from DVARIANT. |
| [ToStrRet](#tostrret) | Returns the contents of a DVARIANT of type VT_BSTR to a STRRET stucture. |
| [ToSystemTime](#tosystemtime) | Returns the contents of DVARIANT of type VT_DATE as a FILETIME structure. |
| [ToULongArray](#toulongarray) | Extracts an array of ULONG values from DVARIANT. |
| [ToULongArrayAlloc](#toulongarrayalloc) | Extracts an array of ULONG values from DVARIANT. |
| [ToULongIntArray](#toulongintarray) | Extracts an array of ULONGINT values from DVARIANT. |
| [ToULongIntArrayAlloc](#toulongintarrayalloc) | Extracts an array of ULONGINT values from DVARIANT. |
| [ToUShortArray](#toushortarray) | Extracts an array of USHORT values from DVARIANT. |
| [ToUShortArrayAlloc](#toushortarrayalloc) | Extracts an array of USHORT values from DVARIANT. |
| [ToUtf8](#toutf8) | Returns the contents of a DVARIANT containing a BSTR as an UTF-8 encoded string. |
| [ToVbDate](#tovbdate) | Returns the contents of a DVARIANT of type VT_DATE as a DATE value. |

#### Helper Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxDVarToStr](#AfxDVarToStr) | Extracts the contents of a DVARIANT to a DWSTRING. |
| [AfxDVarToBuffer](#AfxDVarToBuffer) | Extracts the contents of a variant that contains an array of bytes. |
| [AfxDVarToOptPrm](#AfxDVarToOptPrm) | Returns a DVARIANT suitable to be used with optional parameters. |

---

# <a name="constructors"></a>Constructors

Creates an instance of the DVARIANT class.

```
CONSTRUCTOR
CONSTRUCTOR (BYREF dv AS DVARIANT)
CONSTRUCTOR (BYVAL v AS VARIANT)
CONSTRUCTOR (BYVAL pwsz AS WSTRING PTR)
CONSTRUCTOR (BYREF dws AS DWSTRING)
CONSTRUCTOR (BYREF bs AS BSTRING)
CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)
CONSTRUCTOR (BYVAL cy AS CURRENCY)
CONSTRUCTOR (BYVAL dec AS DECIMAL)
DECLARE CONSTRUCTOR (BYVAL _value AS LONG)
CONSTRUCTOR (BYREF pDisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
CONSTRUCTOR (BYREF pUnk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
CONSTRUCTOR (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
CONSTRUCTOR (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
CONSTRUCTOR (BYVAL _value AS LONGINT, BYREF strType AS STRING)
CONSTRUCTOR (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD)
CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cv* | A DVARIANT. |
| *v* | A VARIANT. |
| *pvar* | Pointer to a VARIANT. |
| *cy* | A currency structure. |
| *dec* | A decimal structure. |
| *b* | A boolean value (TRUE or FALSE). |
| *pwsz* | Pointer to an unicode string. You can also pass a Free Basic ansi string or a string literal. |
| *bs* | A BSTRING. |
| *dws* | A DWSTRING. |
| *pDisp* | Pointer to a DISPATCH interface. |
| *pUnk* | Pointer to a UNKNOWN interface. |
| *_value* | A numeric value or variable. |
| *_pvar* | Pointer to a variable. This will create a VT_BYREF variant of the specified type. |
| *_vtype* | The variant type, e.g. VT_I4, VT_UI4. |
| *strType* | The variant type as a string: "BOOL", "BYTE", "UBYTE", "SHORT", "USHORT, "INT", UINT", "LONG", "ULONG", "LONGINT", "SINGLE, "DOUBLE", "NULL". |
| *fAddRef* | TRUE or FALSE. If TRUE, increases the reference count of the passed interface. |

#### Examples
```
DIM dv AS DVARIANT = "Test string"   ' Creates a VT_BSTR (8) variant
DIM dv AS DVARIANT = 12345           ' Creates a VT_I4 (3) variant
DIM dv AS DVARIANT = 123456.78       ' Creates a VT_R8 (5) variant
```
We can use the constructors to pass values to parameters in procedures without assigning them first to a variable, e.g.:
```
SUB Foo (BYREF dv AS DVARIANT)
   PRINT dv
END SUB

Foo "Test string"
Foo 12345
Foo DVARIANT(12345, "LONG")
```
```
SUB Foo (BYVAL dv AS DVARIANT PTR)
   PRINT AfxDVarToStr(dv)
END SUB
Foo @DVARIANT("Test string")
Foo @DVARIANT(12345)
Foo @DVARIANT(12345, "LONG")```
```
```
SUB Foo (BYVAL v AS VARIANT PTR)
   PRINT AfxVarToStr(v)
END SUB
Foo DVARIANT("Test string")
Foo DVARIANT(12345)
Foo DVARIANT(12345, "LONG")
```
---

# <a name="operators"></a>Operators

Procedures that perform a certain function with their operands. They do the same actions that the native FreeBasic operators but with variants. For detailed descriptions see the FreeBasic documentation.

Global Operators

```
OPERATOR & (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
OPERATOR * (BYREF dv AS DVARIANT) AS VARIANT PTR
```

Cast Operators

```
OPERATOR Cast () AS VARIANT
OPERATOR Cast () AS VARIANT PTR
OPERATOR CAST () BYREF AS WSTRING
```

Assignment operators

```
OPERATOR Let (BYREF dv AS DVARIANT)
OPERATOR Let (BYVAL v AS VARIANT)
OPERATOR Let (BYVAL pvar AS VARIANT PTR)
OPERATOR LET (BYVAL pwszStr AS WSTRING PTR)
OPERATOR Let (BYVAL cy AS CURRENCY)
OPERATOR Let (BYVAL dec AS DECIMAL)
OPERATOR Let (BYREF bs AS BSTRING)
OPERATOR Let (BYREF dws AS DWSTRING)
DECLARE OPERATOR LET (BYVAL _value AS LONG)
OPERATOR Let (BYREF pDisp AS IDispatch PTR)
OPERATOR Let (BYREF pUnk AS IUnknown PTR)
OPERATOR Let (BYVAL _value AS LONGINT)
OPERATOR Let (BYVAL _value AS DOUBLE)
OPERATOR += (BYREF dv AS DVARIANT)
OPERATOR -= (BYREF dv AS DVARIANT)
OPERATOR *= (BYREF dv AS DVARIANT)
OPERATOR /= (BYREF dv AS DVARIANT)
OPERATOR \= (BYREF dv AS DVARIANT)
OPERATOR Mod= (BYREF dv AS DVARIANT)
OPERATOR Imp= (BYREF dv AS DVARIANT)
OPERATOR Eqv= (BYREF dv AS DVARIANT)
OPERATOR ^= (BYREF dv AS DVARIANT)
```

Arithmetic operators

```
OPERATOR + (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
OPERATOR - (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
OPERATOR * (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
OPERATOR / (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
OPERATOR \ (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DVARIANT
OPERATOR ^ (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS DOUBLE
OPERATOR Mod (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
OPERATOR - (BYREF dv AS DVARIANT) AS DVARIANT
```

Relational operators

```
OPERATOR = (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
OPERATOR <> (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
OPERATOR < (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
OPERATOR > (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
OPERATOR <= (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
OPERATOR >= (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS BOOLEAN
```

Bitwise operators

```
OPERATOR And (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
OPERATOR Eqv (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
OPERATOR Imp (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
OPERATOR Not (BYREF dv AS DVARIANT) AS INTEGER
OPERATOR Or (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
OPERATOR Xor (BYREF dv1 AS DVARIANT, BYREF dv2 AS DVARIANT) AS INTEGER
```
```
SUB Foo (BYVAL v AS VARIANT)
   PRINT AfxVarToStr(@v)
END SUB
DIM dv AS DVARIANT = "Test string"
Foo dv
```
Using the pointer syntax:
```
SUB Foo (BYVAL v AS VARIANT)
   PRINT AfxVarToStr(@v)
END SUB
DIM pdv AS DVARIANT PTR  = NEW DVARIANT("Test string")
Foo *pdv
Delete pdv
```
Using the constructors
```
SUB Foo (BYVAL v AS VARIANT)
   PRINT AfxVarToStr(@v)
END SUB
Foo DVARIANT(12345, "LONG")
```

#### Note to the CAST operator

The CAST operators allow to transparently pass the underlying VARIANT to a procedure.
They aren't called directly.

```
SUB Foo (BYREF v AS VARIANT)
   PRINT AfxVarToStr(@v)
END SUB
Foo DVARIANT(12345, "LONG")
```
```
SUB Foo2 (BYVAL v AS VARIANT)
   PRINT AfxVarToStr(@v)
END SUB
Foo2 DVARIANT(12345, "LONG")
```
```
SUB Foo3 (BYREF dv AS DVARIANT)
   PRINT dv
END SUB
Foo3 DVARIANT(12345, "LONG")
```
```
SUB Foo4 (BYVAL dv AS DVARIANT PTR)
   PRINT AfxDVarToStr(dv)
END SUB
Foo4 @DVARIANT(12345, "LONG")
```

#### Remarks

I haven't added a cast to return a numeric value because with procedures like PRINT that can use both a number or a string the compiler will fail, not knowing which cast it should use. If you want to convert it to a number, use VAL(DVARIANT).

---

## <a name="vtype"></a>vType

Returns the VARIANT type.

```
FUNCTION vType () AS VARTYPE
```

The following table shows the available data types and where these values can be used.

| Type                | Description                              | VARIANT | Typedesc | Property set | Safe array |
| ------------------- | ---------------------------------------- | ------- | -------- | ------------ | ---------- |
| VT_EMPTY            | Not specified.                           |    X    |          |      X       |            |
| VT_NULL             | Null.                                    |    X    |          |      X       |            |
| VT_I1               | BYTE. A character.                       |    X    |    X     |      X       |     X      |
| VT_UI1              | UBYTE. An unsigned character.            |    X    |    X     |      X       |     X      |
| VT_I2               | SHORT. A 2-byte integer.                 |    X    |    X     |      X       |     X      |
| VT_UI2              | USHORT. An unsigned short.               |    X    |    X     |      X       |     X      |
| VT_I4               | LONG. A 4-byte integer.                  |    X    |    X     |      X       |     X      |
| VT_UI4              | ULONG. An unsigned long.                 |    X    |    X     |      X       |     X      |
| VT_I8               | LONGINT. A 64-bit integer.               |    X    |    X     |      X       |     X      |
| VT_UI8              | ULONGINT. A 64-bit unsigned integer.     |    X    |    X     |      X       |     X      |
| VT_INT              | LONG. An integer.                        |    X    |    X     |      X       |     X      |
| VT_UINT             | ULOG. An unsigned integer.               |    X    |    X     |              |     X      |
| VT_R4               | SINGLE. A 4-byte real.                   |    X    |    X     |      X       |     X      |
| VT_R8               | DOUBLE. A 8-byte real.                   |    X    |    X     |      X       |     X      |
| VT_CY               | CY. Currency.                            |    X    |    X     |      X       |     X      |
| VT_DATE             | DOUBLE. A date.                          |    X    |    X     |      X       |     X      |
| VT_BSTR             | BSTR. A string.                          |    X    |    X     |      X       |     X      |
| VT_DISPATCH         | IDispatch PTR. An IDispatch pointer.     |    X    |    X     |              |     X      |
| VT_ERROR            | SCODE. An SCODE value.                   |    X    |    X     |      X       |     X      |
| VT_BOOL             | BOOLEAN. A Boolean value<br>(True = -1, False = 0)   |    X    |    X     |      X       |     X      |
| VT_VARIANT          | VARIANT PTR. A variant pointer.          |    X    |    X     |      X       |     X      |
| VT_UNKNOWN          | IUnknown PTR. An IUnknown pointer.       |    X    |    X     |              |     X      |
| VT_DECIMAL          | DECIMAL PTR. A 16-byte fixed-pointer value.           |    X    |    X     |              |     X      |
| VT_VOID             | NULL. A C-style void.                    |         |          |      X       |            |
| VT_HRESULT          | HRESULT. An HRESULT value.               |         |          |      X       |            |
| VT_PTR              | A pointer type.                          |         |          |      X       |            |
| VT_SAFEARRAY        | SAFEARRAY PTR. A safe array.<br>Use VT_ARRAY in VARIANT.   |         |    X     |              |            |
| VT_CARRAY           | A C-style array.                         |         |    X     |              |            |
| VT_USERDEFINED      | A user-defined type.                     |         |          |      X       |            |
| VT_LPSTR            | ZSTRING. A null-terminated string.       |         |    X     |      X       |            |
| VT_LPWSTR           | WSTRING. A wide null-terminated string.  |         |    X     |      X       |            |
| VT_RECORD           | A user-defined type.                     |    X    |    X     |              |     X      |
| VT_INT_PTR          | A signed machine register size width.    |         |          |      X       |            |
| VT_UINT_PTR         | An unsigned machine register size width. |         |          |      X       |            |
| VT_FILETIME         | FILETIME. A FILETIME value.              |         |          |      X       |            |
| VT_BLOB             | Length-prefixed bytes.                   |         |          |      X       |            |
| VT_STREAM           | IStream PTR. The name of the stream follows.          |         |          |      X       |            |
| VT_STORAGE          | The name of the storage follows.         |         |          |      X       |            |
| VT_STREAMED_OBJECT  | The stream contains an object.           |         |          |      X       |            |
| VT_STORED_OBJECT    | The storage contains an object.          |         |          |      X       |            |
| VT_BLOB_OBJECT      | The blob contains an object.             |         |          |      X       |            |
| VT_CF               | A clipboard format.                      |         |          |      X       |            |
| VT_CLSID            | CLSID. A class ID.                       |         |          |      X       |            |
| VT_VERSIONED_STREAM | A stream with a GUID version.            |         |          |      X       |            |
| VT_BSTR_BLOB        | Reserved for system use.                 |         |          |              |            |
| VT_VECTOR           | A simple counted array.                  |         |          |      X       |            |
| VT_ARRAY            | SAFEARRAY PTR. A SAFEARRAY pointer.      |    X    |          |              |            |
| VT_BYREF            | A void pointer for local use.            |    X    |          |              |            |
| VT_RESERVED         | Reserved.                                |         |          |              |            |
| VT_ILLEGAL          |                                          |         |          |              |            |
| VT_ILLEGALMASKED    |                                          |         |          |              |            |
| VT_TYPEMASK         |                                          |         |          |              |            |

---

## <a name="attach"></a>Attach

Attaches a variant to the class.

```
FUNCTION Attach (BYVAL pvar AS VARIANT PTR) AS HRESULT
FUNCTION Attach (BYREF v AS VARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvar* | Pointer to the variant to attach. |
| *v* | The variant to attach. |

#### Remark

Marks the source variant as VT_EMPTY instead of clearing it with **VariantClear** because we aren't making a duplicate of the contents, but transferring ownership.

#### Return value

Returns S_OK (0) on success or an HRESULT error code on failure.

---

## <a name="detach"></a>Detach

Detaches the variant data from this class and transfers ownership to the passed variant.

```
FUNCTION Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
FUNCTION Detach (BYREF v AS VARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvar* | Pointer to the variant where the contents of the variant data will be moved. |
| *v* | Variant where the contents of the variant data will be moved. |

#### Remark

This method transfers ownership of the underlying variant and marks it as empty.

#### Return value

Returns S_OK (0) or an HRESULT error code.

---

## <a name="changetype"></a>ChangeType

Converts the variant from one type to another.

```
FUNCTION ChangeType (BYVAL vtNew AS VARTYPE, BYVAL wFlags AS USHORT = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vtNew* | The new variant type. |
| *wFlags* | *VARIANT_NOVALUEPROP* : Prevents the function from coercing an object to a fundamental type by getting the Value property. Applications should set this flag only if necessary, because it makes their behavior inconsistent with other applications.<br>*VARIANT_ALPHABOOL* : Converts a VT_BOOL value to a string containing either "True" or "False".<br>*VARIANT_NOUSEROVERRIDE* : For conversions to or from VT_BSTR, passes LOCALE_NOUSEROVERRIDE to the core coercion routines.<br>*VARIANT_LOCALBOOL* : For conversions from VT_BOOL to VT_BSTR and back, uses the language specified by the locale in use on the local computer. |

#### Return value

Returns S_OK (0) or an HRESULT error code.

---

## <a name="changetypeex"></a>ChangeTypeEx

Converts the variant from one type to another.

```
FUNCTION ChangeTypeEx (BYVAL vtNew AS VARTYPE, BYVAL lcid AS LCID = 0, BYVAL wFlags AS USHORT = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vtNew* | The new variant type. |
| *lcid* | The locale identifier. The LCID is useful when the type of the source or destination VARIANTARG is VT_BSTR, VT_DISPATCH, or VT_DATE. |
| *wFlags* | *VARIANT_NOVALUEPROP* : Prevents the function from attempting to coerce an object to a fundamental type by getting the Value property. Applications should set this flag only if necessary, because it makes their behavior inconsistent with other applications.<br>*VARIANT_ALPHABOOL* : Converts a VT_BOOL value to a string containing either "True" or "False".<br>*VARIANT_NOUSEROVERRIDE* : For conversions to or from VT_BSTR, passes LOCALE_NOUSEROVERRIDE to the core coercion routines.<br>*VARIANT_LOCALBOOL* : For conversions from VT_BOOL to VT_BSTR and back, uses the language specified by the locale in use on the local computer. |

#### Return value

Returns S_OK (0) or an HRESULT error code.

---

## <a name="getdim"></a>GetDim

Gets the number of dimensions in the array.

```
FUNCTION GetDim () AS ULONG
```

#### Return value

Returns the number of dimensions for variants of type VT_ARRAY; returns 0 otherwise.

---

## <a name="getlbound"></a>GetLBound

Gets the lower bound for the specified dimension of the safe array.

```
FUNCTION GetLBound (BYVAL nDim AS UINT = 1) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDim* | The dimension of the array. |

#### Return value

Returns the lower bound for the specified dimension of the safe array for variants of type VT_ARRAY; returns 0 otherwise.

---

## <a name="getubound"></a>GetUBound

Gets the upper bound for the specified dimension of the safe array.

```
FUNCTION GetUBound (BYVAL nDim AS UINT = 1) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDim* | The dimension of the array. |

#### Return value

Returns the upper bound for the specified dimension of the safe array for variants of type VT_ARRAY; returns 0 otherwise.

---

## <a name="getelementcount"></a>GetElementCount

Gets the number of elements in the array.

```
FUNCTION GetElementCount () AS ULONG
```

#### Return value

Returns the number of elements for variants of type VT_ARRAY; returns 1 otherwise.

---

## <a name="dectocy"></a>DecToCY

Converts a DVARIANT of type decimal to a CY structure.

```
FUNCTION DecToCY () AS CY
```

#### Return value

Returns the contents of a VT_DECIMAL variant as a CY structure.

---

## <a name="dectodouble"></a>DecToDouble

Converts a DVARIANT of type decimal to a double.

```
FUNCTION DecToDouble () AS DOUBLE
```

#### Return value

Returns the contents of a VT_DECIMAL variant as a DOUBLE.

---

## <a name="round"></a>Round

Rounds a variant to the specified number of decimal places.

```
FUNCTION Round (BYREF dv AS DVARIANT, BYVAL cDecimals AS LONG) AS DVARIANT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dv* | The DVARIANT to round. |
| *cDecimals* | The number of decimal places. |

#### Return value

A DVARIANT containing the rounded result.

---

## <a name="formatnumber"></a>FormatNumber

Formats a DVARIANT containing numbers into a string form.

```
FUNCTION FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *iNumDig* | The number of digits to pad to after the decimal point. Specify -1 to use the system default value. |
| *ilncLead* | Specifies whether to include the leading digit on numbers.<br>-2 : Use the system default.<br>-1 : Include the leading digit.<br> 0 : Do not include the leading digit. |
| *iUseParens* | Specifies whether negative numbers should use parentheses.<br>-2 : Use the system default.<br>-1 : Use parentheses.<br>0 : Do not use parentheses. |
| *iGroup* | Specifies whether thousands should be grouped. For example 10,000 versus 10000.<br>-2 : Use the system default.<br>-1 : Group thousands.<br> 0 : Do not group thousands. |
| *dwFlags* | VAR_CALENDAR_HIJRI is the only flag that can be set. |

#### Return value

A DWSTRING containing the formatted value.

#### Remarks

This function uses the user's default locale while calling **VarTokenizeFormatString** and **VarFormatFromTokens**.

---

## <a name="getbooleanelem"></a>GetBooleanElement

Extracts a single boolean element from a safe array of booleans.

```
FUNCTION GetBooleanElement (BYVAL iElem AS ULONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getdoubleelem"></a>GetDoubleElement

Extracts a single DOUBLE element from a safe array of doubles.

```
FUNCTION GetDoubleElement (BYVAL iElem AS ULONG) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getlongelem"></a>GetLongElement

Extracts a single LONG element from a safe array of longs.

```
FUNCTION GetLongElement (BYVAL iElem AS ULONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getlongintelem"></a>GetLongIntElement

Extracts a single LONGINT element from a safe array of long integers.

```
FUNCTION GetLongIntElement (BYVAL iElem AS ULONG) AS LONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getshortelem"></a>GetShortElement

Extracts a single SHORT element from a safe array of short integers.

```
FUNCTION GetShortElement (BYVAL iElem AS ULONG) AS SHORT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getstringelem"></a>GetStringElement

Extracts a single BSTR element from a safe array of unicode strings.

```
FUNCTION GetStringElement (BYVAL iElem AS ULONG) AS BSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getulongelem"></a>GetULongElement

Extracts a single ULONG element from a safe array of unsigned longs.

```
FUNCTION GetULongElement (BYVAL iElem AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getulongintelem"></a>GetULongIntElement

Extracts a single ULONGINT element from a safe array of unsigned long integers.

```
FUNCTION GetULongIntElement (BYVAL iElem AS ULONG) AS ULONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getushortelem"></a>GetUShortElement

Extracts a single USHORT element from a safe array of unsigned shorts.

```
FUNCTION GetUShortElement (BYVAL iElem AS ULONG) AS USHORT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="getvariantelem"></a>GetVariantElement

Extracts a single Variant element from a safe array of variants.

```
FUNCTION GetVariantElement (BYVAL iElem AS ULONG) AS DVARIANT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

---

## <a name="put"></a>Put

Assigns values to a DVARIANT.

```
SUB Put (BYREF wsz AS WSTRING)
SUB Put (BYREF dws AS DWSTRING)
SUB Put (BYREF bs AS BSTRING)
FUNCTION Put (BYREF dv AS DVARIANT) AS HRESULT
FUNCTION Put (BYVAL v AS VARIANT) AS HRESULT
FUNCTION Put (BYVAL pvar AS VARIANT PTR) AS HRESULT
FUNCTION Put (BYREF pDisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
FUNCTION Put (BYREF pUnk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
SUB Put (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
SUB Put (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
SUB Put (BYVAL _value AS LONGINT, BYREF strType AS STRING)
SUB Put (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
SUB Put (BYVAL _pv AS ANY PTR, BYVAL _vType AS WORD)
SUB Put (BYVAL _pv AS ANY PTR, BYREF strType AS STRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wsz* | A Unicode string. You can also pass a FreeBasic ansi string or a string literal. |
| *dws* | A DWSTRING variable. |
| *bs* | A BSTRING variable. |
| *dv* | A DVARIANT variable. |
| *v* | A VARIANT variable. |
| *v* | A VARIANT variable. |
| *pvar* | Pointer to a VARIANT variable. |
| *pDisp* | Pointer to a DISPATCH interface. |
| *pUnk* | Pointer to a UNKNOWN interface. |
| *_pv* | Pointer to a variable. This will create a VT_BYREF variant of the specified type. |
| *_vtype* | The variant type, e.g. VT_I4, VT_UI4. |
| *strType* | The variant type as a string: "BOOL", "BYTE", "UBYTE", "SHORT", "USHORT, "INT", UINT", "LONG", "ULONG", "LONGINT", "SINGLE, "DOUBLE", "NULL". |
| *fAddRef* | TRUE or FALSE. If TRUE, increases the reference count of the passed interface. |

---

## <a name="putnull"></a>PutNull

Assigns a null value to the DVARIANT.

```
SUB PutNull
```
---

## <a name="putbool"></a>PutBool

Assigns a boolean value to the DVARIANT.

```
SUB PutBool (BYVAL _value AS BOOL)
```
---
## <a name="PutBoolean"></a>PutBoolean

Assigns a boolean value to the DVARIANT.

```
SUB PutBoolean (BYVAL _value AS BOOLEAN)
```
---

## <a name="putbyte"></a>PutByte

Assigns a byte value to the DVARIANT.

```
SUB PutByte (BYVAL _value AS BYTE)
```
---

## <a name="putubyte"></a>PutUByte

Assigns an unsigned ubyte value to the DVARIANT.

```
SUB PutUByte (BYVAL _value AS UBYTE)
```
---

## <a name="putshort"></a>PutShort

Assigns a short integer value to the DVARIANT.

```
SUB PutShort (BYVAL _value AS SHORT)
```
---

## <a name="putushort"></a>PutUShort

Assigns an unsigned short integer value to the DVARIANT.

```
SUB PutUShort (BYVAL _value AS USHORT)
```
---

## <a name="putint"></a>PutInt

Assigns an INT_ (long) value to the DVARIANT.

```
SUB PutInt (BYVAL _value AS INT_)
```

### Remark

Don't confuse an INT_ (LONG) with the Free Basic INTEGER data type.

---

## <a name="putuint"></a>PutUInt

Assigns an UINT (unsigned long) value to the DVARIANT.

```
SUB PutUInt (BYVAL _value AS UINT)
```

### Remark

Don't confuse an UINT (ULONG) with the Free Basic UINTEGER data type.

---

## <a name="putlong"></a>PutLong

Assigns a LONG value to the DVARIANT.

```
SUB PutLong (BYVAL _value AS LONG)
```
---

## <a name="putulong"></a>PutULong

Assigns a ULONG value to the DVARIANT.

```
SUB PutULong (BYVAL _value AS ULONG)
```
---

## <a name="putlongint"></a>PutLongInt

Assigns a LONGINT value to the DVARIANT.

```
SUB PutULong (BYVAL _value AS LONGINT)
```
---

## <a name="putulongint"></a>PutULongInt

Assigns a ULONGINT value to the DVARIANT.

```
SUB PutULongInt (BYVAL _value AS ULONGINT)
```
---

## <a name="putsingle"></a>PutSingle

Assigns a SINGLE value to the DVARIANT.

```
SUB PutSingle (BYVAL _value AS SINGLE)
```
---

## <a name="putfloat"></a>PutFloat

Assigns a SINGLE value to the DVARIANT.

```
SUB PutFloat (BYVAL _value AS SINGLE)
```
---

## <a name="putdouble"></a>PutDouble

Assigns a DOUBLE value to the DVARIANT.

```
SUB PutDouble (BYVAL _value AS DOUBLE)
```
---

## <a name="putbooleanarray"></a>PutBooleanArray

Initializes DVARIANT from an array of Boolean values.

```
FUNCTION PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of Boolean values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_BOOL variant.

---

## <a name="putshortarray"></a>PutShortArray

Initializes DVARIANT from an array of signed 16-bit integer values.

```
FUNCTION PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of SHORT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_I2 variant.

---

## <a name="putushortarray"></a>PutUshortArray

Initializes DVARIANT from an array of 16-bit unsigned integer values.

```
FUNCTION PutUshortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of USHORT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI2 variant.

---

## <a name="putlongarray"></a>PutLongArray

Initializes DVARIANT from an array of signed 32-bit integer values.

```
FUNCTION PutLongArray (BYVAL prgf AS LONG PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of LONG values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_I4 variant.

---

## <a name="putulongarray"></a>PutULongArray

Initializes DVARIANT from an array of 32-bit unsigned integer values.

```
FUNCTION PutULongArray (BYVAL prgf AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of ULONG values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI4 variant.

---

## <a name="putlongintarray"></a>PutLongIntArray

Initializes DVARIANT from an array of signed 64-bit integer values.

```
FUNCTION PutLongIntArray (BYVAL prgf AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of LONGINT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_I8 variant.

---

## <a name="putulongintarray"></a>PutULongIntArray

Initializes DVARIANT from an array of unsigned 64-bit integer values.

```
FUNCTION PutULongIntArray (BYVAL prgf AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of ULONGINT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI8 variant.

---

## <a name="putdoublearray"></a>PutDoubleArray

Initializes DVARIANT from an array of unsigned 64-bit integer values.

```
FUNCTION PutDoubleArray (BYVAL prgf AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of DOUBLE values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_R8 variant.

---

## <a name="putstringarray"></a>PutStringArray

Initializes DVARIANT from an array of unicode strings.

```
FUNCTION PutStringArray (BYVAL prgsz AS PDWSTRING, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgsz* | Pointer to source array of unicode strings. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_BSTR variant.

---

## <a name="putbuffer"></a>PutBuffer

Initializes DVARIANT with the contents of a buffer.

```
FUNCTION PutBuffer (BYVAL pv AS VOID PTR, BYVAL cb AS UINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | Pointer to the source buffer. |
| *cb* | The length of the buffer, in bytes. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI1 variant.

### Example
```
DIM dv AS DVARIANT
DIM strIn AS STRING = "Test string"
dv.PutBuffer(STRPTR(strIn), LEN(strIn))

DIM nBytes AS LONG = dv.GetElementCount
DIM strOut AS STRING = SPACE(nBytes)
dv.ToBuffer(STRPTR(strOut), nBytes)
PRINT strOut
```
---

## <a name="putdatestring"></a>PutDateString

Initializes DVARIANT VT_DATE from a string.

```
FUNCTION PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, _
   BYVAL dwFlags AS ULONG = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszDate* | The date value to convert, e.g. "2018-08-20 19:42". |
| *lcid* | The locale identifier. |
| *dwFlags* | One or more of the following flags.<br>*LOCALE_NOUSEROVERRIDE* : Uses the system default locale settings, rather than custom locale settings.<br>*VAR_CALENDAR_HIJRI* : If set then the Hijri calendar is used. Otherwise the calendar set in the control panel is used.<br>*VAR_TIMEVALUEONLY* : Omits the date portion of a VT_DATE and returns only the time. Applies to conversions to or from dates.<br>*VAR_DATEVALUEONLY* : Omits the time portion of a VT_DATE and returns only the date. Applies to conversions to or from dates. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.

---

## <a name="putdec"></a>PutDec

Initializes DVARIANT with the contents of a DECIMAL structure.

```
FUNCTION PutDec (BYCAL dec AS DECIMAL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dec* | A DECIMAL structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DECIMAL OR VT_BYREF variant.

---

## <a name="putdecfromcy"></a>PutDecFromCY

Converts a currency value to a variant of type VT_DECIMAL.

```
FUNCTION PutDecFromCY (BYVAL cyIn AS CY) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cyIn* | The currency value to convert. |

#### Return value

This function can return one of these values.

| Value      | Meaning     |
| ---------- | ----------- |
| *S_OK* | Success. |
| *DISP_E_TYPEMISMATCH* | The argument could not be coerced to the specified type. |
| *E_INVALIDARG* | One of the arguments is not valid. |
| *E_OUTOFMEMORY* | Insufficient memory to complete the operation. |

#### Remarks

Creates a VT_DECIMAL variant.

---

## <a name="putdecfromdouble"></a>PutDecFromDouble

Converts a double value to a variant of type VT_DECIMAL.

```
FUNCTION PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbIn* | The DOUBLE value to convert. |

#### Return value

This function can return one of these values.

| Value      | Meaning     |
| ---------- | ----------- |
| *S_OK* | Success. |
| *DISP_E_TYPEMISMATCH* | The argument could not be coerced to the specified type. |
| *E_INVALIDARG* | One of the arguments is not valid. |
| *E_OUTOFMEMORY* | Insufficient memory to complete the operation. |

#### Remarks

Creates a VT_DECIMAL variant.

---

## <a name="putdecfromstr"></a>PutDecFromStr

Initializes DVARIANT as VT_DECIMAL from a string.

```
FUNCTION PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, _
   BYVAL dwFlags AS ULONG = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszIn* | The string value to convert. |
| *lcid* | The locale identifier. |
| *dwFlags* | One or more of the following flags.<br>*LOCALE_NOUSEROVERRIDE* : Uses the system default locale settings, rather than custom locale settings.<br>*VAR_TIMEVALUEONLY* : Omits the date portion of a VT_DATE and returns only the time. Applies to conversions to or from dates.<br>*VAR_DATEVALUEONLY* : Omits the time portion of a VT_DATE and returns only the date. Applies to conversions to or from dates. |

#### Return value

This function can return one of these values.

| Value      | Meaning     |
| ---------- | ----------- |
| *S_OK* | Success. |
| *DISP_E_TYPEMISMATCH* | The argument could not be coerced to the specified type. |
| *E_INVALIDARG* | One of the arguments is not valid. |
| *E_OUTOFMEMORY* | Insufficient memory to complete the operation. |

#### Remarks

Creates a VT_DECIMAL variant.

---

## <a name="putfiletime"></a>PutFileTime

Initializes DVARIANT with the contents of a FILETIME structure.

```
FUNCTION PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pft* | Pointer to a FILETIME structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.

---

## <a name="putfiletimearray"></a>PutFileTimeArray

Initializes DVARIANT with an array of FILETIME structures.

```
FUNCTION PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgft* | Pointer to an array of FILETIME structures. |
| *cElems* | The number of elements in the array pointed to by *prgft*. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_DATE variant.

---

## <a name="putguid"></a>PutGuid

Initializes DVARIANT from a GUID.

```
FUNCTION PutGuid (BYVAL rguid AS GUID PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *rguid* | Reference to the source GUID. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI1 variant.

---

## <a name="putpropvariant"></a>PutPropVariant

Initializes DVARIANT from the contents of a PROPVARIANT structure.

```
FUNCTION PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPropVar* | Pointer to a source PROPVARIANT structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Normally, the data stored in the PROPVARIANT is copied to the VARIANT without a datatype change. However, in the following cases, there is no direct VARIANT support for the datatype, and they are converted as shown.

VT_BLOB, VT_STREAM<br>
Converted to VT_UNKNOWN. The punkVal member will contain a pointer to an IStream that contains the source data.

VT_LPSTR, VT_LPWSTR, VT_CLSID<br>
Converted to VT_BSTR,

VT_FILETIME<br>
Converted to VT_DATE.

VT_VECTOR OR x<br>
Converted to VT_ARRAY OR x

The following types cannot be converted with this function.

VT_STORAGE<br>
VT_BLOB_OBJECT<br>
VT_STREAMED_OBJECT<br>
VT_STORED_OBJECT<br>
VT_CF<br>
VT_VECTOR OR VT_CF

---

## <a name="putrecord"></a>PutRecord

Initializes DVARIANT with a reference to an UDT.

```
FUNCTION PutRecord (BYVAL pIRecordInfo AS IRecordInfo PTR, BYVAL pRec AS VOID PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pIRecordInfo* | Pointer to the IRecordInfo interface. |
| *pRec* | Pointer to the UDT. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_RECORD variant.

---

## <a name="putref"></a>PutRef

Assigns a value by reference (a pointer to a variable).

```
FUNCTION PutRef (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD) AS HRESULT
FUNCTION PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *_pvar* | Pointer to a variable. |
| *_vType* | Type of the variant: VT_BOOL, VT_I1, VT_UI1, VT_I2, VT_UI2, VT_INT, VT_UINT, VT_I4, VT_UI4, VT_I8, VT_UI8, VT_R4, VT_R8, VT_BSTR, VT_UNKNOWN, VT_DISPATCH, VT_DECIMAL, VT_CY, VT_DATE, VT_VARIANT, VT_SAFEARRAY, VT_ERROR. |
| *strType* | Type of the variant: "BOOL", "BYTE", "UBYTE", "SHORT", "USHORT", "INT", "UINT", "LONG", "ULONG", "LONGINT", "ULONGINT", "SINGLE", "DOUBLE", "BSTR", "UNKNOWN", "DISPATCH", "DECIMAL", "CY", "DATE", "VARIANT", "SAFEARRAY", "ERROR". |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BYREF variant of the sepecified type.

---

## <a name="putresource"></a>PutResource

Initializes the DVARIANT based on a string resource imbedded in an executable file.

```
FUNCTION PutResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hinst* | The instance handle. |
| *id* | Integer identifier of the string to be loaded. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BSTR variant. If the resource does not exist, this function initializes the VARIANT as VT_EMPTY and returns a failure code.

---

## <a name="putsafearray"></a>PutSafeArray

Initializes DVARIANT from a safe array.

```
FUNCTION PutSafeArray (BYVAL parray AS SAFEARRAY PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *parray* | Pointer to safe array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY variant.

---

## <a name="putstrret"></a>PutStrRet

Initializes DVARIANT with the string stored in a STRRET structure.

```
FUNCTION PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstrret* | Pointer to a STRRET structure. |
| *pidl* | PIDL of the item whose details are being retrieved. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BSTR variant. This function frees the resources used for the STRRET contents.

---

## <a name="putsystemtime"></a>PutSystemTime

Initializes DVARIANT with the contents of a SYSTEMTIME structure.

```
FUNCTION PutSystemTime (BYVAL pst AS SYSTEMTIME PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pst* | Pointer to a SYSTEMTIME structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.

---

## <a name="pututf8"></a>PutUtf8

Initializes DVARIANT with the contents of an UTF-8 string.

```
FUNCTION PutUtf8 (BYREF strUtf8 AS STRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *strUtf8* | The UTF-8 encoded string. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BSTR variant.

---

## <a name="putvariantarrayelem"></a>PutVariantArrayElement

Initializes DVARIANT with a value stored in another VARIANT structure.

```
FUNCTION PutVariantArrayElement (BYVAL pvarIn AS VARIANT PTR, BYVAL iElem AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvarIn* | Reference to the source VARIANT structure. |
| *iElem* | Index of one of the source VARIANT structure elements. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

This helper function works for VARIANT structures of the following types: 

    VT_BSTR
    VT_BOOL
    VT_I2
    VT_I4
    VT_I8
    VT_U12
    VT_U14
    VT_U18
    VT_DATE
    VT_ARRAY | (any one of VT_BSTR, VT_BOOL, VT_I2, VT_I4, VT_I8, VT_U12, VT_U14, VT_U18, VT_DATE)

Additional types may be supported in the future.

This function extracts a single value from the source VARIANT structure and uses that value to initialize the output VARIANT structure. The calling application must use VariantClear to free the VARIANT referred to by pvar when it is no longer needed.

If the source VARIANT is an array, *iElem* must be less than the number of elements in the array.

If the source VARIANT has a single value, *iElem* must be 0.

If the source VARIANT is empty, this function always returns an error code.

You can use **GetElementCount** to obtain the number of elements in the array or array.

---

## <a name="putvbdate"></a>PutVbDate

Initializes DVARIANT with the contents of a DATE value.

```
FUNCTION PutVbDate (BYREF vbDate AS DATE_) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vbDate* | The DATE value. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.

---

## <a name="tobooleanarray"></a>ToBooleanArray

Extracts an array of boolean values from DVARIANT.

```
FUNCTION ToBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to a buffer that contains *crgn* boolean values. When this function returns, the buffer has been initialized with elements extracted from the source VARIANT structure. |
| *crgn* | The number of elements in the buffer pointed to by *prgf*. |

#### Return value

The count of WINBOOL elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of boolean values.

If the source VARIANT is of type VT_ARRAY OR VT_BOOL, this function extracts up to *crgn* WINBOOL values and places them into the buffer pointed to by *prgf*. If the VARIANT contains more elements than will fit into the *prgf* buffer, this function returns 0.

---

## <a name="tobooleanarrayalloc"></a>ToBooleanArrayAlloc

Extracts an array of boolean values from DVARIANT.

```
FUNCTION ToBooleanArrayAlloc (BYVAL pprgf AS WINBOOL PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to a WINBOOL PTR variable that will recive a pointer to an array of WINBOOL values extracted from the source DVARIANT. |

#### Return value

The count of WINBOOL elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of WINBOOL values.

If DVARIANT is of type VT_ARRAY OR VT_BOOL, this function extracts an array of WINBOOL values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgf* when it is no longer needed.

---

## <a name="tobstr"></a>ToBstr

Extracts the content of the underlying variant and returns it as a BSTRING.

```
FUNCTION ToBstr () AS BSTRING
```

#### Return value

The contents of the variant as a BSTRING.

#### Example

```
DIM dv AS DVARIANT = "Test string"
DIM bs AS BSTRING = dv.ToBstr
```
---

## <a name="tobuffer"></a>ToBuffer

Extracts the contents of a DVARIANT of type VT_ARRRAY OR VT_UI1 to a buffer.

```
FUNCTION ToBuffer (BYVAL pv AS VOID PTR, BYVAL cb AS UINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | Pointer to a buffer of length *cb* bytes. When this function returns, contains the first *cb* bytes of the extracted buffer value. |
| *cb* | The size of the *pv* buffer, in bytes. The buffer should be the same size as the data to be extracted, or smaller. |

#### Return value

Returns one of the following values:

| Value      | Meaning |
| ---------- | ----------- |
| *S_OK* | Data successfully extracted. |
| *E_INVALIDARG* | The VARIANT was not of type VT_ARRRAY OR VT_UI1. |
| *E_FAIL* |The VARIANT buffer value had fewer than cb bytes. |

#### Remarks

This function is used when the calling application expects a VARIANT to hold a buffer value. The calling application should check that the value has the expected length before it calls this function.

If DVARIANT has type VT_ARRAY OR VT_UI1, this function extracts the first *cb* bytes from the structure and places them in the buffer pointed to by *pv*.

If the stored value has fewer than *cb* bytes, then function fails and the buffer is not modified.

If the value has more than *cb* bytes, then function succeeds and truncates the value.

To retrieve the size of the array call **GetElementCount**.

### Example
```
DIM dv AS DVARIANT
DIM strIn AS STRING = "Test string"
dv.PutBuffer(STRPTR(strIn), LEN(strIn))

DIM nBytes AS LONG = dv.GetElementCount
DIM strOut AS STRING = SPACE(nBytes)
dv.ToBuffer(STRPTR(strOut), nBytes)
PRINT strOut
```
---

## <a name="tobufferstring"></a>ToBuffer (STRING)

Extracts the contents of a DVARIANT of type VT_ARRRAY OR VT_UI1 to a string used as a buffer.

```
FUNCTION ToBuffer () AS STRING
```

#### Return value

A string with the contents of the array.

---

## <a name="todosdatetime"></a>ToDosDateTime

Extracts a date and time value in Microsoft MS-DOS format from a DVARIANT of type VT_DATE.

```
FUNCTION ToDosDateTime (BYVAL pwDate AS USHORT PTR, BYVAL pwTime AS USHORT PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwDate* | When this function returns, contains the extracted USHORT that represents a MS-DOS date. |
| *pwTime* | When this function returns, contains the extracted contains the extracted WORD that represents a MS-DOS time. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold a datetime value.

If DVARIANT is of type VT_DATE, this function extracts the datetime value.

If DVARIANT is not of type VT_DATE, the function attempts to convert the value in the VARIANT structure into the right format. If a conversion is not possible, it returns a failure code.

---

## <a name="todoublearray"></a>ToDoubleArray

Extracts an array of DOUBLE values from DVARIANT.

```
FUNCTION ToDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains crgn DOUBLE variables. When this function returns, the buffer has been initialized with DOUBLE elements extracted from DVARIANT. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of DOUBLE values.

If the source VARIANT is of type VT_ARRAY OR VT_R8, this function extracts up to *crgn* DOUBLE values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

---

## <a name="todoublearrayalloc"></a>ToDoubleArrayAlloc

Extracts an array of DOUBLE values from DVARIANT.

```
FUNCTION ToDoubleArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgf* | Pointer to a DOUBLE PTR variable that will recive a pointer to an array of DOUBLE values extracted from the source DVARIANT. |

#### Return value

The count of DOUBLE elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of DOUBLE values.

If DVARIANT is of type VT_ARRAY OR VT_R8, this function extracts an array of DOUBLE values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer neede

---

## <a name="tofiletime"></a>ToFileTime

Returns the contents of a DVARIANT of type VT_DATE as a FILETIME structure.

```
FUNCTION ToFileTime (BYVAL stfOut AS AFX_PSTIME_FLAGS) AS FILETIME
```

| Parameter  | Description |
| ---------- | ----------- |
| *stfOut* | Specifies one of the following time flags:<br>*PSTF_UTC (0)* : Indicates coordinated universal time.<br>*PSTF_LOCAL (1)* : Indicates local time. |

#### Return value

A FILETIME structure.

---

## <a name="toguid"></a>ToGuid

Returns the contents of a DVARIANT containing a GUID string as a GUID structure.

```
FUNCTION ToGuid () AS GUID
```

#### Return value

Returns the contents of a DVARIANT containing a GUID string as an unicode GUID string.

---

## <a name="toguidbstr"></a>ToGuidBStr

Returns the contents of a DVARIANT containing a GUID string as an unicode GUID string.

```
FUNCTION ToGuidBStr () AS BSTRING
```

#### Return value

A GUID string.

---

## <a name="toguidstr"></a>ToGuidStr

Returns the contents of a DVARIANT containing a GUID string as an unicode GUID string.

```
FUNCTION ToGuidStr () AS DWSTRING
```

#### Return value

A GUID string.

---

## <a name="toguidwstr"></a>ToGuidWStr

Returns the contents of a DVARIANT containing a GUID string as an unicode GUID string.

```
FUNCTION ToGuidWStr () AS DWSTRING
```

#### Return value

A GUID string.

---

## <a name="tolongarray"></a>ToLongArray

Extracts an array of LONG values from DVARIANT.

```
FUNCTION ToLongArray (BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* LONG variables. When this function returns, the buffer has been initialized with Int32 elements extracted from DVARIANT. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of Int32 values.

If the source VARIANT is of type VT_ARRAY OR VT_I4, this function extracts up to *crgn* Int32 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

---

## <a name="tolongarrayalloc"></a>ToLongArrayAlloc

Extracts an array of LONG values from DVARIANT.

```
FUNCTION ToLongArrayAlloc (BYVAL pprgn AS LONG PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a LONG PTR variable that will recive a pointer to an array of LONG values extracted from the source DVARIANT. |

#### Return value

The count of LONG elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of LONG values.

If DVARIANT is of type VT_ARRAY OR VT_I4, this function extracts an array of LONG values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

---

## <a name="tolongintarray"></a>ToLongIntArray

Extracts an array of LONGINT values from DVARIANT.

```
FUNCTION ToLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* LONGINT variables. When this function returns, the buffer has been initialized with Int64 elements extracted from DVARIANT. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of Int64 values.

If the source VARIANT is of type VT_ARRAY OR VT_UI4, this function extracts up to *crgn* Int64 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

---

## <a name="tolongintarrayalloc"></a>ToLongIntArrayAlloc

Extracts an array of LONGINT values from DVARIANT.

```
FUNCTION ToLongIntArrayAlloc (BYVAL pprgn AS LONGINT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a LONGINT PTR variable that will recive a pointer to an array of LONGINT values extracted from the source DVARIANT.
 |

#### Return value

The count of LONGINT elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of LONGINT values.

If DVARIANT is of type VT_ARRAY OR VT_I8, this function extracts an array of LONGINT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

---

## <a name="toshortarray"></a>ToShortArray

Extracts an array of Int16 values from DVARIANT.

```
FUNCTION ToShortArray (BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* Int16 variables. When this function returns, the buffer has been initialized with Int16 elements extracted from DVARIANT. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of Int16 elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of Int16 values.

If the source VARIANT is of type VT_ARRAY OR VT_I2, this function extracts up to *crgn* Int16 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

---

## <a name="toshortarrayalloc"></a>ToShortArrayAlloc

Extracts an array of SHORT values from DVARIANT.

```
FUNCTION ToShortArrayAlloc (BYVAL pprgn AS SHORT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a SHORT PTR variable that will recive a pointer to an array of SHORT values extracted from the source DVARIANT. |

#### Return value

The count of SHORT elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of SHORT values.

If DVARIANT is of type VT_ARRAY OR VT_I2, this function extracts an array of SHORT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

---

## <a name="tostr"></a>ToStr

Extracts the content of the underlying variant and returns it as a DWSTRING.

```
FUNCTION ToStr () AS DWSTRING
```

#### Return value

The contents of the variant as a DWSTRING.

#### Example

```
DIM dv AS DVARIANT = "Test string"
DIM dws AS DWSTRING = dv.ToStr
```
---

## <a name="tostringarray"></a>ToStringArray

Extracts data from a vector structure into a PWSTR array.

```
FUNCTION ToStringArray (BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgsz* | Pointer to a buffer that contains *crgn* PWSTR values. When this function returns, the buffer has been initialized with elements extracted from the source VARIANT structure. |
| *crgsz* | The number of elements in the buffer pointed to by *prgsz*. |

#### Return value

The count of PWSTR elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of PWSTR values. If the VARIANT contains more elements than will fit into the *prgsz* buffer, this function returns 0.

---

## <a name="tostringarrayalloc"></a>ToStringArrayAlloc

Extracts an array of PWSTR values from DVARIANT.

```
FUNCTION ToStringArrayAlloc (BYVAL pprgsz AS PWSTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgsz* | Pointer to a PWSTR PTR variable that will recive a pointer to an array of PWSTR values extracted from the source DVARIANT. |

#### Return value

The count of PWSTR elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of PWSTR values.

This function extracts an array of PWSTR values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to free the memory used by each of the strings and to release the array pointed to by pprgn when it is no longer needed.

---

## <a name="tostrret"></a>ToStrRet

Returns the contents of a DVARIANT of type VT_BSTR to a STRRET stucture.

```
FUNCTION ToStrRet () AS STRRET
```

#### Return value

A STRRET structure.

---

## <a name="tosystemtime"></a>ToSystemTime

Returns the contents of DVARIANT of type VT_DATE as a FILETIME structure.

```
FUNCTION ToSystemTime () AS SYSTEMTIME
```

#### Return value

A SYSTEMTIME structure.

---

## <a name="toulongarray"></a>ToULongArray

Extracts an array of ULONG values from DVARIANT.

```
FUNCTION ToULongArray (BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* ULONG variables. When this function returns, the buffer has been initialized with UInt32 elements extracted from DVARIANT. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of UInt32 elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of UInt32 values.

If the source VARIANT is of type VT_ARRAY OR VT_UI4, this function extracts up to *crgn* UInt32 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

---

## <a name="toulongarrayalloc"></a>ToULongArrayAlloc

Extracts an array of ULONG values from DVARIANT.

```
FUNCTION ToULongArrayAlloc (BYVAL pprgn AS ULONG PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a ULONG PTR variable that will recive a pointer to an array of ULONG values extracted from the source DVARIANT. |

#### Return value

The count of ULONG elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of ULONG values.

If DVARIANT is of type VT_ARRAY OR VT_UI4, this function extracts an array of ULONG values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

---

## <a name="toulongintarray"></a>ToULongIntArray

Extracts an array of ULONGINT values from DVARIANT.

```
FUNCTION ToULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* ULONGINT variables. When this function returns, the buffer has been initialized with UInt64 elements extracted from DVARIANT. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of UInt64 elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of UInt64 values.

If the source VARIANT is of type VT_ARRAY OR VT_UI8, this function extracts up to *crgn* UInt64 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

---

## <a name="toulongintarrayalloc"></a>ToULongIntArrayAlloc

Extracts an array of ULONGINT values from DVARIANT.

```
FUNCTION ToULongIntArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a ULONGINT PTR variable that will recive a pointer to an array of ULONGINT values extracted from the source DVARIANT. |

#### Return value

The count of ULONGINT elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of ULONGINT values.

If DVARIANT is of type VT_ARRAY OR VT_UI8, this function extracts an array of ULONGINT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

---

## <a name="toushortarray"></a>ToUShortArray

Extracts an array of UInt16 values from DVARIANT.

```
FUNCTION ToUShortArray (BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* UInt16 variables. When this function returns, the buffer has been initialized with UInt16 elements extracted from DVARIANT. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of UInt16 elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of UInt16 values.

If the source VARIANT is of type VT_ARRAY OR VT_I2, this function extracts up to *crgn* UInt16 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

---

## <a name="toushortarrayalloc"></a>ToUShortArrayAlloc

Extracts an array of USHORT values from DVARIANT.

```
FUNCTION ToUShortArrayAlloc (BYVAL pprgn AS USHORT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a USHORT PTR variable that will recive a pointer to an array of USHORT values extracted from the source DVARIANT. |

#### Return value

The count of USHORT elements extracted from the DVARIANT.

#### Remarks

This helper function is used when the calling application expects a DVARIANT to hold an array of USHORT values.

If DVARIANT is of type VT_ARRAY OR VT_UI2, this function extracts an array of SHORT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

---

## <a name="toutf8"></a>ToUtf8

Returns the contents of a DVARIANT containing a BSTR as an UTF-8 encoded string.

```
FUNCTION ToUtf8 () AS STRING
```

#### Return value

The UTF-8 string.

---

## <a name="tovbdate"></a>ToVbDate

Returns the contents of a DVARIANT of type VT_DATE as a DATE value.

```
FUNCTION ToVbDate () AS DATE_
```

#### Return value

A DATE_ value (double).

## <a name="AfxDVarToStr"></a>AfxDVarToStr

Extracts the contents of a DVARIANT to a DWSTRING.

```
FUNCTION AfxDVarToStr (BYREF dv AS DVARIANT) AS DWSTRING
```
---

## <a name="AfxDVarToStr"></a>AfxDVarToStr

Extracts the contents of a DVARIANT to a DWSTRING.

```
FUNCTION AfxDVarToStr (BYVAL pdv AS DVARIANT PTR) AS DWSTRING
#define DVAR_TOSTR(cv) AfxDVarToStr(cv)
```
---

## <a name="AfxDVarToBuffer"></a>AfxDVarToBuffer

Extracts the contents of a variant that contains an array of bytes.

```
FUNCTION AfxDVarToBuffer (BYREF cvIn AS DVARIANT, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIn* | Reference to the source DVARIANT. |
| *pv* | Pointer to a buffer of length *cb* bytes. When this function returns, contains the first *cb* bytes of the extracted buffer value. |
| *cb* | The size of the *pv* buffer, in bytes. The buffer should be the same size as the data to be extracted, or smaller. |

#### Return value

Returns one of the following values:

| HRESULT code | Description |
| ---------- | ----------- |
| **S_OK** | Data successfully extracted. |
| **E_INVALIDARG** | The VARIANT was not of type VT_ARRRAY | VT_UI1. |
| **E_FAIL** | The VARIANT buffer value had fewer than cb bytes. |

---

## <a name="AfxDVarOptPrm"></a>AfxDVarOptPrm

Returns a DVARIANT suitable to be used with optional parameters.

```
FUNCTION AfxDVarOptPrm () AS DVARIANT
#define DVAR_OPTPRM AfxDVarOptPrm
```

#### Remarks

If you want to call a method that has optional variant parameters, you still have to pass the parameters, but in a way that the methods know that they were omitted. Specifically, you have to pass them as type VT_ERROR, code DISP_E_PARAM_NOT_FOUND.

```
DIM v AS VARIANT = TYPE(VT_ERROR, 0, 0, 0, DISP_E_PARAMNOTFOUND)
```
---
