# BSTRING Class

The `BSTRING` and `DWSTRING` classes implement a dynamic unicode null terminated string. Free Basic has a dynamic string data type (STRING) and a fixed length unicode data type (WSTRING), but it lacks a dynamic unicode string. `BSTRING` and `DWSTRING` behave as if they were native data types, working directly with the intrinsic FreeBasic string functions and operators.

While `DWSTRING`manages its own memory allocations, `BSTRING`is a wrapper on top of of the OLE string (aka BSTR) data type and the memory management is done by the COM library. It is better to use `DWSTRING`for general purposes, since it is faster, reserving the use of `BSTRING` for COM programming.

**Include file**: AfxNova/BSTRING.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructors) | Initialize the class with the specified value. |
| [Operator \*](#operator*) | One * returns the address of the `BSTRING` buffer.<br> Two ** returns the address of the start of the string data. |
| [Operator Cast](#operatorcast) | Returns a pointer to the `BSTRING` buffer or the string data.<br>Casting is automatic. You don't have to call this operator. |
| [Attach](#attach) | Attaches a BSTR to the BSTRING class. |
| [Detach](#detach) | Detaches the underlying BSTR from the BSTRING class and returns it as the result of the function. |
| [Utf8](#utf8) | Converts from UTF8 to Unicode and from Unicode to UTF8. |
| [vptr](#vptr) | Frees the underlying BSTR and returns a pointer to a new empty allocated string. |
| [wchar](#wchar) | Returns the string data as a new unicode string allocated with CoTaskMemAlloc. |

# <a name="constructors"></a>Constructors

Initialize the class with the specified value.

```
CONSTRUCTOR
CONSTRUCTOR (BYREF wszStr AS CONST WSTRING)
CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR)
CONSTRUCTOR (BYREF bs AS BSTRING)
CONSTRUCTOR (BYREF dws AS DWSTRING)
CONSTRUCTOR (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
CONSTRUCTOR (BYVAL nChars AS LONG, BYREF wszFill AS CONST WSTRING)
CONSTRUCTOR (BYVAL n AS LONGINT)
CONSTRUCTOR (BYVAL n AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | A WSTRING. |
| *pwszStr* | A pointer to a WSTRING. |
| *dws* | A DWSTRING. |
| *bs* | A BSTRING. |
| *ansiStr* | An ansi string or string literal. |
| *nCodePage* | The code page to be used for ansi to unicode conversions. |
| *n* | A number. |

1. Creates an empty Unicode string buffer with an initial null-terminated string.
```
Example: DIM bs AS BSTRING
```
2. Initializes a `BSTRING` instance from a wide string pointer.
```
DIM wsz AS WSTRING * 30 = "This is a test string"
DIM bsText AS BSTRING = wsz
```
3. Initializes a `BSTRING` from an ANSI or UTF-8 encoded string, with optional code page support.
```
Example: DIM s AS STRING = "Hello, world"
         DIM bs AS BSTRING = s
Example: DIM bs AS BSTRING = BSTRING("Hello, utf", CP_UTF8)
```
For a list of code pages see: [Code Page Identifiers](https://msdn.microsoft.com/en-us/library/windows/desktop/dd317756(v=vs.85).aspx)

5. Creates a copy of an existing `BSTRING`.
```
Example: DIM bs1 AS BSTRING = "Test string"
         DIM bs2 AS BSTRING = bs1
```
6. Creates a `BSTRING` with a fixed-length buffer, initialized with a fill character.
```
Example: DIM bs AS BSTRING = BSTRING(260, " ")
```
7. Initializes a `BSTRING` with a numeric value, allowing automatic conversion.
```
Example: DIM bsNum AS BSTRING = 12345
Example: DIM bsFloat AS BSTRING = 3.1415
```

#### Remarks

`BSTRING` works transparently with literals and FreeBasic native strings, e.g.
```
DIM bs AS BSTRING = "One"
DIM s AS STRING = "Three"
bs = bs & " Two " & s
PRINT bs
```
It can be used with Windows API functions, e.g.
```
DIM nLen AS LONG = SendMessageW(hWin, WM_GETTEXTLENGTH, 0, 0)
DIM bsText AS BSTRING = WSPACE(nLen + 1)
SendMessageW(hWin, WM_GETTEXT, nLen + 1, cast(LPARAM, *bsText))
bsText = LEFT(bsText, LEN(bsText) - 1)
PRINT bsText
```
We can use arrays of `DWSTRING` strings transparently, e.g.
```
DIM rg(1 TO 10) AS BSTRING
FOR i AS LONG = 1 TO 10
   rg(i) = "string " & i
NEXT
FOR i AS LONG = 1 TO 10
   PRINT rg(i)
NEXT
```
A two-dimensional array
```
DIM rg2 (1 TO 2, 1 TO 2) AS BSTRING
rg2(1, 1) = "string 1 1"
rg2(1, 2) = "string 1 2"
rg2(2, 1) = "string 2 1"
rg2(2, 2) = "string 2 2"
PRINT rg2(2, 1)
```

REDIM PRESERVE / ERASE
```
REDIM rg(0) AS BSTRING
rg(0) = "string 0"
REDIM PRESERVE rg(0 TO 2) AS BSTRING
rg(1) = "string 1"
rg(2) = "string 2"
PRINT rg(0)
PRINT rg(1)
PRINT rg(2)
ERASE rg
```
You can also use it with files:

Using FreeBasic intrinsic statements:
```
DIM bs AS BSTRING = "Дмитрий Дмитриевич Шостакович"
DIM f AS LONG = FREEFILE
OPEN "Test.txt" FOR OUTPUT ENCODING "utf16" AS #f
PRINT #f, bs
CLOSE #f
```
Using the Windows API:
```
' // Writing to a file
DIM bsFilename AS BSTRING = "тест.txt"
DIM bsText AS BSTRING = "Дмитрий Дмитриевич Шостакович"
DIM hFile AS HANDLE = CreateFileW(bsFilename, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)
IF hFile THEN
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS LONG = WriteFile(hFile, bsText, LEN(bsText) * 2, @dwBytesWritten, NULL)
   CloseHandle(hFile)
END IF
```
```
' // Read the file
hFile = CreateFileW(bsFilename, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
IF hFile THEN
   DIM dwFileSize AS DWORD = GetFileSize(hFile, NULL)
   IF dwFileSize THEN
      DIM bsOut AS BSTRING = WSPACE(dwFileSize \ 2)
      DIM bSuccess AS LONG = ReadFile(hFile, *bsOut, dwFileSize, NULL, NULL)
      CloseHandle(hFile)
      PRINT bsOut
   END IF
END IF
```

Notice that, contrarily to **CreateFileW**, FreeBasic's OPEN statement doesn't allow to use unicode for the file name.

---

# Operators

## <a name="operator*"></a>Operator *

Deferences the `BSTRING`.<br>One * returns the address of the underlying `BSTR` handle (same as STRPTR).<br>Two ** returns the address of the start of the string data (same as *STRPTR).

```
OPERATOR * (BYREF bs AS BSTRING) AS WSTRING PTR
```

---

# Casting and Conversions

## <a name="operatorcast"></a>Operator Cast

Returns a pointer to the underlying `BSTR` or the string data. These operators aren't called directly.

```
OPERATOR CAST () BYREF AS CONST WSTRING
OPERATOR CAST () AS ANY PTR
```

---

## <a name="utf8"></a>Utf8

Converts from UTF8 to Unicode and from Unicode to UTF8.

```
PROPERTY Utf8() AS STRING
PROPERTY Utf8 (BYREF utf8String AS STRING)
```

---

## <a name="vptr"></a>vptr

Frees the underlying `BSTR` and returns a pointer to a new empty allocated `BSTR`.

Used to pass the underlying `BSTR` to an OUT BYVAL BSTR PTR parameter. If we pass a `BSTRING` to a procedure with an OUT BSTR parameter without first freeing it we will have a memory leak.

```
FUNCTION vptr () AS BSTR PTR
```

---

## <a name="wchar"></a>wchar

Returns the string data as a new unicode string allocated with **CoTaskMemAlloc**.

Useful when we need to pass a pointer to a null terminated wide string to a function or method that will release it. If we pass a WSTRING it will GPF. If the length of the input string is 0, **CoTaskMemAlloc** allocates a zero-length item and returns a valid pointer to that item. If there is insufficient memory available, **CoTaskMemAlloc** returns NULL.

```
FUNCTION wchar () AS WSTRING PTR
```

---

# Methods

## <a name="attach"></a>Attach

Attaches a `BSTR` to the `BSTRING` class.

```
SUB Attach (BYVAL pbstrSrc AS BSTR)
```

---

## <a name="detach"></a>Detach

Detaches the underlying `BSTR` from the `BSTRING` class and returns it as the result of the function. The returned pointer must be freed by **SysFreeString**.

```
FUNCTION Detach () AS BSTR
```

This method frees the *m_bstr* member of the `BSTRING` class and returns it as the result of the function. Because it no longer belongs to the class, it must be freed by **SysFreeString**.

---

