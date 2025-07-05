# String Procedures

This is a collection of assorted, ready-to-use string procedures to manipulate FreeBasic strings and my own `DWSTRING` (Dynamic Unicode string) and `BSTRING` (OLE strings, aka BSTR).

The procedures that need to be fast have been hard coded, whereas the ones that need flexibility use `CRegExp`, a wrapper class on top of the VBScript regular expressions engine.

[`CRegExp`](https://github.com/JoseRoca/WinFBX2/blob/main/docs/String%20Management/CRegExp%20Class.md) can also be used to build other string procedures not included in this collection.

**Include file**: AfxNova/DWSTRProcs.inc

Additionally, we can call external variadic functions written in C. They can be very useful to do string formatting.

**StringCbPrintf** (A/W):
https://msdn.microsoft.com/en-us/library/windows/desktop/ms647510(v=vs.85).aspx
```
DIM wszOut AS WSTRING * 260
DIM wszFmt AS WSTRING * 260 = "%s %d + %d = %d."
DIM wszText AS WSTRING * 260 = "The answer is"
DIM hr AS HRESULT = StringCbPrintfW(@wszOut, SIZEOF(wszOut), @wszFmt, @wszText, 1, 2, 3)
PRINT wszOut
' Output: "The answer is 1 + 2 = 3."
```
**StringCbPrintf_l** (A/W) is similar to **StringCbPrintf** but includes a parameter for locale information.

**StringCbPrintfEx** (A/W) adds to the functionality of **StringCbPrintf** by returning a pointer to the end of the destination string as well as the number of bytes left unused in that string. Flags may also be passed to the function for additional control.

**StringCbPrintf_lEx** (A/W) is similar to **StringCbPrintfEx** but includes a parameter for locale information.

### String procedures list

| Name       | Description |
| ---------- | ----------- |
| [DWStrAcode](#dwstracode) | Translates Unicode chars to ansi bytes. |
| [DWStrBase64DecodeA](#dwstrbase64decodea) | Converts the contents of a Base64 mime encoded string to an ASCII string. |
| [DWStrBase64DecodeW](#dwstrbase64decodew) | Converts the contents of a Base64 mime encoded string to an Unicode string. |
| [DWStrBase64EncodeA](#dwstrbase64encodea) | Converts the contents of an ASCII string to Base64 mime encoding. |
| [DWStrBase64EncodeW](#dwstrbase64encodew) | Converts the contents of an Unicode string to Base64 mime encoding. |
| [DWStrClip](#dwstrclip) | Returns a string with the specified number of characters removed from the left, right or mid section of the string. |
| [DWStrCryptBinaryToString](#dwstrcryptbinarytostring) | Converts an array of bytes into a formatted string. |
| [DWStrCryptStringToBinary](#dwstrcryptstringtobinary) | Converts a formatted string into an array of bytes. |
| [DWStrCSet](#dwstrcset) | Returns a string containing a centered (padded) string. |
| [DWStrCSetAbs](#dwstrcsetabs) | Returns a string containing a centered string within the space of another string. |
| [DWStrDelete](#dwstrdelete) | Deletes a specified number of characters from a string expression. |
| [DWStrEscape](#dwstrescape) | Escapes any potential regex syntax characters in a string. |
| [DWStrExtract](#dwstrextract) | Extracts characters from a string up to (but not including) the specified matching. |
| [DWStrFormatByteSize](#dwstrformatbytesize) | Converts a numeric value into a string that represents the number expressed as a size value in bytes, kilobytes, megabytes, or gigabytes, depending on the size. |
| [DWStrFormatKBSize](#dwstrformatkbsize) | Converts a numeric value into a string that represents the number expressed as a size value in kilobytes. |
| [DWStrFromTimeInterval](#dwstrfromtimeinterval) | Converts a time interval, specified in milliseconds, to a string. |
| [DWStrInsert](#dwstrinsert) | Inserts a string at a specified position within another string expression. |
| [DWStrIsNumeric](#dwstrisnumeric) | Returns True if the passed string is numeric. |
| [DWStrJoin](#dwstrjoin) | Returns a string consisting of all of the strings in an array, each separated by a delimiter. |
| [DWStrLCase](#dwstrlcase) | Returns a lowercased version of a string. |
| [DWStrLSet](#dwstrlset) | Returns a string containing a left-justified (padded) string. |
| [DWStrLSetAbs](#dwstrlsetabs) | Left-aligns a string within the space of another string. |
| [DWStrMCase](#dwstrmcase) | Returns a mixed case version of its string argument. |
| [DWStrParse](#dwstrparse) | Returns a delimited field from a string expression. |
| [DWStrParseCount](#dwstrparsecount) | Returns the count of delimited fields from a string expression. |
| [DWStrPathName](#dwstrpathname) | Parses a path to extract component parts. |
| [DWStrPathScan](#dwstrpathscan) | Searches a path for a file name. |
| [DWStrRemain](#dwstrremain) | Returns the portion of a string following the first occurrence of a string. |
| [DWStrRemove](#dwstrremove) | Returns a new string with substrings removed. |
| [DWStrRepeat](#dwstrrepeat) | Returns a string consisting of multiple copies of the specified string. |
| [DWStrReplace](#dwstrreplace) | Replaces all the occurrences of a string with another string. |
| [DWStrRetain](#dwstrretain) | Returns a string containing only the characters contained in a specified match string. |
| [DWStrReverse](#dwstrreverse) | Reverses the contents of a string expression. |
| [DWStrRSet](#dwstrrset) | Returns a string containing a right justified string. |
| [DWStrRSetAbs](#dwstrrsetabs) | Right-aligns a string within the space of another string. |
| [DWStrShrink](#dwstrshrink) | Shrinks a string to use a consistent single character delimiter. |
| [DWStrSplit](#dwstrsplit) | Splits a string into tokens, which are sequences of contiguous characters separated by any of the characters that are part of delimiters. |
| [DWStrSpn](#dwstrspn) | Returns the index of the initial portion of a string which consists only of characters that are part of a specified set of characters. |
| [DWStrTally](#dwstrtally) | Count the number of occurrences of a string within a string |
| [DWStrUCase](#dwstrucase) | Returns an uppercased version of a string. |
| [DWStrUCode](#dwstrucode) | Translates ansi bytes to Unicode bytes. |
| [DWStrVerify](#dwstrverify) | Determine whether each character of a string is present in another string. |
| [DWStrWrap](#dwstrwrap) | Adds paired characters to the beginning and end of a string. |
| [DWStrUnWrap](#dwstrunwrap) | Removes paired characters to the beginning and end of a string. |

### Surrogates

Checking and fixing broken surrogates is diabled by default for speed reasons. If you want to enable it, use `"#define __AFX2_CHECKSURROGATES__"` (without the quotes) at the beginning of your application, before including the files of this framework. Surrogate checking and fixing is centralized in te internal **AppendBuffer** method of the `DWSTRING` class.

| Name       | Description |
| ---------- | ----------- |
| [DWStrChrW](#dwstrchrw) | Returns a wide-character string from a codepoint. |
| [DWStrCodePointToSurrogatePair](#dwstrcodepointtosurrogatepair) | Converts a Unicode code point (above U+FFFF) back into its high and low surrogate pair. |
| [DWStrHasSurrogates](#dwstrhassurrogates) | Checks if the passed string has surrogates. |
| [DWStrIsValidSurrogatePair](#dwstrisvalidsurrogatepair) | Checks whether a UTF-16 encoded string contains valid high-low surrogate pairs. |
| [DWStrScanForSurrogates](#dwstrscanforsurrogates) | Scans a string to search for surrogates. |
| [DWStrSurrogatePairToCodePoint](#dwstrsurrogatepairtocodepoint) | Converts a surrogate pair to a Unicode code point. |

### Surrogate macros

The wollowing macros are available in the Windows headers (winnls.bi):
```
const HIGH_SURROGATE_START = &hd800
const HIGH_SURROGATE_END = &hdbff
const LOW_SURROGATE_START = &hdc00
const LOW_SURROGATE_END = &hdfff
#define IS_HIGH_SURROGATE(wch) (((wch) >= HIGH_SURROGATE_START) andalso ((wch) <= HIGH_SURROGATE_END))
#define IS_LOW_SURROGATE(wch) (((wch) >= LOW_SURROGATE_START) andalso ((wch) <= LOW_SURROGATE_END))
#define IS_SURROGATE_PAIR(hs, ls) (IS_HIGH_SURROGATE(hs) andalso IS_LOW_SURROGATE(ls))
```
**IS_HIGH_SURROGATE**: Determines if a character is a UTF-16 high surrogate code point, ranging from &hD800 to &hDBFF, inclusive.

**IS_LOW_SURROGATE**: Determines if a character is a UTF-16 low surrogate code point, ranging from &hDC00 to &hDFFF, inclusive.

**IS_SURROGATE_PAIR**: Determines if the specified code units form a UTF-16 surrogate pair.

---

### <a name="dwstracode"></a>DWStrAcode

Translates Unicode chars to ansi bytes.

```
FUNCTION DWStrAcode (BYVAL pwszStr AS WSTRING PTR, BYVAL nCodePage AS LONG = 0) AS STRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *pwszStr* | The Unicode string to translate. |
| *nCodePage* | The code page used in the conversion, e.g. 1251 for Russian. If the code page is omited, the function will use CP_ACP (0), which is the system default Windows ANSI code page.|

#### Return value

The translated string.

---

### <a name="dwstrclip"></a>DWStrClip

Returns a string with the specified number of characters removed from the left, middle or right side of the string.

```
FUNCTION FUNCTION DWStrClip (BYREF wszSide AS CONST WSTRING, BYREF wszSourceString AS CONST WSTRING, _
   BYVAL nCount AS LONG, BYVAL nStart AS LONG = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSide* | The "LEFT", "MID" or "RIGHT". |
| *wszSourceString:* | The source string. |
| *nCount* | The number of characters to be removed. If it is 0 or negative, the source string is inaltered. |
| *nStart* | The starting position. It is only used with the "MID" option. |

#### Usage examples
```
DIM dws AS DWSTRING = DWStrClip("LEFT", "1234567890", 3)     ' Output: "4567890"
DIM dws AS DWSTRING = DWStrClip("RIGHT", "1234567890", 3)    ' Output: "1234567"
DIM dws AS DWSTRING = DWStrClip("MID", "1234567890", 3, 4)   ' Output: "1237890"
```
---

### <a name="dwstrcset"></a>DWStrCSet

Returns a string containing a centered (padded) string.

```
FUNCTION DWStrCSet (BYREF wszSourceString AS CONST WSTRING, BYVAL nStringLength AS LONG, _
   BYREF wszPadCharacter AS CONST WSTRING = " ") AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to be justified. |
| *nStringLength* | The length of the new string. |
| *wszPadCharacter* | The character to be used for padding. If it is not specified, the string will be padded with spaces. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrCSet("FreeBasic", 20, "*")   ' Output: "*****FreeBasic******"
```
---

### <a name="dwstrcsetabs"></a>DWStrCSetAbs

Returns a string containing a centered string within the space of another string.

```
FUNCTION DWStrCSetAbs (BYREF wszPadString AS CONST WSTRING, BYREF wszString AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPadString* | The padding string. |
| *wszString* | The string to be centered inside the padding string. |

#### Usage example

```
DIM dwsPad AS DWSTRING = "COOL COOL COOL COOL COOL"
DIM dws AS DWSTRING = "..FreeBasic is.."
PRINT DWStrCSetAbs(dwsPad, dws)
' Output: "COOL..FreeBasic is..COOL"
```
---

### <a name="dwstrdelete"></a>DWStrDelete

Deletes a specified number of characters from a string expression.

```
FUNCTION DWStrDelete (BYREF wszSourceString AS CONST WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *nStart* | The one-based starting position. |
| *nCount* | The number of characters to be removed. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrDelete("1234567890", 4, 3)   ' Output: "1237890"
```
---

### <a name="DWStrEscape"></a>DWStrEscape

Escapes any potential regex syntax characters in a pattern string and returns a new string that can be safely used as a literal pattern.

```
FUNCTION DWStrEscape (BYREF wszStr AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | The pattern string. |

---

### <a name="DWStrExtract"></a>DWStrExtract

Extracts characters from a string up to (but not including) a string.

```
FUNCTION DWStrExtract (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, _
   BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
FUNCTION DWStrExtract (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, _
   BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position to begin the search. |
| *wszSourceString* | The source string. |
| *wszMatchStr* | The string expression to be removed. If *wszMatchStr* is not present in *wszSourceString*, all of *wszSourceString* is returned intact. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

*nStart* is the optional starting position to begin extracting. If *nStart* is not specified, it will start at position 1. If *nStart* is zero, or beyond the length of *wszSourceString*, a nul string is returned. If *nStart* is negative, the starting position is counted from right to left: if -1, the search begins at the last character; if -2, the second to last, and so forth.

#### Overloaded methods:

Returns the portion of a string following the occurrence of a specified delimiter up to the second delimiter. If no match is found, an empty string is returned. If *nStart* is 0 or greater than the length of *wszSourceString*, an empty string is returned.
```
FUNCTION DWStrExtract OVERLOAD (BYREF wszSourceString AS CONST WSTRING, BYREF leftDelimiter AS CONST WSTRING, _
   BYREF rightDelimiter AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
FUNCTION DWStrExtract OVERLOAD (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, _
   BYREF leftDelimiter AS CONST WSTRING, BYREF rightDelimiter AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position to begin the search. |
| *wszSourceString* | The source string. |
| *leftDelimiter* | The left delimiter. |
| *rightDelimiter* | The right delimiter. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Usage examples

```
DIM dws AS DWSTRING = "abacadabra"
PRINT DWStrExtract(dws, "cad")
' Output: "aba" - match on "cad"
```
*wszMatchStr* can specify a list of single characters, enclosed between [], to be searched for individually, a match on any one of which will cause the extract operation to be performed up to that character.
```
DIM dws AS DWSTRING = "abacadabra"
PRINT DWStrExtract(dws, "[dr]")
' Output: "abaca" - match on "d"
```
```
DIM dwsText AS DWSTRING = "blah blah text between parentheses) blah blah"
PRINT DWStrExtract(dwsText, "(", ")")   ' Output: "text between parentheses"
```
---

### <a name="dwstrformatbytesize"></a>DWStrFormatByteSize

Converts a numeric value into a string that represents the number expressed as a size value in bytes, kilobytes, megabytes, or gigabytes, depending on the size.

```
FUNCTION DWStrFormatByteSize (BYVAL ull AS LONGLONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *ull* | The numeric value to be converted. |

---

### <a name="dwstrformatkbsize"></a>DWStrFormatKBSize

Converts a numeric value into a string that represents the number expressed as a size value in kilobytes.

```
FUNCTION DWStrFormatKBSize (BYVAL ull AS LONGLONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *ull* | The numeric value to be converted. |

---

### <a name="dwstrfromtimeinterval"></a>DWStrFromTimeInterval

Converts a time interval, specified in milliseconds, to a string.

```
FUNCTION DWStrFromTimeInterval (BYVAL dwTimeMS AS DWORD, BYVAL digits AS LONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwTimeMS* | The time interval, in milliseconds. |
| *digits* | The maximum number of significant digits to be represented in the output string. |

Some examples for *digits*:

| dwTimeMS   | digits      | dwsOut      |
| ---------- | ----------- | ----------- |
| 34000 | 3 | 34 sec |
| 34000 | 2 | 34 sec |
| 34000 | 1 | 30 sec |
| 74000 | 3 | 1 min 14 sec |
| 74000 | 2 | 1 min 10 sec |
| 74000 | 2 | 1 min |

---

### <a name="dwstrinsert"></a>DWStrInsert

Inserts a string at a specified position within another string expression.

```
FUNCTION DWStrInsert (BYREF wszSourceString AS CONST WSTRING, BYREF wszInsertString AS CONST WSTRING, _
   BYVAL nPosition AS LONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *wszInsertString* | The string to be inserted. |
| *nPosition* | The one-based starting position. If *nPosition* is greater than the length of *wszSourceString* or <= zero then *wszInsertString* is appended to *wszSourceString*. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrInsert("1234567890", "--", 6)
```
---

### <a name="dwstrisnumeric"></a>DWStrIsNumeric

Retuns True if the passed string is muneric.

```
FUNCTION DWStrIsNumeric (BYREF wszSourcestring AS CONST WSTRING) AS BOOLEAN
```
#### Example

```
DWStrIsNumeric("1.2345678901234567e+029")   ' Output: "true"
```
Explanation of the pattern used: "^[\+\-]?\d*\.?\d+(?:[Ee][\+\-]?\d+)?$"
```
^ ? Anchors the match to the start of the string.
[\+\-]? ? Matches an optional plus (+) or minus (-) sign at the beginning (for signed numbers).
\d* ? Matches zero or more digits before the decimal point (allows integers or leading zero suppression).
\.? ? Matches an optional decimal point (if present, allows floating-point numbers).
\,? ? Matches an optional decimal point (in Spain, a comma is used instead of a colon).
\d+ ? Matches at least one digit after the decimal (ensuring valid numeric values).
(?:[Ee][\+\-]?\d+)? ? Handles scientific notation:
  E or e for exponent notation.
  [\+\-]? for optional sign after the exponent indicator.
  \d+ ensures at least one digit in the exponent.
$ ? Anchors the match to the end of the string, ensuring a full numeric match.
```
---

### <a name="dwstrjoin"></a>DWStrJoin

Returns a string consisting of all of the strings in an array, each separated by a delimiter. If the delimiter is a null (zero-length) string then no separators are inserted between the string sections. If the delimiter expression is the 3-byte value of '","', which may be expressed in your source code as the string literal """,""" or as Chr(34,44,34) then a leading and trailing double-quote is added to each string section. This ensures that the returned string contains standard comma-delimited quoted fields that can be easily parsed.

```
#macro DWStrJoin(array, dest, delim)
```

| Parameter  | Description |
| ---------- | ----------- |
| *array* | The one-dimensional array to join. |
| *dest* | The destination string. |
| *delim* | The delimiter character. |

#### Return value

A string containing the joined string.

#### Usage example

```
DIM rg(1 TO 10) AS DWSTRING
FOR i AS LONG = 1 TO 10
   rg(i) = "string " & i
NEXT
DIM dws AS DWSTRING
DWStrJoin(rg, dws, """,""")
PRINT dws
' Instead of DWSTRING, any other data type can be used: BSTRING, STRING...
```
---

### <a name="dwstrlcase"></a>DWStrLCase

Returns a lowercased version of a string.

```
FUNCTION DWStrLCase (BYVAL pwszStr AS WSTRING PTR, _
   BYVAL pwszLocaleName AS WSTRING PTR = LOCALE_NAME_USER_DEFAULT, _
   BYVAL dwMapFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszStr* | The string. |
| *pwszLocaleName* | Optional. Pointer to a locale name or one of these pre-defined values: LOCALE_NAME_INVARIANT, LOCALE_NAME_SYSTEM_DEFAULT, LOCALE_NAME_USER_DEFAULT |
| *dwMapFlags* | Optional. Flag specifying the type of transformation to use during string mapping or the type of sort key to generate. |

For a table of language culture names see: [Table of Language Culture Names, Codes, and ISO Values](https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20))

For a complete list see: [LCMapStringEx function](https://docs.microsoft.com/en-us/windows/desktop/api/winnls/nf-winnls-lcmapstringex)

#### Remarks

The string conversion functions available in FreeBasic are not fully suitable for some languages. For example, the Turkish word "karışıklığı" is uppercased as "KARıŞıKLıĞı" instead of "KARIŞIKLIĞI", and "KARIŞIKLIĞI" is lowercased to "karişikliği" instead of "karışıklığı". Notice the "ı", that is not an "i".

For Turkey, use:
```
DWStrUCase("karışıklığı", "tr-TR")
DWStrLCase("KARIŞIKLIĞI", "tr-TR")
```

#### Return value

The lowercased string.

---

### <a name="dwtrlset"></a>DWStrLSet

Returns a string containing a left justified string.

```
FUNCTION DWStrLSet (BYREF wszSourceString AS CONST WSTRING, BYVAL nStringLength AS LONG, _
   BYREF wszPadCharacter AS CONST WSTRING = " ") AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to be justified. |
| *nStringLength* | The length of the new string. |
| *wszPadCharacter* | The character to be used for padding. If it is not specified, the string will be padded with spaces. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrLSet("FreeBasic", 20, "*")
```
---

### <a name="dwstrlsetabs"></a>DWStrLSetAbs

Left-aligns a string within the space of another string. If *wszStr* is empty, the function leaves the padding positions unchanged from their original content, rather than replacing them with spaces as `LSET` does. If *wszStr* is longer than *wszSourceString*, **dwstrLSetAbs** truncates it from the right until it fits in the result string.

```
FUNCTION DWStrLSetAbs (BYREF wszSourceString AS CONST WSTRING, BYREF wszStr AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The source string. |
| *wszStr* | The string to be left aligned inside the source string. |

#### Usage example

```
DIM dws AS DWSTRING = "NameBasic=SuperBasic"
PRINT DWStrLSetAbs(dws, "FreeBasic")  ' Output: "FreeBasic=SuperBasic"
```
---

### <a name="dwstrmcase"></a>DWStrMCase

Returns a mixed case version of its string argument.

```
FUNCTION DWStrMCase (BYREF wszSourceString AS WSTRING) AS DWSTRING
```
#### Example
```
DIM dws AS DWSTRING = strMCase("Cats aren't AL.WAYS good.")
' Output: "Cats Aren'T Al.Ways Good."
' Note: It mimincs PowerBASIC's MCase$ function.
```

### <a name="dwstrparse"></a>DWStrParse 

Returns a delimited field from a string expression.

```
FUNCTION DWStrParse (BYREF wszSourceString AS CONST WSTRING, BYVAL index AS LONG = 1, _
   BYREF wszDelimiter AS CONST WSTRING = ",") AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to be parsed. |
| *nPosition* | Starting position. If *nPosition* is zero or is outside of the actual field count, an empty string is returned. If *nPosition* is negative, fields are searched from the right to left of *wszSourceString*. |
| *wszDelimiter* | A string of one or more characters that must be fully matched to be successful. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrParse("one,two,three")           ' Output: "one"
DIM dws AS DWSTRING = DWStrParse("one;two,three", 1, ";")   ' Output: "one"
DIM dws AS DWSTRING = DWStrParse("1;2,3", 2, ",;")          ' Output: "2"
```
---

### <a name="DWStrParseCount"></a>DWStrParseCount 

Returns the count of delimited fields from a string expression.

```
FUNCTION DWStrParseCount (BYREF wszSourceString AS CONST WSTRING, BYREF wszDelimiter AS CONST WSTRING = ",") AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to parse. If *wszSourceString* is empty (a null string) or contains no delimiter character(s), the string is considered to contain exactly one subfield. |
| *wszDelimiter* | One or more character delimiters that must be fully matched. Delimiters are case-sensitive. |

#### Usage example

```
DIM nCount AS LONG = DWStrParseCount("one,two,three", ",")   ' Output: "3"
DIM nCount AS LONG = DWStrParseCount("1;2,3", ",;")          ' Output: "3"

```
---

### <a name="dwstrpathname"></a>DWStrPathName 

Parses a path to extract component parts.

```
FUNCTION DWStrPathName (BYREF wszOption AS CONST WSTRING, BYREF wszFileSpec AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszOption* | One of the following words which is used to specify the requested part: PATH, NAME, EXTN, NAMEX. |
| *wszFileSpec* | The path to be scanned. |

| Keyword    | Action      |
| ---------- | ----------- |
| **PATH** | Returns the path portion of the path/file Name. That is the text up to and including the last backslash (\) or colon (:). |
| **NAME** | Returns the name portion of the path/file Name. That is the text to the right of the last backslash (\) or colon (:), ending just before the last period (.). |
| **EXTN** | Returns the extension portion of the path/file name. That is the last period (.) in the string plus the text to the right of it. |
| **NAMEX** | Returns the name and the extension parts combined. |

---


### <a name="dwstrpathscan"></a>DWStrPathScan

Searches a path for a file name.

```
FUNCTION DWStrPathScan (BYREF wszOption AS CONST WSTRING, BYREF wszFileSpec AS CONST WSTRING,_
   BYREF wszOtherDirs AS CONST WSTRING = "") AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszOption* | One of the following words which is used to specify the requested part: PATH, NAME, EXTN, NAMEX. |
| *wszFileSpec* | The path to be scanned. |
| *wszOtherDirs* | An optional path string which includes one or more paths to be searched to find wszFileSpec. If multiple path names are included in this string, they musteach be separated by a semicolon (;) delimiter. If wszOtherDirs: is not given, or it is a null (zero-length) string, the following directories are searched:<br>- The directory from which the application was loaded.<br>- The current directory.<br>- The standard directories such as System32 and the directories specified in the PATH environment variable.<br>To expedite the process or enable DWStrPathScan to search a wider range of directories,use the wszOtherDirs parameter to specify one or more directories to be searched first. |

#### Return value

If the file is found, it returns either the full path/file name, or a selected part of it.If the file is not found, a null (zero-length)  is returned. If you wish to simply parsea text file name, without regard to its validation on disk, you should use the companionfunction *DWStrPathName*.

---

### <a name="dwstrremain"></a>DWStrRemain

Returns the portion of a string following the first occurrence of a character or group of characters. If *wszMatchString* is not present in *wszSourceString* (or is null) then a zero-length empty string is returned.

```
FUNCTION DWStrRemain (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, _
   BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
FUNCTION DWStrRemain (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, _
   BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | Starting position to begin the search. If *nStart* is not specified, the search will begin at position 1. If *nStart* is zero, a nul string is returned. If *nStart* is negative, the starting position is counted from right to left: -1 for the last character, -2 for the second to last, etc. |
| *wszSourceString* | The main string. |
| *wszMatchString* | The string to search for. |

#### Usage examples

```
DIM dws AS DWSTRING = "I think, therefore I am"
dws = DWStrRemain(dws, ",", TRUE)
PRINT dws
' Output: " therefore I am"
```
```
DIM dws AS DWSTRING = "One, two, three"
dws = DWStrRemain(5, dws, ",", TRUE)
PRINT dws
' Output: " three"
```
---

### <a name="DWStrRemove"></a>DWStrRemove

Returns a new string with strings removed.

```
FUNCTION DWStrRemove (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, _
   BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
FUNCTION DWStrRemove (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, _
   BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position to begin the search. |
| *wszSourceString* | The source string. |
| *wszMatchStr* | The string expression to be removed. If *wszMatchStr* is not present in *wszSourceString*, all of *wszSourceString* is returned intact. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Overloaded methods:

Return a copy of a string with a substring enclosed between the specified delimiters removed.

```
FUNCTION DWStrRemove (BYREF wszSourceString AS CONST WSTRING, BYREF leftDelimiter AS CONST WSTRING, _
   BYREF rightDelimiter AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
FUNCTION DWStrRemove OVERLOAD (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, _
   BYREF leftDelimiter AS CONST WSTRING, BYREF rightDelimiter AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The starting position to begin the search. |
| *wszSourceString* | The source string. |
| *leftDelimiter* | The left delimiter. |
| *rightDelimiter* | The right delimiter. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Usage examples

```
DWStrRemove("Hello World. Welcome to the Freebasic World", "World")   'Output: "Hello . Welcome to the Freebasic"
DWStrRemove("abacadabra", "bac")        ' Output: "aaabra"
DWStrRemove("abacadabra", "[bac]")      ' Output: "dr"
```
```
DIM dwsText AS DWSTRING = "As Long var1(34), var2(  73 ), var3(any)"
PRINT WstrRemove(19, dwsText, "(", ")", TRUE)   ' Returns "var2, var3"
```
---

### <a name="DWStrRepeat"></a>DWStrRepeat

Returns a string consisting of multiple copies of the specified string. This function is similar to FreeBasic `STRING`, but `STRING` only makes multiple copies of a single character.

```
FUNCTION DWStrRepeat (BYVAL nCount AS LONG, BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCount* | The number of copies. |
| *wszStr* | The string to be copied. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrRepeat(5, "Paul")
```
---

### <a name="dwstrreplace"></a>DWStrReplace

Replaces all the occurrences of *wszMatchStr* in *wszSourceString* with the contents of *wszReplaceWith*.

```
FUNCTION DWStrReplace (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, _
   BYREF wszReplaceString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string from which you want to replace the specified string. |
| *wszMatchStr* | The string expression to be replaced. |
| *wszReplaceWith* | The replacement string. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Usage examples

```
DWStrReplace("Hello World", "World", "Earth")   ' Output: "Hello Earth"
DWStrReplace("abacadabra", "bac", "***")        ' Output: "a***adabra"
DWStrReplace("abacadabra", "[bac]", "*")        ' Output: "*****d**r*"
```
---

### <a name="DWStrRetain"></a>DWStrRetain

Returns a string containing only the characters contained in a specified match string. If *wszMatchString* is an empty string, **strRetain** returns an empty string.

```
FUNCTION DWStrRetain (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, _
   BYVAL IgnoreCase AS BOOLEAN = TRUE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The source string. |
| *wszMatchStr* | The string expression to be searched. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Usage examples

```
DIM dws AS DWSTRING = "abacadabra"
dws = DWStrRetain(dws, "B", TRUE)
PRINT dws
' Output: "bb"
```
```
DIM dws AS DWSTRING = "<p>1234567890<ak;lk;l>1234567890</p>"
dws = DWStrRetain(dws, "<;/p>", TRUE)
PRINT dws
' Output: "<p><;;></p>"
```
```
DIM dws AS DWSTRING = "<p>1234567890<ak;lk;l>1234567890</p>"
dws = DWStrRetain(dws, "0123456789", TRUE)
PRINT dws
' Output: "12345678901234567890"
```
---

### <a name="dwstrreverse"></a>DWStrReverse

Reverses the contents of a string expression.

```
FUNCTION DWStrReverse (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to be reversed. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrReverse("garden")   ' Output: "nedrag"
```
---

### <a name="dwstrrset"></a>DWStrRSet

Returns a string containing a right justified string.

```
FUNCTION DWStrRSet (BYREF wszSourceString AS CONST WSTRING, BYVAL nStringLength AS LONG, _
   BYREF wszPadCharacter AS CONST WSTRING = " ") AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to be justified. |
| *nStringLength* | The length of the new string. |
| *wszPadCharacter* | The character to be used for padding. If it is not specified, the string will be padded with spaces. |

#### Usage example

```
DIM dws AS DWSTRING = DWStrRSet("FreeBasic", 20, "*")
```
---


### <a name="dwtrrsetabs"></a>DWStrRSetAbs

Right-aligns a string within the space of another string. If *wszStr* is empty, the function leaves the padding positions unchanged from their original content, rather than replacing them with spaces as `RSET` does. If *wszStr* is longer than *wszSourceString*, **DWStrRSetAbs** truncates it from the right until it fits in the result string.

```
FUNCTION DWStrRSetAbs (BYREF wszSourceString AS CONST WSTRING, BYREF wszStr AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The source string. |
| *wszStr* | The string to be right aligned inside the source string. |

#### Usage example

```
DIM dws AS DWSTRING = "NameBasic=NameBasic"
PRINT DWStrRSetAbs(dws, "FreeBasic")  ' Output: "NameBasic=FreeBasic"
```

### <a name="DWStrShrink"></a>DWStrShrink

Shrinks a string to use a consistent single character delimiter.

```
FUNCTION DWStrShrink(BYREF wszSourceString AS CONST WSTRING, BYREF wszMask AS CONST WSTRING = "", _
   BYVAL IgnoreCase AS LONG = TRUE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *wszMask* | One or more character delimiters to shrink. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Remarks

This function creates a string with consecutive words separated by a consistent single character, making it easier to parse. If *wszMask* is not specified, all leading and trailing spaces are removed and all occurrences of two or more spaces are changed to a single space. If *wszMask* contains one or more characters to shrink, all the leading and trailing occurences of them are removes and all occurrences of one or more characters of the mask are replaced with the first character of the mask.

#### Usage example

```
DIM dws AS DWSTRING = DWStrShrink(",,, one , two     three, four,", " ,")
' Output: "one two three four".
```
---

### <a name="DWStrSplit"></a>DWStrSplit

Splits a string into tokens, which are sequences of contiguous characters separated by any of the characters that are part of delimiters. Each token is added to a DWSTRING (my own dynamic Unicode string data type for FreeBasic) and delimited by a carriage return and line feed. The returned string will be parsed later to get the individual tokens.

```
FUNCTION DWStrSplit (BYREF wszStr AS WSTRING, BYREF wszDelimiters AS WSTRING = " ", _
   BYREF wszSeparator AS WSTRING = CHR(34, 44, 34), BYVAL maxsplits AS LONG = -1) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | The string to split. |
| *wszDelimiters* | The delimiter characters to use when splitting the string. |
| *wszSeparator* | Optional. Specifies the separator to use in the returned tokens. If the delimiter expression is the 3-byte value of "," which may be expressed in your source code as the string literal """,""" or as Chr(34,44,34) then a leading and trailing double-quote is added to each token. |
| *maxsplits* | Optional. Specifies how many splits to do. Default value is -1, which is "all occurrences" |

#### Return value

A list of tokens separated by the optional seoarator specified in *wszSeparator*.

#### Usage example

```
DIM wsz AS WSTRING * 260 = "- This, a sample string."
DIM dwsTokens AS DWSTRING = DWStrSplit(wsz, " ,.-", , -1)
PRINT "len dwsTokens: ", len(dwsTokens)
PRINT dwsTokens
' Output: "This","a","sample","string"
' Passing " # " in wszSeparator:
' DIM dwsTokens AS DWSTRING = DWStrSplit(wsz, " ,.-", " # ", -1)
' Output: "This # a # sample # string"
```
---

### <a name="dwstrspn"></a>DWStrSpn

Returns the index of the initial portion of a string which consists only of characters that are part of a specified set of characters.

```
FUNCTION DWStrSpn (BYREF wszText AS CONST WSTRING, BYREF wszSet AS CONST WSTRING, BYVAL IgnoreCase AS LONG = TRUE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to be scanned. |
| *wszSet* | The set of characters for which to search. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Usage example

```
DIM dwsText AS DWSTRING = "129th"
DIM dwsSet AS DWSTRING = "1234567890"
DIM n AS LONG = DWStrSpn(dwsText, dwsSet)
PRINT "The initial number has " & WSTR(n) & " digits"   ' Output: "3"
```

#### Remarks

The Windows API **StrSpnW** function and the C **wcsspn** function can also be used.

---

### <a name="DWStrTally"></a>DWStrTally

Count the number of occurrences of a string or a list of characters within a string.

```
FUNCTION DWStrTally (BYREF wszSourceString AS CONST WSTRING, BYREF wszMatchString AS CONST WSTRING, _
   BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *wszMatchStr* | The string expression to be searched. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |


#### Return value

The number of occurrences of *wszMatchStr* in *wszSourceString*.

#### Usage examples

```
DIM dws AS DWSTRING = "abacadabra"
DIM nCount AS LONG = DWStrTally(dws, "bac")
PRINT nCount
' Output: Returns 1, counting the string "bac"
```
```
DIM dws AS DWSTRING = "abacadabra"
DIM nCount AS LONG = DWStrTally(dws, "b|a|c")   ' // [bac] is the same that [b|a|c]
PRINT nCount
' Output: Returns 8, counting all "b", "a", and "c" characters.
' The | is the OR operator in regular expressions. It means "match b OR a OR c".
```
---

### <a name="DWStrUCase"></a>DWStrUCase

Returns an uppercased version of a string.

```
FUNCTION DWStrUCase (BYVAL pwszStr AS WSTRING PTR, _
   BYVAL pwszLocaleName AS WSTRING PTR = LOCALE_NAME_USER_DEFAULT, _
   BYVAL dwMapFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszStr* | The string to uppercase. |
| *pwszLocaleName* | Optional. Pointer to a locale name or one of these pre-defined values: LOCALE_NAME_INVARIANT, LOCALE_NAME_SYSTEM_DEFAULT, LOCALE_NAME_USER_DEFAULT |
| *dwMapFlags* | Optional. Flag specifying the type of transformation to use during string mapping or the type of sort key to generate. |

For a table of language culture names see: [Table of Language Culture Names, Codes, and ISO Values](https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20))

For a complete list see: [LCMapStringEx function](https://docs.microsoft.com/en-us/windows/desktop/api/winnls/nf-winnls-lcmapstringex)

#### Remarks

The string conversion functions available in FreeBasic are not fully suitable for some languages. For example, the Turkish word "karışıklığı" is uppercased as "KARıŞıKLıĞı" instead of "KARIŞIKLIĞI", and "KARIŞIKLIĞI" is lowercased to "karişikliği" instead of "karışıklığı". Notice the "ı", that is not an "i".

For Turkey, use:
```
wstrUcase("karışıklığı", "tr-TR")
DWStrLCase("KARIŞIKLIĞI", "tr-TR")
```

#### Return value

The uppercased string.

---

### <a name="DWStrUCode"></a>DWStrUCode

Translates ansi bytes to Unicode chars.

```
FUNCTION DWStrUCode (BYREF ansiStr AS CONST STRING, BYVAL nCodePage AS LONG = 0) AS DWSTRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *ansiStr* | The ANSI or UTF8 string to translate. |
| *nCodePage* | The code page used in the conversion, e.g. 1251 for Russian. If CP_UTF8 is specified, it is assumed that *ansiStr* contains an UTF8-encoded string. If the code page is omited, the function will use CP_ACP (0), which is the system default Windows ANSI code page.|

#### Return value

The translated string.

---

### <a name="dwstrunwrap"></a>DWStrUnWrap

Removes paired characters to the beginning and end of a string.

```
FUNCTION DWStrUnWrap (BYREF wszSourceString AS CONST WSTRING, BYREF wszLeftChar AS CONST WSTRING, _
   BYREF wszRightChar AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *wszLeftChar* | The left character. |
| *wszRightChar* | The right character. |

```
FUNCTION DWStrUnWrap (BYREF wszSourceString AS CONST WSTRING, BYREF wszChar AS CONST WSTRING = CHR(34)) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *wszChar* | The same character for both the left and right sides. |

#### Remarks

If only one wrap character/string is specified then that character or string is used for both sides. If no wrap character/string is specified then double quotes are used.

#### Usage examples

```
DWStrUnWrap("<Paul>", "<", ">") results in Paul
DWStrUnWrap("'Paul'", "'") results in Paul
DWStrUnWrap("""Paul""") results in Paul
```
---

### <a name="dwstrverify"></a>DWStrVerify

Determine whether each character of a string is present in another string.

```
FUNCTION DWStrVerify (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING, _
   BYREF wszMatchString AS CONST WSTRING, BYVAL IgnoreCase AS BOOLEAN = TRUE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position. |
| *wszSourceString* | The main string. |
| *wszMatchStr* | The string expression to be searched. |
| *IgnoreCase* | Boolean. If False, the search is case-sensitive; otherwise, it is case-insensitive. |

#### Return value

Returns zero if each character in *wszSourceString* is present in *wszMatchStr*; otherwise, it returns the position of the first non-matching character in *wszSourceString*. If *nStart* is zero, a negative number of a value greater that the length of *wszSourceString*, zero is returned.

#### Usage example

```
DIM dws AS DWSTRING = "123.65,22.5"
DIM nPos AS LONG = DWStrVerify(5, dws, "0123456789", TRUE)
PRINT nPos
' Output: "7"
' Returns "7" since 5 starts it past the first non-digit ("." at position 4)
```

### <a name="DWStrWrap"></a>DWStrWrap

Adds paired characters to the beginning and end of a string.

```
FUNCTION DWStrWrap (BYREF wszSourceString AS CONST WSTRING, BYREF wszLeftChar AS CONST WSTRING, _
   BYREF wszRightChar AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *wszLeftChar* | The left character. |
| *wszRightChar* | The right character. |

```
FUNCTION DWStrWrap (BYREF wszSourceString AS CONST WSTRING, BYREF wszChar AS CONST WSTRING = CHR(34)) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The main string. |
| *wszChar* | The same character for both the left and right sides. |

#### Remarks

If only one wrap character/string is specified then that character or string is used for both sides. If no wrap character/string is specified then double quotes are used.

#### Usage examples

```
DWStrWrap("Paul", "<", ">") results in <Paul>
DWStrWrap("Paul", "'") results in 'Paul'
DWStrWrap("Paul") results in "Paul"
```
---

### <a name="dwstrbase64decodea"></a>DWStrBase64DecodeA

Converts the contents of a Base64 mime encoded string to an ascii string.

```
FUNCTION DWStrBase64DecodeA (BYREF strData AS STRING) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *strData* | The string to decode. |

#### Return value

The decoded string on success, or a null string on failure.

#### Remaks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

If we want to encode a unicode string, we must convert it to utf8 before calling AfxBase64EncodeA, e.g.

````
DIM dws AS DWSTRING = "おはようございます – Good morning!"
DIM s AS STRING = wstrBase64EncodeA(dws.Utf8)
````

To decode it, we can use

````
DIM dwsOut AS dwsTRING = DWSTRING(DWStrBase64DecodeA(s), CP_UTF8)
````
or
```
DIM dwsOut AS DWSTRING
dws.utf8 = DWStrBase64DecodeA(s)
```
---

### <a name="dwstrbase64decodew"></a>DWStrBase64DecodeW

Converts the contents of a Base64 mime encoded string to an unicode string.

```
FUNCTION DWStrBase64DecodeW (BYREF dwsData AS DWSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwsData* | The string to decode. |

#### Return value

The decoded string on success, or a null string on failure.

#### Remaks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

---

### <a name="dwstrbase64encodea"></a>wstrBase64EncodeA

Converts the contents of an ascii string to Base64 mime encoding.

```
FUNCTION wstrBase64EncodeA (BYREF strData AS STRING) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *strData* | The string to encode. |

#### Return value

The encoded string on succeess, or a null string on failure.

#### Remarks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

If we want to encode a unicode string, we must convert it to utf8 before calling AfxBase64Encode, e.g.

````
DIM dws AS DWSTRING = "おはようございます – Good morning!"
DIM s AS STRING = wstrBase64EncodeA(dws.Utf8)
````

To decode it, we can use

````
DIM dwsOut AS DWSTRING = DWSTRING(DWStrBase64DecodeA(s), CP_UTF8)
````
or
````
DIM dwsOut AS DWSTRING
dws.utf8 = DWStrBase64DecodeA(s)
````
---

### <a name="dwstrbase64encodew"></a>DWStrBase64EncodeW

Converts the contents of an unicode string to Base64 mime encoding.

```
FUNCTION wstrBase64EncodWeA (BYREF dwsData AS DWSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwsData* | The string to encode. |

#### Return value

The encoded string on succeess, or a null string on failure.

#### Remarks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

---

### <a name="dwstrcryptbinarytostring"></a>DWStrCryptBinaryToString

Converts an array of bytes into a formatted string.

```
FUNCTION FUNCTION DWStrCryptBinaryToStringA ( _
   BYVAL pbBinary AS CONST UBYTE PTR, _
   BYVAL cbBinary AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pszString AS LPSTR, _
   BYVAL pcchString AS DWORD PTR _
) AS WINBOOL
```

```
FUNCTION DWStrCryptBinaryToStringW ( _
   BYVAL pbBinary AS CONST UBYTE PTR, _
   BYVAL cbBinary AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pszString AS LPWSTR, _
   BYVAL pcchString AS DWORD PTR _
) AS WINBOOL
```

| Parameter  | Description |
| ---------- | ----------- |
| *pbBinary* | A pointer to the array of bytes to be converted into a string. |
| *cbBinary* | The number of elements in the *pbBinary* array. |
| *dwFlags* | Specifies the format of the resulting formatted string. |
| *pszString* | A pointer to a buffer that receives the converted string. To calculate the number of characters that must be allocated to hold the returned string, set this parameter to NULL. The function will place the required number of characters, including the terminating NULL character, in the value pointed to by *pcchString*. |
| *pcchString* | A pointer to a DWORD variable that contains the size, in characters, of the *pszString* buffer. If *pszString* is NULL, the function calculates the length of the return string (including the terminating null character) in characters and returns it in this parameter. If *pszString* is not NULL and big enough, the function converts the binary data into a specified string format including the terminating null character, but *pcchString* receives the length in characters, not including the terminating null character. |

Values available for the *dwFlags* parameter:

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_BASE64HEADER | Base64, with certificate beginning and ending headers. |
| CRYPT_STRING_BASE64 | Base64, without headers. |
| CRYPT_STRING_BINARY | Pure binary copy. |
| CRYPT_STRING_BASE64REQUESTHEADER | Base64, with request beginning and ending headers. |
| CRYPT_STRING_HEX | Hexadecimal only. |
| CRYPT_STRING_HEXASCII | Hexadecimal, with ASCII character display. |
| CRYPT_STRING_BASE64X509CRLHEADER | Base64, with X.509 CRL beginning and ending headers. |
| CRYPT_STRING_HEXADDR | Hexadecimal, with address display. |
| CRYPT_STRING_HEXASCIIADDR | Hexadecimal, with ASCII character and address display. |
| CRYPT_STRING_HEXRAW | A raw hexadecimal string. Not supported in Windows Server 2003 and Windows XP. |
| CRYPT_STRING_STRICT | Enforce strict decoding of ASN.1 text formats. Some ASN.1 binary BLOBS can have the first few bytes of the BLOB incorrectly interpreted as Base64 text. In this case, the rest of the text is ignored. Use this flag to enforce complete decoding of the BLOB. Not suported in Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP. |

In addition to the values above, one or more of the following values can be specified to modify the behavior of the function.

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_NOCRLF | Do not append any new line characters to the encoded string. The default behavior is to use a carriage return/line feed (CR/LF) pair (0x0D/0x0A) to represent a new line. Not supported in Windows Server 2003 and Windows XP. |
| CRYPT_STRING_NOCR | Only use the line feed (LF) character (0x0A) for a new line. The default behavior is to use a CR/LF pair (0x0D/0x0A) to represent a new line. |

#### Return value

If the function succeeds, the function returns nonzero (CTRUE).
If the function fails, it returns zero (FALSE).

#### Remarks

With the exception of when **CRYPT_STRING_BINARY** encoding is used, all strings are appended with a new line sequence. By default, the new line sequence is a CR/LF pair (0x0D/0x0A). If the *dwFlags* parameter contains the **CRYPT_STRING_NOCR** flag, then the new line sequence is a LF character (0x0A). If the *dwFlags* parameter contains the **CRYPT_STRING_NOCRLF** flag, then no new line sequence is appended to the string.

---

### <a name="dwstrcryptstringtobinary"></a>DWStrCryptStringToBinary

Converts a formatted string into an array of bytes.

```
FUNCTION DWStrCryptStringToBinaryA ( _
   BYVAL pszString AS LPCSTR, _
   BYVAL cchString AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pbBinary AS UBYTE PTR, _
   BYVAL pcbBinary AS DWORD PTR, _
   BYVAL pdwSkip AS DWORD PTR, _
   BYVAL pdwFlags AS DWORD PTR _
) AS WINBOOL
```

```
FUNCTION DWStrCryptStringToBinaryW ( _
   BYVAL pszString AS LPCWSTR, _
   BYVAL cchString AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pbBinary AS UBYTE PTR, _
   BYVAL pcbBinary AS DWORD PTR, _
   BYVAL pdwSkip AS DWORD PTR, _
   BYVAL pdwFlags AS DWORD PTR _
) AS WINBOOL
```

| Parameter  | Description |
| ---------- | ----------- |
| *pszString* | A pointer to a string that contains the formatted string to be converted. |
| *cchString* | The number of characters of the formatted string to be converted, not including the terminating NULL character. If this parameter is zero, pszString is considered to be a null-terminated string. |
| *dwFlags* | Indicates the format of the string to be converted. |
| *pbBinary* | A pointer to a buffer that receives the returned sequence of bytes. If this parameter is NULL, the function calculates the length of the buffer needed and returns the size, in bytes, of required memory in the DWORD pointed to by *pcbBinary*. |
| *pcbBinary* | A pointer to a DWORD variable that, on entry, contains the size, in bytes, of the *pbBinary* buffer. After the function returns, this variable contains the number of bytes copied to the buffer. If this value is not large enough to contain all of the data, the function fails and GetLastError returns **ERROR_MORE_DATA**. If *pbBinary* is NULL, the DWORD pointed to by *pcbBinary* is ignored. |
| *pdwSkip* | A pointer to a DWORD value that receives the number of characters skipped to reach the beginning of the actual base64 or hexadecimal strings. This parameter is optional and can be NULL if it is not needed. |
| *pdwFlags* | A pointer to a DWORD value that receives the flags actually used in the conversion. These are the same flags used for the *dwFlags* parameter. In many cases, these will be the same flags that were passed in the *dwFlags* parameter. If *dwFlags* contains one of the flags inicated below, this value will receive a flag that indicates the actual format of the string. This parameter is optional and can be NULL if it is not needed. |

Values available for the *dwFlags* parameter:

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_BASE64HEADER | Base64, with certificate beginning and ending headers. |
| CRYPT_STRING_BASE64 | Base64, without headers. |
| CRYPT_STRING_BINARY | Pure binary copy. |
| CRYPT_STRING_BASE64REQUESTHEADER | Base64, with request beginning and ending headers. |
| CRYPT_STRING_BASE64 | Base64, without headers. |
| CRYPT_STRING_HEX | Hexadecimal only. |
| CRYPT_STRING_HEXASCII | Hexadecimal, with ASCII character display. |
| CRYPT_STRING_BASE64_ANY | Tries the following, in order: CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64. |
| CRYPT_STRING_ANY | Tries the following, in order: CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64, CRYPT_STRING_BINARY. |
| CRYPT_STRING_HEX_ANY | Tries the following, in order: CRYPT_STRING_HEXADDR, CRYPT_STRING_HEXASCIIADDR, CRYPT_STRING_HEX, CRYPT_STRING_HEXRAW, CRYPT_STRING_HEXASCII. |
| CRYPT_STRING_HEX | Hexadecimal only. |
| CRYPT_STRING_BASE64X509CRLHEADER | Base64, with X.509 CRL beginning and ending headers. |
| CRYPT_STRING_HEXADDR | Hexadecimal, with address display. |
| CRYPT_STRING_HEXASCIIADDR | Hexadecimal, with ASCII character and address display. |
| CRYPT_STRING_HEXRAW | A raw hexadecimal string. Not supported in Windows Server 2003 and Windows XP. |
| CRYPT_STRING_STRICT | Enforce strict decoding of ASN.1 text formats. Some ASN.1 binary BLOBS can have the first few bytes of the BLOB incorrectly interpreted as Base64 text. In this case, the rest of the text is ignored. Use this flag to enforce complete decoding of the BLOB. Not suported in Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP. |

Values available for the *pdwFlags* parameter:

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_ANY | This variable will receive one of the following values. Each value indicates the actual format of the string. CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64, CRYPT_STRING_BINARY. |
| CRYPT_STRING_BASE64_ANY | This variable will receive one of the following values. Each value indicates the actual format of the string. CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64. |
| CRYPT_STRING_HEX_ANY | This variable will receive one of the following values. Each value indicates the actual format of the string. CRYPT_STRING_HEXADDR, CRYPT_STRING_HEXASCIIADDR, CRYPT_STRING_HEX, CRYPT_STRING_HEXRAW, CRYPT_STRING_HEXASCII. |

---

### <a name="dwstrhassurrogates"></a>DWStrHasSurrogates

Checks if the specified string has surrogates.
```
FUNCTION DWStrHasSurrogates (BYREF wszStr AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | The string to parse. |

#### Return value

BOOLEAN. True if the string has rurrogates; False, otherwise.

---

### <a name="dwstrisvalidsurrogatepair"></a>DWStrIsValidSurrogatePair

Checks whether a UTF-16 encoded string contains valid high-low surrogate pairs.
```
FUNCTION DWStrIsValidSurrogatePair (BYVAL high AS USHORT, BYVAL low AS USHORT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *high* | The high surrogate part. |
| *low* | The low surrogate part. |

#### Return value

BOOLEAN. True if the surrogate pair is valid; False, otherwise.

---

### <a name="dwstrsurrogatepairtocodepoint"></a>DWStrSurrogatePairToCodePoint

Converts a surrogate pair to a Unicode code point.
```
FUNCTION DWStrSurrogatePairToCodePoint (BYVAL high AS USHORT, BYVAL low AS USHORT) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *high* | The high surrogate part. |
| *low* | The low surrogate part. |

#### Return value

BOOLEAN. True if the surrogate pair is valid; False, otherwise.

---

### <a name="wstrcodepointtosurrogatpair"></a>DWStrCodePointToSurrogatePair

Converts a Unicode code point (above U+FFFF) back into its high and low surrogate pair.
```
SUB DWStrCodePointToSurrogatePair (BYVAL codePoint AS ULONG, BYREF high AS USHORT, BYREF low AS USHORT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *high* | The high surrogate part. |
| *low* | The low surrogate part. |

#### Return value

This method does not return a value.

---

### <a name="dwstrscanforsurrogates"></a>DWStrScanForSurrogates

Scans a UTF-16 buffer (passed as a pointer to WSTRING) in chunks of 64 characters. Returns the 0-based index (relative to *memAddr*) of the first broken surrogate found, or -1 if none is found.

| Parameter  | Description |
| ---------- | ----------- |
| *memAddr* | Pointer to the passes UTF-16 buffer. |
| *nChars* | Number of UTF-16 code units (USHORTs) to scan. |
| *searchBrokenOnly* |  Optional (default TRUE): if True, only broken surrogates are signaled. If False, returns the position of the first surrogate (valid or not).  |

| Parameter  | Description |
| ---------- | ----------- |
| *memAddr* | The high surrogate part. |
| *nChars* | The low surrogate part. |
| *searchBrokenOnly* | The low surrogate part. |

```
FUNCTION DWStrScanForSurrogates( _
   BYVAL memAddr AS WSTRING PTR, _
   BYVAL nChars AS LONG, _
   BYVAL searchBrokenOnly AS BOOLEAN = TRUE) AS LONG
```
#### Result code

Returns the 0-based index (relative to *memAddr*) of the first broken surrogate found, or -1 if none is found. If the optional parameter *searchBrokenOnly* is set to false, then it returns the 0-based index of the first surroghate found, broken or not.

---

### <a name="dwstrchrw"></a>DWStrChrW

Returns a wide-character string from a codepoint.
```
FUNCTION DWStrChrW (BYVAL codepoint AS UInteger) AS DWSTRING
   If codepoint <= &hFFFF Then RETURN WCHR(codepoint)
   ' Convert to UTF-16 surrogate pair for higher codepoints
   DIM AS USHORT highSurrogate = &hD800 OR ((codepoint - &h10000) SHR 10)
   DIM AS USHORT lowSurrogate = &hDC00 OR ((codepoint - &h10000) AND &h3FF)
   RETURN WCHR(highSurrogate) + WCHR(lowSurrogate)
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *codePoint* | The value of a Unicode code point. |

#### Return value

The codepoint returned is the sum of a surrogate pair.

---
