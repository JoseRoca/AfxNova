# CRegExp Class

`CRegExp` is a wrapper class on top of the VBScript Regular Expressions.

It hides the complexity of COM programming (creating an instance of the COM server, managing pointers to interfaces, avoiding the use of BSTR and VARIANT data types, and bypassing low-level virtual table calls), offering a clean interface to FreeBasic users by exposing only the standard `WSTRING` data type and my own `DWSTRING` class, which implements a dynamic Unicode string type.

In a typical search and replace operation, you must provide the exact text for which you are searching. That technique may be adequate for simple search and replace tasks in static text, but it lacks flexibility and makes searching dynamic text difficult, if not impossible.

With regular expressions, you can:

* Test for a pattern within a string. For example, you can test an input string to see if a telephone number pattern or a credit card number pattern occurs within the string. This is called data validation.

* Replace text. You can use a regular expression to identify specific text in a document and either remove it completely or replace it with other text.

* Extract a substring from a string based upon a pattern match. You can find specific text within a document or input field.

**Include file**: AfxNova/CRegExp.inc

**Tutorial**: [Introduction to Regular Expressions](#Introduction)

**Recipes**: [Recipes](#recipes)

### Properties

| Name  | Description |
| ---------- | ----------- |
| [Global](#global) | Sets or returns a boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| [IgnoreCase](#ignorecase) | Sets or returns a boolean value that indicates if a pattern search is case-sensitive or not. |
| [MatchLen](#matchlen) | Returns the length of a match found in a search string. |
| [MatchPos](#matchpos) | Returns the position in a search string where a match occurs. |
| [MatchValue](#matchvalue) | Returns the value or text of a match found in a search string. |
| [Multiline](#multiline) | Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines. |
| [Pattern](#pattern) | Sets or returns the regular expression pattern being searched for. |
| [SubMatchesCount](#submatchescount) | Returns the number of submatches. |

### Methods

| Name  | Description |
| ---------- | ----------- |
| [Execute](#execute) | Executes a regular expression search against a specified string. |
| [Extract](#extract) | Extracts a substring using VBScript regular expressions search patterns. |
| [Find](#find) | Find function with VBScript regular expressions search patterns. |
| [FindEx](#findex) | Global, multiline find function with VBScript regular expressions search patterns. |
| [MatchCount](#matchcount) | Returns the number of matches found. |
| [RegExpPtr](#regexpptr) | Returns a direct pointer to the **IRegExp2** interface. |
| [Remove](#remove) | Returns a copy of a string with text removed using a regular expression as the search string. |
| [Replace](#replace) | Replaces text found in a regular expression search. |
| [SubMatchValue](#submatchvalue) | Retrieves the content of the specified submatch. |
| [Test](#test) | Executes a regular expression search against a specified string and returns a boolean value that indicates if a pattern match was found. |

## Error and result codes

| Name       | Description |
| ---------- | ----------- |
| [GetErrorInfo](#geterrorinfo) | Returns a localized description of the specified error code. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [SetResult](#setresult) | Sets the last result code. |

---

### <a name="geterrorinfo"></a>GetErrorInfo

Returns a localized description of the specified error code. If the error is omited, it will return the value returned by the Windows API function **GetLastError**.
```
PRIVATE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
```

### <a name="getlastresult"></a>GetLastResult

Returns the last result code
```
FUNCTION GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
```
---

### <a name="setresult"></a>SetResult

Sets the last result code.
```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
```
| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

---

### <a name="execute"></a>Execute

Executes a regular expression search against a specified string.

```
FUNCTION Execute (BYREF wszSourceString AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to be parsed. |

#### Return value

BOOLEAN. True on success or False on failure.

#### Example
```
#include once "AfxNova/CRegExp2.inc"
USING AfxNova
DIM pRegExp AS CRegExp
pRegExp.Pattern = "is."
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute("IS1 is2 IS3 is4") = FALSE THEN
   PRINT "No match found"
ELSE
   DIM nCount AS LONG = pRegExp.MatchesCount
   FOR i AS LONG = 0 TO nCount - 1
      PRINT "Value: ", pRegExp.MatchValue(i)
      PRINT "Position: ", pRegExp.MatchPos(i)
      PRINT "Length: ", pRegExp.MatchLen(i)
      PRINT
   NEXT
END IF
```

---

### <a name="extract"></a>Extract

Extracts a substring using VBScript regular expressions search patterns.

```
FUNCTION Extract (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
FUNCTION Extract (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The position in the string at which the search will begin. The first character starts at position 1. |
| *wszSourceString* | The string to be parsed. |

#### Return value

DWSTRING. Returns the extracted string on exit or an empty string on failure.

#### Usage examples

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
dwsText = pRegExp.Extract(dwsText)
PRINT dwsText
' Output: "a234"
```
```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
dwsText = pRegExp.Extract(15, dwsText)
PRINT dwsText
' Output: "x345"
```

---

### <a name="find"></a>Find

Find function with VBScript regular expressions search patterns.

```
FUNCTION Find (BYREF wszSourceString AS CONST WSTRING) AS LONG
FUNCTION Find (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The position in the string at which the search will begin. The first character starts at position 1. |
| *wszSourceString* | The string to be parsed. |

#### Return value

Returns the position of the match or 0 if not found. The length of the match can be retrieved calling the **MatchLen** property.

#### Usage examples

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
DIM nPos AS LONG = pRegExp.Find(dwsText)
PRINT nPos
' Output: 11
```
```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
DIM nPos AS LONG = pRegExp.Find(15, dwsText)
PRINT nPos
' Output: 26
```

---

### <a name="findex"></a>FindEx

Global, multiline find function with VBScript regular expressions search patterns.

```
FUNCTION FindEx (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to be parsed. |

#### Return value

Returns a list of comma separated "index, length" value pairs. The pairs are separated by a semicolon.

#### Usage example

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
DIM dwsOut AS DWSTRING = pRegExp.FindEx(dwsText)
PRINT dwsOut
' Output: 11,4;26,4
```

---

### <a name="matchcount"></a>MatchCount

Returns the number of matches found.

```
FUNCTION MatchCount () AS LONG
```

---

### <a name="regexpptr"></a>RegExpPtr

Returns a direct pointer to the IRegExp2 interface.

```
FUNCTION RegExpPtr () AS IRegExp2 PTR
```
#### Remarks

Since it is a direct pointer, you don't have to release it calling the **Release** method.

---

### <a name="remove"></a>Remove

Returns a copy of a string with text removed using a regular expression as the search string.

```
FUNCTION Remove (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to parse. |

#### Usage examples

```
DIM pRegExp AS CRegExp
pRegExp.Pattern = "ab"
pRegExp.Global = TRUE
PRINT pRegExp.Remove("abacadabra")
' Output: "acadra"
```
```
DIM pRegExp AS CRegExp
pRegExp.Pattern = "[bAc]"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
PRINT pRegExp.Remove("abacadabra")
' Output: "dr"
```

---

### <a name="replace"></a>Replace

Replaces text found in a regular expression search.

```
FUNCTION Replace (BYREF wszSourceString AS CONST WSTRING, BYREF wszReplaceString AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to parse. |
| *wszReplaceString* | The replacement string. |

#### Remarks

The **Replace** method returns a copy of *wszSourceString* with the text replaced with *wszReplaceString*. If no match is found, a copy of *wszSourceString* is returned unchanged.

#### Examples

```
#INCLUDE ONCE "AfxNova/CRegExp.inc"
USING AfxNova

DIM pRegExp AS CRegExp
pRegExp.Pattern = "fox"
pRegExp.IgnoreCase = TRUE
DIM dwsText AS DWSTRING = "The quick brown fox jumped over the lazy dog."
' Make replacement
DIM dwsRes AS DWSTRING = pRegExp.Replace(dwsText, "cat")
PRINT dwsRes
' Output: The quick brown cat jumped over the lazy dog.
```
In addition, the **Replace** method can replace subexpressions in the pattern.

The following call to the function shown in the previous example swaps the first pair of words in the original string:

```
DIM pRegExp AS CRegExp
pRegExp.Pattern = "(\S+)(\s+)(\S+)"
pRegExp.IgnoreCase = TRUE
DIM dwsText AS DWSTRING = "The quick brown fox jumped over the lazy dog."
' Make replacement
DIM dwsRes AS DWsTRING = pRegExp.Replace(dwsText, "$3$2$1")
PRINT dwsRes
' Outpput: "quick The brown fox jumped over the lazy dog."
```

Suppose that you have a telephone directory, and all the phone numbers are formatted like this:
555-123-4567. Now, you decide that all the phone numbers should be formatted to look like this: (555) 123-4567.

```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = "(\d{3})-(\d{3})-(\d{4})"
DIM dwsText AS DWSTRING = "555-123-4567"
DIM dwsRes AS DWSTRING = pRegExp.Replace(dwsText, "($1) $2-$3")
print dwsRes
' Output: "(555) 123-4567"
```
What we have done is to search for 3 digits (\d{3}) followed by a dash, followed by 3 more digits and a dash, followed by 4 digits and add () to the first three digits and change the first dash with a space.  $1, $2, and $3 are examples of a regular expression "back reference." A back reference is simply a portion of the found text that can be saved and then reused.

#### Other usage examples
```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.IgnoreCase = TRUE
pRegExp.Pattern = "World"
PRINT pRegExp.Replace("Hello World", "Earth") 
 'Output: "Hello Earth"
```
```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = "[bac]"
' --or-- pRegExp.Pattern = "b|a|c"
PRINT pRegExp.Replace("abacadabra", "*")
' Output: "*****d**r*"
```
```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = $"(\S+), (\S+)"
PRINT pRegExp.Replace("Roca, José", "$2 $1")
' Output: "José Roca"
```
```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = $"\b0{1,}\."
PRINT pRegExp.Replace("0000.34500044", ".")
' Output: ".34500044"
```

---

### <a name="submatchvalue"></a>SubMatchValue

Retrieves the content of the specified submatch.

```
FUNCTION SubMatchValue (BYVAL MatchIndex AS LONG = 0, BYVAL SubMatchIndex AS LONG = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *MatchIndex* | 0-based index of the match to retrieve. |
| *SubMatchIndex* | 0-based index of the submatch to retrieve. |

---

### <a name="test"></a>Test

Executes a regular expression search against a specified string and returns a boolean value that indicates if a pattern match was found.

```
FUNCTION Test (BYREF wszSourceString AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSourceString* | The string to parse. |

#### Return value

BOOLEAN. True if a pattern match is found; False if no match is found.

#### Usage example

```
DIM pRegExp AS CRegEx
pRegExp.Pattern = "is."
pRegExp.IgnoreCase = TRUE
IF pRegExp.Test("IS1 is2 IS3 is4") = FALSE THEN
  PRINT "No match found"
ELSE
   PRINT "Match found"
END IF
```

---

### <a name="global"></a>Global

Sets or returns a boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one.

```
PROPERTY Global () AS BOOLEAN
PROPERTY Global (BYVAL bGlobal AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bGlobal* | True if the search applies to the entire string, False if it does not. Default is False. |

---

### <a name="ignorecase"></a>IgnoreCase

Sets or returns a boolean value that indicates if a pattern search is case-sensitive or not.

```
PROPERTY IgnoreCase () AS BOOLEAN
PROPERTY IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bIgnoreCase* | False if the search is case-sensitive, True if it is not. Default is False. |

---

### <a name="matchlen"></a>MatchLen

Returns the length of a match found in a search string.

```
PROPERTY MatchLen (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

---

### <a name="matchpos"></a>MatchPos

Returns the position in a search string where a match occurs.

```
PROPERTY MatchPos (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

#### Remarks

The **MatchPos** property uses a zero-based offset from the beginning of the search string. In other words, the first character in the string is identified as character zero (0).

---

### <a name="matchvalue"></a>MatchValue

Returns the value or text of a match found in a search string.

```
PROPERTY MatchValue (BYVAL index AS LONG = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

---

### <a name="multiline"></a>Multiline

Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines.

```
PROPERTY Multiline () AS BOOLEAN
PROPERTY Multiline (BYVAL bMultiline AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bMultiline* | True if the search is performed across multiple lines, False if it is not. Default is False. |

---

### <a name="pattern"></a>Pattern

Sets or returns the regular expression pattern being searched for.

```
PROPERTY Pattern () AS DWSTRING
PROPERTY Pattern (BYREF wszPattern AS CONST WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPattern* | Regular string expression being searched for. May include any of the regular expression characters defined in the table in the **Settings** section. |

#### Settings

Special characters and sequences are used in writing patterns for regular expressions. The following table describes and gives an example of the characters and sequences that can be used.

| Character  | Description |
| ---------- | ----------- |
| \\ | Marks the next character as either a special character or a literal. For example, "n" matches the character "n". "\\n" matches a newline character. The sequence "\\\\" matches "\\" and "\\(" matches "(". |
| ^ | Matches the beginning of input. |
| $ | Matches the end of input. |
| * | Matches the preceding character zero or more times. For example, "zo*" matches either "z" or "zoo". |
| + | Matches the preceding character one or more times. For example, "zo+" matches "zoo" but not "z". |
| ? | Matches the preceding character zero or one time. For example, "a?ve?" matches the "ve" in "never". |
| . | Matches any single character except a newline character. |
| (pattern) | Matches pattern and remembers the match. The matched substring can be retrieved from the resulting Matches collection, using Item \[0]...\[n]. To match parentheses characters ( ), use "\\(" or "\\)". |
| x\|y | Matches either x or y. For example, "z\|wood" matches "z" or "wood". "(z\|w)oo" matches "zoo" or "wood". |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, "o{2}" does not match the "o" in "Bob," but matches the first two o's in "foooood". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, "o{2,}" does not match the "o" in "Bob" and matches all the o's in "foooood." "o{1,}" is equivalent to "o+". "o{0,}" is equivalent to "o*". |
| { n , m } | m and n are nonnegative integers. Matches at least n and at most m times. For example, "o{1,3}" matches the first three o's in "fooooood." "o{0,1}" is equivalent to "o?". |
| \[xyz] | A character set. Matches any one of the enclosed characters. For example, "\[abc]" matches the "a" in "plain".  |
| \[^xyz] | A negative character set. Matches any character not enclosed. For example, "\[^abc]" matches the "p" in "plain".  |
| \[a-z] | A range of characters. Matches any character in the specified range. For example, "\[a-z]" matches any lowercase alphabetic character in the range "a" through "z". |
| \[^m-z] | A negative range characters. Matches any character not in the specified range. For example, "\[m-z]" matches any character not in the range "m" through "z".  |
| \\b | Matches a word boundary, that is, the position between a word and a space. For example, "er\b" matches the "er" in "never" but not the "er" in "verb". |
| \\B | Matches a non-word boundary. "ea\*r\B" matches the "ear" in "never early". |
| \\d | Matches a digit character. Equivalent to \[0-9]. |
| \\D | Matches a non-digit character. Equivalent to \[^0-9]. |
| \\f | Matches a form-feed character. |
| \\n | Matches a newline character. |
| \\r | Matches a carriage return character. |
| \\s | Matches any white space including space, tab, form-feed, etc. Equivalent to "\[ \f\n\r\t\v]". |
| \\S | Matches any nonwhite space character. Equivalent to "\[^ \f\n\r\t\v]".  |
| \\t | Matches a tab character. |
| \\v | Matches a vertical tab character. |
| \\w | Matches any word character including underscore. Equivalent to "\[A-Za-z0-9_]". |
| \\W | Matches any non-word character. Equivalent to "\[^A-Za-z0-9_]". |
| \\num | Matches num, where num is a positive integer. A reference back to remembered matches. For example, "(.)\1" matches two consecutive identical characters. |
| \\n | Matches n, where n is an octal escape value. Octal escape values must be 1, 2, or 3 digits long. For example, "\11" and "\011" both match a tab character. "\0011" is the equivalent of "\001" & "1". Octal escape values must not exceed 256. If they do, only the first two digits comprise the expression. Allows ASCII codes to be used in regular expressions. |
| \\xn | Matches n, where n is a hexadecimal escape value. Hexadecimal escape values must be exactly two digits long. For example, "\x41" matches "A". "\x041" is equivalent to "\x04" & "1". Allows ASCII codes to be used in regular expressions. |

#### Example

```
#INCLUDE ONCE "AfxNova/CRegExp2.inc"
USING AfxNova

DIM pRegExp AS CRegExp
pRegExp.Pattern = "is."
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute("IS1 is2 IS3 is4") = FALSE THEN
   print "No match found"
ELSE
   DIM nCount AS LONG = pRegExp.MatchesCount
   FOR i AS LONG = 0 TO nCount - 1
      print "Value: ", pRegExp.MatchValue(i)
      print "Position: ", pRegExp.MatchPos(i)
      print "Length: ", pRegExp.MatchLen(i)
      print
   NEXT
END IF
```

---

### <a name="submatchescount"></a>SubMatchesCount

Returns the number of submatches.

```
PROPERTY SubMatchesCount (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

---

# <a name="Introduction"></a>Introduction to Regular Expressions

Unless you have worked with regular expressions before, the term and the concept may be unfamiliar to you. However, they may not be as unfamiliar as you think.

### The Concept of Regular Expressions

Think about how you search for files on your hard disk. You most likely use the ? and * characters to find files. The ? character matches a single character in a file name, while the * matches zero or more characters. A pattern such as 'data?.dat' would find the following files:

data1.dat
data2.dat
datax.dat
dataN.dat

Using the * character instead of the ? character expands the number of files found. 'data*.dat' matches all of the following:

data.dat
data1.dat
data2.dat
data12.dat
datax.dat
dataXYZ.dat

While this method of searching for files can certainly be useful, it is also very limited. The limited ability of the ? and * wildcard characters give you an idea of what regular expressions can do, but regular expressions are much more powerful and flexible.

### Uses for Regular Expressions

In a typical search and replace operation, you must provide the exact text for which you are searching. That technique may be adequate for simple search and replace tasks in static text, but it lacks flexibility and makes searching dynamic text difficult, if not impossible. 

Flexibility of Regular Expressions

With regular expressions, you can: 

* Test for a pattern within a string. For example, you can test an input string to see if a telephone number pattern or a credit card number pattern occurs within the string. This is called data validation.

* Replace text. You can use a regular expression to identify specific text in a document and either remove it completely or replace it with other text.

* Extract a substring from a string based upon a pattern match. You can find specific text within a document or input field

For example, if you need to search an entire web site to remove some outdated material and replace some HTML formatting tags, you can use a regular expression to test each file to see if the material or the HTML formatting tags you are looking for exists in that file. That way, you can narrow down the affected files to only those that contain the material that has to be removed or changed. You can then use a regular expression to remove the outdated material, and finally, you can use regular expressions to search for and replace the tags that need replacing.

### Regular Expression Syntax

A regular expression is a pattern of text that consists of ordinary characters (for example, letters a through z) and special characters, known as metacharacters. The pattern describes one or more strings to match when searching a body of text. The regular expression serves as a template for matching a character pattern to the string being searched.

The following table contains the complete list of metacharacters and their behavior in the context of regular expressions:

| Character  | Description |
| ---------- | ----------- |
| \\ | Marks the next character as either a special character or a literal. For example, "n" matches the character "n". "\\n" matches a newline character. The sequence "\\\\" matches "\\" and "\\(" matches "(". |
| ^ | Matches the beginning of input. |
| $ | Matches the end of input. |
| * | Matches the preceding character zero or more times. For example, "zo*" matches either "z" or "zoo". |
| + | Matches the preceding character one or more times. For example, "zo+" matches "zoo" but not "z". |
| ? | Matches the preceding character zero or one time. For example, "a?ve?" matches the "ve" in "never". |
| . | Matches any single character except a newline character. |
| (pattern) | Matches pattern and remembers the match. The matched substring can be retrieved from the resulting Matches collection, using Item \[0]...\[n]. To match parentheses characters ( ), use "\\(" or "\\)". |
| x\|y | Matches either x or y. For example, "z\|wood" matches "z" or "wood". "(z\|w)oo" matches "zoo" or "wood". |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, "o{2}" does not match the "o" in "Bob," but matches the first two o's in "foooood". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, "o{2,}" does not match the "o" in "Bob" and matches all the o's in "foooood." "o{1,}" is equivalent to "o+". "o{0,}" is equivalent to "o*". |
| { n , m } | m and n are nonnegative integers. Matches at least n and at most m times. For example, "o{1,3}" matches the first three o's in "fooooood." "o{0,1}" is equivalent to "o?". |
| \[xyz] | A character set. Matches any one of the enclosed characters. For example, "\[abc]" matches the "a" in "plain".  |
| \[^xyz] | A negative character set. Matches any character not enclosed. For example, "\[^abc]" matches the "p" in "plain".  |
| \[a-z] | A range of characters. Matches any character in the specified range. For example, "\[a-z]" matches any lowercase alphabetic character in the range "a" through "z". |
| \[^m-z] | A negative range characters. Matches any character not in the specified range. For example, "\[m-z]" matches any character not in the range "m" through "z".  |
| \\b | Matches a word boundary, that is, the position between a word and a space. For example, "er\b" matches the "er" in "never" but not the "er" in "verb". |
| \\B | Matches a non-word boundary. "ea\*r\B" matches the "ear" in "never early". |
| \\d | Matches a digit character. Equivalent to \[0-9]. |
| \\D | Matches a non-digit character. Equivalent to \[^0-9]. |
| \\f | Matches a form-feed character. |
| \\n | Matches a newline character. |
| \\r | Matches a carriage return character. |
| \\s | Matches any white space including space, tab, form-feed, etc. Equivalent to "\[ \f\n\r\t\v]". |
| \\S | Matches any nonwhite space character. Equivalent to "\[^ \f\n\r\t\v]".  |
| \\t | Matches a tab character. |
| \\v | Matches a vertical tab character. |
| \\w | Matches any word character including underscore. Equivalent to "\[A-Za-z0-9_]". |
| \\W | Matches any non-word character. Equivalent to "\[^A-Za-z0-9_]". |
| \\num | Matches num, where num is a positive integer. A reference back to remembered matches. For example, "(.)\1" matches two consecutive identical characters. |
| \\n | Matches n, where n is an octal escape value. Octal escape values must be 1, 2, or 3 digits long. For example, "\11" and "\011" both match a tab character. "\0011" is the equivalent of "\001" & "1". Octal escape values must not exceed 256. If they do, only the first two digits comprise the expression. Allows ASCII codes to be used in regular expressions. |
| \\xn | Matches n, where n is a hexadecimal escape value. Hexadecimal escape values must be exactly two digits long. For example, "\x41" matches "A". "\x041" is equivalent to "\x04" & "1". Allows ASCII codes to be used in regular expressions. |

Examples of Regular Expressions

"^\\s\*$" Match a blank line.<br>
"\\d{2}-\\d{5}" Validate an ID number consisting of 2 digits, a hyphen, and another 5 digits.

#### Constructing a Regular Expression

Regular expressions are constructed in the same way that arithmetic expressions are created. That is, small expressions are combined using a variety of metacharacters and operators to create larger expressions.

You construct a regular expression by putting the various components of the expression pattern between a pair of quotation marks ("") delimit regular expressions. For example: "expression". The regular expression pattern (expression) is stored in the **Pattern** property.

The components of a regular expression can be individual characters, sets of characters, ranges of characters, choices between characters, or any combination of all of these components.

Once you have constructed a regular expression, it is evaluated much like an arithmetic expression, that is, it is evaluated from left to right and follows an order of precedence. 

#### Order of Precedence

From highest to lowest, the order of precedence of the regular expression operators:

| Operator(s) | Description |
| ---------- | ----------- |
| \\ | Escape |
| (), (?:), (?=), \[] | Parentheses and brackets |
| \*, +, ?, {n}, {n,}, {n,m} | Quantifiers |
| ^, $, \\anymetacharacter | Anchors and Sequences |
| \| | Alternation |

Characters have higher precedence than the alternation operator, which allows 'm|food' to match "m" or "food". To match "mood" or "food", use parentheses to create a subexpression, which results in '(m|f)ood'.

#### Ordinary Characters

Ordinary characters consist of all printable and non-printable characters that are not explicitly designated as metacharacters. This includes all uppercase and lowercase alphabetic characters, all digits, all punctuation marks, and some symbols. 

The simplest form of a regular expression is a single, ordinary character that matches itself in a searched string. For example, the single-character pattern 'A' matches the letter 'A' wherever it appears in the searched string. Here are some examples of single-character regular expression patterns:

"a"<br>
"7"<br>
"M"

You can combine a number of single characters to form a larger expression. For example, the following regular expression is nothing more than an expression created by combining the single-character expressions 'a', '7', and 'M'. 

"a7M"

Notice that there is no concatenation operator. All that is required is that you just put one character after another.

#### Special Characters

There are a number of metacharacters that require special treatment when trying to match them. To match these special characters, you must first escape those characters, that is, precede them with a backslash character (\\). 

| Special character | Meaning |
| ---------- | ----------- |
| $ | Matches the position at the end of an input string. If the RegExp object's Multiline property is set, $ also matches the position preceding '\\n' or '\\r'. To match the $ character itself, use \\$. |
| ( ) | Marks the beginning and end of a subexpression. Subexpressions can be captured for later use. To match these characters, use \\( and \\). |
| * | Matches the preceding character or subexpression zero or more times. To match the * character, use \\\*. |
| + | Matches the preceding character or subexpression one or more times. To match the + character, use \\+. |
| . | Matches any single character except the newline character \\n. To match ., use \\. |
| \[ ] | Marks the beginning of a bracket expression. To match these characters, use \\\[ and \\]. |
| ? | Matches the preceding character or subexpression zero or one time, or indicates a non-greedy quantifier. To match the ? character, use \\?. |
| \ | Marks the next character as either a special character, a literal, a backreference, or an octal escape. For example, 'n' matches the character 'n'. '\\n' matches a newline character. The sequence '\\\\' matches "\\" and '\\(' matches "(". |
| / | Denotes the start or end of a literal regular expression. To match the '/' character, use '\\/'. |
| ^ | Matches the position at the beginning of an input string except when used in a bracket expression where it negates the character set. To match the ^ character itself, use \\^. |
| { } | Marks the beginning of a quantifier expression. To match these characters, use \\{ and \}. |
| \| | Indicates a choice between two items. To match \|, use \\\|. |

#### Escape Sequences that Represent Non-printing Characters

There are a number of useful non-printing characters that must be used occasionally.

| Character  | Meaning     |
| ---------- | ----------- |
| \\cx | Matches the control character indicated by x. For example, \\cM matches a Control-M or carriage return character. The value of x must be in the range of A-Z or a-z. If not, c is assumed to be a literal 'c' character. |
| \\f | Matches a form-feed character. Equivalent to \\x0c and \\cL. |
| \\n | Matches a newline character. Equivalent to \\x0a and \\cJ. |
| \\r | Matches a carriage return character. Equivalent to \\x0d and \\cM. |
| \\s | Matches any white space character including space, tab, form-feed, and so on. Equivalent to \[\f\n\r\t\v]. |
| \\S | Matches any non-white space character. Equivalent to \[^\f\n\r\t\v]. |
| \\t | Matches a tab character. Equivalent to \\x09 and \\cI. |
| \\v | Matches a vertical tab character. Equivalent to \\x0b and \\cK. |

#### Character Matching

The period (.) matches any single printing or non-printing character in a string, except a newline character (\n). The following regular expression matches 'aac', 'abc', 'acc', 'adc', and so on, as well as 'a1c', 'a2c', a-c', and a#c': 

"a.c"

If you are trying to match a string containing a file name where a period (.) is part of the input string, you do so by preceding the period in the regular expression with a backslash (\\) character. To illustrate, the following regular expression matches 'filename.ext':

"filename\\.ext"

These expressions are still pretty limited. They only let you match any single character. Many times, it is useful to match specified characters from a list. For example, if you have an input text that contains chapter headings that are expressed numerically as Chapter 1, Chapter 2, and so on, you might want to find those chapter headings. 

#### Bracket Expressions

You can create a list of matching characters by placing one or more individual characters within square brackets (\[ and ]). When characters are enclosed in brackets, the list is called a bracket expression. Within brackets, as anywhere else, ordinary characters represent themselves, that is, they match an occurrence of themselves in the input text. Most special characters lose their meaning when they occur inside a bracket expression. Here are some exceptions: 

The ']' character ends a list if it is not the first item. To match the ']' character in a list, place it first, immediately following the opening '\['.

The '\\' character continues to be the escape character. To match the '\\' character, use '\\\\'.

Characters enclosed in a bracket expression match only a single character for the position in the regular expression where the bracket expression appears. The following regular expression matches 'Chapter 1', 'Chapter 2', 'Chapter 3', 'Chapter 4', and 'Chapter 5':

"Chapter \[12345]"

Notice that the word 'Chapter' and the space that follows are fixed in position relative to the characters within brackets. The bracket expression then, is used to specify only the set of characters that matches the single character position immediately following the word 'Chapter' and a space. That is the ninth character position.

If you want to express the matching characters using a range instead of the characters themselves, you can separate the beginning and ending characters in the range using the hyphen (-) character. The character value of the individual characters determines their relative order within a range. The following regular expression contains a range expression that is equivalent to the bracketed list shown above.

"Chapter \[1-5]"

When a range is specified in this manner, both the starting and ending values are included in the range. It is important to note that the starting value must precede the ending value in Unicode sort order.

If you want to include the hyphen character in your bracket expression, you must do one of the following: 

Escape it with a backslash: 

\[\\-]

Put the hyphen character at the beginning or the end of the bracketed list. The following expressions matches all lowercase letters and the hyphen: 

\[-a-z]
\[a-z-]

Create a range where the beginning character value is lower than the hyphen character and the ending character value is equal to or greater than the hyphen. Both of the following regular expressions satisfy this requirement: 

\[!--]
\[!-~]

You can also find all the characters not in the list or range by placing the caret (^) character at the beginning of the list. If the caret character appears in any other position within the list, it matches itself, that is, it has no special meaning. The following regular expression matches chapter headings with numbers greater than 5':

"Chapter \[^12345]"

In the examples shown above, the expression matches any digit character in the ninth position except 1, 2, 3, 4, or 5. So, for example, 'Chapter 7' is a match and so is 'Chapter 9'. 

The same expressions above can be represented using the hyphen character (-):

"Chapter \[^1-5]"

A typical use of a bracket expression is to specify matches of any upper- or lowercase alphabetic characters or any digits. The following expression specifies such a match:

"\[A-Za-z0-9]"

#### Quantifiers

Sometimes, you do not know how many characters there are to match. In order to accommodate that kind of uncertainty, regular expressions support the concept of quantifiers. These quantifiers let you specify how many times a given component of your regular expression must occur for your match to be true.

| Character  | Meaning     |
| ---------- | ----------- |
| \* | Matches the preceding character or subexpression zero or more times. For example, 'zo*' matches "z" and "zoo". * is equivalent to {0,}. |
| + | Matches the preceding character or subexpression one or more times. For example, 'zo+' matches "zo" and "zoo", but not "z". + is equivalent to {1,}. |
| ? | Matches the preceding character or subexpression zero or one time. For example, 'do(es)?' matches the "do" in "do" or "does". ? is equivalent to {0,1}. |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, 'o{2}' does not match the 'o' in "Bob," but matches the two o's in "food". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, 'o{2,}' does not match the 'o' in "Bob" and matches all the o's in "foooood". 'o{1,}' is equivalent to 'o+'. 'o{0,}' is equivalent to 'o*'. |
| {n,m} | m and n are nonnegative integers, where n <= m. Matches at least n and at most m times. For example, 'o{1,3}' matches the first three o's in "fooooood". 'o{0,1}' is equivalent to 'o?'. Note that you cannot put a space between the comma and the numbers. |

With a large input document, chapter numbers could easily exceed nine, so you need a way to handle two or three digit chapter numbers. Quantifiers give you that capability. The following regular expression matches chapter headings with any number of digits:

"Chapter \[1-9][0-9]*"

Notice that the quantifier appears after the range expression. Therefore, it applies to the entire range expression which, in this case, specifies only digits from 0 through 9, inclusive.

The '+' quantifier is not used here because there does not necessarily need to be a digit in the second or subsequent position. The '?' character also is not used because it limits the chapter numbers to only two digits. You want to match at least one digit following 'Chapter' and a space character.

If you know that your chapter numbers are limited to only 99 chapters, you can use the following expression to specify at least one, but not more than 2 digits.

"Chapter \[0-9]{1,2}"

The disadvantage to the expression shown above is that if there is a chapter number greater than 99, it will still only match the first two digits. Another disadvantage is that somebody could create a Chapter 0 and it would match. A better expression for matching only two digits are the following:
 
"Chapter \[1-9][0-9]?"

-or-
 	
"Chapter \[1-9]\[0-9]{0,1}"

The '\*', '+', and '?' quantifiers are all what are referred to as greedy, that is, they match as much text as possible. Sometimes that is not at all what you want to happen. Sometimes, you just want a minimal match. 

Say, for example, you are searching an HTML document for an occurrence of a chapter title enclosed in an H1 tag. That text appears in your document as:

\<H1>Chapter 1 – Introduction to Regular Expressions\</H1>

The following expression matches everything from the opening less than symbol (<) to the greater than symbol (>) at the end of the closing H1 tag.

"\<\.\*>"

If all you really wanted to match was the opening H1 tag, the following, non-greedy expression matches only \<H1>.

"\<\.\*?>"

By placing the '?' after a '\*', '+', or '?' quantifier, the expression is transformed from a greedy to a non-greedy, or minimal, match.

#### Anchors

So far, the examples you've seen have been concerned only with finding chapter headings wherever they occur. Any occurrence of the string 'Chapter' followed by a space, followed by a number, could be an actual chapter heading, or it could also be a cross-reference to another chapter. Since true chapter headings always appear at the beginning of a line, you'll need to devise a way to find only the headings and not find the cross-references.

Anchors provide that capability. Anchors allow you to fix a regular expression to either the beginning or end of a line. They also allow you to create regular expressions that occur either within a word or at the beginning or end of a word. The following table contains the list of regular expression anchors and their meanings:

| Character  | Meaning     |
| ---------- | ----------- |
| ^ | Matches the position at the beginning of the input string. If the RegExp object's Multiline property is set, ^ also matches the position following '\\n' or '\\r'. |
| $ | Matches the position at the end of the input string. If the RegExp object's Multiline property is set, $ also matches the position preceding '\\n' or '\\r'. |
| \\b | Matches a word boundary, that is, the position between a word and a space. |
| \\B | Matches a nonword boundary. |

You cannot use a quantifier with an anchor. Since you cannot have more than one position immediately before or after a newline or word boundary, expressions such as '^\*' are not permitted.

To match text at the beginning of a line of text, use the '^' character at the beginning of the regular expression. Do not confuse this use of the '^' with the use within a bracket expression. 

To match text at the end of a line of text, use the '$' character at the end of the regular expression.

To use anchors when searching for chapter headings, the following regular expression matches a chapter heading with up to two following digits that occurs at the beginning of a line:
 
"^Chapter \[1-9]\[0-9]{0,1}"

Not only does a true chapter heading occur at the beginning of a line, it is also the only text on the line, so it also must be at the end of a line as well. The following expression ensures that the match specified only matches chapters and not cross-references. It does so by creating a regular expression that matches only at the beginning and end of a line of text.
 
"^Chapter \[1-9]\[0-9]{0,1}$"

Matching word boundaries is a little different but adds a very important capability to regular expressions. A word boundary is the position between a word and a space. A nonword boundary is any other position. The following expression matches the first three characters of the word 'Chapter' because they appear following a word boundary:
 
"\\bCha"

The position of the '\\b' operator is critical. If it is positioned at the beginning of a string to be matched, it looks for the match at the beginning of the word; if it is positioned at the end of the string, it looks for the match at the end of the word. For example, the following expression matches 'ter' in the word 'Chapter' because it appears before a word boundary:
 
"ter\\b"

The following expression matches 'apt' as it occurs in 'Chapter', but not as it occurs in 'aptitude':
 
"\\Bapt"

The string 'apt' occurs on a nonword boundary in the word 'Chapter' but on a word boundary in the word 'aptitude'. For the \\B nonword boundary operator, position is not important because the match is not relative to the beginning or end of a word.
 
#### Alternation and Grouping

Alternation allows use of the '\|' character to allow a choice between two or more alternatives. Expanding the chapter heading regular expression, you can expand it to cover more than just chapter headings. However, it is not as straightforward as you might think. When alternation is used, the largest possible expression on either side of the '\|' character is matched. You might think that the following expression match either 'Chapter' or 'Section' followed by one or two digits occurring at the beginning and ending of a line:

Example

"^Chapter|Section [1-9][0-9]{0,1}$"

Unfortunately, the regular expression shown above matches either the word 'Chapter' at the beginning of a line, or 'Section' and whatever numbers follow that, at the end of the line. If the input string is 'Chapter 22', the expression shown above only matches the word 'Chapter'. If the input string is 'Section 22', the expression matches 'Section 22'. But that is not the intent here so there must be a way to make that regular expression more responsive to what you're trying to do and there is.

You can use parentheses to limit the scope of the alternation, that is, make sure that it applies only to the two words, 'Chapter' and 'Section'. However, parentheses are also used to create subexpressions and possibly capture them for later use, something that is covered in the section on backreferences. By taking the regular expressions shown above and adding parentheses in the appropriate places, you can make the regular expression match either 'Chapter 1' or 'Section 3'. 

The following regular expression uses parentheses to group 'Chapter' and 'Section' so the expression works properly.

"^(Chapter|Section) \[1-9]\[0-9]{0,1}$"

Although these expressions work properly, the parentheses around 'Chapter\|Section' also cause either of the two matching words to be captured for future use. Since there is only one set of parentheses in the expression shown above, there is only one captured submatch. This submatch can be referred to using the Submatches collection.

In the above example, you merely want to use the parentheses to group a choice between the words 'Chapter' and 'Section'. To prevent the match from being saved for possible later use, place '?:' before the regular expression pattern inside the parentheses. The following modification provides the same capability without saving the submatch:

"^(?:Chapter|Section) \[1-9]\[0-9]{0,1}$"

In addition to the '?:' metacharacters, there are two other non-capturing metacharacters used for something called lookahead matches. A positive lookahead, specified using ?=, matches the search string at any point where a matching regular expression pattern in parentheses begins. A negative lookahead, specified using '?!', matches the search string at any point where a string not matching the regular expression pattern begins.

For example, suppose you have a document containing references to Windows 3.1, Windows 95, Windows 98, and Windows NT. Suppose further that you need to update the document by finding all the references to Windows 95, Windows 98, and Windows NT and changing those reference to Windows 2000. You can use the following regular expression, which is an example of a positive lookahead, to match Windows 95, Windows 98, and Windows NT:

"Windows(?=95 |98 |NT )"

Once the match is found, the search for the next match begins immediately following the matched text, not including the characters included in the look-ahead. For example, if the expressions shown above matched 'Windows 98', the search resumes after 'Windows' not after '98'.

---

# <a name="recipes"></a>Recipes

#### Finding similar words

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "My cat found a dead bat over the mat"
pRegExp.Pattern = $"\b[bcm]at\b"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
'-or- pRegPattern = "\b(b|c|m)at\b"
pRegExp.Execute(dwsText)
FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
   PRINT pRegExp.MatchValue(i)
NEXT
```
As an alternative, we can use the operator "\|":

```
pRegExp.Pattern = $"\b(b|c|m)at\b"
```

\b is a word boundary, [bcm] one of b, c, or m, followed by at and finally another word boundary (\b),

```
Output:
cat
bat
mat
```
---

#### Finding variations on words

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "Hello, my name is John Doe, what's your name?"
pRegExp.Pattern = $"\bJoh?n(athan)? Doe\b"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
pRegExp.Execute(dwsText)
IF pRegExp.MatchesCount THEN print pRegExp.MatchValue
' Output "John Doe"
```

\b is a word boundary (matches whole words only), Jo is the common part of John, Jon and Jonathan, h? is optional (it is followed by the metacharacter ?), n is a letter common in the three names, (anathan)? is an optional group of characters that will match Jonathan, a space followed by Doe, and another word boundary (\b).

---

#### Removing characters

```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = "[abc]"
DIM dws AS DWSTRING = pRegExp.Remove("abacadabra")
PRINT dws ' - prints "dr"
```

Removes any of the characters, a, b or c, found in the string.

---

#### Removing words

```
DIM pRegExp AS CRegExp
pRegExp.Pattern = "ab"
pRegExp.Global = TRUE
DIM dws AS DWSTRING = pRegExp.Remove("abacadabra")
PRINT dws
'Output: acadra
```
```
DIM pRegExp AS CRegExp
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
pRegExp.Pattern = $"\bworld\b"
DIM dws AS DWSTRING = pRegExp.Remove("World, worldx, world")
PRINT dws
' Output: ", worldx,"
```

\b is word boundary.

---

#### Replacing characters

```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = "[bac]"
DIM dws AS DWSTRING = pRegExp.Replace("abacadabra", "*")
PRINT dws
' Output: "*****d**r*"
```

Replaces any of the characters, a, b or c, found in the string with an asterisk.

---

#### Replacing words

```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.IgnoreCase = TRUE
pRegExp.Pattern = "World"
DIM dws AS DWSTRING = pRegExp.Replace("Hello World", "Earth")
PRINT dws
' Output: "Hello Earth"
```

Replaces any occurrence of two consecutive identical words in a string of text with a single occurrence of the same word. The TRUE in the fourth parameter indicates that the operation is not case sensitive.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "Is is the cost of of gasoline going up up?"
pRegExp.Global = TRUE
pRegExp.IgnoreCase = TRUE
pRegExp.Pattern = $"\b([a-z]+) \1\b"
DIM dws AS DWSTRING = pRegExp.Replace(dwsText, "$1")
PRINT dws
' Output: "Is the cost of gasoline going up?"
```

Adds a space after the dots that are immediately followed by a word.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a text.With some dots.Between words. This one not."
pRegExp.Pattern = "(\.)(\w)"
pRegExp.Global = TRUE
pRegExp.IgnoreCase = TRUE
DIM dws AS DWSTRING = pRegExp.Replace(dwsText, "$1 $2")
PRINT dws
' Output: This is a text. With some dots. Between words. This one not."
```

Changes the format of a telephone number.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "555-123-4567"
pRegExp.Pattern = "(\d{3})-(\d{3})-(\d{4})"
pRegExp.Global = TRUE
pRegExp.IgnoreCase = TRUE
DIM dws AS DWSTRING = pRegExp.Replace(dwsText, "($1) $2-$3")
PRINT dws
' Output: (555) 123-4567
```

Removes leading zeroes.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "0000.34500044"
pRegExp.Pattern = $"\b0{1,}\."
pRegExp.Global = TRUE
pRegExp.IgnoreCase = TRUE
DIM dws AS DWSTRING = pRegExp.Replace(dwsText, ".")
print dws
' Output: ".34500044"
```

Replaces everything between quotes with three asterisks.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a ""quoted"" string"
pRegExp.Pattern = """[^""]*"""
pRegExp.Global = TRUE
DIM dws AS DWSTRING = pRegExp.Replace(dwsText, "***")
print dws
' Output: "This is a "***"" string"
```

Replaces a tab with a pipe character.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a " & CHR(9) & " test string"
'DIM dwsText AS DWSTRING = !"This is a \t test string" ' alternative way
pRegExp.Pattern = $"\t"
pRegExp.Global = TRUE
DIM dws AS DWSTRING = pRegExp.Replace(dwsText, "|")
print dws
' Output: "This is a | test string"
```
---

#### Searching words and substrings

Searches for the first occurence of a word that starts with a letter ans is followed by three numbers. If found, it returns its position.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
pRegExp.Global = TRUE
DIM nPos AS LONG = pRegExp.Find(dwsText)
PRINT nPos
' Output: 11
```

Searches all the occurences of a word that starts with a letter ans is followed by three numbers. Returns a list of comma separated "index, length" value pairs. The pairs are separated by a semicolon.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
pRegExp.Global = TRUE
DIM dws AS DWSTRING = pRegExp.FindEx(dwsText)
PRINT dws
' Output: 11,4; 26,4
```
---

#### Extracting words and substrings

Searches for the first occurrence of a word. Case-sensitive and not global.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a test string"
pRegExp.Pattern = "test"
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN PRINT pRegExp.MatchValue
' Output: "test"
```

Searches for the first occurrence of a word. Case-insensitive and not global.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a test string"
pRegExp.Pattern = "test"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = FALSE
IF pRegExp.Execute(dwsText) THEN PRINT pRegExp.MatchValue
' Output: "test"
```

Searches for the first occurrence of a substring. Case-sensitive and not global.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a test string"
pRegExp.Pattern = "is a test"
pRegExp.IgnoreCase = FALSE
pRegExp.Global = FALSE
IF pRegExp.Execute(dwsText) THEN PRINT pRegExp.MatchValue
' Output: "is a test"
```

Searches for the first occurrence of a substring. Case-insensitive and not global.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a test string"
pRegExp.Pattern = "is a test"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = FALSE
IF pRegExp.Execute(dwsText) THEN PRINT pRegExp.MatchValue
' Output: "is a test"
```

Searches for the all the occurrences of a word. Case-sensitive and global.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a test testx string"
pRegExp.Pattern = "test"
pRegExp.IgnoreCase = FALSE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output:
' test
' test
```

Searches for the all the occurrences of a word. Case-insensitive and global.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a test testx string"
pRegExp.Pattern = "test"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output:
' test
' test
```

Searches for the all the occurrences of a whole word. Case-insensitive and global.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "This is a test testx string"
pRegExp.Pattern = $"\bTest\b"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output: "test"
```

Case sensitive, double search (c.t and d.g), whole words. Retrieves cut, cat, i.e. whole words with three letters that begin with c and end with t or begin with d and end with g.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "I have a cat and a dog, because I love cats and dogs"
pRegExp.Pattern = "c.t|d.g"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output:
' cat
' dog
' cat
' dog
```

Case insensitive, double search (c.t and d.g), whole words. Retrieves cat, dog, i.e. whole words with three letters that begin with c and end with t or begin with d and end with g.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "I have a cat and a dog, because I love cats and dogs"
pRegExp.Pattern = $"\bc.t\b|\bd.g\b"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output:
' cat
' dog
```

We can search for more than a word at the same time.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "The contests will be in January, July and November"
pRegExp.Pattern = $"\b(january|february|march|april|may|june|july|" & _
    $"august|september|october|november|december)\b"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output:
' January
' July
' November
```

Extracts a quoted string.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "29:Sep:2017 ""This is an Example!"""
pRegExp.Pattern = """.*?"""
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN PRINT pRegExp.MatchValue
' Output:
' "This is an Example!"
```

Extracts all the alphabetic words.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "29:Sep:2017 ""This is an Example!"""
pRegExp.Pattern = "(?:[a-z][a-z]+)"
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output:
' Sep
' This
' is
' an
' Example
```

Extracts the year.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "29:Sep:2017 ""This is an Example!"""
pRegExp.Pattern = $"((?:(?:[1]{1}\d{1}\d{1}\d{1})|(?:[2]{1}\d{3})))(?![\d])"
IF pRegExp.Execute(dwsText) THEN
   PRINT pRegExp.MatchValue
END IF
' Output:
' 2017
```

Extracts integers.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "29:Sep:2017 ""This is an Example!"""
pRegExp.Pattern = $"(\d+)"
pRegExp.Global = TRUE
IF pRegExp.Execute(dwsText) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
' Output:
' 29
' 2017
```

Extracts text between parentheses.

```
' // extract the first match
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah blah (text beween parentheses) blah blah"
pRegExp.Pattern = $"([^(]*?)(?=\))"
DIM dws AS DWSTRING = pRegExp.Extract(dwsText)
PRINT dws
```
```
' // extract the first match after the 11th position
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah (xxx) blah (text beween parentheses) blah blah"
pRegExp.Pattern = $"([^(]*?)(?=\))"
DIM dws AS DWSTRING = pRegExp.Extract(11, dwsText)
PRINT dws
```
```
Output:
"text between parentheses"
```

Extracts text between curly braces.

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah (xxx) blah {text beween curly braces} blah blah"
pRegExp.Pattern = $"([^{]*?)(?=\})"
DIM dws AS DWSTRING = pRegExp.Extract(dwsText)
PRINT dws
```
```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "blah (xxx) blah {text beween curly braces} blah blah"
pRegExp.Pattern = $"([^{]*?)(?=\})"
DIM dws AS DWSTRING = pRegExp.Extract(11, dwsText)
PRINT dws
```
```
Output:
"text between curly braces"
```
---

#### Checking if a string is numeric

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "1.2345678901234567e+029"
pRegExp.Pattern = $"^[\+\-]?\d*\.?\d+(?:[Ee][\+\-]?\d+)?$"
PRINT pRegExp.Test(dwsText)
' Output: True
```

```
Pattern: "^[\+\-]?\d*\.?\d+(?:[Ee][\+\-]?\d+)?$"
```

The initial "^" and the final "$" match the start and the end of the string, to ensure the check spans the whole string.

The "\[\+\-]?" part is the inital plus or minus sign with the "?" multiplier that allows zero or one instance of it.

The "\d*" is a digit, zero or more times.

"\.?" is the decimal point, zero or one time.

The "\d+" part matches a digit one or more times.

The "(?:\[Ee]\[\+\-]?\d+)?" matches "e+", "e-", "E+" or "E-" followed by one or more digits, with the "?" multiplier that allows zero or one instance of it.

---

#### Checking if an url is valid

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "https://www.google.es/search?q=msdn+jscript+check+if+an+url+is+valid&client=firefox-b&dcr=0&ei=hM_OWfb2BMzSU-mhr7AG&start=20&sa=N&biw=947&bih=394"
pRegExp.Pattern = $"(ftp|http|https):\/\/(\w+:{0,1}\w*@)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%@!\-\/]))?"
PRINT pRegExp.Test(dwsText)
' Output: True
```
---

#### Breaking down an URI in its parts

```
DIM pRegExp AS CRegExp
DIM dwsText AS DWSTRING = "http://msdn.microsoft.com:80/scripting/default.htm"
pRegExp.Pattern = $"(\w+):\/\/([^/:]+)(:\d*)?([^# ]*)"
PRINT pRegExp.Execute(dwsText)
FOR i AS LONG = 0 TO pRegExp.SubMatchesCount - 1
   PRINT pRegExp.SubMatchValue(0, i)
NEXT
' Output:
' http
' msdn.microsoft.com
' :80
' /scripting/default.htm
```
---
