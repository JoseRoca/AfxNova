' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxJson.inc
' Purpose:  Extensions to the FreeBasic headers for Windows
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
#include once "AfxNova/DWSTRING.inc"
#Include Once "AfxNova/BSTRING.inc"
#Include Once "AfxNova/DVARIANT.inc"
#Include Once "AfxNova/DSafeArray.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' Unquote a Json string
' This correctly reconstructs characters from JSON escapes, including surrogate pairs.
' JSON encodes astral characters as two consecutive \uXXXX sequences; decoding them into
' their two UTF-16 code units preserves the character in a DWSTRING.
' Usage with WebView2
' Direct string outputs: If your JS returns a string, ExecuteScript gives you a JSON
' string like "\"José \\ud83d\\ude03\"".
' Pass it to JsonUnquoteW to get the plain text.
' Non-string outputs: If your JS returns a number, boolean, null, array, or object, the
' result won’t be quoted.
' For strings only, use the unquoter; otherwise, handle as needed.
' Assume resultJson is the UTF-16 JSON returned by WebView2 ExecuteScript.
' DIM resultJson   ' e.g., "\"José \\ud83d\\ude03\""
' DIM dwsRes AS DWSTRING = JSonUnquoteW(!"\"José \\ud83d\\ude03\"")
' IF LEN(dwsRes) THEN
'    AfxMsg dwsRes
' ELSE
'    Not a JSON string; handle other JSON types or keep the raw JSON.
'    AfxMsg "Non-string JSON: " & resultJson
' END IF
' ========================================================================================
PRIVATE FUNCTION JSonUnquoteW (BYREF wszJson AS WSTRING) AS DWSTRING
   DIM n AS INTEGER = LEN(wszJson)
   IF n < 2 THEN RETURN ""
   IF ASC(MID(wszJson, 1, 1)) <> 34 ORELSE ASC(MID(wszJson, n, 1)) <> 34 THEN RETURN ""
   DIM result AS DWSTRING
   DIM i AS INTEGER = 2   ' skip opening quote
   DIM last AS INTEGER = n - 1   ' before closing quote
   WHILE i <= last
      DIM ch AS ULONG = ASC(MID(wszJson, i, 1))
      IF ch = 92 THEN ' backslash
         i += 1
         IF i > last THEN EXIT WHILE
         DIM e AS ULONG = ASC(MID(wszJson, i, 1))
         SELECT CASE e
            CASE 34  : result += WCHR(34)   ' \"
            CASE 92  : result += WCHR(92)   ' \\
            CASE 47  : result += WCHR(47)   ' \/
            CASE ASC("b") : result += WCHR(8)
            CASE ASC("f") : result += WCHR(12)
            CASE ASC("n") : result += WCHR(10)
            CASE ASC("r") : result += WCHR(13)
            CASE ASC("t") : result += WCHR(9)
            CASE ASC("u")
                ' Expect 4 hex digits -> one UTF-16 code unit
               IF i + 4 <= last THEN
                  DIM v AS ULONG
                  FOR k AS LONG = 1 to 4
                     DIM h AS ULONG = ASC(MID(wszJson, i + k, 1))
                     SELECT CASE h
                        CASE 48 to 57  : v = (v shl 4) or (h - 48)
                        CASE 65 to 70  : v = (v shl 4) or (h - 55)   ' A-F
                        CASE 97 to 102 : v = (v shl 4) or (h - 87)   ' a-f
                        CASE ELSE
                           EXIT FOR
                     END SELECT
                  NEXT
                  result += WCHR(v)
                  i += 4
               END IF
            CASE ELSE
                ' Unknown escape, keep as-is
               result += WCHR(e)
         END SELECT
      ELSE
         result += WCHR(ch)
      END IF
      i += 1
    WEND
    RETURN result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Json token type enumeration
' ========================================================================================
ENUM JsonTokenType
   JSON_NONE = 0
   JSON_OBJECT_START
   JSON_OBJECT_END
   JSON_ARRAY_START
   JSON_ARRAY_END
   JSON_STRING
   JSON_NUMBER
   JSON_BOOL
   JSON_NULL
   JSON_COLON
   JSON_COMMA
END ENUM

' ========================================================================================
' Json token struct
' ========================================================================================
TYPE JsonToken
   kind  AS JsonTokenType
   value AS DWSTRING  ' only for string/number/bool/null
END TYPE

' ########################################################################################
' Class JSonReader
' ########################################################################################
TYPE JsonReader

 private:
   m_buf AS DWSTRING   ' buffer
   m_pos AS INTEGER    ' position

   DECLARE SUB SkipWhitespace()
   DECLARE FUNCTION ReadString() AS DWSTRING
   DECLARE FUNCTION ReadNumber() AS DWSTRING

public:
   DECLARE CONSTRUCTOR (BYREF source AS WSTRING)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION ReadNext (BYREF tok AS JsonToken) AS BOOLEAN

END TYPE

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR JsonReader (BYREF source AS WSTRING)
   m_buf = source
   m_pos = 1
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR JsonReader
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' // Skip white space
' ========================================================================================
PRIVATE SUB JsonReader.SkipWhitespace
   WHILE m_pos <= LEN(m_buf)
      SELECT CASE ASC(MID(m_buf, m_pos, 1))
         CASE 9, 10, 13, 32
            m_pos += 1
         CAsE ELSE
            EXIT WHILE
      END SELECT
   WEND
END SUB
' ========================================================================================

' ========================================================================================
' // Read string by slicing the raw JSON string and unquoting via JSonUnquoteW
' ========================================================================================
' // Read string by slicing and then unquoting
PRIVATE FUNCTION JsonReader.ReadString () AS DWSTRING
   DIM startPos AS INTEGER = m_pos   ' points at opening quote
   m_pos += 1
   WHILE m_pos <= LEN(m_buf)
      DIM ch AS ULONG = ASC(MID(m_buf, m_pos, 1))
      IF ch = 34 THEN                 ' "
         m_pos += 1
         EXIT WHILE
      ELSEIF ch = 92 THEN             ' \
         IF m_pos + 1 <= LEN(m_buf) THEN
            DIM e AS ULONG = ASC(MID(m_buf, m_pos + 1, 1))
            IF e = ASC("u") AND m_pos + 5 <= LEN(m_buf) THEN
               m_pos += 6             ' \uXXXX
            ELSE
               m_pos += 2             ' simple escape
            END IF
         ELSE
            EXIT WHILE
         END IF
      ELSE
         m_pos += 1
      END IF
   WEND
   DIM rawJson AS DWSTRING = MID(m_buf, startPos, m_pos - startPos)
   RETURN JSonUnquoteW(rawJson)
END FUNCTION
' ========================================================================================

' ========================================================================================
' // Read number: [-]? (0|[1-9]\d*) (.\d+)? ([eE][+-]?\d+)?
' ========================================================================================
' // Read number: [-]? (0|[1-9]\d*) (.\d+)? ([eE][+-]?\d+)?
PRIVATE FUNCTION JsonReader.ReadNumber () AS DWSTRING
   DIM result AS DWSTRING
   DIM startPos AS INTEGER = m_pos
   ' optional leading minus
   IF m_pos <= LEN(m_buf) THEN
      DIM ch AS ULONG = ASC(MID(m_buf, m_pos, 1))
      IF ch = 45 THEN ' -
         result += WCHR(45)
         m_pos += 1
      END IF
   END IF
   ' integer part
   DIM digits AS INTEGER = 0
   IF m_pos <= LEN(m_buf) THEN
      DIM ch AS ULONG = ASC(MID(m_buf, m_pos, 1))
      IF ch = 48 THEN ' 0
         result += WCHR(48)
         m_pos += 1
         digits = 1
      ELSEIF ch >= 49 AND ch <= 57 THEN ' 1-9
         DO WHILE m_pos <= LEN(m_buf)
            ch = ASC(MID(m_buf, m_pos, 1))
            IF ch >= 48 AND ch <= 57 THEN
               result += WCHR(ch)
               m_pos += 1
               digits = 1
            ELSE
               EXIT DO
            END IF
         LOOP
      END IF
   END IF
   ' fractional part
   IF m_pos <= LEN(m_buf) THEN
      DIM ch AS ULONG = ASC(MID(m_buf, m_pos, 1))
      IF ch = 46 THEN ' .
         DIM fracDigits AS INTEGER = 0
         result += WCHR(46)
         m_pos += 1
         DO WHILE m_pos <= LEN(m_buf)
            ch = ASC(MID(m_buf, m_pos, 1))
            IF ch >= 48 AND ch <= 57 THEN
               result += WCHR(ch)
               m_pos += 1
               fracDigits = 1
            ELSE
               EXIT DO
            END IF
         LOOP
         IF fracDigits = 1 THEN digits = 1
      END IF
   END IF
   ' exponent
   IF m_pos <= LEN(m_buf) THEN
      DIM ch AS ULONG = ASC(MID(m_buf, m_pos, 1))
      IF ch = 69 OR ch = 101 THEN ' E/e
         DIM expDigits AS INTEGER = 0
         result += WCHR(ch)
         m_pos += 1
         ' optional sign
         IF m_pos <= LEN(m_buf) THEN
            ch = ASC(MID(m_buf, m_pos, 1))
            IF ch = 43 OR ch = 45 THEN ' +/-
               result += WCHR(ch)
               m_pos += 1
            END IF
         END IF
         ' digits
         DO WHILE m_pos <= LEN(m_buf)
            ch = ASC(MID(m_buf, m_pos, 1))
            IF ch >= 48 AND ch <= 57 THEN
               result += WCHR(ch)
               m_pos += 1
               expDigits = 1
            ELSE
               EXIT DO
            END IF
         LOOP
         IF expDigits = 0 THEN
            ' invalid exponent: roll back exponent portion
            ' (simple approach: truncate back to before 'e' by resetting)
            m_pos = startPos + LEN(result) - 1  ' before 'e'
            ' and trim result back to before 'e'
            DO WHILE LEN(result) > 0 AND RIGHT(result, 1) <> WCHR(69) AND RIGHT(result, 1) <> WCHR(101)
               result = LEFT(result, LEN(result) - 1)
            LOOP
            IF LEN(result) > 0 THEN result = LEFT(result, LEN(result) - 1)
         END IF
      END IF
   END IF
   ' ensure we consumed something numeric (at least one digit)
   IF digits = 0 THEN
      ' Failsafe: back out entirely so caller can recover
      m_pos = startPos
      RETURN ""
   END IF
   RETURN result
END FUNCTION
' ========================================================================================

' ========================================================================================
' // Read next token
' ========================================================================================
PRIVATE FUNCTION JsonReader.ReadNext (BYREF tok AS JsonToken) AS BOOLEAN
   this.SkipWhitespace
   IF m_pos > LEN(m_buf) THEN RETURN FALSE
   DIM startPos AS INTEGER = m_pos
   tok.value = ""  ' clear previous value
   DIM ch AS ULONG = ASC(MID(m_buf, m_pos, 1))
   SELECT CASE ch
      CASE 123 : tok.kind = JSON_OBJECT_START : m_pos += 1
      CASE 125 : tok.kind = JSON_OBJECT_END   : m_pos += 1
      CASE 91  : tok.kind = JSON_ARRAY_START  : m_pos += 1
      CASE 93  : tok.kind = JSON_ARRAY_END    : m_pos += 1
      CASE 44  : tok.kind = JSON_COMMA        : m_pos += 1
      CASE 58  : tok.kind = JSON_COLON        : m_pos += 1
      CASE 34  ' "
         tok.kind = JSON_STRING
         tok.value = this.ReadString
      CASE 48 TO 57, 45
         tok.kind = JSON_NUMBER
         tok.value = this.ReadNumber
      CASE ELSE
         ' literals: true / false / null (JSON requires lowercase)
         IF LCASE(MID(m_buf, m_pos, 4)) = "true" THEN
            tok.kind = JSON_BOOL : tok.value = "true" : m_pos += 4
         ELSEIF LCASE(MID(m_buf, m_pos, 5)) = "false" THEN
            tok.kind = JSON_BOOL : tok.value = "false" : m_pos += 5
         ELSEIF LCASE(MID(m_buf, m_pos, 4)) = "null" THEN
            tok.kind = JSON_NULL : tok.value = "null" : m_pos += 4
         ELSE
            tok.kind = JSON_NONE
            m_pos = LEN(m_buf) + 1
         END IF
   END SELECT
   ' Defensive: if no progress, consume one char to avoid infinite loops
   IF m_pos = startPos THEN
      tok.kind = JSON_NONE
      m_pos += 1
   END IF
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Json test
' ========================================================================================
''#CONSOLE ON
'#define UNICODE
'#define _WIN32_WINNT &h0602
'#include once "AfxNova/AfxJson.inc"
'USING AfxNova

'SUB TestJsonReader()
'   ' Craft some JSON with tricky bits:
'   ' - Unicode \u00E9 (é) and surrogate pair \uD83D\uDE03 (??)
'   ' - Quotes, backslashes, slash, control codes
'   ' - Numbers with exponents, booleans, null
'   DIM sample AS DWSTRING = _
'       "{""name"":""Jos" & WCHR(&h00E9) & " " & WCHR(&hD83D) & WCHR(&hDE03) & """, " & _
'       """quote"":""He said """"Hello/World""""!"", " & _
'       """newline"":""Line1" & WCHR(10) & "Line2"", " & _
'       """pi"":3.14159, ""exp"":-2.5e+3, " & _
'       """ok"":true, ""missing"":null}"

'   DIM rdr AS JsonReader = JsonReader(sample)
'   DIM tok AS JsonToken
'   DIM idx AS INTEGER = 0

'   PRINT "Testing JSON Reader with JSonUnquoteW-backed ReadString"
'   PRINT STRING(60,"-")

'   WHILE rdr.ReadNext(tok)
'      idx += 1
'      PRINT idx; ". "; 
'      SELECT CASE tok.kind
'         CASE JSON_STRING
'            PRINT "STRING: "; tok.value
'         CASE JSON_NUMBER
'            PRINT "NUMBER: "; tok.value
'         CASE JSON_BOOL
'            PRINT "BOOL: "; tok.value
'         CASE JSON_NULL
'            PRINT "NULL"
'         CASE JSON_OBJECT_START
'            PRINT "{"
'         CASE JSON_OBJECT_END
'            PRINT "}"
'         CASE JSON_ARRAY_START
'            PRINT "["
'         CASE JSON_ARRAY_END
'            PRINT "]"
'         CASE JSON_COLON
'            PRINT ":"
'         CASE JSON_COMMA
'            PRINT ","
'         CASE ELSE
'            PRINT "UNKNOWN"
'      END SELECT
'   WEND

'   PRINT STRING(60,"-")
'   PRINT "Done."
'END SUB

'TestJsonReader
' ========================================================================================


' ########################################################################################
' Class JSonWriter
' ########################################################################################
TYPE JsonWriter
 
Private:
   buf As BSTRING
   firstItemStack(ANY) AS BOOLEAN
   indentSize AS LONG = 4
   depth AS LONG
   inlineThreshold As Integer = 40 ' max chars for inline array/object

   DECLARE FUNCTION IsSmallInline (ByRef tmp As String) As Boolean
   DECLARE SUB AppendEscaped (ByRef s As WString)
   DECLARE SUB AppendCommaIfNeeded()
   DECLARE SUB AppendNewlineAndIndent()
   DECLARE SUB ValueNull()
   DECLARE SUB ValueBool(ByVal b As Boolean)

Public:

   DECLARE SUB SetIndentSize(ByVal n As LONG)
   DECLARE SUB SetInlineThreshold(ByVal n As LONG)
   DECLARE SUB BeginObject()
   DECLARE SUB EndObject()
   DECLARE SUB BeginArray()
   DECLARE SUB EndArray()
   DECLARE SUB Name(ByRef s As WString)
   DECLARE SUB Value(ByRef s As WString)
   DECLARE SUB Value(ByVal n As LongInt)
   DECLARE SUB Value(ByVal n As Double)
   DECLARE SUB ValueVariant(ByRef dv As DVARIANT)
   DECLARE SUB ValueSafeArray(ByRef sa As DSafeArray)
   DECLARE FUNCTION ToBString() AS BSTRING
   DECLARE FUNCTION ToUtf8() AS STRING
   DECLARE SUB Clear()

END TYPE

' ========================================================================================
' Purpose: Determines whether a JSON value can be written “inline” (i.e., without expanding
' across multiple lines) based on its size and structure.
' Logic: Checks if the string length is less than or equal to a predefined inlineThreshold.
' Ensures the string does not contain a { (object start) or [ (array start).
' Returns TRUE if both conditions are met, FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION JsonWriter.IsSmallInline (BYREF tmp AS STRING) AS BOOLEAN
   RETURN LEN(tmp) <= inlineThreshold ANDALSO INSTR(tmp, "{") = 0 ANDALSO INSTR(tmp, "[") = 0
END FUNCTION
' ========================================================================================

PRIVATE SUB JsonWriter.AppendEscaped (ByRef s As WString)
   buf &= !"\""
   For i As Integer = 0 To Len(s) - 1
       Dim ch As UShort = s[i]
       Select Case ch
           Case 34 : buf &= !"\\\""
           Case 92 : buf &= !"\\\\"
           Case 8  : buf &= !"\\b"
           Case 12 : buf &= !"\\f"
           Case 10 : buf &= !"\\n"
           Case 13 : buf &= !"\\r"
           Case 9  : buf &= !"\\t"
           Case Else
               If ch < 32 Then
                   buf &= !"\\u" & Hex(ch, 4)
               Else
                   buf &= WChr(ch)
               End If
       End Select
   Next
   buf &= !"\""
END SUB

PRIVATE SUB JsonWriter.AppendCommaIfNeeded()
    IF UBound(firstItemStack) >= 0 THEN
        DIM idx AS LONG = UBound(firstItemStack)
        IF firstItemStack(idx) THEN
            firstItemStack(idx) = False
        ELSE
            buf &= ","
            ' Pretty-print: newline + indent before the next value
            AppendNewlineAndIndent()
        END IF
    END IF
END SUB

PRIVATE SUB JsonWriter.AppendNewlineAndIndent()
   buf &= !"\n" & WSTRING(depth * indentSize, WSTR(" "))
END SUB

PRIVATE SUB JsonWriter.ValueNull()
   AppendCommaIfNeeded()
   buf &= "null"
END SUB

PRIVATE SUB JsonWriter.ValueBool(ByVal b As Boolean)
   AppendCommaIfNeeded()
   If b Then buf &= "true" Else buf &= "false"
END SUB

' Config
PRIVATE SUB JsonWriter.SetIndentSize(ByVal n As LONG)
   If n >= 0 Then indentSize = n
END SUB

PRIVATE SUB JsonWriter.SetInlineThreshold(ByVal n As LONG)
   If n >= 0 Then inlineThreshold = n
END SUB

' Structure control with inline suppression
PRIVATE SUB JsonWriter.BeginObject()
   AppendCommaIfNeeded()
   buf &= "{"
   depth += 1
   ReDim Preserve firstItemStack(UBound(firstItemStack) + 1)
   firstItemStack(UBound(firstItemStack)) = True
   AppendNewlineAndIndent()
END SUB

PRIVATE SUB JsonWriter.EndObject()
   depth -= 1
   buf &= !"\n" & WString(depth * indentSize, WSTR(" ")) & "}"
   If UBound(firstItemStack) >= 0 Then
       ReDim Preserve firstItemStack(UBound(firstItemStack) - 1)
   End If
End Sub

PRIVATE SUB JsonWriter.BeginArray()
   buf &= "["
   depth += 1
   ReDim Preserve firstItemStack(UBound(firstItemStack) + 1)
   firstItemStack(UBound(firstItemStack)) = True
   AppendNewlineAndIndent()
END SUB

PRIVATE SUB JsonWriter.EndArray()
   depth -= 1
   buf &= !"\n" & WString(depth * indentSize, WSTR(" ")) & "]"
   If UBound(firstItemStack) >= 0 Then
       ReDim Preserve firstItemStack(UBound(firstItemStack) - 1)
   End If
END SUB

' Name inside object
PRIVATE SUB JsonWriter.Name(ByRef s As WString)
   AppendCommaIfNeeded()
   AppendEscaped(s)
   buf &= ": "
   firstItemStack(UBound(firstItemStack)) = True
END SUB

' Primitive values
PRIVATE SUB JsonWriter.Value(ByRef s As WString)
   AppendCommaIfNeeded()
   AppendEscaped(s)
END SUB

PRIVATE SUB JsonWriter.Value(ByVal n As LongInt)
   AppendCommaIfNeeded()
   buf &= WStr(n)
END SUB

PRIVATE SUB JsonWriter.Value(ByVal n As Double)
   AppendCommaIfNeeded()
   buf &= WStr(n)
END SUB

' Variant entry point
PRIVATE SUB JsonWriter.ValueVariant(ByRef dv As DVARIANT)
print dv.vType, VT_BOOL
   SELECT CASE dv.vType()
       CASE VT_EMPTY, VT_NULL
           ValueNull()
       CASE VT_BOOL
           ValueBool(CBool(dv))
       CASE VT_BSTR
           Value(dv.ToBStr)
       CASE VT_R8, VT_R4
           Value(CDbl(dv))
       CASE VT_I1, VT_I2, VT_I4, VT_I8, _
            VT_UI1, VT_UI2, VT_UI4, VT_UI8, _
            VT_INT, VT_UINT
           Value(CLngInt(dv))
       CASE VT_ARRAY Or VT_VARIANT, _
            VT_ARRAY Or VT_BSTR, _
            VT_ARRAY Or VT_R8, _
            VT_ARRAY Or VT_I4
           Dim sa As DSafeArray = *dv
           ValueSafeArray(sa)
       CASE ELSE
           Value(dv.ToStr())
   END SELECT
END SUB

' ========================================================================================
' SAFEARRAY serialization with inline suppression.
' Example:
'   DIM dvArr AS DSafeArray
'   dvArr.Create(VT_VARIANT, 3, 0)
'   dvArr.PutVar(0, DVARIANT("first ??"))
'   dvArr.PutVar(1, DVARIANT(42))
'   dvArr.PutVar(2, DVARIANT(3.14159))
'   DIM jw AS JsonWriter
'   jw.SetIndentSize(2) ' 2-space indent
'   jw.BeginObject()
'      jw.Name("title") : jw.Value("Test")
'      jw.Name("data")  : jw.ValueSafeArray(dvArr)
'   jw.EndObject()
'   PRINT jw.ToBString()
' Example:
'   DIM tiny AS DSafeArray
'   tiny.Create(VT_VARIANT, 3, 0)
'   tiny.PutVar(0, DVARIANT(1))
'   tiny.PutVar(1, DVARIANT(2))
'   tiny.PutVar(2, DVARIANT(3))
'   DIM jw As JsonWriter
'   jw.SetIndentSize(2)
'   jw.SetInlineThreshold(20)  ' tighten threshold
'   jw.BeginObject()
'      jw.Name("a") : jw.ValueSafeArray(tiny)
'      jw.Name("b") : jw.Value("longer text here will break inline")
'  jw.EndObject()
' Output:
' {
'  "a": [
'    1,
'    2,
'    3
'  ],
'  "b": "longer text here will break inline"
'}
' ========================================================================================
PRIVATE SUB JsonWriter.ValueSafeArray(ByRef sa As DSafeArray)
   IF sa.UBound < sa.LBound THEN
       AppendCommaIfNeeded()
       buf &= "[]"
       EXIT SUB
   END IF

    ' First, render into temp
    DIM temp As STRING
    DIM origBuf As BSTRING = buf
    buf.Clear()
    buf &= "["
    FOR i AS LONG = sa.LBound TO sa.UBound
        IF i > sa.LBound THEN buf &= ","  ' local temp, not pretty — compact form
        SELECT CASE sa.GetType
            CASE VT_BSTR
                AppendEscaped(sa.GetStr(i))
            CASE ELSE
                ' Use local rendering without pretty; safe because this is the temp buffer
                ValueVariant(sa.GetVar(i))
        END SELECT
    NEXT
    buf &= "]"
    temp = buf.Utf8

    ' Restore main buffer
    buf = origBuf

    IF IsSmallInline(temp) THEN
        ' Comma is needed here if this array is not the first item in the parent
        AppendCommaIfNeeded()
        buf &= temp
    ELSE
        ' Pretty path: let BeginArray() manage commas and first-item state
        BeginArray()
        FOR i AS LONG = sa.LBound TO sa.UBound
            ' No explicit newline here; AppendCommaIfNeeded will inject ",\n  " as needed
            ValueVariant(sa.GetVar(i))
        NEXT
        EndArray()
    END IF
END SUB


' Output & reset
PRIVATE FUNCTION JsonWriter.ToBString() AS BSTRING
   RETURN buf
END FUNCTION

PRIVATE FUNCTION JsonWriter.ToUtf8() AS STRING
   RETURN buf.Utf8
END FUNCTION

PRIVATE SUB JsonWriter.Clear()
   buf.Clear()
   depth = 0
   REDIM firstItemStack(-1)
END SUB

' // More examples:
' Primitives + string escaping
' DIM esc AS DWSTRING = "Quote: """ & " Newline: " & WCHR(10) & "Tab: " & WCHR(9) & "End"
' DIM jw AS JsonWriter
' jw.SetIndentSize(2)
' jw.BeginObject()
'   jw.Name("s")    : jw.Value(esc)
'   jw.Name("i64")  : jw.Value(9223372036854775807)  ' max signed 64
'   jw.Name("pi")   : jw.Value(3.141592653589793)
'   jw.Name("ok")   : jw.ValueVariant(DVARIANT(TRUE, "BOOLEAN"))   ' VT_BOOL ? true
'   jw.Name("none") : jw.ValueVariant(DVARIANT())                  ' VT_EMPTY ? null
' jw.EndObject()
' print jw.ToUtf8()

' Unicode text (BMP): accents, em dash, Greek
' DIM uni AS DWSTRING = "España — café — ?x"
' DIM jw AS JsonWriter
' jw.SetIndentSize(2)
' jw.BeginObject()
'   jw.Name("title") : jw.Value("Unicode test")
'   jw.Name("text")  : jw.Value(uni)
' jw.EndObject()
' AfxMsg jw.ToBString

'Dim jw As JsonWriter
'jw.SetIndentSize(2)
'jw.BeginObject()
'  jw.Name("app") : jw.Value("AfxNova")
'  jw.Name("version") : jw.Value(1)
'  jw.Name("author")
'  jw.BeginObject()
'    jw.Name("name")  : jw.Value("Jose Roca")
'    jw.Name("site")  : jw.Value("https://github.com/JoseRoca/AfxNova")
'    jw.Name("active"): jw.ValueVariant(DVARIANT(TRUE, "BOOLEAN"))
'  jw.EndObject()
'  jw.Name("notes") : jw.Value("All UTF-16 internally")
'jw.EndObject()
'print jw.ToBString

'DIM jw AS JsonWriter
'jw.SetIndentSize(0) ' 0 = no pretty-print
'jw.BeginObject()
'  jw.Name("k")  : jw.Value("v")
'  jw.Name("n")  : jw.Value(123)
'  jw.Name("b")  : jw.ValueVariant(DVARIANT(FALSE, "BOOLEAN"))
'  jw.Name("nil"): jw.ValueVariant(DVARIANT())
'jw.EndObject()
'print jw.ToBString


END NAMESPACE
