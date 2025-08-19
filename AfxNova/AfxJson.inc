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

   DECLARE SUB SkipWhitespace
   DECLARE FUNCTION ReadString () AS DWSTRING
   DECLARE FUNCTION ReadNumber () AS DWSTRING

public:
   DECLARE CONSTRUCTOR (BYREF source AS WSTRING)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION ReadNext (BYREF tok AS JsonToken) AS BOOLEAN

END TYPE

' ========================================================================================
' Creates an instance of the JSonReader class
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
' Advances the position past spaces, tabs, CR and LF.
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
' Read string by slicing the raw JSON string and unquoting via JSonUnquoteW
' ========================================================================================
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
' Parses a JSON number
' ========================================================================================
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
' Reads the next token
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


' ########################################################################################
' Class JSonWriter
' ########################################################################################
TYPE JsonWriter
 
Private:
   m_buf AS BSTRING
   m_firstItemStack(ANY) AS BOOLEAN
   m_indentSize AS LONG = 4
   m_depth AS LONG
   m_inlineThreshold As Integer = 40 ' max chars for inline array/object

   DECLARE FUNCTION IsSmallInline (BYREF tmp AS STRING) AS BOOLEAN
   DECLARE SUB AppendEscaped (BYREF s AS WSTRING)
   DECLARE SUB AppendCommaIfNeeded
   DECLARE SUB AppendNewlineAndIndent
   DECLARE SUB ValueNull
   DECLARE SUB ValueBool(BYVAL b AS BOOLEAN)

Public:

   DECLARE SUB SetIndentSize(BYVAL n As LONG)
   DECLARE SUB SetInlineThreshold(BYVAL n As LONG)
   DECLARE SUB BeginObject
   DECLARE SUB EndObject
   DECLARE SUB BeginArray
   DECLARE SUB EndArray
   DECLARE SUB Name(BYREF s AS WSTRING)
   DECLARE SUB Value(BYREF s AS WSTRING)
   DECLARE SUB Value(BYVAL n AS LONGINT)
   DECLARE SUB Value(BYVAL n AS DOUBLE)
   DECLARE SUB ValueVariant(BYREF dv AS DVARIANT)
   DECLARE SUB ValueSafeArray(BYREF sa AS DSafeArray)
   DECLARE FUNCTION ToBString () AS BSTRING
   DECLARE FUNCTION ToUtf8 () AS STRING
   DECLARE SUB Clear

END TYPE

' ========================================================================================
' Purpose: Determines whether a JSON value can be written “inline” (i.e., without expanding
' across multiple lines) based on its size and structure.
' Logic: Checks if the string length is less than or equal to a predefined inlineThreshold.
' Ensures the string does not contain a { (object start) or [ (array start).
' Returns TRUE if both conditions are met, FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION JsonWriter.IsSmallInline (BYREF tmp AS STRING) AS BOOLEAN
   RETURN LEN(tmp) <= m_inlineThreshold ANDALSO INSTR(tmp, "{") = 0 ANDALSO INSTR(tmp, "[") = 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a quoted, JSON-escaped form of s into buf, encoding control chars < 32 as \uXXXX.
' ========================================================================================
PRIVATE SUB JsonWriter.AppendEscaped (BYREF s AS WSTRING)
   m_buf &= !"\""
   FOR i AS LONG = 0 TO Len(s) - 1
      DIM ch AS USHORT = s[i]
      SELECT CASE ch
         CASE 34 : m_buf &= !"\\\""
         CASE 92 : m_buf &= !"\\\\"
         CASE 8  : m_buf &= !"\\b"
         CASE 12 : m_buf &= !"\\f"
         CASE 10 : m_buf &= !"\\n"
         CASE 13 : m_buf &= !"\\r"
         CASE 9  : m_buf &= !"\\t"
         CASE ELSE
            IF ch < 32 THEN
               m_buf &= !"\\u" & HEX(ch, 4)
            ELSE
               m_buf &= WCHR(ch)
            END IF
      END SELECT
   NEXT
   m_buf &= !"\""
END SUB
' ========================================================================================

' ========================================================================================
' Manages comma insertion between sibling elements based on firstItemStack; injects newline+indent for pretty-print paths.
' ========================================================================================
PRIVATE SUB JsonWriter.AppendCommaIfNeeded
   IF UBOUND(m_firstItemStack) >= 0 THEN
      DIM idx AS LONG = UBOUND(m_firstItemStack)
      IF m_firstItemStack(idx) THEN
         m_firstItemStack(idx) = FALSE
      ELSE
         m_buf &= ","
         ' Pretty-print: newline + indent before the next value
         AppendNewlineAndIndent
      END IF
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Adds line break and depth×indentSize spaces.
' ========================================================================================
PRIVATE SUB JsonWriter.AppendNewlineAndIndent
   m_buf &= !"\n" & WSTRING(m_depth * m_indentSize, WSTR(" "))
END SUB
' ========================================================================================

' ========================================================================================
' Emits literal null, respecting comma/indent rules.
' ========================================================================================
PRIVATE SUB JsonWriter.ValueNull
   AppendCommaIfNeeded
   m_buf &= "null"
END SUB
' ========================================================================================

' ========================================================================================
' Emits true or false.
' ========================================================================================
PRIVATE SUB JsonWriter.ValueBool(BYVAL b AS BOOLEAN)
   AppendCommaIfNeeded
   If b THEN m_buf &= "true" ELSE m_buf &= "false"
END SUB
' ========================================================================================

' ========================================================================================
' Controls spaces per indent level (0 disables pretty-printing newlines/indents for structural entries).
' ========================================================================================
PRIVATE SUB JsonWriter.SetIndentSize(BYVAL n As LONG)
   If n >= 0 Then m_indentSize = n
END SUB
' ========================================================================================

' ========================================================================================
' Sets the maximum character length used to decide if arrays/objects may be emitted inline.
' ========================================================================================
PRIVATE SUB JsonWriter.SetInlineThreshold(BYVAL n As LONG)
   If n >= 0 THEN m_inlineThreshold = n
END SUB
' ========================================================================================

' ========================================================================================
' Structures control with inline suppression
' ========================================================================================
PRIVATE SUB JsonWriter.BeginObject
   AppendCommaIfNeeded
   m_buf &= "{"
   m_depth += 1
   REDIM PRESERVE m_firstItemStack(UBOUND(m_firstItemStack) + 1)
   m_firstItemStack(UBOUND(m_firstItemStack)) = TRUE
   AppendNewlineAndIndent
END SUB
' ========================================================================================

' ========================================================================================
' Emits “{”, increase depth, push a “first item” flag, and insert newline+indent.
' ========================================================================================
PRIVATE SUB JsonWriter.EndObject
   m_depth -= 1
   m_buf &= !"\n" & WSTRING(m_depth * m_indentSize, WSTR(" ")) & "}"
   IF UBOUND(m_firstItemStack) >= 0 THEN
      REDIM PRESERVE m_firstItemStack(UBOUND(m_firstItemStack) - 1)
   END IF
END SUB
' ========================================================================================

' ========================================================================================
'Opens the array with "[".
' ========================================================================================
PRIVATE SUB JsonWriter.BeginArray
   m_buf &= "["
   m_depth += 1
   REDIM PRESERVE m_firstItemStack(UBOUND(m_firstItemStack) + 1)
   m_firstItemStack(UBOUND(m_firstItemStack)) = TRUE
   AppendNewlineAndIndent
END SUB
' ========================================================================================

' ========================================================================================
' Closes the array with proper dedent and “]”, then pop the stack.
' ========================================================================================
PRIVATE SUB JsonWriter.EndArray
   m_depth -= 1
   m_buf &= !"\n" & WSTRING(m_depth * m_indentSize, WSTR(" ")) & "]"
   IF UBOUND(m_firstItemStack) >= 0 THEN
       REDIM PRESERVE m_firstItemStack(UBOUND(m_firstItemStack) - 1)
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Writes a JSON object key with correct escaping, colon, and spacing.
' ========================================================================================
PRIVATE SUB JsonWriter.Name (BYREF s AS WSTRING)
   AppendCommaIfNeeded
   AppendEscaped(s)
   m_buf &= ": "
   m_firstItemStack(UBOUND(m_firstItemStack)) = TRUE
END SUB
' ========================================================================================

' ========================================================================================
' Emits a JSON string with escaping for quotes, backslashes, control chars, and nonprintables via \uXXXX.
' ========================================================================================
PRIVATE SUB JsonWriter.Value(BYREF s AS WSTRING)
   AppendCommaIfNeeded
   AppendEscaped(s)
END SUB
' ========================================================================================

' ========================================================================================
' Emits an integer number using WSTR conversion (no quotes).
' ========================================================================================
PRIVATE SUB JsonWriter.Value (BYVAL n AS LONGINT)
   AppendCommaIfNeeded
   m_buf &= WSTR(n)
END SUB
' ========================================================================================

' ========================================================================================
' Emits a floating-point number using WSTR conversion (no quotes).
' ========================================================================================
PRIVATE SUB JsonWriter.Value (BYVAL n AS DOUBLE)
   AppendCommaIfNeeded
   m_buf &= WSTR(n)
END SUB
' ========================================================================================

' ========================================================================================
' Serializes a DVARIANT using JSON-compatible mapping.
' ========================================================================================
PRIVATE SUB JsonWriter.ValueVariant (BYREF dv AS DVARIANT)
   SELECT CASE dv.vType
      CASE VT_EMPTY, VT_NULL
         ValueNull
      CASE VT_BOOL
         ValueBool(CBOOL(dv))
      CASE VT_BSTR
         Value(dv.ToBStr)
      CASE VT_R8, VT_R4
         Value(CDBL(dv))
      CASE VT_I1, VT_I2, VT_I4, VT_I8, _
          VT_UI1, VT_UI2, VT_UI4, VT_UI8, _
          VT_INT, VT_UINT
         Value(CLNGINT(dv))
      CASE VT_ARRAY Or VT_VARIANT, _
          VT_ARRAY Or VT_BSTR, _
          VT_ARRAY Or VT_R8, _
          VT_ARRAY Or VT_I4
         DIM sa AS DSafeArray = *dv
         ValueSafeArray(sa)
      CASE ELSE
         Value(dv.ToStr)
   END SELECT
END SUB
' ========================================================================================

' ========================================================================================
' Serializes a 1-D SAFEARRAY as JSON array with smart inline vs pretty layout.
' ========================================================================================
PRIVATE SUB JsonWriter.ValueSafeArray (BYREF sa AS DSafeArray)
   IF sa.UBound < sa.LBound THEN
      AppendCommaIfNeeded
      m_buf &= "[]"
      EXIT SUB
   END IF
   ' First, render into temp
   DIM temp AS STRING
   DIM origBuf AS BSTRING = m_buf
   m_buf.Clear
   m_buf &= "["
   FOR i AS LONG = sa.LBound TO sa.UBound
      IF i > sa.LBound THEN m_buf &= ","  ' local temp, not pretty — compact form
      SELECT CASE sa.GetType
         CASE VT_BSTR
            AppendEscaped(sa.GetStr(i))
         CASE ELSE
            ' Use local rendering without pretty; safe because this is the temp buffer
            ValueVariant(sa.GetVar(i))
      END SELECT
   NEXT
   m_buf &= "]"
   temp = m_buf.Utf8

   ' Restore main buffer
   m_buf = origBuf

   IF IsSmallInline(temp) THEN
      ' Comma is needed here if this array is not the first item in the parent
      AppendCommaIfNeeded
      m_buf &= temp
   ELSE
      ' Pretty path: let BeginArray manage commas and first-item state
      BeginArray
      FOR i AS LONG = sa.LBound TO sa.UBound
         ' No explicit newline here; AppendCommaIfNeeded will inject ",\n  " as needed
         ValueVariant(sa.GetVar(i))
      NEXT
      EndArray
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Returns the current buffer as BSTRING (UTF-16).
' ========================================================================================
PRIVATE FUNCTION JsonWriter.ToBString () AS BSTRING
   RETURN m_buf
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current buffer as utf-8.
' ========================================================================================
PRIVATE FUNCTION JsonWriter.ToUtf8 () AS STRING
   RETURN m_buf.Utf8
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the buffer, depth, and firstItemStack to an empty state.
' ========================================================================================
PRIVATE SUB JsonWriter.Clear
   m_buf.Clear
   m_depth = 0
   REDIM m_firstItemStack(-1)
END SUB
' ========================================================================================

END NAMESPACE
