' ########################################################################################
' Platform: Microsoft Windows
' Filename: CRegExp.inc
' Purpose:  VB Script Regular expressons
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
#include once "AfxNova/AfxRegExp.bi"
#include once "AfxNova/DWSTRING.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' CRegExp class
' ========================================================================================
TYPE CRegExp

Private:
   m_Result AS LONG                               ' // Last result code
   m_pRegExp AS Afx_IRegExp2 PTR                  ' // IRegExp2 interface
   m_pMatches AS Afx_IMatchCollection2 PTR        ' // IMatchCollection2 interface

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS LONG
   DECLARE FUNCTION SetResult (BYVAL Result AS LONG) AS LONG
   DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS DWSTRING
   DECLARE FUNCTION SafeRelease (BYREF pv AS ANY PTR) AS ULONG
   DECLARE PROPERTY RegExpPtr () AS Afx_IRegExp2 PTR
   DECLARE PROPERTY Pattern () AS DWSTRING
   DECLARE PROPERTY Pattern (BYREF wszPattern AS CONST WSTRING)
   DECLARE PROPERTY Global () AS BOOLEAN
   DECLARE PROPERTY Global (BYVAL bGlobal AS BOOLEAN)
   DECLARE PROPERTY IgnoreCase () AS BOOLEAN
   DECLARE PROPERTY IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
   DECLARE PROPERTY Multiline () AS BOOLEAN
   DECLARE PROPERTY Multiline (BYVAL bMultiline AS BOOLEAN)
   DECLARE FUNCTION Execute (BYREF wszSourceString AS CONST WSTRING) AS BOOLEAN
   DECLARE FUNCTION MatchesCount () AS LONG
   DECLARE FUNCTION MatchValue (BYVAL index AS LONG = 0) AS DWSTRING
   DECLARE FUNCTION MatchPos (BYVAL index AS LONG = 0) AS LONG
   DECLARE FUNCTION MatchLen (BYVAL index AS LONG = 0) AS LONG
   DECLARE FUNCTION SubMatchesCount (BYVAL index AS LONG = 0) AS LONG
   DECLARE FUNCTION SubMatchValue (BYVAL MatchIndex AS LONG = 0, BYVAL SubMatchIndex AS LONG = 0) AS DWSTRING
   DECLARE FUNCTION Test (BYREF wszSourceString AS CONST WSTRING) AS BOOLEAN
   DECLARE FUNCTION Find (BYREF wszSourceString AS CONST WSTRING) AS LONG
   DECLARE FUNCTION Find (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING) AS LONG
   DECLARE FUNCTION FindEx (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   DECLARE FUNCTION Replace (BYREF wszSourceString AS CONST WSTRING, BYREF wszReplaceString AS CONST WSTRING) AS DWSTRING
   DECLARE FUNCTION Remove (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   DECLARE FUNCTION Extract (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   DECLARE FUNCTION Extract (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING) AS DWSTRING

   ' // Helper function
   DECLARE FUNCTION VariantToString (BYVAL pvarIn AS VARIANT PTR) AS DWSTRING

END TYPE
' ========================================================================================

' ========================================================================================
' * CRegExp default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CRegExp
   ' // Initialize the COM library
   CoInitialize(NULL)
   ' // Create an instance of the IRegExp2 interface
   DIM classID AS CLSID, riid AS IID
   DIM wszProgId AS WSTRING * 260 = "VBScript.RegExp"
   CLSIDFromProgID(@wszProgID, @ClassID)
   DIM wszriid AS WSTRING * 260 = AFX_IID_IRegExp2
   CLSIDFromString(@wszriid, @riid)
   DIM hr AS HRESULT = CoCreateInstance(@ClassID, NULL, CLSCTX_SERVER, @riid, @m_pRegExp)
   IF m_pRegExp = NULL THEN
      CoUninitialize
      EXIT CONSTRUCTOR
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' CRegExp destructor
' ========================================================================================
PRIVATE DESTRUCTOR CRegExp
   ' // Release the interfaces
   IF m_pMatches THEN m_pMatches->Release
   IF m_pRegExp THEN m_pRegExp->Release
   ' // Uninitialize the COM library
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code
' ========================================================================================
PRIVATE FUNCTION CRegExp.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CRegExp.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CRegExp.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" + HEX(nError) & CHR(13, 10) + dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the passed interface pointer and sets it to null.
' ========================================================================================
PRIVATE FUNCTION CRegExp.SafeRelease (BYREF pv AS ANY PTR) AS ULONG
   IF pv = NULL THEN RETURN 0
   FUNCTION = cast(IUnknown PTR, pv)->lpvtbl->Release(pv)
   pv = NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the Afx_IRegExp2 interface
' ========================================================================================
PRIVATE PROPERTY CRegExp.RegExpPtr () AS Afx_IRegExp2 PTR
   PROPERTY = m_pRegExp
END PROPERTY
' ========================================================================================

' ========================================================================================
' * The regular expression pattern to use.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Pattern () AS DWSTRING
   DIM bstrPattern AS BSTR
   m_Result = m_pRegExp->get_Pattern(@bstrPattern)
   DIM dwsPattern AS DWSTRING = bstrPattern
   SysFreeString(bstrPattern)
   PROPERTY = dwsPattern
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Pattern (BYREF wszPattern AS CONST WSTRING)
   DIM bstrPattern AS BSTR = SysAllocString(@wszPattern)
   m_Result = m_pRegExp->put_Pattern(bstrPattern)
   SysFreeString(bstrPattern)
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Global seach for all occurrences of the pattern.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Global () AS BOOLEAN
   DIM bGlobal AS VARIANT_BOOL
   m_Result = m_pRegExp->get_Global(@bGlobal)
   PROPERTY = bGlobal
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Global (BYVAL bGlobal AS BOOLEAN)
   m_Result = m_pRegExp->put_Global(cast(VARIANT_BOOL, bGlobal))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Ignore case.
' ========================================================================================
PRIVATE PROPERTY CRegExp.IgnoreCase () AS BOOLEAN
   DIM bIgnoreCase AS VARIANT_BOOL
   m_Result = m_pRegExp->get_IgnoreCase(@bIgnoreCase)
   PROPERTY = bIgnoreCase
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
   m_Result = m_pRegExp->put_IgnoreCase(cast(VARIANT_BOOL, bIgnoreCase))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Multiline search.
' ========================================================================================
PRIVATE PROPERTY CRegExp.Multiline () AS BOOLEAN
   DIM bMultiline AS VARIANT_BOOL
   m_Result = m_pRegExp->get_Multiline(@bMultiline)
   PROPERTY = bMultiline
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CRegExp.Multiline (BYVAL bMultiline AS BOOLEAN)
   m_Result = m_pRegExp->put_Multiline(cast(VARIANT_BOOL, bMultiline))
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Executes a regular expression search against a specified string.
' Example:
' #include once "Afx2/CRegExp2.inc"
' USING Afx2
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "is."
' pRegExp.IgnoreCase = TRUE
' pRegExp.Global = TRUE
' IF pRegExp.Execute("IS1 is2 IS3 is4") = FALSE THEN
'    PRINT "No match found"
' ELSE
'    DIM nCount AS LONG = pRegExp.MatchesCount
'    FOR i AS LONG = 0 TO nCount - 1
'       PRINT "Value: ", pRegExp.MatchValue(i)
'       PRINT "Position: ", pRegExp.MatchPos(i)
'       PRINT "Length: ", pRegExp.MatchLen(i)
'       PRINT
'    NEXT
' END IF
' ========================================================================================
PRIVATE FUNCTION CRegExp.Execute (BYREF wszSourceString AS CONST WSTRING) AS BOOLEAN
   IF m_pMatches THEN this.SafeRelease(m_pMatches)
   DIM bstrSrc AS BSTR = SysAllocString(@wszSourceString)
   m_Result = m_pRegExp->Execute(bstrSrc, cast(Afx_IDispatch PTR PTR, @m_pMatches))
   SysFreeString(bstrSrc)
   IF m_Result = S_OK THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of matches.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchesCount () AS LONG
   IF m_pMatches = NULL THEN m_Result = E_POINTER: RETURN 0
   DIM nCount AS LONG
   m_Result = m_pMatches->get_Count(@nCount)
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the value or text of a match found in a search string.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchValue (BYVAL index AS LONG) AS DWSTRING
   IF m_pMatches = NULL THEN m_Result = E_POINTER: RETURN ""
   DIM bstrValue AS BSTR
   DIM pMatch AS Afx_IMatch2 PTR
   m_Result = m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch))
   IF pMatch THEN
      pMatch->get_Value(@bstrValue)
      this.SafeRelease(pMatch)
   END IF
   DIM dwsValue AS DWSTRING = bstrValue
   SysFreeString(bstrValue)
   RETURN dwsValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the position in a search string where a match occurs.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchPos (BYVAL index AS LONG) AS LONG
   IF m_pMatches = NULL THEN m_Result = E_POINTER: RETURN 0
   DIM nFirstIndex AS LONG
   DIM pMatch AS Afx_IMatch2 PTR
   m_Result = m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch))
   IF pMatch THEN
      ' // FirstIndex property indicates the number of characters in the string to
      ' // the left of the match. If the match was found at the very start of the
      ' // string, FirstIndex will be zero. If the match starts at the second character
      ' // in the string, FirstIndex will be one, etc. Therefore, we increase the value by one.
      IF pMatch->get_FirstIndex(@nFirstIndex) = S_OK THEN nFirstIndex += 1
      this.SafeRelease(pMatch)
   END IF
   RETURN nFirstIndex
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the length of a match found in a search string.
' ========================================================================================
PRIVATE FUNCTION CRegExp.MatchLen (BYVAL index AS LONG) AS LONG
   IF m_pMatches = NULL THEN m_Result = E_POINTER: RETURN 0
   DIM nLen AS LONG
   DIM pMatch AS Afx_IMatch2 PTR
   m_Result = m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch))
   IF pMatch THEN
      pMatch->get_Length(@nLen)
      this.SafeRelease(pMatch)
   END IF
   FUNCTION = nLen
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of submatches.
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "(\w+)@(\w+)\.(\w+)"
' pRegExp.IgnoreCase = TRUE
' DIM dwsText AS DWSTRING = "Please send mail to dragon@xyzzy.com. Thanks!"
' IF pRegExp.Execute(dwsText) = FALSE THEN
'    PRINT "No match found"
' ELSE
'    ' // Get the number of submatches
'    DIM nCount AS LONG = pRegExp.SubMatchesCount(0)
'    PRINT "Submatches: ", nCount
'    FOR i AS LONG = 0 TO nCount - 1
'       PRINT pRegExp.SubMatchValue(0, i)
'    NEXT
' END IF
' Output:
' dragon
' xyzzy
' com
' ========================================================================================
PRIVATE FUNCTION CRegExp.SubMatchesCount (BYVAL index AS LONG) AS LONG
   IF m_pMatches = NULL THEN m_Result = E_POINTER: RETURN 0
   DIM nCount AS LONG
   DIM pMatch AS Afx_IMatch2 PTR
   m_Result = m_pMatches->get_Item(index, cast(Afx_IDispatch PTR PTR, @pMatch))
   IF pMatch THEN
      DIM pSubMatches AS Afx_ISubMatches PTR
      m_Result = pMatch->get_SubMatches(cast(Afx_IDispatch PTR PTR, @pSubMatches))
      this.SafeRelease(pMatch)
      IF pSubMatches THEN
         m_Result = pSubMatches->get_Count(@nCount)
         this.SafeRelease(pSubMatches)
      END IF
   END IF
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value or text of a submatch found in a search string.
' ========================================================================================
PRIVATE FUNCTION CRegExp.SubMatchValue (BYVAL MatchIndex AS LONG, BYVAL SubMatchIndex AS LONG) AS DWSTRING
   DIM vValue AS VARIANT
   IF m_pMatches = NULL THEN m_Result = E_POINTER: RETURN ""
   DIM pMatch AS Afx_IMatch2 PTR
   m_Result = m_pMatches->get_Item(MatchIndex, cast(Afx_IDispatch PTR PTR, @pMatch))
   IF pMatch THEN
      DIM pSubMatches AS Afx_ISubMatches PTR
      m_Result = pMatch->get_SubMatches(cast(Afx_IDispatch PTR PTR, @pSubMatches))
      this.SafeRelease(pMatch)
      IF pSubMatches THEN
         m_Result = pSubMatches->get_Item(SubMatchIndex, @vValue)
         this.SafeRelease(pSubMatches)
      END IF
   END IF
   RETURN this.VariantToString(@vValue)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Executes a regular expression search against a specified string and returns a Boolean
' value that indicates if a pattern match was found.
' - wszSourceString : The source string.
' Note: The Global property has no effect on the Test method.
' Example:
' DIM pRegExp AS CRegExp
' pRegExp.Pattern = "is."
' pRegExp.IgnoreCase = TRUE
' IF pRegExp.Test("IS1 is2 IS3 is4") = FALSE THEN
'   PRINT "No match found"
' ELSE
'    PRINT "Match found"
' END IF
' ========================================================================================
PRIVATE FUNCTION CRegExp.Test (BYREF wszSourceString AS CONST WSTRING) AS BOOLEAN
   DIM bMatch AS VARIANT_BOOL
   DIM bstrSrc AS BSTR = SysAllocString(@wszSourceString)
   m_Result = m_pRegExp->Test(bstrSrc, @bMatch)
   SysFreeString(bstrSrc)
   RETURN bMatch
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Find function with VBScript regular expressions search patterns.
' Parameter:
' - wszSourceString = The text to be parsed.
' Return value:
' - Returns the position of the match or 0 if not found.
'   The length of the match can be retrieved calling pRegExp.MatchLen.
' Usage example:
'   DIM pRegExp AS CRegExp
'   pRegExp.Global = TRUE
'   pRegExp.IgnoreCase = TRUE
'   DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM nPos AS LONG = pRegExp.Find(bsText)
'   PRINT nPos
' ' Output: 11
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYREF wszSourceString AS CONST WSTRING) AS LONG
   IF m_pMatches THEN this.SafeRelease(m_pMatches)
   DIM bstrSrc AS BSTR = SysAllocString(@wszSourceString)
   m_Result = m_pRegExp->Execute(bstrSrc, cast(Afx_IDispatch PTR PTR, @m_pMatches))
   SysFreeString(bstrSrc)
   RETURN this.MatchPos(0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Find function with VBScript regular expressions search patterns.
' Parameters:
' - nStart: The position in the string at which the search will begin. The first character
'   starts at position 1.
' - wszSourceString = The text to be parsed.
' Return value:
' - Returns the position of the match or 0 if not found.
'   The length of the match can be retrieved calling pRegExp.MatchLen.
' Usage example:
'   DIM pRegExp AS CRegExp
'   pRegExp.Global = TRUE
'   pRegExp.IgnoreCase = TRUE
'   DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM nPos AS LONG = pRegExp.Find(15, dwsText)
'   PRINT nPos
' Output: 26
' ========================================================================================
PRIVATE FUNCTION CRegExp.Find (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING) AS LONG
   IF m_pMatches THEN this.SafeRelease(m_pMatches)
   DIM bstrSrc AS BSTR = SysAllocString(wszSourceString)
   DIM bs AS BSTRING = MID(wszSourceString, nStart)
   m_Result = m_pRegExp->Execute(bs, cast(Afx_IDispatch PTR PTR, @m_pMatches))
   SysFreeString(bstrSrc)
   RETURN this.MatchPos(0) + nStart - 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Global, multiline find function with VBScript regular expressions search patterns.
' Parameters:
' - wszSourceString = The text to be parsed.
' Return value:
' - Returns a list of comma separated "index, length" value pairs. The pairs are separated
'   by a semicolon.
' Usage Example:
'   DIM pRegExp AS CRegExp
'   DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
'   pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'   DIM dwsOut AS DWSTRING = pRegExp.FindEx(dwsText)
'   PRINT dwsOut
' Output: 11,4;26,4
' ========================================================================================
PRIVATE FUNCTION CRegExp.FindEx (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   DIM dwsOut AS DWSTRING, pMatches AS Afx_IMatchCollection2 PTR
   m_pRegExp->put_Global(TRUE)
   m_pRegExp->put_Multiline(TRUE)
   DIM bstrSrc AS BSTR = SysAllocString(@wszSourceString)
   m_Result = m_pRegExp->Execute(bstrSrc, cast(Afx_IDispatch PTR PTR, @pMatches))
   SysFreeString(bstrSRc)
   IF pMatches THEN
      DIM nCount AS LONG
      pMatches->get_Count(@nCount)
      FOR i AS LONG = 0 TO nCount - 1
         DIM pMatch AS Afx_IMatch2 PTR
         m_Result = pMatches->get_Item(i, cast(Afx_IDispatch PTR PTR, @pMatch))
         IF pMatch THEN
            DIM nFirstIndex AS LONG
            pMatch->get_FirstIndex(@nFirstIndex)
            DIM nLen AS LONG
            pMatch->get_Length(@nLen)
            IF i < nCount - 1 THEN
               dwsOut += STR(nFirstIndex + 1) & "," & STR(nLen) & ";"
            ELSE
               dwsOut += STR(nFirstIndex + 1) & "," & STR(nLen)
            END IF
            this.SafeRelease(pMatch)
         END IF
      NEXT
      this.SafeRelease(pMatches)
   END IF
   RETURN dwsOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with text removed using a regular expression as the search string.
' Parameters:
' - wszSourceString : The source string.
' Examples:
'    DIM pRegExp AS CRegExp
'    pRegExp.Pattern = "ab"
'    pRegExp.Global = TRUE
'    PRINT pRegExp.Remove("abacadabra")
' Output: "acadra"
' ----------------------------------------------------------------------------------------
'    DIM pRegExp AS CRegExp
'    pRegExp.Pattern = "[bAc]"
'    pRegExp.IgnoreCase = TRUE
'    pRegExp.Global = TRUE
'    PRINT pRegExp.Remove("abacadabra")
' Output: "dr"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Remove (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   DIM vd AS VARIANT
   DIM bstrSrc AS BSTR = SysAllocString(@wszSourceString)
   DIM bstrDest AS BSTR
   m_Result = m_pRegExp->Replace(bstrSrc, vd, @bstrDest)
   SysFreeString bstrSrc
   DIM dwsOut AS DWSTRING
   IF bstrDest THEN
      dwsOut = bstrDest
      SysFreeString(bstrDest)
   ElSE
      dwsOut = wszSourceString
   END IF
   RETURN dwsOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with text replaced using a regular expression as the search string.
' Parameters:
' - wszSourceString : The source string.
' - wszReplaceString : The replacement string.
' Examples:
'    DIM pRegExp AS CRegExp
'    pRegExp.Global = TRUE
'    pRegExp.IgnoreCase = TRUE
'    pRegExp.Pattern = "World"
'    PRINT pRegExp.Replace("Hello World", "Earth") 
 'Output: "Hello Earth"
 ' ----------------------------------------------------------------------------------------
'    DIM pRegExp AS CRegExp
'    pRegExp.Global = TRUE
'    pRegExp.Pattern = "[bac]"
'    --or-- pRegExp.Pattern = "b|a|c"
'    PRINT pRegExp.Replace("abacadabra", "*")
' Output: "*****d**r*"
' ----------------------------------------------------------------------------------------
'    DIM pRegExp AS CRegExp
'    pRegExp.Global = TRUE
'    pRegExp.Pattern = $"(\d{3})-(\d{3})-(\d{4})"
'    PRINT pRegExp.Replace("555-123-4567", "($1) $2-$3")
' Output: "(555) 123-4567"
' ----------------------------------------------------------------------------------------
'    DIM pRegExp AS CRegExp
'    pRegExp.Global = TRUE
'    pRegExp.Pattern = $"(\S+), (\S+)"
'    PRINT pRegExp.Replace("Roca, José", "$2 $1")
' Output: "José Roca"
' ----------------------------------------------------------------------------------------
'    DIM pRegExp AS CRegExp
'    pRegExp.Global = TRUE
'    pRegExp.Pattern = $"\b0{1,}\."
'    PRINT pRegExp.Replace("0000.34500044", ".")
' Output: ".34500044"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Replace (BYREF wszSourceString AS CONST WSTRING, BYREF wszReplaceString AS CONST WSTRING) AS DWSTRING
   DIM bstrSrc AS BSTR = SysAllocString(wszSourceString)
   DIM vd AS VARIANT : vd.vt = VT_BSTR : vd.bstrVal = SysAllocString(@wszReplaceString)
   DIM bstrDest AS BSTR
   m_Result = m_pRegExp->Replace(bstrSrc, vd, @bstrDest)
   SysFreeString(bstrSrc)
   VariantClear(@vd)
   DIM dwsOut AS DWSTRING
   IF bstrDest THEN
      dwsOut = bstrDest
      SysFreeString(bstrDest)
   ElSE
      dwsOut = wszSourceString
   END IF
   RETURN dwsOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Extracts a substring using VBScript regular expressions search patterns.
' Parameters:
' - wszSourceString = The text to be parsed.
' Return value:
' - Returns the extracted string on exit or an empty string on failure.
' Usage example:
'    DIM pRegExp AS CRegExp
'    DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
'    pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'    dwsText = pRegExp.Extract(dwsText)
'    PRINT dwsText
' Output: "a234"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   IF m_pMatches THEN this.SafeRelease(m_pMatches)
   DIM bstrSrc AS BSTR = SysAllocString(@wszSourceString)
   m_Result = m_pRegExp->Execute(bstrSrc, cast(Afx_IDispatch PTR PTR, @m_pMatches))
   SysFreeString(bstrSrc)
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   RETURN MID(wszSourceString, nPos, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Extracts a substring using VBScript regular expressions search patterns.
' Parameters:
' - nStart: The position in the string at which the search will begin. The first character
'   starts at position 1.
' - wszSourceString = The text to be parsed.
' Return value:
' - Returns the extracted string on exit or an empty string on failure.
' Usage example:
'    DIM pRegExp AS CRegExp
'    DIM dwsText AS DWSTRING = "blah blah a234 blah blah x345 blah blah"
'    pRegExp.Pattern = "[a-z][0-9][0-9][0-9]"
'    dwsText = pRegExp.Extract(15, dwsText)
'    PRINT dwsText
' Output: "x345"
' ========================================================================================
PRIVATE FUNCTION CRegExp.Extract (BYVAL nStart AS LONG, BYREF wszSourceString AS CONST WSTRING) AS DWSTRING
   IF m_pMatches THEN this.SafeRelease(m_pMatches)
   DIM bsSrc AS BSTRING = MID(wszSourceString, nStart)
   m_Result = m_pRegExp->Execute(bsSrc, cast(Afx_IDispatch PTR PTR, @m_pMatches))
   DIM nPos AS LONG = this.MatchPos(0)
   DIM nLen AS LONG = this.MatchLen(0)
   RETURN MID(bsSrc, nPos, nLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the variant value of a variant structure to a string.
' Parameters:
' - pvarIn   : [in] Reference to a source variant structure.
' Return value:
' If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
' Usage: pRegExp.VariantToString(@variant_name)
' ========================================================================================
PRIVATE FUNCTION CRegExp.VariantToString (BYVAL pvarIn AS VARIANT PTR) AS DWSTRING
   DIM dws AS DWSTRING
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN RETURN ""
   DIM pVariantToStringAlloc AS FUNCTION (BYVAL pVar AS VARIANT PTR, BYVAL ppwszBuf AS WSTRING PTR PTR) AS HRESULT
   pVariantToStringAlloc = DyLibSymbol(pLib, "VariantToStringAlloc")
   IF pVariantToStringAlloc = NULL THEN
      m_Result = E_FAIL
   ELSE
      DIM pwszBuf AS WSTRING PTR
      m_Result = pVariantToStringAlloc(pvarIn, @pwszBuf)
      IF pwszBuf THEN
         dws = *pwszBuf
         CoTaskMemFree pwszBuf
      END IF
   END IF
   DyLibFree(pLib)
   RETURN dws
END FUNCTION
' ========================================================================================

END NAMESPACE
