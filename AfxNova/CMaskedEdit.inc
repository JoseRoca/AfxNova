' ########################################################################################
' Platform: Microsoft Windows
' File: CMaskedEdit.inc
' Contents: Masked edit control
' Based on CMFCMaskedEdit (Microsoft Foundation Classes)
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
#include once "/crt/ctype.bi"
#include once "AfxNova/CWindow.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' CMaskedEdit class
' ========================================================================================
TYPE CMaskedEdit

   Public:
      m_hCtl  AS HWND                   ' // Control handle

   Private:
      m_strMask AS DWSTRING             ' // The mask string
      m_strInputTemplate AS DWSTRING    ' // "_" char = character entry
      m_str AS DWSTRING                 ' // Initial value
      m_chMaskInputTemplate AS USHORT   ' // Default char
      m_strValid AS DWSTRING            ' // Valid string characters
      m_bSelectByGroup AS BOOLEAN
      m_strUndo AS DWSTRING             ' // Backup string for undo

   Public:
      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1)
      DECLARE FUNCTION Create (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1) AS HWND
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hWindow () AS HWND
      DECLARE SUB GetGroupBounds(BYREF nBegin AS LONG, BYREF nEnd AS LONG, BYVAL nStartPos AS LONG = 0, BYVAL bForward AS BOOLEAN = TRUE)
      DECLARE SUB EnableMask (BYVAL lpszMask AS WSTRING PTR, BYVAL lpszInputTemplate AS WSTRING PTR, _
              BYREF chMaskInputTemplate AS DWSTRING = "_", BYVAL lpszValid AS WSTRING PTR = NULL)
      DECLARE SUB SetValidChars (BYVAL lpszValid AS WSTRING PTR)
      DECLARE FUNCTION IsMaskedChar (BYVAL chChar AS USHORT, BYVAL chMaskChar AS USHORT) AS BOOLEAN
      DECLARE FUNCTION CheckChar (BYVAL chChar AS USHORT, BYVAL nPos AS LONG) AS BOOLEAN
      DECLARE FUNCTION SetValue (BYVAL lpszString AS WSTRING PTR, BYVAL bWithDelimiters AS BOOLEAN = TRUE) AS BOOLEAN
      DECLARE FUNCTION DoUpdate (BYVAL bRestoreLastGood AS BOOLEAN, BYVAL nBeginOld AS LONG = -1, BYVAL nEndOld AS LONG = -1) AS BOOLEAN
      DECLARE SUB OnCharPrintchar(BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)
      DECLARE SUB OnKeyDown (BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)
      DECLARE SUB OnCharBackspace (BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)
      DECLARE SUB OnCharDelete (BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)
      DECLARE FUNCTION SetText (BYREF dwsText AS DWSTRING) AS BOOLEAN
      DEClARE FUNCTION SetText (BYREF dwsText AS DWSTRING, BYVAL bSetMaskedCharsOnly AS BOOLEAN) AS BOOLEAN
      DECLARE FUNCTION GetText (BYVAL bGetMaskedCharsOnly AS BOOLEAN = FALSE) AS DWSTRING
      DECLARE FUNCTION GetTextLength (BYVAL bGetMaskedCharsOnly AS BOOLEAN = FALSE) AS LONG
      DECLARE SUB SetPos (BYVAL nPos AS LONG = -1)

      DECLARE STATIC FUNCTION CMaskedEditProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM, BYVAL uIdSubclass AS UINT_PTR, BYVAL dwRefData AS DWORD_PTR) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' Processes messages for the subclassed control.
' ========================================================================================
FUNCTION CMaskedEdit.CMaskedEditProc ( _
   BYVAL hwnd   AS HWND, _                 ' // Control window handle
   BYVAL uMsg   AS UINT, _                 ' // Type of message
   BYVAL wParam AS WPARAM, _               ' // First message parameter
   BYVAL lParam AS LPARAM, _               ' // Second message parameter
   BYVAL uIdSubclass AS UINT_PTR, _        ' // The subclass ID
   BYVAL dwRefData AS DWORD_PTR _          ' // Pointer to reference data
   ) AS LRESULT

   SELECT CASE uMsg

      CASE WM_GETDLGCODE
         ' // All keyboard input
         FUNCTION = DLGC_WANTARROWS OR DLGC_WANTCHARS ' DLGC_WANTALLKEYS
         EXIT FUNCTION

      ' // Returns the class pointer
      CASE WM_USER + 100
         RETURN dwRefData

      CASE 9999
         ' // Custom message to set the focus in the first editable group
         DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
         IF pthis = NULL THEN EXIT FUNCTION
         pthis->SetPos
         EXIT FUNCTION

      CASE WM_KEYDOWN
         DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
         IF pthis = NULL THEN EXIT FUNCTION
         IF LEN(pthis->m_strMask) <> 0 AND LEN(pthis->m_strInputTemplate) <> 0 THEN
            pthis->OnKeyDown(wParam, HIWORD(lParam), LOWORD(lParam))
            EXIT FUNCTION
         END IF

      CASE WM_CHAR
         ' // Handle Ctr+V - Decimal 22, Hex 16 - by sending a WM_PASTE message
         IF wParam = 22 THEN
            SendMessageW hwnd, WM_PASTE, 0, 0
            EXIT FUNCTION
         END IF
         ' // Handle Ctrl+C - Decimal 3, Hex 3 - by sending a WM_COPY message
         IF wParam = 3 THEN
            SendMessageW hwnd, WM_COPY, 0, 0
            EXIT FUNCTION
         END IF
         ' // Handle Ctrl+X - Decimal 24, Hex 18 - by sending a WM_CUT message
         IF wParam = 24 THEN
            SendMessageW hwnd, WM_CUT, 0, 0
            EXIT FUNCTION
         END IF
         ' // Let handle back space in WM_KEYDOWN
         ' // If mask is enabled, let handle backspace in WM_KEYDOWN, else do default
         IF wParam = 8 THEN
            DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
            IF pthis = NULL THEN EXIT FUNCTION
            IF LEN(pthis->m_strMask) <> 0 AND LEN(pthis->m_strInputTemplate) <> 0 THEN EXIT FUNCTION
         END IF
         IF wParam = 127 THEN
            DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
            IF pthis = NULL THEN EXIT FUNCTION
            IF LEN(pthis->m_strMask) = 0 OR LEN(pthis->m_strInputTemplate) = 0 THEN EXIT FUNCTION
         END IF
         ' // Handle Ctrl+Z - Decimal 26, Hex 1A - Undo
         IF wParam = 26 THEN
            DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
            IF pthis = NULL THEN EXIT FUNCTION
            DIM nPos AS LONG = Edit_GetSelStart(pthis->hWindow)
            IF LEN(pthis->m_strUndo) THEN pthis->SetText(pthis->m_strUndo)
            Edit_SetSel(pthis->hWindow, nPos, nPos)
            EXIT FUNCTION
         END IF
         DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
         IF pthis = NULL THEN EXIT FUNCTION
         pthis->m_strUndo = pthis->GetText(FALSE)
         IF LEN(pthis->m_strMask) <> 0 AND LEN(pthis->m_strInputTemplate) <> 0 THEN
            IF iswprint(wParam) <> 0 AND (GetKeyState(VK_CONTROL) AND &H80) <> &H80 THEN
               pthis->OnCharPrintchar(wParam, HIWORD(lParam), LOWORD(lParam))
            END IF
            EXIT FUNCTION
         END IF

      CASE WM_CUT, WM_CLEAR
         DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
         IF pthis = NULL THEN EXIT FUNCTION
         pthis->m_strUndo = pthis->GetText(FALSE)
         DIM nStart AS LONG = Edit_GetSelStart(pthis->hWindow)
         DIM nEnd AS LONG = Edit_GetSelEnd(pthis->hWindow)
         DIM dwsText AS DWSTRING = pthis->GetText(FALSE)
         IF LEN(dwsText) = 0 THEN EXIT FUNCTION
         dwsText = MID(dwsText, nStart + 1, nEnd - nStart)
         IF uMsg = WM_CUT THEN AfxSetClipboardText(dwsText)
         pthis->OnCharDelete(ASC(" "), 0, 0)
         dwsText = pthis->GetText(FALSE)
         IF LEN(pthis->m_strInputTemplate) = 0 THEN EXIT FUNCTION
         pthis->SetText(pthis->m_strInputTemplate)
         FOR i AS LONG = 1 TO LEN(dwsText)
            Edit_SetSel(pthis->m_hCtl, i - 1, i - 1)
            IF MID(dwsText, i, 1) <> MID(pthis->m_strInputTemplate, i, 1) THEN
               IF ASC(MID(dwsText, i, 1)) <> pthis->m_chMaskInputTemplate THEN
                  pthis->OnCharPrintchar(ASC(MID(dwsText, i, 1)), 0, 0)
               END IF
            END IF
         NEXT
         Edit_SetSel(pthis->hWindow, nStart, nStart)
         pthis->SetPos
         EXIT FUNCTION

      CASE WM_PASTE
         DIM pthis AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, dwRefData)
         IF pthis = NULL THEN EXIT FUNCTION
         pthis->m_strUndo = pthis->GetText(FALSE)
         DIM dwsText AS DWSTRING = AfxGetClipboardText
         IF LEN(dwsText) = 0 THEN EXIT FUNCTION
         FOR i AS LONG = 1 TO LEN(dwsText)
            pthis->OnCharPrintchar((ASC(MID(dwsText, i, 1))), 0, 0)
         NEXT
         EXIT FUNCTION

      CASE WM_DESTROY
         ' // REQUIRED: Remove control subclassing
         RemoveWindowSubclass hwnd, @CMaskedEditProc, uIdSubclass

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CMaskedEdit
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CMaskedEdit (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1)

   IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR ES_AUTOHSCROLL
   IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE

   ' // Create the control
   IF pWindow THEN
      m_hCtl = pWindow->AddControl("EDIT", pWindow->hWindow, cID, "", _
               x, y, nWidth, nHeight, dwStyle, dwExStyle, NULL)
      ' // Subclass the control
      IF m_hCtl THEN pWindow->SetControlSubclass(m_hCtl, @CMaskedEditProc, cID, CAST(DWORD_PTR, @this))
   END IF

   m_bSelectByGroup = TRUE

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an instance of the control.
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.Create (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR,  _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = -1, BYVAL dwExStyle AS DWORD = -1) AS HWND

   ' // Check if it has been created already
   IF m_hCtl THEN RETURN m_hCtl

   IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR ES_AUTOHSCROLL
   IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE

   ' // Create the control
   IF pWindow THEN
      m_hCtl = pWindow->AddControl("EDIT", pWindow->hWindow, cID, "", _
               x, y, nWidth, nHeight, dwStyle, dwExStyle, NULL)
      ' // Subclass the control
      IF m_hCtl THEN pWindow->SetControlSubclass(m_hCtl, @CMaskedEditProc, cID, CAST(DWORD_PTR, @this))
   END IF

   m_bSelectByGroup = TRUE

END FUNCTION
' ========================================================================================

' ========================================================================================
' CMaskedEdit class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CMaskedEdit
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the control
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CMaskedEdit.GetGroupBounds(BYREF nBegin AS LONG, BYREF nEnd AS LONG, BYVAL nStartPos AS LONG, BYVAL bForward AS BOOLEAN = TRUE)
   IF LEN(m_strMask) = 0 OR LEN(m_strInputTemplate) = 0 THEN EXIT SUB
   IF LEN(m_strInputTemplate) THEN   ' /// use mask
      IF LEN(m_str) <> LEN(m_strMask) THEN EXIT SUB
      IF nStartPos < 0 THEN EXIT SUB
      IF nStartPos > LEN(m_strInputTemplate) THEN EXIT SUB
      IF bForward THEN
         ' // If nStartPos is in the middle of a group
         ' // Reverse search for the begin of a group
         DIM i AS LONG = nStartPos
         IF nStartPos > 0 THEN
            IF m_strInputTemplate[nStartPos-1] = ASC("_") THEN
               DO
                  i -= 1
               LOOP WHILE (i > 0) AND m_strInputTemplate[i] = ASC("_")
            END IF
         END IF
         IF i = LEN(m_strInputTemplate) THEN
            nBegin = -1   ' // no group
            nEnd = 0
            EXIT SUB
         END IF
         ' // i points between groups or to the begin of a group
         ' // Search for the begin of a group
         IF m_strInputTemplate[i] <> ASC("_") THEN
            i = INSTR(i + 1, **m_strInputTemplate, "_") - 1
            IF i = -1 THEN
               nBegin = -1   ' // no group
               nEnd = 0
               EXIT SUB
            END IF
         END IF
         nBegin = i
         ' // Search for the end of a group
         WHILE i < LEN(m_strInputTemplate) AND m_strInputTemplate[i] = ASC("_")
            i += 1
         WEND
         nEnd = i
      ELSE   ' // backward
         ' // If nStartPos is in the middle of a group
         ' // Search for the end of a group
         DIM i AS LONG = nStartPos
         WHILE i < LEN(m_str) AND m_strInputTemplate[i] = ASC("_")
            i += 1
         WEND
         IF i = 0 THEN
            nBegin = -1   ' // no group
            nEnd = 0
            EXIT SUB
         END IF
         ' // i points between groups or to the end of a group
         ' // Reverse search for the end of a group
         IF m_strInputTemplate[i-1] <> ASC("_") THEN
            DO
               i -= 1
            LOOP WHILE i > 0 AND m_strInputTemplate[i-1] <> ASC("_")
            IF i = 0 THEN
               nBegin = -1   ' // no group
               nEnd = 0
               EXIT SUB
            END IF
         END IF
         nEnd = i
         ' // Search for the begin of a group
         DO
            i -= 1
         LOOP WHILE i > 0 AND m_strInputTemplate[i-1] = ASC("_")
         nBegin = i
      END IF
   ELSE   ' // don't use mask
      ' // nStartPos ignored
      nBegin = 0
      nEnd = LEN(m_str)
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Returns TRUE if the symbol is valid
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.CheckChar (BYVAL chChar AS USHORT, BYVAL nPos AS LONG) AS BOOLEAN
   IF LEN(m_strMask) = 0 OR LEN(m_strInputTemplate) = 0 THEN RETURN FALSE
   IF iswprint(chChar) = FALSE THEN RETURN FALSE
   IF nPos < 0 THEN RETURN FALSE
   ' // Use mask
   IF LEN(m_str) <> LEN(m_strMask) THEN RETURN FALSE
   IF m_strInputTemplate[nPos] = ASC("_") THEN
      DIM bIsMaskedChar AS BOOLEAN = this.IsMaskedChar(chChar, m_strMask[nPos])
      ' // Use valid string characters
      IF LEN(m_strValid) THEN
         IF bIsMaskedChar THEN
            RETURN INSTR(**m_strValid, WCHR(chChar)) <> 0
         ELSE
            RETURN FALSE
         END IF
      ELSE
         ' // Don't use valid string characters
         RETURN bIsMaskedChar
      END IF
   ELSE
      RETURN FALSE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes the masked edit control.
' ========================================================================================
PRIVATE SUB CMaskedEdit.EnableMask (BYVAL lpszMask AS WSTRING PTR, BYVAL lpszInputTemplate AS WSTRING PTR, _
BYREF chMaskInputTemplate AS DWSTRING = "_", BYVAL lpszValid AS WSTRING PTR = NULL)
   IF lpszMask = NULL THEN EXIT SUB
   IF lpszInputTemplate = NULL THEN EXIT SUB
   IF iswprint(ASC(chMaskInputTemplate, 1)) = FALSE THEN EXIT SUB
   m_strMask = *lpszMask
   m_strInputTemplate = *lpszInputTemplate
   m_chMaskInputTemplate = ASC(chMaskInputTemplate, 1)
   m_str = *lpszInputTemplate
   IF LEN(m_strMask) <> LEN(m_strInputTemplate) THEN EXIT SUB
   m_strValid = ""
   IF lpszValid THEN m_strValid = *lpszValid
   this.SetText(m_strInputTemplate)
END SUB
' ========================================================================================

' ========================================================================================
' Specifies a string of valid characters that the user can enter.
' ========================================================================================
PRIVATE SUB CMaskedEdit.SetValidChars (BYVAL lpszValid AS WSTRING PTR)
   m_strValid = ""
   IF lpszValid THEN m_strValid = *lpszValid
END SUB
' ========================================================================================

' ========================================================================================
' Called by the class to validate the specified character against the corresponding mask character.
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.IsMaskedChar (BYVAL chChar AS USHORT, BYVAL chMaskChar AS USHORT) AS BOOLEAN
   ' // Check the key against the mask
   SELECT CASE chMaskChar
      CASE ASC("D")   ' // digit only
         IF iswdigit(chChar) THEN RETURN TRUE
      CASE ASC("d")   ' // digit only
         IF iswdigit(chChar) THEN RETURN TRUE
         IF iswspace(chChar) THEN RETURN TRUE
      CASE ASC("+")   ' // '+' or '-' or space
         IF chChar = ASC("+") OR chChar = ASC("-") THEN RETURN TRUE
         IF iswspace(chChar) THEN RETURN TRUE
      CASE ASC("C")   ' // alpha only
         IF iswalpha(chChar) THEN RETURN TRUE
      CASE ASC("c")   ' // alpha or space
         IF iswalpha(chChar) THEN RETURN TRUE
         IF iswspace(chChar) THEN RETURN TRUE
      CASE ASC("A")   ' // alpha numeric only
         IF iswalnum(chChar) THEN RETURN TRUE
      CASE ASC("a")   ' // alpha numeric or space
         IF iswalnum(chChar) THEN RETURN TRUE
         IF iswspace(chChar) THEN RETURN TRUE
      CASE ASC("*")   ' // a printable character
         IF iswprint(chChar) THEN RETURN TRUE
   END SELECT   
   RETURN FALSE   ' // not allowed symbol
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.SetValue (BYVAL lpszString AS WSTRING PTR, BYVAL bWithDelimiters AS BOOLEAN = TRUE) AS BOOLEAN
   IF lpszString = NULL THEN RETURN FALSE
   ' // Make sure the string is not longer than the mask
   DIM dwsSource AS DWSTRING = *lpszString
   IF LEN(m_strMask) THEN
      IF bWithDelimiters THEN
         IF LEN(dwsSource) > LEN(m_strMask) THEN RETURN FALSE
      ELSE
         ' // Count _T('_') in m_strInputTemplate
         DIM nCount AS LONG
         FOR i AS LONG = 0 TO LEN(m_strInputTemplate) - 1
            IF m_strInputTemplate[i] = ASC("_") THEN nCount += 1
         NEXT
         IF LEN(dwsSource) > nCount THEN RETURN FALSE
      END IF
   END IF
   ' // Make sure the value has only valid string characters
   IF LEN(m_strValid) THEN
      DIM bOk AS BOOLEAN = TRUE
      IF bOk THEN
         FOR iPos AS LONG = 0 TO LEN(dwsSource) - 1
            IF LEN(m_strInputTemplate) = 0 OR m_strInputTemplate[iPos] = ASC("_") THEN
               IF LEN(m_strInputTemplate) = 0 OR dwsSource[iPos] <> m_chMaskInputTemplate THEN   ' // allow m_chMaskInputTemplate
                  bOk = INSTR(m_strValid, WCHR(dwsSource[iPos]))
               END IF
            END IF
         NEXT
      END IF
      IF bOk = FALSE THEN RETURN FALSE
   END IF
   ' // Use mask, validate against the mask
   IF LEN(m_strMask) THEN
      IF LEN(m_str) <> LEN(m_strMask) THEN RETURN FALSE
      DIM dwsResult AS DWSTRING = m_strInputTemplate
      ' // Replace '_' with default char
      FOR i AS LONG = 0 TO LEN(dwsResult) - 1
         IF m_strInputTemplate[i] = ASC("_") THEN
            MID(dwsResult, i + 1, 1) = WCHR(m_chMaskInputTemplate)
         END IF
      NEXT
      DIM iSrcChar AS LONG
      DIM iDstChar AS LONG
      WHILE iSrcChar < LEN(dwsSource) AND iDstChar < LEN(m_strInputTemplate)
         ' // iDstChar - character entry position("_" char)
         IF m_strInputTemplate[iDstChar] = ASC("_") THEN
            DIM chChar AS USHORT = dwsSource[iSrcChar]
            IF chChar <> m_chMaskInputTemplate THEN   ' // allow m_chMaskInputTemplate
               IF IsMaskedChar(chChar, m_strMask[iDstChar]) = FALSE THEN RETURN FALSE
            END IF
            MID(dwsResult, iDstChar + 1, 1) = WCHR(chChar)
            iSrcChar += 1
            iDstChar += 1
         ' // iDstChar - delimeter
         ELSE
            IF bWithDelimiters THEN
               IF m_strInputTemplate[iDstChar] <> dwsSource[iSrcChar] THEN RETURN FALSE
               iSrcChar += 1
               iDstChar += 1
            ELSE
               iDstChar += 1
            END IF
         END IF
      WEND
      m_str = dwsResult
   ELSE   ' // Don't use mask
      m_str = dwsSource
   END IF
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.DoUpdate (BYVAL bRestoreLastGood AS BOOLEAN, BYVAL nBeginOld AS LONG = -1, BYVAL nEndOld AS LONG = -1) AS BOOLEAN
   DIM dwsNew AS DWSTRING = AfxGetWindowText(m_hCtl)
   DIM bRet AS BOOLEAN = this.SetValue(dwsNew, TRUE)
   IF bRet = FALSE THEN MessageBeep(CAST(UINT, -1))
   IF bRet = FALSE AND bRestoreLastGood = TRUE THEN
      DIM dwsOld AS DWSTRING = m_str
      SetWindowText(m_hCtl, dwsOld)
      IF nBeginOld <> -1 THEN Edit_SetSel(m_hCtl, nBeginOld, nEndOld)
   END IF
   RETURN bRet
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CMaskedEdit.OnCharPrintchar(BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)
   IF LEN(m_strMask) <> LEN(m_strInputTemplate) THEN RETURN
   DIM chChar AS USHORT = nChar
   IF iswprint(chChar) = FALSE THEN RETURN
   ' // Processing ES_UPPERCASE and ES_LOWERCASE styles
   DIM dwStyle AS DWORD = AfxGetWindowStyle(m_hCtl)
   IF (dwStyle AND ES_UPPERCASE) = ES_UPPERCASE THEN
      chChar = towupper(chChar)
   ELSEIF (dwStyle AND ES_LOWERCASE) = ES_LOWERCASE THEN
      chChar = towlower(chChar)
   END IF
   DIM nStartPos AS LONG, nEndPos AS LONG
   nStartPos = Edit_GetSelStart(m_hCtl)
   nEndPos = Edit_GetSelEnd(m_hCtl)
   ' // Calc group bounds
   DIM nGroupStart AS LONG, nGroupEnd AS LONG
   this.GetGroupBounds(nGroupStart, nGroupEnd, nStartPos)
   ' // Out of range
'   IF (nStartPos < 0 AND nEndPos > LEN(m_str)) OR nStartPos < nGroupStart OR nStartPos > nGroupEnd OR _
'      nEndPos < nGroupStart OR nEndPos > nGroupEnd THEN
'      MessageBeep(CAST(UINT, -1))
'      Edit_SetSel(m_hCtl, nGroupStart, nGroupEnd)
'      RETURN
'   END IF
   ' // No selected chars
   IF nStartPos = nEndPos THEN
      ' // Use m_strMask
      IF LEN(m_strMask) THEN
         ' // Automatically move the cursor to the next group
         IF nEndPos = nGroupEnd OR _   ' // at the end of group
            nStartPos < nGroupStart OR nStartPos > nGroupEnd THEN   ' // not in the middle of a group
            ' // no space for new char
            IF nEndPos >= LEN(m_str) - 1 THEN
               MessageBeep(CAST(UINT, -1))
               RETURN
            ' // can search next group
            ELSEIF nEndPos < LEN(m_str) - 1 THEN
               this.GetGroupBounds(nGroupStart, nGroupEnd, nEndPos + 1, TRUE)
            END IF
            ' // if next group was found
            IF nGroupStart <> -1 AND nGroupStart > nEndPos THEN
               Edit_SetSel(m_hCtl, nGroupStart, nGroupStart)
               nStartPos = nGroupStart
               nEndPos = nGroupStart
            ' // no more groups
            ELSE
               MessageBeep(CAST(UINT, -1))
               RETURN
            END IF
         END IF
         ' // Check char in position
         IF this.CheckChar(chChar, nStartPos) = FALSE THEN
            MessageBeep(CAST(UINT, -1))
            RETURN
         END IF
         ' // Replace char in Editbox and m_str
         Edit_SetSel(m_hCtl, nStartPos, nEndPos + 1)
         DIM dws AS DWSTRING = WCHR(chChar)
         Edit_ReplaceSel(m_hCtl, *dws)
         MID(m_str, nEndPos + 1, 1) = WCHR(chChar)
         Edit_SetSel(m_hCtl, nEndPos + 1, nEndPos + 1)
         ' // Automaticaly move the cursor to the next group
         nStartPos = Edit_GetSelStart(m_hCtl)
         nEndPos = Edit_GetSelEnd(m_hCtl)
         IF nEndPos = nGroupEnd THEN   ' // at the end of group
            ' // can search next group
            IF nEndPos < LEN(m_str) - 1 THEN
               this.GetGroupBounds(nGroupStart, nGroupEnd, nEndPos + 1, TRUE)
            END IF
            ' // if next group was found
            IF nGroupStart <> -1 AND nGroupStart > nEndPos THEN
               Edit_SetSel(m_hCtl, nGroupStart, nGroupStart)
               nStartPos = nGroupStart
               nEndPos = nGroupStart
            END IF
         END IF
      ' // Don't use m_strMask
      ELSE
         ' // Check char in position
         IF this.CheckChar(chChar, nStartPos) = FALSE THEN
            MessageBeep(CAST(UINT, -1))
            RETURN
         END IF
         ' // Don't use m_chMaskInputTemplate
         DIM nBeginOld AS LONG, nEndOld AS LONG
         nBeginOld = Edit_GetSelStart(m_hCtl)
         nEndOld = Edit_GetSelEnd(m_hCtl)
         this.DoUpdate(TRUE, nBeginOld, nEndOld)
      END IF
   ELSE   ' // Have one or more chars selected
      ' // Check char in position
      IF CheckChar(chChar, nStartPos) = FALSE THEN
            MessageBeep(CAST(UINT, -1))
            RETURN
      END IF
      ' // Replace chars in Editbox and m_str
      IF LEN(m_strInputTemplate) THEN   ' // Use m_strInputTemplate
         IF nStartPos > LEN(m_strInputTemplate) THEN RETURN
         IF nEndPos > LEN(m_strInputTemplate) THEN RETURN
         DIM nRange AS LONG = nEndPos - nStartPos
         IF nRange < 1 THEN RETURN
         DIM dwsReplace AS DWSTRING
         dwsReplace = WSTRING(nRange, m_chMaskInputTemplate)
         MID(dwsReplace, 1, 1) = WCHR(chChar)
         Edit_SetSel(m_hCtl, nStartPos, nStartPos + nRange)
         Edit_ReplaceSel(m_hCtl, *dwsReplace)
         Edit_SetSel(m_hCtl, nStartPos, nStartPos)
         FOR i AS LONG = 0 TO LEN(dwsReplace) - 1
            MID(m_str, nStartPos + i + 1, 1) = WCHR(dwsReplace[i])
         NEXT
         Edit_SetSel(m_hCtl, nStartPos + 1, nStartPos + 1)
      ELSE
         ' // Don't use m_chMaskInputTemplate
         DIM nBeginOld AS LONG, nEndOld AS LONG
         nBeginOld = Edit_GetSelStart(m_hCtl)
         nEndOld = Edit_GetSelEnd(m_hCtl)
         this.DoUpdate(TRUE, nBeginOld, nEndOld)
      END IF
   END IF
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CMaskedEdit.OnKeyDown(BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)

   ' // Make sure the mask has entry positions
   DIM nStartBound AS LONG, nEndBound AS LONG
   this.GetGroupBounds(nStartBound, nEndBound)
   IF nStartBound = -1 THEN
      ' // mask has no entry positions
      MessageBeep(CAST(UINT, -1))
      RETURN
   END IF

   SELECT CASE nChar
      CASE VK_END
         ' // Set the cursor at the end of the text
         IF GetKeyState(VK_CONTROL) AND &h80 THEN
            Edit_SetSel(m_hCtl, LEN(m_str), LEN(m_str))
            RETURN
         END IF
         ' // Calc last group bounds
         DIM nGroupStart AS LONG, nGroupEnd AS LONG
         nGroupStart = Edit_GetSelStart(m_hCtl)
         nGroupEnd = Edit_GetSelEnd(m_hCtl)
         this.GetGroupBounds(nGroupStart, nGroupEnd, nGroupEnd, TRUE)
         IF nGroupStart = -1 THEN
            this.GetGroupBounds(nGroupStart, nGroupEnd, LEN(m_str), FALSE)
         END IF
         IF nGroupStart = -1 THEN RETURN
         Edit_SetSel(m_hCtl, nGroupEnd, nGroupEnd)
         RETURN

      CASE VK_HOME
         ' // Set the cursor at the end of the text
         IF GetKeyState(VK_CONTROL) AND &h80 THEN
            Edit_SetSel(m_hCtl, 1, 1)
            RETURN
         END IF
         ' // Calc last group bounds
         DIM nGroupStart AS LONG, nGroupEnd AS LONG
         nGroupStart = Edit_GetSelStart(m_hCtl)
         nGroupEnd = Edit_GetSelEnd(m_hCtl)
         this.GetGroupBounds(nGroupStart, nGroupEnd, nGroupEnd, FALSE)
         IF nGroupStart = -1 THEN
            this.GetGroupBounds(nGroupStart, nGroupEnd, 0, TRUE)
         END IF
         IF nGroupStart = -1 THEN RETURN
         Edit_SetSel(m_hCtl, nGroupStart, nGroupStart)
         RETURN

      CASE VK_UP, VK_LEFT
         ' // Calc last group bounds
         DIM nGroupStart AS LONG, nGroupEnd AS LONG
         nGroupStart = Edit_GetSelStart(m_hCtl)
         nGroupEnd = Edit_GetSelEnd(m_hCtl)
         this.GetGroupBounds(nGroupStart, nGroupEnd, nGroupEnd, FALSE)
         IF nGroupStart = -1 THEN
            this.GetGroupBounds(nGroupStart, nGroupEnd, 0, TRUE)
         END IF
         IF nGroupStart = -1 THEN RETURN
         IF GetKeyState(VK_SHIFT) AND &h80 THEN
            DIM nStart AS LONG, nEnd AS LONG
            nStart = Edit_GetSelStart(m_hCtl)
            nEnd = Edit_GetSelEnd(m_hCtl)
            IF m_bSelectByGroup THEN
               DIM nNewStart AS LONG = MAX(nStart-1, nGroupStart)
               ' // additional
               nNewStart = MIN(nNewStart, nGroupEnd)
               Edit_SetSel(m_hCtl, nNewStart, nEnd)
            ELSE
               Edit_SetSel(m_hCtl, nStart-1, nEnd)
            END IF
         ELSEIF GetKeyState(VK_CONTROL) AND &h80 THEN
            ' // move to the previous group
            DIM nStart AS LONG, nEnd AS LONG
            nStart = Edit_GetSelStart(m_hCtl)
            nEnd = Edit_GetSelEnd(m_hCtl)
            IF nStart < 0 THEN RETURN
            IF nStart > 1 THEN   ' // can search previous group
               this.GetGroupBounds(nGroupStart, nGroupEnd, nStart-1, FALSE)
            END IF
            IF (nGroupStart > -1) AND _  ' // if previous group was found
               (nGroupStart <> nStart OR nGroupEnd <> nEnd) THEN   ' // and it's not the same
               Edit_SetSel(m_hCtl, nGroupStart, nGroupEnd)
            ELSE   ' else // no more groups
               MessageBeep(CAST(UINT, -1))
            END IF
            RETURN
         ELSE
            DIM nStart AS LONG, nEnd AS LONG
            nStart = Edit_GetSelStart(m_hCtl)
            nEnd = Edit_GetSelEnd(m_hCtl)
            ' // move to the previous group
            IF nStart = nEnd AND nStart = nGroupStart THEN
               IF nStart > 1 THEN   ' // can search previous group
                  this.GetGroupBounds(nGroupStart, nGroupEnd, nStart-1, FALSE)
               END IF
               IF nGroupStart <> -1 AND nGroupEnd < nStart THEN   ' // if previous group was found
                  Edit_SetSel(m_hCtl, nGroupEnd, nGroupEnd)
               ELSE   ' // no more groups
                  MessageBeep(CAST(UINT, -1))
               END IF
            ELSE
               DIM nNewStart AS LONG = MAX(nStart-1, nGroupStart)
               ' // additional
               nNewStart = MIN(nNewStart, nGroupEnd)
               Edit_SetSel(m_hCtl, nNewStart, nNewStart)
            END IF
            RETURN
         END IF

      CASE VK_DOWN, VK_RIGHT
         ' // Calc last group bounds
         DIM nGroupStart AS LONG, nGroupEnd AS LONG
         nGroupStart = Edit_GetSelStart(m_hCtl)
         nGroupEnd = Edit_GetSelEnd(m_hCtl)
         this.GetGroupBounds(nGroupStart, nGroupEnd, nGroupEnd, TRUE)
         IF nGroupStart = -1 THEN
            this.GetGroupBounds(nGroupStart, nGroupEnd, LEN(m_str), FALSE)
         END IF
         IF nGroupStart = -1 THEN RETURN
         IF GetKeyState(VK_SHIFT) AND &h80 THEN
            DIM nStart AS LONG, nEnd AS LONG
            nStart = Edit_GetSelStart(m_hCtl)
            nEnd = Edit_GetSelEnd(m_hCtl)
            IF m_bSelectByGroup THEN
               DIM nNewEnd AS LONG = MIN(nEnd + 1, nGroupEnd)
               ' // additional
               nNewEnd = MAX(nNewEnd, nGroupStart)
               Edit_SetSel(m_hCtl, nStart, nNewEnd)
            ELSE
               Edit_SetSel(m_hCtl, nStart, nEnd + 1)
            END IF
            RETURN
         ELSEIF GetKeyState(VK_CONTROL) AND &h80 THEN
            ' // move to the next group
            DIM nStart AS LONG, nEnd AS LONG
            nStart = Edit_GetSelStart(m_hCtl)
            nEnd = Edit_GetSelEnd(m_hCtl)
            IF nStart = -1 THEN RETURN
            IF nEnd < LEN(m_str) -1 THEN ' // can search next group
               this.GetGroupBounds(nGroupStart, nGroupEnd, nEnd + 1, TRUE)
            END IF
            IF (nGroupStart <> -1) AND _   ' // if previous group was found
               (nGroupStart <> nStart OR nGroupEnd <> nEnd) THEN   ' // and it's not the same
               Edit_SetSel(m_hCtl, nGroupStart, nGroupEnd)
            ELSE   ' // no more groups
               MessageBeep(CAST(UINT, -1))
            END IF
            RETURN
         ELSE
            DIM nStart AS LONG, nEnd AS LONG
            nStart = Edit_GetSelStart(m_hCtl)
            nEnd = Edit_GetSelEnd(m_hCtl)
            ' // move to the next group
            IF nStart = nEnd AND nEnd = nGroupEnd THEN
               IF nEnd < LEN(m_str) - 1 THEN   ' // can search next group
                  this.GetGroupBounds(nGroupStart, nGroupEnd, nStart + 1, TRUE)
               END IF
               IF nGroupStart <> -1 AND nGroupStart > nEnd THEN   ' // if next group was found
                  Edit_SetSel(m_hCtl, nGroupStart, nGroupStart)
               ELSE   ' // no more groups
                  MessageBeep(CAST(UINT, -1))
               END IF
            ELSE
               DIM nNewEnd AS LONG = MIN(nEnd + 1, nGroupEnd)
               ' // additional
               nNewEnd = MAX(nNewEnd, nGroupStart)
               Edit_SetSel(m_hCtl, nNewEnd, nNewEnd)
            END IF
            RETURN
         END IF

      CASE VK_BACK
         IF GetKeyState(VK_SHIFT) AND &h80 THEN RETURN
         ' // Special processing
         this.m_strUndo = this.GetText(FALSE)
         this.OnCharBackspace(nChar, nRepCnt, nFlags)
         RETURN

      CASE VK_DELETE
         IF GetKeyState(VK_SHIFT) AND &h80 THEN RETURN
         ' // Special processing
         this.m_strUndo = this.GetText(FALSE)
         this.OnCharDelete(nChar, nRepCnt, nFlags)
         RETURN

   END SELECT

END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CMaskedEdit.OnCharBackspace (BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)

   IF LEN(m_strMask) = 0 OR LEN(m_strInputTemplate) = 0 THEN RETURN
   DIM nStartPos AS LONG, nEndPos AS LONG
   nStartPos = Edit_GetSelStart(m_hCtl)
   nEndPos = Edit_GetSelEnd(m_hCtl)
   IF nStartPos < 0 THEN RETURN
   IF nEndPos < 0 THEN RETURN
   IF nEndPos < nStartPos THEN RETURN

   ' // Calc group bounds
   DIM nGroupStart AS LONG, nGroupEnd AS LONG
   this.GetGroupBounds(nGroupStart, nGroupEnd, nStartPos)

'   // Out of range
'   if ((nStartPos<0) &&(nEndPos > m_str.GetLength()) || (nStartPos < nGroupStart) ||(nStartPos > nGroupEnd) || (nEndPos < nGroupStart) ||(nEndPos > nGroupEnd))
'   {
'      MessageBeep((UINT)-1);
'      CEdit::SetSel(nGroupStart, nGroupEnd);
'      return;
'   }

   ' // No selected chars
   IF nStartPos = nEndPos THEN
      ' // Use m_strMask
      IF LEN(m_strMask) THEN
         ' // Automaticaly move the cursor to the previous group
         IF nEndPos = nGroupStart THEN   ' // at the start of group
            ' // can search previous group
            IF nEndPos > 1 THEN
               this.GetGroupBounds(nGroupStart, nGroupEnd, nEndPos - 1, FALSE)
            END IF
            ' // if previous group was found
            IF nGroupStart <> -1 AND nGroupEnd < nEndPos THEN
               Edit_SetSel(m_hCtl, nGroupEnd, nGroupEnd)
               RETURN
            ' // no more group
            ELSE
               MessageBeep(CAST(UINT, -1))
               RETURN
            END IF
         END IF
         '// Calc the number of literals with the same mask char
         IF nStartPos < 1 THEN RETURN
         IF nEndPos < 1 THEN RETURN
         IF nGroupEnd > LEN(m_strInputTemplate) THEN RETURN
         DIM nSameMaskCharsNum AS LONG = 1
         DIM nIndex AS LONG = nStartPos - 1   ' // an index of the char to delete
         DIM chMaskChar AS USHORT = m_strMask[nIndex]
         DIM bScanMore AS BOOLEAN = TRUE
         WHILE bScanMore = TRUE AND nIndex + nSameMaskCharsNum < nGroupEnd
            IF m_strMask[nIndex + nSameMaskCharsNum] = chMaskChar THEN
               nSameMaskCharsNum += 1
            ELSE
               bScanMore = FALSE
            END IF
         WEND
         ' // Form the shifted string
         IF nIndex < nGroupStart THEN RETURN
         IF nIndex + nSameMaskCharsNum > nGroupEnd THEN RETURN
         DIM dwsReplace AS DWSTRING = MID(m_str, nIndex + 1, nSameMaskCharsNum)
         IF nSameMaskCharsNum > 0 THEN
            dwsReplace = RIGHT(dwsReplace, nSameMaskCharsNum - 1)
            dwsReplace += WCHR(m_chMaskInputTemplate)
         END IF
         ' // Replace the content with the shifted string
         Edit_SetSel(m_hCtl, nIndex, nIndex + nSameMaskCharsNum)
         Edit_ReplaceSel(m_hCtl, *dwsReplace)
         Edit_SetSel(m_hCtl, nIndex, nIndex)
         FOR i AS LONG = 0 TO LEN(dwsReplace) - 1
            MID(m_str, nIndex + i + 1) = WCHR(dwsReplace[i])
         NEXT
      ELSE   ' // Don't use m_chMaskInputTemplate - delete symbol
         DIM nBeginOld AS LONG, nEndOld AS LONG
         nBeginOld = Edit_GetSelStart(m_hCtl)
         nEndOld = Edit_GetSelEnd(m_hCtl)
         this.DoUpdate(TRUE, nBeginOld, nEndOld)
      END IF
   ELSE   ' // Have one or more chars selected
      ' // Replace chars in Editbox and m_str
      IF LEN(m_strInputTemplate) THEN   ' // Use m_strInputTemplate
         IF nStartPos > LEN(m_strInputTemplate) THEN RETURN
         IF nEndPos > LEN(m_strInputTemplate) THEN RETURN
         DIM nRange AS LONG = nEndPos - nStartPos
         IF nRange < 1 THEN RETURN
         DIM dwsReplace AS DWSTRING
         dwsReplace = WSTRING(nRange, m_chMaskInputTemplate)
         Edit_SetSel(m_hCtl, nStartPos, nStartPos + nRange)
         Edit_ReplaceSel(m_hCtl, *dwsReplace)
         Edit_SetSel(m_hCtl, nStartPos, nStartPos)
         FOR i AS LONG = 0 TO LEN(dwsReplace) - 1
            MID(m_str, nStartPos + i + 1, 1) = WCHR(dwsReplace[i])
         NEXT
         Edit_SetSel(m_hCtl, nStartPos, nStartPos)
      ELSE
         ' // Don't use m_chMaskInputTemplate - delete symbols
         DIM nBeginOld AS LONG, nEndOld AS LONG
         nBeginOld = Edit_GetSelStart(m_hCtl)
         nEndOld = Edit_GetSelEnd(m_hCtl)
         this.DoUpdate(TRUE, nBeginOld, nEndOld)
      END IF
   END IF
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CMaskedEdit.OnCharDelete (BYVAL nChar AS UINT, BYVAL nRepCnt AS UINT, BYVAL nFlags AS UINT)
   IF LEN(m_strMask) = 0 OR LEN(m_strInputTemplate) = 0 THEN RETURN
   DIM nStartPos AS LONG, nEndPos AS LONG
   nStartPos = Edit_GetSelStart(m_hCtl)
   nEndPos = Edit_GetSelEnd(m_hCtl)
   IF nStartPos < 1 THEN RETURN
   IF nEndPos < 1 THEN RETURN
   IF nEndPos < nStartPos THEN RETURN
   ' // Calc group bounds
   DIM nGroupStart AS LONG, nGroupEnd AS LONG
   nGroupStart = Edit_GetSelStart(m_hCtl)
   nGroupEnd = Edit_GetSelEnd(m_hCtl)
   this.GetGroupBounds(nGroupStart, nGroupEnd, nGroupStart)

'   // Out of range
'   if ((nStartPos<0) &&(nEndPos > m_str.GetLength()) || (nStartPos < nGroupStart) ||(nStartPos > nGroupEnd) || (nEndPos < nGroupStart) ||(nEndPos > nGroupEnd))
'   {
'      MessageBeep((UINT)-1);
'      CEdit::SetSel(nGroupStart, nGroupEnd);
'      return;
'   }

   ' // No selected chars
   IF nStartPos = nEndPos THEN
      IF LEN(m_strMask) THEN   ' // Don't use m_strMask
         ' // Make sure the cursor is not at the end of group
         IF nEndPos = nGroupEnd THEN
            MessageBeep(CAST(UINT, -1))
            RETURN
         END IF
         ' // Calc the number of literals with the same mask char
         IF nStartPos < 1 THEN RETURN
         IF nEndPos < 1 THEN RETURN
         IF nGroupEnd > LEN(m_strInputTemplate) THEN RETURN
         DIM nSameMaskCharsNum AS LONG = 1
         DIM nIndex AS LONG = nStartPos   ' // an index of the char to delete
         DIM chMaskChar AS USHORT = m_strMask[nIndex]
         DIM bScanMore AS BOOLEAN = TRUE
         WHILE bScanMore = TRUE AND nIndex + nSameMaskCharsNum < nGroupEnd
            IF m_strMask[nIndex + nSameMaskCharsNum] = chMaskChar THEN
               nSameMaskCharsNum += 1
            ELSE
               bScanMore = FALSE
            END IF
         WEND
         ' // Form the shifted string
         IF nIndex < nGroupStart THEN RETURN
         IF nIndex + nSameMaskCharsNum > nGroupEnd THEN RETURN
         DIM dwsReplace AS DWSTRING = MID(m_str, nIndex + 1, nSameMaskCharsNum)
         IF nSameMaskCharsNum > 0 THEN
            dwsReplace = RIGHT(dwsReplace, nSameMaskCharsNum - 1)
            dwsReplace += WCHR(m_chMaskInputTemplate)
         END IF
         ' // Replace the content with the shifted string
         Edit_SetSel(m_hCtl, nIndex, nIndex + nSameMaskCharsNum)
         Edit_ReplaceSel(m_hCtl, *dwsReplace)
         Edit_SetSel(m_hCtl, nIndex, nIndex)
         FOR i AS LONG = 0 TO LEN(dwsReplace) - 1
            MID(m_str, nIndex + i + 1) = WCHR(dwsReplace[i])
         NEXT
      ELSE
         DIM nBeginOld AS LONG, nEndOld AS LONG
         nBeginOld = Edit_GetSelStart(m_hCtl)
         nEndOld = Edit_GetSelEnd(m_hCtl)
         this.DoUpdate(TRUE, nBeginOld, nEndOld)
      END IF
   ELSE   ' // Have one or more chars selected
      ' // Replace chars in Editbox and m_str
      IF LEN(m_strInputTemplate) THEN   ' // Use m_strInputTemplate
         IF nStartPos > LEN(m_strInputTemplate) THEN RETURN
         IF nEndPos > LEN(m_strInputTemplate) THEN RETURN
         DIM nRange AS LONG = nEndPos - nStartPos
         IF nRange < 1 THEN RETURN
         DIM dwsReplace AS DWSTRING
         dwsReplace = WSTRING(nRange, m_chMaskInputTemplate)
         Edit_SetSel(m_hCtl, nStartPos, nStartPos + nRange)
         Edit_ReplaceSel(m_hCtl, *dwsReplace)
         Edit_SetSel(m_hCtl, nStartPos, nStartPos)
         FOR i AS LONG = 0 TO LEN(dwsReplace) - 1
            MID(m_str, nStartPos + i + 1, 1) = WCHR(dwsReplace[i])
         NEXT
         Edit_SetSel(m_hCtl, nStartPos, nStartPos)
      ELSE
         ' // Don't use m_chMaskInputTemplate - delete symbols
         DIM nBeginOld AS LONG, nEndOld AS LONG
         nBeginOld = Edit_GetSelStart(m_hCtl)
         nEndOld = Edit_GetSelEnd(m_hCtl)
         this. DoUpdate(TRUE, nBeginOld, nEndOld)
      END IF
   END IF
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CMaskedEdit.SetPos (BYVAL nPos AS LONG = -1)
   IF nPos = -1 THEN
      ' // Calc last group bounds
      DIM nGroupStart AS LONG, nGroupEnd AS LONG
      nGroupStart = Edit_GetSelStart(m_hCtl)
      nGroupEnd = Edit_GetSelEnd(m_hCtl)
      this.GetGroupBounds(nGroupStart, nGroupEnd, nGroupEnd, FALSE)
      IF nGroupStart = -1 THEN
         this.GetGroupBounds(nGroupStart, nGroupEnd, 0, TRUE)
      END IF
      IF nGroupStart = -1 THEN RETURN
      Edit_SetSel(m_hCtl, nGroupStart, nGroupStart)
      RETURN
   END IF
   Edit_SetSel(m_hCtl, nPos, nPos)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the text of the control
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.SetText (BYREF dwsText AS DWSTRING) AS BOOLEAN
   IF LEN(m_strInputTemplate) = 0 THEN
      SetWindowText(m_hCtl, dwsText)
   ELSEIF dwsText = "" THEN
      ' // Clear the contents of the control
      m_str = m_strInputTemplate
      FUNCTION = SetWindowText(m_hCtl, m_str)
   ELSE
      m_str = m_strInputTemplate
      FUNCTION = SetWindowText(m_hCtl, m_str)
      IF dwsText <> m_strInputTemplate THEN
         FOR i AS LONG = 1 TO LEN(dwsText)
            this.OnCharPrintchar((ASC(MID(dwsText, i, 1))), 0, 0)
         NEXT
         Edit_SetSel(m_hCtl, 0, 0)
      END IF
   END IF
   DIM nPos AS LONG = Edit_GetSelStart(m_hCtl)
   IF nPos = 0 AND LEN(m_str) <> 0 THEN PostMessageW(m_hCtl, 9999, 0, 0)
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets masked text to the control.
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.SetText (BYREF dwsText AS DWSTRING, BYVAL bSetMaskedCharsOnly AS BOOLEAN) AS BOOLEAN
   DIM dws AS DWSTRING = dwsText
   IF bSetMaskedCharsOnly THEN 
      FOR i AS LONG = 1 TO LEN(m_strInputTemplate)
         IF MID(dws, i, 1) = MID(m_strInputTemplate, i, 1) THEN MID(dws, i, 1) = "_"
      NEXT
      dws = DWStrRemove(dws, "_")
   END IF
   FUNCTION = this.SetText(dws)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text of the control
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.GetText (BYVAL bGetMaskedCharsOnly AS BOOLEAN = FALSE) AS DWSTRING
   DIM dws AS DWSTRING = AfxGetWindowText(m_hCtl)
   IF bGetMaskedCharsOnly THEN
      FOR i AS LONG = 1 TO LEN(m_strInputTemplate)
         IF MID(dws, i, 1) = MID(m_strInputTemplate, i, 1) THEN MID(dws, i, 1) = "_"
      NEXT
      dws = DWStrRemove(dws, "_")
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the length of the text of the control
' ========================================================================================
PRIVATE FUNCTION CMaskedEdit.GetTextLength (BYVAL bGetMaskedCharsOnly AS BOOLEAN = FALSE) AS LONG
   DIM dws AS DWSTRING = AfxGetWindowText(m_hCtl)
   IF bGetMaskedCharsOnly THEN
      FOR i AS LONG = 1 TO LEN(m_strInputTemplate)
         IF MID(dws, i, 1) = MID(m_strInputTemplate, i, 1) THEN MID(dws, i, 1) = "_"
      NEXT
      dws = DWStrRemove(dws, "_")
   END IF
   RETURN LEN(dws)
END FUNCTION
' ========================================================================================

' ########################################################################################
'                          H E L P E R    F U N C T I O N S
' ########################################################################################

' ========================================================================================
' Retrieves a pointer to the CRichEditCtrl class from the handle of the RichEdit control.
' hMaskedEdit = Handle of the Masked Edit control.
' Usage example:
' DIM hMasKedEdit AS HWND = GetDlgItem(<parent handle>, <identifier of the control>)
' DIM pMaskedEdit AS CMaskedEdit PTR = AfxCMaskedEditPtr(hMasKedEdit))
' ========================================================================================
PRIVATE FUNCTION AfxCMaskedEditPtr OVERLOAD (BYVAL hMasKedEdit AS HWND) AS CMaskedEdit PTR
   DIM pMaskedEdit AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, SendMessageW(hMasKedEdit, WM_USER + 100, 0, 0))
   RETURN pMaskedEdit
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a pointer to the class given the handle of the parent window of the control and
' the identifier of the control.
' hParent = Handle of the parent window of the rich edit control.
' cID = Identifier of the rich edit control.
' Usage example:
' DIM pMaskedEdit AS CMaskedEdit PTR = AfxCMaskedEditPtr(<parent handle>, <identifier of the control>)
' ========================================================================================
PRIVATE FUNCTION AfxCMaskedEditPtr OVERLOAD (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CMaskedEdit PTR
   DIM pMaskedEdit AS CMaskedEdit PTR = CAST(CMaskedEdit PTR, SendMessageW(GetDlgItem(hParent, cID), WM_USER + 100, 0, 0))
   RETURN pMaskedEdit
END FUNCTION
' ========================================================================================

END NAMESPACE
