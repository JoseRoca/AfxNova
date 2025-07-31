' ########################################################################################
' Platform: Microsoft Windows
' Filename: CTabPage.inc
' Contents: Wrapper class to create generic dialogs to be used as pages of a tab control
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

#include once "AfxNova/CWindow.inc"


' ########################################################################################
'                                 *** CTABPAGE CLASS ***
' ########################################################################################

' =====================================================================================
' CTabPage class
' Creates a generic window used as a tab page of a tab control.
' =====================================================================================
TYPE CTabPage EXTENDS CWindow

   Private:
      m_hTabPage AS HWND                     ' // Tab page handle

   Public:
      DECLARE DESTRUCTOR
      DECLARE PROPERTY hTabPage () AS HWND
      DECLARE FUNCTION InsertPage (BYVAL hTab AS HWND, BYVAL nPage AS LONG, BYREF wszTitle AS WSTRING = "", _
         BYVAL nImage AS LONG = -1, BYVAL lpfnWndProc AS WNDPROC = NULL, _
         BYVAL dwStyle AS DWORD = WS_CLIPSIBLINGS OR WS_CLIPCHILDREN, BYVAL dwExStyle AS DWORD = 0) AS HWND
      DECLARE FUNCTION ResizePage (BYVAL hTab AS HWND = NULL) AS BOOLEAN
      DECLARE FUNCTION ResizePages (BYVAL hTab AS HWND = NULL) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' CTabPage class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTabPage

   DIM i AS LONG, nCount AS LONG, hTab AS HWND, tci AS TCITEMW
   ' // Get the handle of the tab control
   hTab = ..GetParent(m_hTabPage)
   IF hTab THEN
      ' // Get the number of items
      nCount = SendMessageW(hTab, TCM_GETITEMCOUNT, 0, 0)
      ' // Ask to return the value of the lParam member
      tci.mask = TCIF_PARAM
      ' // Get information of the items
      FOR i = 0 TO nCount - 1
         IF SendMessageW(hTab, TCM_GETITEMW, i, CAST(lParam, @tci)) THEN
            IF tci.lParam = @this THEN
               ' // Delete the tab item
               SendMessageW(hTab, TCM_DELETEITEM, i, 0)
               ' // Destroy the window of the tab page
               ..DestroyWindow m_hTabPage
               ' // Set the focus in the first tab
               IF nCount > 1 THEN .SendMessageW hTab, TCM_SETCURFOCUS, 0, 0
               EXIT FOR
            END IF
         END IF
      NEXT
   ELSE
      ' // Destroy the window of the tab page
      ..DestroyWindow m_hTabPage
   END IF

END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Adds a tab page and creates a generic window that will be associated with the page.
' Parameters:
' - hTab        = [in] A handle to the tab control.
' - nPage       = [in] The zero based position of the page to be inserted.
' - wszTitle    = [in, opt] The text to be displayed of the tab area.
' - nImage      = [in, opt] Index in the tab control's image list, or -1 if there is no image for the tab.
' - lpfnWndProc = [in, opt] Address of the window callback procedure.
' Return Value:
' - The handle of the new tab page.
' ========================================================================================
PRIVATE FUNCTION CTabPage.InsertPage (BYVAL hTab AS HWND, BYVAL nPage AS LONG, BYREF wszTitle AS WSTRING = "", _
   BYVAL nImage AS LONG = -1, BYVAL lpfnWndProc AS WNDPROC = NULL, _
   BYVAL dwStyle AS DWORD = WS_CLIPSIBLINGS OR WS_CLIPCHILDREN, BYVAL dwExStyle AS DWORD = 0) AS HWND

   DIM AS LONG x, y, nWidth, nHeight, cItems
   DIM rc AS RECT, tci AS TCITEMW

   IF IsWindow(hTab) = NULL THEN EXIT FUNCTION
   IF lpfnWndProc = NULL THEN lpfnWndProc = CAST(HANDLE, .GetClassLongPtrW(this.hWindow, GWLP_WNDPROC))

   dwStyle = dwStyle OR WS_CHILD
   dwExStyle = dwExStyle OR WS_EX_CONTROLPARENT
   this.ClassStyle = CS_DBLCLKS

   cItems = .SendMessageW(hTab, TCM_GETITEMCOUNT, 0, 0)
   IF nPage < 0 OR nPage > cItems THEN nPage = cItems
   tci.mask    = TCIF_TEXT OR TCIF_IMAGE OR TCIF_PARAM
   tci.pszText = @wszTitle
   tci.iImage  = nImage
   tci.lParam  = CAST(lParam, @this)
   SendMessageW hTab, TCM_INSERTITEMW, nPage, CAST(lParam, @tci)
   ..GetWindowRect(hTab, @rc)
   ..SendMessageW hTab, TCM_ADJUSTRECT, FALSE, CAST(lParam, @rc)
   ..MapWindowPoints NULL, hTab, CAST(LPPOINT, @rc), 2
   ' // Adjust for High DPI because create will resize the values
   rc.Left   /= this.rxRatio
   rc.Right  /= this.rxRatio
   rc.Top    /= this.ryRatio
   rc.Bottom /= this.ryRatio
   ' // Calculate coordinates and size
   x = rc.Left
   y = rc.Top
   nWidth  = max(1, rc.Right - rc.Left)
   nHeight = max(1, rc.Bottom - rc.Top)
   m_hTabPage = this.Create(hTab, wszTitle, lpfnWndProc, x, y, nWidth, nHeight, dwStyle, dwExStyle)
   FUNCTION = m_hTabPage

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle of the tab page
' ========================================================================================
PRIVATE PROPERTY CTabPage.hTabPage () AS HWND
   PROPERTY = m_hTabPage
END PROPERTY
' ========================================================================================

' =====================================================================================
' Resizes all the tab pages associated with a tab control
' =====================================================================================
PRIVATE FUNCTION CTabPage.ResizePage (BYVAL hTab AS HWND = NULL) AS BOOLEAN
   DIM pTabPage AS CTabPage PTR
   ' // Ask to return the value of the lParam member
   DIM tci AS TCITEMW
   tci.mask = TCIF_PARAM
   ' // Get the handle of the tab control
   IF hTab = NULL THEN hTab = ..GetParent(m_hTabPage)
   IF hTab = NULL THEN RETURN FALSE
   ' // Get information of the item
   DIM curTab AS LONG = TabCtrl_GetCurSel(hTab)
   ' // Rseize the tab page
   IF .SendMessageW(hTab, TCM_GETITEMW, curTab, CAST(LPARAM, @tci)) THEN
      IF tci.lParam THEN
         pTabPage = CAST(CTabPage PTR, tci.lParam)
         ' // Retrieve the size of the tab control window
         DIM rcParent AS RECT
         ..GetWindowRect(hTab, @rcParent)
         ' // Calculates the tab control's display area given its window rectangle
         SendMessageW(hTab, TCM_ADJUSTRECT, FALSE, CAST(LPARAM, @rcParent))
         ' // Convert to window coordinates
         ..MapWindowPoints(NULL, hTab, CAST(LPPOINT, @rcParent), 2)
         ' // Move the tab page
         ..MoveWindow(pTabPage->hTabPage, rcParent.Left, rcParent.Top, _
            rcParent.Right - rcParent.Left, rcParent.Bottom - rcParent.Top, CTRUE)
      END IF
   END IF
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' =====================================================================================
' Resizes all the tab pages associated with a tab control
' =====================================================================================
PRIVATE FUNCTION CTabPage.ResizePages (BYVAL hTab AS HWND = NULL) AS BOOLEAN
   ' // Get the handle of the tab control
   IF hTab = NULL THEN hTab = ..GetParent(m_hTabPage)
   IF hTab = NULL THEN RETURN FALSE
   ' // Get the number of items
   DIM nCount AS LONG = .SendMessageW(hTab, TCM_GETITEMCOUNT, 0, 0)
   IF nCount = 0 THEN EXIT FUNCTION
   ' // Ask to return the value of the lParam member
   DIM tci AS TCITEMW
   tci.mask = TCIF_PARAM
   ' // Get information of the items
   DIM pTabPage AS CTabPage PTR
   FOR i AS LONG = 0 TO nCount - 1
      IF SendMessageW(hTab, TCM_GETITEMW, i, CAST(LPARAM, @tci)) THEN
         IF tci.lParam THEN
            pTabPage = CAST(CTabPage PTR, tci.lParam)
            ' // Retrieve the size of the tab control window
            DIM rcParent AS RECT
            ..GetWindowRect(hTab, @rcParent)
            ' // Calculates the tab control's display area given its window rectangle
            SendMessageW(hTab, TCM_ADJUSTRECT, FALSE, CAST(LPARAM, @rcParent))
            ' // Convert to window coordinates
            ..MapWindowPoints(NULL, hTab, CAST(LPPOINT, @rcParent), 2)
            ' // Move the tab page
            ..MoveWindow(pTabPage->hTabPage, rcParent.Left, rcParent.Top, _
               rcParent.Right - rcParent.Left, rcParent.Bottom - rcParent.Top, CTRUE)
         END IF
      END IF
   NEXT
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                *** HELPER FUNCTIONS ***
' ########################################################################################

' =====================================================================================
' Resizes all the tab pages associated with a tab control
' =====================================================================================
PRIVATE FUNCTION AfxResizeTabPages (BYVAL hTab AS HWND) AS BOOLEAN
   IF hTab = NULL THEN EXIT FUNCTION
   DIM nCount AS LONG, i AS LONG, tci AS TCITEMW, pTabPage AS CTabPage PTR
   ' // Get the number of items
   nCount = .SendMessageW(hTab, TCM_GETITEMCOUNT, 0, 0)
   IF nCount = 0 THEN EXIT FUNCTION
   ' // Ask to return the value of the lParam member
   tci.mask = TCIF_PARAM
   ' // Get information of the items
   FOR i = 0 TO nCount - 1
      IF .SendMessageW(hTab, TCM_GETITEMW, i, CAST(LPARAM, @tci)) THEN
         IF tci.lParam THEN
            pTabPage = CAST(CTabPage PTR, tci.lParam)
            ' // Retrieve the size of the tab control window
            DIM rcParent AS RECT
            ..GetWindowRect(hTab, @rcParent)
            ' // Calculates the tab control's display area given its window rectangle
            SendMessageW(hTab, TCM_ADJUSTRECT, FALSE, CAST(LPARAM, @rcParent))
            ' // Convert to window coordinates
            ..MapWindowPoints(NULL, hTab, CAST(LPPOINT, @rcParent), 2)
            ' // Move the tab page
            ..MoveWindow(pTabPage->hTabPage, rcParent.Left, rcParent.Top, _
               rcParent.Right - rcParent.Left, rcParent.Bottom - rcParent.Top, CTRUE)
         END IF
      END IF
   NEXT
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detroys the specified tab page
' Parameters:
' - hTab = Handle to the tab control.
' - idx  = Zero based index of the tab.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxDestroyTabPage (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS BOOLEAN
   IF hTab = NULL THEN EXIT FUNCTION
   DIM tci AS TCITEMW, pTabPage AS CTabPage PTR
   ' // Ask to return the value of the lParam member
   tci.mask = TCIF_PARAM
   IF SendMessageW(hTab, TCM_GETITEMW, idx, CAST(LPARAM, @tci)) = 0 THEN EXIT FUNCTION
   IF tci.lParam THEN
      pTabPage = CAST(CTabPage PTR, tci.lParam)
      IF pTabPage THEN
         Delete pTabPage
         FUNCTION = TRUE
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detroys all the tab pages
' Parameter:
' - hTab = Handle to the tab control.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxDestroyAllTabPages (BYVAL hTab AS HWND) AS BOOLEAN
   IF hTab = NULL THEN EXIT FUNCTION
   DIM tci AS TCITEMW, pTabPage AS CTabPage PTR
   ' // Ask to return the value of the lParam member
   tci.mask = TCIF_PARAM
   DO
      IF SendMessageW(hTab, TCM_GETITEMW, 0, CAST(LPARAM, @tci)) = 0 THEN EXIT DO
      IF tci.lParam THEN
         pTabPage = CAST(CTabPage PTR, tci.lParam)
         IF pTabPage THEN
            Delete pTabPage
            FUNCTION = TRUE
         END IF
      END IF
   LOOP
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CTabPage class given the handle of the tab control to which the
' tab page is associated and the zero-based tab index. If nTabIdx is ommited, the function
' will return the pointer of the selected tab, if any.
' ========================================================================================
PRIVATE FUNCTION AfxCTabPagePtr (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS CTabPage PTR
   IF hTab = NULL THEN EXIT FUNCTION
   IF idx = -1 THEN idx = SendMessageW(hTab, TCM_GETCURSEL, 0, 0)
   IF idx = -1 THEN EXIT FUNCTION   ' No tab selected
   ' // Ask to return the value of the lParam member
   DIM tci AS TCITEMW
   tci.mask = TCIF_PARAM
   IF .SendMessageW(hTab, TCM_GETITEMW, idx, CAST(LPARAM, @tci)) THEN
      IF tci.lParam THEN
         FUNCTION = CAST(CTabPage PTR, tci.lParam)
      END IF
   END IF
END FUNCTION
' ========================================================================================
