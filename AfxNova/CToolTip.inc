' ########################################################################################
' Platform: Microsoft Windows
' Filename: CToolTip.inc
' Purpose:  Windows common controls
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2026 José Roca
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
#include once "win/uxtheme.bi"
#include once "AfxNova/DWSTRING.inc"

NAMESPACE AfxNova

' ========================================================================================
' CToolTip class
' ========================================================================================
TYPE CToolTip

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC FUNCTION ComCtlVersion () AS LONG
   DECLARE STATIC SUB Activate (BYVAL hTooltip AS HWND, BYVAL fActivate AS BOOLEAN)
   DECLARE STATIC SUB Activate (BYVAL hTooltip AS HWND)
   DECLARE STATIC SUB Deactivate (BYVAL hTooltip AS HWND)
   DECLARE STATIC FUNCTION AddTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION AdjustRect (BYVAL hTooltip AS HWND, BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR) AS BOOLEAN
   DECLARE STATIC SUB DelTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE STATIC FUNCTION EnumTools (BYVAL hTooltip AS HWND, BYVAL iTool AS DWORD, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION GetBubbleSize (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS DWORD
   DECLARE STATIC FUNCTION GetCurrentTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION GetDelayTime (BYVAL hTooltip AS HWND, BYVAL flag AS DWORD) AS LONG
   DECLARE STATIC SUB GetMargin (BYVAL hTooltip AS HWND, BYVAL lprc AS RECT PTR)
   DECLARE STATIC FUNCTION GetMargin (BYVAL hTooltip AS HWND) AS RECT
   DECLARE STATIC FUNCTION GetMaxTipWidth (BYVAL hTooltip AS HWND) AS LONG
   DECLARE STATIC SUB GetText (BYVAL hTooltip AS HWND, BYVAL numChars AS DWORD, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE STATIC FUNCTION GetTipBkColor (BYVAL hTooltip AS HWND) AS COLORREF
   DECLARE STATIC FUNCTION GetTipTextColor (BYVAL hTooltip AS HWND) AS COLORREF
   DECLARE STATIC SUB GetTitle (BYVAL hTooltip AS HWND, BYVAL pTTGetTitle AS TTGETTITLE PTR)
   DECLARE STATIC FUNCTION GetToolCount (BYVAL hTooltip AS HWND) AS LONG
   DECLARE STATIC FUNCTION GetToolInfo (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION HitTest (BYVAL hTooltip AS HWND, BYVAL lphti AS TTHITTESTINFO PTR) AS BOOLEAN
   DECLARE STATIC SUB NewToolRect (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE STATIC SUB Pop (BYVAL hTooltip AS HWND)
   DECLARE STATIC SUB Popup (BYVAL hTooltip AS HWND)
   DECLARE STATIC FUNCTION RelayEvent (BYVAL hTooltip AS HWND, BYVAL lpmsg AS tagMSG PTR) AS LONG
   DECLARE STATIC SUB SetDelayTime (BYVAL hTooltip AS HWND, BYVAL dwDuration AS DWORD, BYVAL iTime AS SHORT)
   DECLARE STATIC SUB SetMargin (BYVAL hTooltip AS HWND, BYVAL lprc AS RECT PTR)
   DECLARE STATIC SUB SetMargin (BYVAL hTooltip AS HWND, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG)
   DECLARE STATIC FUNCTION SetMaxTipWidth (BYVAL hTooltip AS HWND, BYVAL iWidth AS LONG) AS LONG
   DECLARE STATIC SUB SetTipBkColor (BYVAL hTooltip AS HWND, BYVAL clr AS COLORREF)
   DECLARE STATIC SUB SetTipTextColor (BYVAL hTooltip AS HWND, BYVAL clr AS COLORREF)
   DECLARE STATIC FUNCTION SetTitle (BYVAL hTooltip AS HWND, BYVAL nIcon AS LONG, BYVAL pwszTitle AS WSTRING PTR) AS BOOLEAN
   DECLARE STATIC SUB SetToolInfo (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE STATIC SUB SetWindowTheme (BYVAL hTooltip AS HWND, BYVAL pwszTheme AS WSTRING PTR)
   DECLARE STATIC SUB SetDarkMode (BYVAL hTooltip AS HWND)
   DECLARE STATIC SUB TrackActivate (BYVAL hTooltip AS HWND, BYVAL bActivate AS BOOLEAN, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE STATIC SUB TrackPosition (BYVAL hTooltip AS HWND, BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   DECLARE STATIC SUB Update (BYVAL hTooltip AS HWND)
   DECLARE STATIC SUB UpdateTipText (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   DECLARE STATIC FUNCTION WindowFromPoint (BYVAL hTooltip AS HWND, BYVAL lppt AS POINT PTR) AS HWND

   DECLARE STATIC FUNCTION AddTooltip (BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING = "", BYVAL bBalloon AS BOOLEAN = FALSE, BYVAL bCentered AS BOOLEAN = FALSE) AS HWND
   DECLARE STATIC SUB SetTooltipText (BYVAL hTooltip AS HWND, BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING)
   DECLARE STATIC SUB DeleteTooltip (BYVAL hTooltip AS HWND, BYVAL hwnd AS HWND)

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CToolTip
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CToolTip
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Returns the version of ComCtrl.dll version multiplied by 100, e.g. 601 for version 6.01.
' Example: DIM ver AS LONG = AfxGetFileVersion("COMCTL32.DLL")
' ========================================================================================
PRIVATE FUNCTION CToolTip.ComCtlVersion () AS LONG
   DIM pvsffi AS VS_FIXEDFILEINFO PTR, dwHandle AS DWORD
   DIM cbLen AS DWORD = GetFileVersionInfoSizeW("COMCTL32.DLL", @dwHandle)
   IF cbLen = 0 THEN EXIT FUNCTION
   DIM pVerInfo AS HANDLE = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, cbLen)
   IF pVerInfo = NULL THEN EXIT FUNCTION
   IF GetFileVersionInfoW("COMCTL32.DLL", dwHandle, cbLen, pVerInfo) THEN
      IF VerQueryValueW(pVerInfo, "\", @pvsffi, @cbLen) THEN
         DIM wMajor AS WORD = HIWORD(pvsffi->dwFileVersionMS)
         DIM wMinor AS WORD = LOWORD(pvsffi->dwFileVersionMS)
         FUNCTION = (wMajor + wMinor / 100) * 100
      END IF
   END IF
   HeapFree(GetProcessHeap, 0, pVerInfo)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Activates or deactivates a ToolTip control.
' ========================================================================================
PRIVATE SUB CToolTip.Activate (BYVAL hTooltip AS HWND, BYVAL fActivate AS BOOLEAN)
   SendMessageW(hTooltip, TTM_ACTIVATE, fActivate, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CToolTip.Activate (BYVAL hTooltip AS HWND)
   SendMessageW(hTooltip, TTM_ACTIVATE, TRUE, 0)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CToolTip.Deactivate (BYVAL hTooltip AS HWND)
   SendMessageW(hTooltip, TTM_ACTIVATE, FALSE, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Registers a tool with a ToolTip control.
' See the AddTooltip method.
' ========================================================================================
PRIVATE FUNCTION CToolTip.AddTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   IF lpti = NULL THEN RETURN FALSE
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   RETURN SendMessageW(hTooltip, TTM_ADDTOOLW, 0, CAST(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates a ToolTip control's text display rectangle from its window rectangle, or the
' ToolTip window rectangle needed to display a specified text display rectangle.
' ========================================================================================
PRIVATE FUNCTION CToolTip.AdjustRect (BYVAL hTooltip AS HWND, BYVAL fLarger AS BOOLEAN, BYVAL lprc AS RECT PTR) AS BOOLEAN
   RETURN SendMessageW(hTooltip, TTM_ADJUSTRECT, cast(WPARAM, fLarger), cast(LPARAM, lprc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes a tool from a ToolTip control.
' Seet the DeleteToolTip method.
' ========================================================================================
PRIVATE SUB CToolTip.DelTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   IF lpti = NULL THEN EXIT SUB
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   SendMessageW(hTooltip, TTM_DELTOOLW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the information that a ToolTip control maintains about the current tool--that
' is, the tool for which the ToolTip is currently displaying text.
' ========================================================================================
PRIVATE FUNCTION CToolTip.EnumTools (BYVAL hTooltip AS HWND, BYVAL iTool AS DWORD, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   IF lpti = NULL THEN RETURN FALSE
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   RETURN SendMessageW(hTooltip, TTM_ENUMTOOLSW, cast(WPARAM, iTool), cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the width and height of a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetBubbleSize (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS DWORD
   IF lpti = NULL THEN RETURN 0
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   RETURN SendMessageW(hTooltip, TTM_GETBUBBLESIZE, 0, cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the information for the current tool in a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetCurrentTool (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   IF lpti = NULL THEN RETURN FALSE
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   RETURN SendMessageW(hTooltip, TTM_GETCURRENTTOOLW, 0, cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the information for the current tool in a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetDelayTime (BYVAL hTooltip AS HWND, BYVAL flag AS DWORD) AS LONG
   RETURN SendMessageW(hTooltip, TTM_GETDELAYTIME, cast(WPARAM, flag), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the top, left, bottom, and right margins set for a ToolTip window. A margin is
' the distance, in pixels, between the ToolTip window border and the text contained within
' the ToolTip window.
' ========================================================================================
PRIVATE SUB CToolTip.GetMargin (BYVAL hTooltip AS HWND, BYVAL lprc AS RECT PTR)
   SendMessageW(hTooltip, TTM_GETMARGIN, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetMargin (BYVAL hTooltip AS HWND) AS RECT
   DIM rc AS RECT
   SendMessageW(hTooltip, TTM_GETMARGIN, 0, cast(LPARAM, @rc))
   RETURN rc
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the maximum width for a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetMaxTipWidth (BYVAL hTooltip AS HWND) AS LONG
   RETURN SendMessageW(hTooltip, TTM_GETMAXTIPWIDTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the information a ToolTip control maintains about a tool.
' ========================================================================================
PRIVATE SUB CToolTip.GetText (BYVAL hTooltip AS HWND, BYVAL numChars AS DWORD, BYVAL lpti AS TOOLINFOW PTR)
   IF lpti = NULL THEN EXIT SUB
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   SendMessageW(hTooltip, TTM_GETTEXTW, numChars, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' * Retrieves the background color in a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetTipBkColor (BYVAL hTooltip AS HWND) AS COLORREF
   RETURN SendMessageW(hTooltip, TTM_GETTIPBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the text color in a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetTipTextColor (BYVAL hTooltip AS HWND) AS COLORREF
   RETURN SendMessageW(hTooltip, TTM_GETTIPTEXTCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves information concerning the title of a tooltip control.
' ========================================================================================
PRIVATE SUB CToolTip.GetTitle (BYVAL hTooltip AS HWND, BYVAL pTTGetTitle AS TTGETTITLE PTR)
   SendMessageW(hTooltip, TTM_GETTITLE, 0, cast(LPARAM, pTTGetTitle))
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves a count of the tools maintained by a ToolTip control.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetToolCount (BYVAL hTooltip AS HWND) AS LONG
   RETURN SendMessageW(hTooltip, TTM_GETTOOLCOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the information that a ToolTip control maintains about a tool.
' ========================================================================================
PRIVATE FUNCTION CToolTip.GetToolInfo (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR) AS BOOLEAN
   IF lpti = NULL THEN RETURN FALSE
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   RETURN SendMessageW(hTooltip, TTM_GETTOOLINFOW, 0, cast(LPARAM, lpti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tests a point to determine whether it is within the bounding rectangle of the specified
' tool and, if it is, retrieves information about the tool.
' ========================================================================================
PRIVATE FUNCTION CToolTip.HitTest (BYVAL hTooltip AS HWND, BYVAL lphti AS TTHITTESTINFO PTR) AS BOOLEAN
   RETURN SendMessageW(hTooltip, TTM_HITTEST, 0, cast(LPARAM, lphti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a new bounding rectangle for a tool.
' ========================================================================================
PRIVATE SUB CToolTip.NewToolRect (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   IF lpti = NULL THEN EXIT SUB
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   SendMessageW(hTooltip, TTM_NEWTOOLRECTW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Removes a displayed ToolTip window from view.
' ========================================================================================
PRIVATE SUB CToolTip.Pop (BYVAL hTooltip AS HWND)
   SendMessageW hTooltip, TTM_POP, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Causes the ToolTip to display at the coordinates of the last mouse message.
' ========================================================================================
PRIVATE SUB CToolTip.Popup (BYVAL hTooltip AS HWND)
   SendMessageW hTooltip, TTM_POPUP, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Passes a mouse message to a ToolTip control for processing.
' ========================================================================================
PRIVATE FUNCTION CToolTip.RelayEvent (BYVAL hTooltip AS HWND, BYVAL lpmsg AS tagMSG PTR) AS LONG
   RETURN SendMessageW(hTooltip, TTM_RELAYEVENT, 0, cast(LPARAM, lpmsg))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the initial, pop-up, and reshow durations for a ToolTip control.
' ========================================================================================
PRIVATE SUB CToolTip.SetDelayTime (BYVAL hTooltip AS HWND, BYVAL dwDuration AS DWORD, BYVAL iTime AS SHORT)
   SendMessageW(hTooltip, TTM_SETDELAYTIME, dwDuration, MAKELONG(iTime, 0))
END SUB
' ========================================================================================

' ========================================================================================
' * Sets the top, left, bottom, and right margins for a ToolTip window. A margin is the
' distance, in pixels, between the ToolTip window border and the text contained within the
' ToolTip window.
' ========================================================================================
PRIVATE SUB CToolTip.SetMargin (BYVAL hTooltip AS HWND, BYVAL lprc AS RECT PTR)
   SendMessageW(hTooltip, TTM_SETMARGIN, 0, cast(LPARAM, lprc))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CToolTip.SetMargin (BYVAL hTooltip AS HWND, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG)
   DIM rc AS RECT = (nLeft, nTop, nRight, nBottom)
   SendMessageW(hTooltip, TTM_SETMARGIN, 0, cast(LPARAM, @rc))
END SUB
' ========================================================================================

' ========================================================================================
' * Sets the maximum width for a ToolTip window.
' ========================================================================================
PRIVATE FUNCTION CToolTip.SetMaxTipWidth (BYVAL hTooltip AS HWND, BYVAL iWidth AS LONG) AS LONG
   RETURN SendMessageW(hTooltip, TTM_SETMAXTIPWIDTH, 0, cast(LPARAM, iWidth))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the background color in a ToolTip window.
' Does nothing when using visual styles.
' ========================================================================================
PRIVATE SUB CToolTip.SetTipBkColor (BYVAL hTooltip AS HWND, BYVAL clr AS COLORREF)
   SendMessageW(hTooltip, TTM_SETTIPBKCOLOR, cast(WPARAM, clr), 0)
END SUB
' ========================================================================================

' ========================================================================================
' * Sets the text color in a ToolTip window.
' Does nothing when using visual styles.
' ========================================================================================
PRIVATE SUB CToolTip.SetTipTextColor (BYVAL hTooltip AS HWND, BYVAL clr AS COLORREF)
   SendMessageW(hTooltip, TTM_SETTIPTEXTCOLOR, cast(WPARAM, clr), 0)
END SUB
' ========================================================================================

' ========================================================================================
' * Adds a standard icon and title string to a ToolTip.
' If you use visual styles, you will get a bigger caption size that if you don't, and you
' can also use the large predefined icons.
' ========================================================================================
PRIVATE FUNCTION CToolTip.SetTitle (BYVAL hTooltip AS HWND, BYVAL nIcon AS LONG, BYVAL pwszTitle AS WSTRING PTR) AS BOOLEAN
   RETURN SendMessageW(hTooltip, TTM_SETTITLEW, nIcon, cast(LPARAM, pwszTitle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the information that a ToolTip control maintains for a tool.
' ========================================================================================
PRIVATE SUB CToolTip.SetToolInfo (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   IF lpti = NULL THEN EXIT SUB
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   SendMessageW(hTooltip, TTM_SETTOOLINFOW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the visual style of a ToolTip control.
' ========================================================================================
PRIVATE SUB CToolTip.SetWindowTheme (BYVAL hTooltip AS HWND, BYVAL pwszTheme AS WSTRING PTR)
   SendMessageW(hTooltip, TTM_SETWINDOWTHEME, 0, cast(LPARAM, pwszTheme))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the dark mode visual style of a ToolTip control.
' You need to use visual styles.
' ========================================================================================
PRIVATE SUB CToolTip.SetDarkMode (BYVAL hTooltip AS HWND)
   DIM wszTheme AS WSTRING * 20 = "DarkMode_Explorer"
   SendMessageW(hTooltip, TTM_SETWINDOWTHEME, 0, cast(LPARAM, @wszTheme))
END SUB
' ========================================================================================

' ========================================================================================
' Activates or deactivates a tracking ToolTip.
' ========================================================================================
PRIVATE SUB CToolTip.TrackActivate (BYVAL hTooltip AS HWND, BYVAL bActivate AS BOOLEAN, BYVAL lpti AS TOOLINFOW PTR)
   IF lpti = NULL THEN EXIT SUB
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   SendMessageW(hTooltip, TTM_TRACKACTIVATE, cast(WPARAM, bActivate), cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the position of a tracking tooltip.
' ========================================================================================
PRIVATE SUB CToolTip.TrackPosition (BYVAL hTooltip AS HWND, BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   SendMessageW(hTooltip, TTM_TRACKPOSITION, 0, cast(LPARAM, MAKELONG(xPos, yPos)))
END SUB
' ========================================================================================

' ========================================================================================
' Forces the current tooltip to be redrawn.
' ========================================================================================
PRIVATE SUB CToolTip.Update (BYVAL hTooltip AS HWND)
   SendMessageW(hTooltip, TTM_UPDATE, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the ToolTip text for a tool.
' ========================================================================================
PRIVATE SUB CToolTip.UpdateTipText (BYVAL hTooltip AS HWND, BYVAL lpti AS TOOLINFOW PTR)
   IF lpti = NULL THEN EXIT SUB
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 64 ELSE lpti->cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN lpti->cbSize = 44 ELSE lpti->cbSize = 48
#endif
   SendMessageW(hTooltip, TTM_UPDATETIPTEXTW, 0, cast(LPARAM, lpti))
END SUB
' ========================================================================================

' ========================================================================================
' * Creates a tooltip for a control.
' Parameters:
' - hwnd      = Handle of the window or control
' - wszText   = Tooltip text
' - bBalloon  = Ballon tip (TRUE or FALSE)
' - bCentered = Centered (TRUE or FALSE)
' Return Value:
'   The handle of the tooltip control
' ========================================================================================
PRIVATE FUNCTION CToolTip.AddTooltip (BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING = "", BYVAL bBalloon AS BOOLEAN = FALSE, BYVAL bCentered AS BOOLEAN = FALSE) AS HWND
   IF IsWindow(hwnd) = 0 THEN EXIT FUNCTION
   ' // Creates the tooltip control
   DIM dwStyle AS DWORD = WS_POPUP OR TTS_ALWAYSTIP
   IF bBalloon THEN dwStyle = dwStyle OR TTS_BALLOON
   DIM hTooltip AS ..HWND
   hTooltip = CreateWindowExW(0, "tooltips_class32", "", dwStyle, 0, 0, 0, 0, NULL, NULL, NULL, NULL)
   IF hTooltip = NULL THEN EXIT FUNCTION
   ' // You must explicitly define a tooltip control as topmost. Otherwise, it might be covered by the parent window.
   SetWindowPos(hTooltip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE OR SWP_NOACTIVATE)
   ' // Registers the window with the tooltip control
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.uFlags = TTF_IDISHWND OR TTF_SUBCLASS
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.uFlags = TTF_SUBCLASS
      tti.hwnd = hwnd
      tti.uId = 0
      GetClientRect(hwnd, @tti.rect)
   END IF
   IF bCentered THEN tti.uFlags = tti.uFlags OR TTF_CENTERTIP
   tti.hinst = GetModuleHandleW(NULL)
   tti.lpszText = CAST(LPWSTR, @wszText)
   SendMessageW(hTooltip, TTM_ADDTOOLW, 0, CAST(LPARAM, @tti))
   RETURN hTooltip
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets/replaces the text of a tooltip control
' Parameters:
' - hTooltip = Handle of the tooltip control
' - hwnd     = Handle of the window or control
' - wszText  = Tooltip text
' ========================================================================================
PRIVATE SUB CToolTip.SetTooltipText (BYVAL hTooltip AS HWND, BYVAL hwnd AS HWND, BYREF wszText AS CONST WSTRING)
   IF hTooltip = NULL OR hwnd = NULL THEN EXIT SUB
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.hwnd = hwnd
      tti.uId = 0
   END IF
   ' // Retrieve the tooltip information
   SendMessageW(hTooltip, TTM_GETTOOLINFOW, 0, CAST(LPARAM, @tti))
   IF SendMessageW(hTooltip, TTM_GETTOOLINFOW, 0, CAST(LPARAM, @tti)) THEN
      tti.lpszText = CAST(LPWSTR, @wszText)
      SendMessageW(hTooltip, TTM_SETTOOLINFOW, 0, CAST(LPARAM, @tti))
   END IF
   ' // Note: Windows provides the TTM_UPDATETIPTEXT message, but needs Vista or superior.
END SUB
' ========================================================================================

' ========================================================================================
' * Removes a tool from a tooltip control.
' Parameters:
' - hTooltip = Handle of the tooltip control
' - hwnd     = Handle of the window or control
' ========================================================================================
PRIVATE SUB CToolTip.DeleteTooltip (BYVAL hTooltip AS HWND, BYVAL hwnd AS HWND)
   IF hTooltip = NULL OR hwnd = NULL THEN EXIT SUB
   ' // 32-bit: The size of the TOOLINFOW structure is 48 bytes in
   '    version 6 of comctl32.dll, and 44 bytes in lower versions.
   ' // 64-bit: The size of the TOOLINFOW structure is 72 bytes in
   '    version 6 of comctl32.dll, and 64 bytes in lower versions.
   DIM tti AS TOOLINFOW
#ifdef __FB_64BIT__
   IF CToolTip.ComCtlVersion < 600 THEN tti.cbSize = 64 ELSE tti.cbSize = 72
#else
   IF CToolTip.ComCtlVersion < 600 THEN tti.cbSize = 44 ELSE tti.cbSize = 48
#endif
   IF (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD) = WS_CHILD THEN
      tti.hwnd = GetParent(hwnd)
      tti.uId = CAST(UINT_PTR, hwnd)
   ELSE
      tti.hwnd = hwnd
      tti.uId = 0
   END IF
   ' // Remove the tooltip
   SendMessageW(hTooltip, TTM_DELTOOL, 0, CAST(LPARAM, @tti))
END SUB
' ========================================================================================

END NAMESPACE
