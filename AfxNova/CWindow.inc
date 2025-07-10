' ########################################################################################
' Platform: Microsoft Windows
' Filename: CWindow.inc
' Contents: A wrapper class to create a SDK main window and add controls to it.
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
#include once "AfxNova/AfxCtl.inc"
#include once "AfxNova/AfxWin.inc"
USING AfxNova

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWINDOW_DEBUG_ 1 in your application before including this file.
' To capture and display the messages sent by the Windows function OutputDebugStringW, you
' can use the DebugView tool. See: https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
' ========================================================================================
#ifndef _CWINDOW_DEBUG_
   #define _CWINDOW_DEBUG_ 0
#ENDIF
#ifndef _CWINDOW_DP_
   #define _CWINDOW_DP_ 1
   #MACRO CWINDOW_DP(st)
      #IF (_CWINDOW_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' // For compatibility with the Power Basic syntax
#define CBHNDL(wp, lp) cast(HWND, lp)
#define CBCTL(wp, lp) LOWORD(wp)
#define CBCTLMSG(wp, lp) HIWORD(wp)
#define CBNMCODE(wp, lp) LOWORD(wp) 
#define CBNMHWND(wp, lp) cast(HWND, lp)
#define CBNMID(wp, lp) HIWORD(wp)
#define CBNMHDR(wp, lp) cast(NMHDR PTR, lp)
' // For cast notification messages, e.g.
' DIM p AS NMDATETIMECHANGE PTR = CBNMSGPTR(NMDATETIMECHANGE, wParam, lParam)
' print p->nmhdr.idfrom
' print p->dwFlags
#define CBNMSGPTR(msg, wp, lp) cast(msg PTR, CBNMHDR(wp, lp))
' // To copy notification messages, e.g.
' DIM dt AS NMDATETIMECHANGE
' CBNMTYPESET(dt, wParam, lParam)
' print dt.nmhdr.idfrom
' print dt.dwFlags
#define CBNMTYPESET(tp, wp, lp) memcpy @tp, CAST(ANY PTR, lp), SIZEOF(tp)

NAMESPACE AfxNova

' // Layout manager - Anchor flags
ENUM AFX_ANCHORPOINT
   AFX_ANCHOR_NONE =  0
   AFX_ANCHOR_WIDTH
   AFX_ANCHOR_RIGHT
   AFX_ANCHOR_CENTER_HORZ
   AFX_ANCHOR_HEIGHT
   AFX_ANCHOR_HEIGHT_WIDTH
   AFX_ANCHOR_HEIGHT_RIGHT
   AFX_ANCHOR_BOTTOM
   AFX_ANCHOR_BOTTOM_WIDTH
   AFX_ANCHOR_BOTTOM_RIGHT
   AFX_ANCHOR_CENTER_HORZ_BOTTOM
   AFX_ANCHOR_CENTER_VERT
   AFX_ANCHOR_CENTER_VERT_RIGHT
   AFX_ANCHOR_CENTER
END ENUM

' // Layout manager - Anchor properties
TYPE AFX_ANCHORPROPERTY
   hwndCtl AS HWND
   anchor  AS LONG
   rc      AS RECT
   centerX AS LONG
   centerY AS LONG
END TYPE

' // Controls user data
TYPE AFX_CTLUSERDATA
   hCtl AS HWND
   Values (0 TO 9) AS LONG_PTR
END TYPE

' ========================================================================================
' CWindow class
' ========================================================================================
TYPE CWindow

Private:
   m_hwnd AS HWND                                  ' // Window handle
   m_hInstance AS HINSTANCE                        ' // Instance handle
   m_pUserWndProc AS WNDPROC                       ' // User window procedure
   m_hFont AS HFONT                                ' // Default font handle
   m_hAccel AS HACCEL                              ' // Accelerator table handle
   m_DPI AS LONG = 96                              ' // Design-time DPI
   m_rx AS SINGLE = 1                              ' // Horizontal scaling ratio
   m_ry AS SINGLE = 1                              ' // Vertical scaling ratio
   m_wAtom AS ATOM                                 ' // Class atom
   m_wszClassName AS WSTRING * 260                 ' // Class name
   m_DefaultFontSize AS LONG                       ' // Default font size
   m_wszDefaultFontName AS WSTRING * LF_FACESIZE   ' // Default font name
   m_hRichEditLib AS HMODULE                       ' // Rich Edit moudle handle
'   m_pScrollWindow AS CScrollWindow PTR            ' // Pointer to the scroll window class
   DIM m_rgUserData(0 TO 99) AS LONG_PTR           ' // User data
   DIM m_rgAccelEntries(ANY) AS ACCEL              ' // Accelerator table
   DIM m_rgProps(ANY) AS AFX_ANCHORPROPERTY        ' // Array of anchor property structures
   DIM m_rgCtrlUserData(ANY) AS AFX_CTLUSERDATA    ' // Control's user data

Public:
   DECLARE CONSTRUCTOR (BYREF wszClassName AS CONST WSTRING = "")
   DECLARE DESTRUCTOR
   DECLARE PROPERTY DPI () AS SINGLE
   DECLARE PROPERTY DPI (BYVAL dpi AS SINGLE)
   DECLARE PROPERTY UserData (BYVAL idx AS LONG) AS LONG_PTR
   DECLARE PROPERTY UserData (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
   DECLARE FUNCTION DoEvents (BYVAL nCmdShow AS LONG = 0) AS LONG
   DECLARE FUNCTION Create (BYVAL hParent AS HWND = NULL, BYREF wszTitle AS WSTRING = "", BYVAL lpfnWndProc AS WNDPROC = NULL, _
      BYVAL x AS LONG = CW_USEDEFAULT, BYVAL y AS LONG = CW_USEDEFAULT, BYVAL nWidth AS LONG = CW_USEDEFAULT, BYVAL nHeight AS LONG = CW_USEDEFAULT, _
      BYVAL dwStyle AS DWORD = WS_OVERLAPPEDWINDOW OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS, _
      BYVAL dwExStyle AS DWORD = WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE) AS HWND
   DECLARE FUNCTION CreateOverlapped (BYVAL hParent AS HWND = NULL, BYREF wszTitle AS WSTRING = "", BYVAL lpfnWndProc AS WNDPROC = NULL, _
      BYVAL x AS LONG = CW_USEDEFAULT, BYVAL y AS LONG = CW_USEDEFAULT, BYVAL nWidth AS LONG = CW_USEDEFAULT, BYVAL nHeight AS LONG = CW_USEDEFAULT, _
      BYVAL dwExStyle AS UINT = WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE) AS HWND
   DECLARE SUB SetClientSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
   DECLARE SUB Center (BYVAL hwnd AS HWND = NULL, BYVAL hwndParent AS HWND = NULL)
   DECLARE FUNCTION GetWindowProc () AS WNDPROC
   DECLARE PROPERTY hWindow () AS HWND
   DECLARE PROPERTY hWindow (BYVAL hwnd AS HWND)
   DECLARE PROPERTY InstanceHandle () AS HINSTANCE
   DECLARE PROPERTY InstanceHandle (BYVAL hInst AS HINSTANCE)
   DECLARE PROPERTY ClassStyle () AS ULONG_PTR
   DECLARE PROPERTY ClassStyle (BYVAL dwStyle AS ULONG_PTR)
   DECLARE PROPERTY WindowStyle () AS ULONG_PTR
   DECLARE PROPERTY WindowStyle (BYVAL dwStyle AS ULONG_PTR)
   DECLARE PROPERTY WindowExStyle () AS ULONG_PTR
   DECLARE PROPERTY WindowExStyle (BYVAL dwExStyle AS ULONG_PTR)
   DECLARE PROPERTY DefaultFontName () AS STRING
   DECLARE PROPERTY DefaultFontName (BYREF wszFontName AS WSTRING)
   DECLARE PROPERTY DefaultFontSize () AS LONG
   DECLARE PROPERTY DefaultFontSize (BYVAL nPointSize AS LONG)
   DECLARE PROPERTY Font () AS HFONT
   DECLARE PROPERTY Font (BYVAL hFont AS HFONT)
   DECLARE PROPERTY Brush () AS HBRUSH
   DECLARE PROPERTY Brush (BYVAL hbrBackground AS HBRUSH)
   DECLARE PROPERTY BigIcon (BYVAL hIcon AS HICON)
   DECLARE PROPERTY SmallIcon (BYVAL hIcon AS HICON)
   DECLARE PROPERTY rxRatio () AS SINGLE
   DECLARE PROPERTY rxRatio (BYVAL rx AS SINGLE)
   DECLARE PROPERTY ryRatio (BYVAL rx AS SINGLE)
   DECLARE PROPERTY ryRatio () AS SINGLE
   DECLARE FUNCTION ScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DECLARE FUNCTION ScaleY (BYVAL cy AS SINGLE) AS SINGLE
   DECLARE FUNCTION UnScaleX (BYVAL cx AS SINGLE) AS SINGLE
   DECLARE FUNCTION UnScaleY (BYVAL cy AS SINGLE) AS SINGLE
   DECLARE PROPERTY Width () AS LONG
   DECLARE PROPERTY Height () AS LONG
   DECLARE PROPERTY ScreenX () AS LONG
   DECLARE PROPERTY ScreenY () AS LONG
   DECLARE PROPERTY ClientHeight () AS LONG
   DECLARE PROPERTY ClientWidth () AS LONG
   DECLARE PROPERTY ControlWidth (BYVAL hwnd AS HWND) AS LONG
   DECLARE PROPERTY ControlHeight (BYVAL hwnd AS HWND) AS LONG
   DECLARE PROPERTY ControlClientWidth (BYVAL hwnd AS HWND) AS LONG
   DECLARE PROPERTY ControlClientHeight (BYVAL hwnd AS HWND) AS LONG
   DECLARE SUB GetWindowRect OVERLOAD (BYVAL lpRect AS LPRECT)
   DECLARE FUNCTION GetWindowRect OVERLOAD () AS RECT
   DECLARE SUB GetClientRect OVERLOAD (BYVAL lpRect AS LPRECT)
   DECLARE FUNCTION GetClientRect OVERLOAD () AS RECT
   DECLARE SUB GetControlWindowRect OVERLOAD (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
   DECLARE FUNCTION GetControlWindowRect OVERLOAD (BYVAL hwnd AS HWND) AS RECT
   DECLARE SUB GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
   DECLARE FUNCTION GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND) AS RECT
   DECLARE SUB GetWorkArea OVERLOAD (BYVAL lpRect AS LPRECT)
   DECLARE FUNCTION GetWorkArea OVERLOAD () AS RECT
   DECLARE FUNCTION ControlHandle (BYVAL cID AS LONG) AS HWND
'      DECLARE PROPERTY ScrollWindowPtr () AS CScrollWindow PTR
'      DECLARE PROPERTY ScrollWindowPtr (BYVAL pScrollWindow AS CScrollWindow PTR)
   DECLARE SUB Resize ()
   DECLARE FUNCTION SetWindowPos (BYVAL hwnd AS HWND, BYVAL hwndInsertAfter AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL uFlags AS UINT) AS BOOLEAN
   DECLARE FUNCTION MoveWindow (BYVAL hwnd AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL bRepaint AS BOOLEAN) AS BOOLEAN
   DECLARE FUNCTION CreateFont (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, BYVAL lWeight AS LONG = 0, _
           BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT
   DECLARE FUNCTION SetFont (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, BYVAL lWeight AS LONG = 0, _
           BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS BOOLEAN
   DECLARE FUNCTION AddControl (BYREF wszClassName AS WSTRING, BYVAL hParent AS HWND = NULL, BYVAL cID AS LONG_PTR = 0, _
           BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
           BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS LONG = -1, BYVAL dwExStyle AS LONG = -1, BYVAL lpParam AS LONG_PTR = 0, _
           BYVAL pWndProc AS WNDPROC = NULL) AS HWND

   ' // Keyboard accelerators
   DECLARE PROPERTY AccelHandle () AS HACCEL
   DECLARE PROPERTY AccelHandle (BYVAL hAccel AS HACCEL)
   DECLARE SUB AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYVAL wKey AS WORD, BYVAL cmd AS WORD)
   DECLARE SUB AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYREF wszKey AS WSTRING, BYVAL cmd AS WORD)
   DECLARE FUNCTION CreateAcceleratorTable () AS HACCEL
   DECLARE SUB DestroyAcceleratorTable ()

   ' // Layout manager
   DECLARE FUNCTION AnchorControl (BYVAL cID AS LONG, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   DECLARE FUNCTION AnchorControl (BYVAL hCtl AS HWND, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   DECLARE FUNCTION AnchorControl (BYVAL hParent AS HWND, BYVAL cID AS LONG, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   DECLARE FUNCTION GetAnchorItem (BYVAL hwndCtl AS HWND) AS LONG
   DECLARE FUNCTION AdjustControls () AS BOOLEAN

   '// Control's user data
   DECLARE FUNCTION GetUserDataItem (BYVAL cid AS LONG) AS LONG
   DECLARE FUNCTION ControlSetUser (BYVAL cid AS LONG, BYVAL index AS LONG, BYVAL userValue AS LONG_PTR) AS BOOLEAN
   DECLARE FUNCTION ControlGetUser (BYVAL cid AS LONG, BYVAL index AS LONG) AS LONG_PTR

Private:
   DECLARE STATIC FUNCTION LayoutEnumChildProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG

END TYPE
' ========================================================================================

' ########################################################################################
'                                 *** CWINDOW CLASS ***
' ########################################################################################

' // CWindow class forward declarations
DECLARE FUNCTION CWindow_RegisterClass (BYREF wszClassName AS CONST WSTRING, BYVAL hInstance AS HINSTANCE, BYVAL lpfnWndProc AS WNDPROC) AS ATOM
DECLARE FUNCTION CWindow_WindowProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL hwnd AS HWND) AS CWindow PTR
DECLARE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL lParam AS LPARAM) AS CWindow PTR
DECLARE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL pCreateStruct AS CREATESTRUCT PTR) AS CWindow PTR
DECLARE FUNCTION AfxCWindowOwnerPtr OVERLOAD (BYVAL hwnd AS HWND) AS CWindow PTR

' ========================================================================================
' Registers the window class.
' Parameter:
' - lpfnWndProc = Pointer to the window procedure.
' Return Value:
'   An atom that uniquely identifies the class.
' Remarks:
'   The values of several of the members of this structure can be retrieved with
'   GetClassLongPtrW and set with SetClassLongPtrW.
' Note: The class uses CS_HREDRAW OR CS_VREDRAW as default window styles. Without them,
' the background is not repainted and the controls leave garbage in it when resized. With
' them, windows with many controls cause heavy flicker. To avoid flicker, you can change
' the windows style using e.g. pWindow.ClassStyle = CS_DBLCLKS and take care yourself of
' repainting.
' ========================================================================================
PRIVATE FUNCTION CWindow_RegisterClass (BYREF wszClassName AS CONST WSTRING, BYVAL hInstance AS HINSTANCE, BYVAL lpfnWndProc AS WNDPROC) AS ATOM

   DIM    wAtom AS ATOM                     ' // Atom
   DIM    wcexw AS WNDCLASSEXW              ' // WNDCLASSEXW structure
   STATIC nCount AS LONG = 0                ' // Counter
   STATIC m_wszClassName AS WSTRING * 256   ' // class name

   IF LEN(wszClassName) THEN
      m_wszClassName = wszClassName
   ELSE
      m_wszClassName = "FBWindowClass"
      m_wszClassName = m_wszClassName & ":" & WSTR(nCount)
   END IF

   ' // Fill the WNDCLASSEXW structure
   WITH wcexw
      .cbSize        = SIZEOF(wcexw)
      .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
      .lpfnWndProc   = lpfnWndProc
      .cbClsExtra    = 0
      .cbWndExtra    = SIZEOF(HANDLE)
      .hInstance     = hInstance
      .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_ARROW))
      .hbrBackground = CAST(HBRUSH, COLOR_3DFACE + 1)
      .lpszMenuName  = NULL
      .lpszClassName = @m_wszClassName
      .hIcon         = 0
      .hIconSm       = 0
   END WITH

   ' // Register the class
   wAtom = RegisterClassExW(@wcexw)
   ' // Increment the class counter
   IF wAtom THEN nCount = nCount + 1
   ' // Return the atom
   FUNCTION = wAtom

END FUNCTION
' ========================================================================================

' ========================================================================================
' Processes window messages
' // Uses IsDialogMessage in the message pump
' // Note: To process arrow keys, characters, enter, insert, backspace or delete keys, set USEDLGMSG = 0.
' // Or you can leave it as is and process the WM_GETDLGCODE message:
' // CASE WM_GETDLGCODE
' //    FUNCTION = DLGC_WANTALLKEYS
' // If you are only interested in arrow keys and characters...
' // CASE WM_GETDLGCODE
' //    FUNCTION = DLGC_WANTARROWS OR DLGC_WANTCHARS
' ========================================================================================
PRIVATE FUNCTION CWindow.DoEvents (BYVAL nCmdShow AS LONG = 0) AS LONG

   DIM uMsg AS MSG
   IF m_hwnd = NULL THEN EXIT FUNCTION
   ' // Show the window and update its client area
   IF nCmdShow = 0 THEN ..ShowWindow(m_hwnd, SW_SHOW) ELSE ..ShowWindow(m_hwnd, nCmdShow)
   ..UpdateWindow m_hwnd
   ' // Message loop
   WHILE ..GetMessageW(@uMsg, NULL, 0, 0)
      ' // Processes accelerator keys for menu commands
      IF m_hAccel = NULL OR .TranslateAcceleratorW(m_hwnd, m_hAccel, @uMsg) = 0 THEN
         ' // Determines whether a message is intended for the specified
         ' // dialog box and, if it is, processes the message.
         IF ..IsDialogMessageW(m_hwnd, @uMsg) = 0 THEN
            ' // Translates virtual-key messages into character messages.
            ..TranslateMessage @uMsg
            ' // Dispatches a message to a window procedure.
            ..DispatchMessageW @uMsg
         END IF
      END IF
   WEND
   FUNCTION = uMsg.wParam

END FUNCTION
' ========================================================================================

' ========================================================================================
' Internal CWindow callback procedure.
' ========================================================================================
PRIVATE FUNCTION CWindow_WindowProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM pWindow AS CWindow PTR

   SELECT CASE AS CONST uMsg

      CASE WM_CREATE
         ' // Retrieve a reference to the CWindow class
         DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
         pWindow = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
         RETURN CallWindowProcW(pWindow->GetWindowProc, hwnd, uMsg, wParam, lParam)

      CASE WM_SYSCOMMAND
         ' // Capture this message and send a WM_CLOSE message
         ' // Note: Needed with some OCXs, that otherwise remain in memory
         IF (wParam AND &HFFF0) = SC_CLOSE THEN
            SendMessageW hwnd, WM_CLOSE, 0, 0
         END IF

      CASE WM_COMMAND
         SELECT CASE LOWORD(wParam)
            CASE IDCANCEL
               IF HIWORD(wParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
               END IF
         END SELECT

      CASE WM_SIZE
         ' // Adjust anchored controls and redraw the main window
         pWindow = AfxCWindowPtr(hwnd)
         IF pWindow THEN
            IF wParam = SIZE_RESTORED OR wParam = SIZE_MAXIMIZED THEN
               IF pWindow->AdjustControls THEN
                  InvalidateRect hwnd, NULL, CTRUE
                  UpdateWindow hwnd
               END IF
            END IF
         END IF

      CASE WM_NOTIFY
'         DIM ptnmhdr AS NMHDR PTR
'         ptnmhdr = lParam
'         SELECT CASE @ptnmhdr.idFrom
'            ' ...
'            ' ...
'         END SELECT

   END SELECT

   pWindow = AfxCWindowPtr(hwnd)
   IF pWindow THEN
      IF pWindow->GetWindowProc THEN
         DIM res AS LRESULT = CallWindowProcW(pWindow->GetWindowProc, hwnd, uMsg, wParam, lParam)
         IF res = 0 THEN RETURN 0
      END IF
   END IF

   RETURN DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' CWindow class constructor
' Usage:
'    DIM pWindow AS CWindow
' -or-
'    DIM pWindow AS CWindow = "MyClassName" (change it as needed)
' Remarks:
' Using the fist syntax, CWindow will use "FBWindowClass:" and a number as the window class.
' Using the second syntax, CWindows will use the passed string as the class name.
' ========================================================================================
PRIVATE CONSTRUCTOR CWindow (BYREF wszClassName AS CONST WSTRING = "")

   ' // Class name
   m_wszClassName = wszClassName
   ' // Instance handle
   m_hInstance = ..GetModuleHandleW(NULL)
   ' // Scale windows according to the DPI setting.
'   IF AfxUseDpiScaling = 0 THEN this.DPI = 96 ELSE this.DPI = -1   ' May fail in Windows 10
   IF AfxIsProcessDpiAware = FALSE THEN this.DPI = 96 ELSE this.DPI = -1
   ' // Default font name
   IF LEN(m_wszDefaultFontName) = 0 THEN
      IF AfxWindowsVersion >= 600 AND CLNG(AfxIsProcessDpiAware) THEN m_wszDefaultFontName = "Segoe UI" ELSE m_wszDefaultFontName = "Tahoma"
   END IF
   ' // Default font size
   IF m_DefaultFontSize = 0 THEN
      IF AfxWindowsVersion >= 600 AND CLNG(AfxIsProcessDpiAware) THEN m_DefaultFontSize = 9 ELSE m_DefaultFontSize = 8
   END IF
   ' // Initialize the common controls library
   DIM icc AS INITCOMMONCONTROLSEX
   icc.dwSize = SIZEOF(icc)
   icc.dwICC  = ICC_NATIVEFNTCTL_CLASS OR ICC_COOL_CLASSES OR ICC_BAR_CLASSES OR _
                ICC_TAB_CLASSES OR ICC_USEREX_CLASSES OR ICC_WIN95_CLASSES OR _
                ICC_STANDARD_CLASSES OR ICC_ANIMATE_CLASS OR ICC_DATE_CLASSES OR _
                ICC_HOTKEY_CLASS OR ICC_INTERNET_CLASSES OR ICC_LISTVIEW_CLASSES OR _
                ICC_PAGESCROLLER_CLASS OR ICC_PROGRESS_CLASS OR ICC_TREEVIEW_CLASSES OR _
                ICC_UPDOWN_CLASS
   ..InitCommonControlsEx(@icc)

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CWindow class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWindow

   IF m_hFont THEN ..DeleteObject m_hFont
   IF m_hAccel THEN ..DestroyAcceleratorTable(m_hAccel)
   IF m_wszClassName <> "" THEN .UnregisterClassW(m_wszClassName, m_hInstance)
   IF m_hRichEditLib THEN ..FreeLibrary m_hRichEditLib
'   IF m_pScrollWindow THEN Delete m_pScrollWindow

END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the DPI (dots per pixel) used by the application.
' ========================================================================================
PRIVATE PROPERTY CWindow.DPI () AS SINGLE
   PROPERTY = m_DPI
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the DPI (dots per pixel) to be used by the application.
' Pass -1 to use the value returned by the GetDeviceCaps API function.
' Note: Set this value if you want to make your application High-DPI aware.
' The main window, controls and fonts will be scaled if the user changes the DPI setting.
' Don't change the DPI value once the main window has been created.
' ========================================================================================
PRIVATE PROPERTY CWindow.DPI (BYVAL nDPI AS SINGLE)

   DIM hDC AS HDC
   m_DPI = nDPI
   ' // Get the screen device context
   hDC = ..GetDC(NULL)
   IF m_DPI < 0 THEN
      m_DPI = ..GetDeviceCaps(hDC, LOGPIXELSX)
      nDPI = m_DPI
   END IF
   ' // Resolution ratio = current resolution / 96
   IF nDPI < 0 THEN
      m_rx = (..GetDeviceCaps(hDC, LOGPIXELSX) / 96)
      m_ry = (..GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ELSE
      m_rx = nDPI / 96
      m_ry = nDPI / 96
   END IF
   ..ReleaseDC NULL, hDC

END PROPERTY
' ========================================================================================

' =====================================================================================
' Retrieves a value from the user data area of a CWindow.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE PROPERTY CWindow.UserData (BYVAL idx AS LONG) AS LONG_PTR
   IF idx < 0 OR idx > 99 THEN EXIT PROPERTY
   PROPERTY = m_rgUserData(idx)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets a value in the user data area of a CWindow.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE PROPERTY CWindow.UserData (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
   IF idx < 0 OR idx > 99 THEN EXIT PROPERTY
   m_rgUserData(idx) = newValue
END PROPERTY
' =====================================================================================

' ========================================================================================
' Window creation
' Parameters:
' - hParent     = Parent window handle
' - wszTitle    = Window caption
' - lpfnWndProc = Address of the callback function
' - x           = Horizontal position
' - y           = Vertical position
' - nWidth      = Window width
' - nHeight     = Window height
' - dwStyle     = Window style
' - dwExStyle   = Extended style
' Remarks: As the last parameter we are passing a pointer to the class to allow its use
' in the WM_CREATE message, e.g.
'    CASE WM_CREATE
'       DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
'       DIM pWindow AS CWindow PTR = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
'       IF pWindow THEN pWindow->AddControl("Button", hwnd, IDCANCEL, "&Close", 350, 250, 75, 23)
' -or-
'    CASE WM_CREATE
'       DIM pWindow AS CWindow PTR = AfxCWindowPtr(CAST(CREATESTRUCT PTR, lParam))
'       IF pWindow THEN pWindow->AddControl("Button", hwnd, IDCANCEL, "&Close", 350, 250, 75, 23)
' -or-
'    CASE WM_CREATE
'       DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
'       IF pWindow THEN pWindow->AddControl("Button", hwnd, IDCANCEL, "&Close", 350, 250, 75, 23)
' ========================================================================================
PRIVATE FUNCTION CWindow.Create (BYVAL hParent AS HWND = NULL, BYREF wszTitle AS WSTRING = "", BYVAL lpfnWndProc AS WNDPROC = NULL, _
   BYVAL x AS LONG = CW_USEDEFAULT, BYVAL y AS LONG = CW_USEDEFAULT, BYVAL nWidth AS LONG = CW_USEDEFAULT, BYVAL nHeight AS LONG = CW_USEDEFAULT, _
   BYVAL dwStyle AS UINT = WS_OVERLAPPEDWINDOW OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS, _
   BYVAL dwExStyle AS UINT = WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE) AS HWND

   IF m_hwnd <> NULL THEN EXIT FUNCTION
   ' // User window callback procedure
   m_pUserWndProc = lpfnWndProc
   ' // Internal window callback procedure
   lpfnWndProc = @CWindow_WindowProc
   ' // Register the class
   m_wAtom = CWindow_RegisterClass(m_wszClassName, m_hInstance, lpfnWndProc)
   IF m_wAtom = 0 THEN EXIT FUNCTION
   ' // Create a default font
   IF m_hFont = NULL THEN m_hFont = this.CreateFont(m_wszDefaultFontName, m_DefaultFontSize, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET)
   ' // Create the window
   m_hwnd = ..CreateWindowExW(dwExStyle, CAST(LPCWSTR, CAST(ULONG_PTR, CAST(WORD, m_wAtom))), wszTitle, dwStyle, _
            IIF(x = CW_USEDEFAULT, CW_USEDEFAULT, x * m_rx), _
            IIF(y = CW_USEDEFAULT, CW_USEDEFAULT, y * m_ry), _
            IIF(nWidth = CW_USEDEFAULT, CW_USEDEFAULT, nWidth * m_rx), _
            IIF(nHeight = CW_USEDEFAULT, CW_USEDEFAULT, nHeight * m_ry), _
            hParent, NULL, m_hInstance, CAST(HANDLE, @this))
   IF m_hwnd = NULL THEN EXIT FUNCTION
   ' // Get the class name
   ..GetClassNameW(m_hwnd, @m_wszClassName, SIZEOF(m_wszClassName) / 2)
   FUNCTION = m_hwnd
   ' // Set the font
   IF m_hFont THEN ..SendMessageW m_hwnd, WM_SETFONT, CAST(WPARAM, m_hFont), FALSE
   ' // Store the class pointer
   ..SetWindowLongPtrW(m_hwnd, 0, CAST(LONG_PTR, @this))

END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an overlapped window. An overlapped window has a title bar and a border.
' ========================================================================================
PRIVATE FUNCTION CWindow.CreateOverlapped (BYVAL hParent AS HWND = NULL, BYREF wszTitle AS WSTRING = "", BYVAL lpfnWndProc AS WNDPROC = NULL, _
   BYVAL x AS LONG = CW_USEDEFAULT, BYVAL y AS LONG = CW_USEDEFAULT, BYVAL nWidth AS LONG = CW_USEDEFAULT, BYVAL nHeight AS LONG = CW_USEDEFAULT, _
   BYVAL dwExStyle AS UINT = WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE) AS HWND

   RETURN this.Create(hParent, wszTitle, lpfnWndProc, x, y, nWidth, nHeight, WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU, dwExStyle)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjusts the bounding rectangle of the window based on the desired size of the client area.
' ========================================================================================
PRIVATE SUB CWindow.SetClientSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
   AfxSetWindowClientSize(m_hwnd, nWidth, nHeight, m_rx, m_ry)
END SUB
' ========================================================================================

' =====================================================================================
' Centers a window on the screen or over another window.
' It also ensures that the placement is done within the work area.
' Parameters:
' - hwnd = Handle of the window.
' - hwndParent = [optional] Handle of the parent window.
' =====================================================================================
PRIVATE SUB CWindow.Center (BYVAL hwnd AS HWND = NULL, BYVAL hwndParent AS HWND = NULL)
   IF hwnd = NULL THEN hwnd = m_hwnd
   AfxCenterWindow(hwnd, hwndParent)
END SUB
' =====================================================================================

' =====================================================================================
' Return the Window procedure function pointer
' =====================================================================================
PRIVATE FUNCTION CWindow.GetWindowProc () AS WNDPROC
   FUNCTION = m_pUserWndProc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the window handle
' =====================================================================================
PRIVATE PROPERTY CWindow.hWindow () AS HWND
   PROPERTY = m_hwnd
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the window handle
' =====================================================================================
PRIVATE PROPERTY CWindow.hWindow (BYVAL hwnd AS HWND)
   IF m_hwnd = NULL THEN m_hwnd = hwnd
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the instance handle
' =====================================================================================
PRIVATE PROPERTY CWindow.InstanceHandle () AS HINSTANCE
   PROPERTY = m_hInstance
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the instance handle
' =====================================================================================
PRIVATE PROPERTY CWindow.InstanceHandle (BYVAL hInst AS HINSTANCE)
   IF m_hInstance = NULL THEN m_hInstance = hInst
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window class style.
' =====================================================================================
PRIVATE PROPERTY CWindow.ClassStyle () AS ULONG_PTR
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = ..GetClassLongPtrW(m_hwnd, GCL_STYLE)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window class style.
' =====================================================================================
PRIVATE PROPERTY CWindow.ClassStyle (BYVAL dwStyle AS ULONG_PTR)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   ..SetClassLongPtrW(m_hwnd, GCL_STYLE, dwStyle)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window style.
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowStyle () AS ULONG_PTR
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = ..GetWindowLongPtrW(m_hwnd, GWL_STYLE)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the window style.
' Example to create a popup window: pWindow.WindowStyle = WS_POPUPWINDOW OR WS_CAPTION
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowStyle (BYVAL dwStyle AS ULONG_PTR)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   ..SetWindowLongPtrW(m_hwnd, GWL_STYLE, dwStyle)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window extended style.
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowExStyle () AS ULONG_PTR
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = ..GetWindowLongPtrW(m_hwnd, GWL_EXSTYLE)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the window extended style.
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowExStyle (BYVAL dwExStyle AS ULONG_PTR)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   ..SetWindowLongPtrW(m_hwnd, GWL_EXSTYLE, dwExStyle)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets/sets the default font name
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontName () AS STRING
   PROPERTY = m_wszDefaultFontName
END PROPERTY
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontName (BYREF wszFontName AS WSTRING)
   IF VARPTR(wszFontName) <> NULL AND LEN(wszFontName) > 0 THEN m_wszDefaultFontName = wszFontName
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets/sets the size in points of the default font.
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontSize () AS LONG
   PROPERTY = m_DefaultFontSize
END PROPERTY
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontSize (BYVAL nPointSize AS LONG)
   m_DefaultFontSize = nPointSize
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the handle of the font used as default.
' =====================================================================================
PRIVATE PROPERTY CWindow.Font () AS HFONT
   PROPERTY = m_hFont
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the handle of the font used as default.
' =====================================================================================
PRIVATE PROPERTY CWindow.Font (BYVAL hFont AS HFONT)
   IF m_hFont THEN ..DeleteObject m_hFont
   m_hFont = hFont
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the background brush.
' =====================================================================================
PRIVATE PROPERTY CWindow.Brush () AS HBRUSH
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = CAST(HBRUSH, ..GetClassLongPtrW(m_hwnd, GCLP_HBRBACKGROUND))
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the background brush.
' Handle to the class background brush. This member can be a handle to the physical
' brush to be used for painting the background, or it can be a color value. A color
' value must be one of the standard system colors (the value 1 must be added
' to the chosen color), e.g. COLOR_WINDOW + 1.
' You can also use CreateSolidBrush to create a logical brush with a solid color, e.g.
' CreateSolidBrush(BGR(0, 0, 255)
' =====================================================================================
PRIVATE PROPERTY CWindow.Brush (BYVAL hbrBackground AS HBRUSH)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   ..SetClassLongPtrW m_hwnd, GCLP_HBRBACKGROUND, CAST(LONG_PTR, hbrBackground)
   ..InvalidateRect m_hwnd, NULL, CTRUE
   ..UpdateWindow m_hwnd
END PROPERTY
' =====================================================================================

' =====================================================================================
' Associates a new large icon with the main window. The system displays the large icon
' in the ALT+TAB dialog box.
' Parameters:
' - hIcon = Handle to the new large icon.
'   If this parameter is NULL, the icon is removed.
' =====================================================================================
PRIVATE PROPERTY CWindow.BigIcon (BYVAL hIcon AS HICON)
   ..SendMessageW(m_hwnd, WM_SETICON, CAST(WPARAM, ICON_BIG), CAST(LPARAM, hIcon))
END PROPERTY
' =====================================================================================

' =====================================================================================
' Associates a new small icon with the main window. The system displays the small icon
' in the in the window caption.
' Parameters:
' - hIcon = Handle to the new small icon.
'   If this parameter is NULL, the icon is removed.
' =====================================================================================
PRIVATE PROPERTY CWindow.SmallIcon (BYVAL hIcon AS HICON)
   ..SendMessageW(m_hwnd, WM_SETICON, CAST(WPARAM, ICON_SMALL), CAST(LPARAM, hIcon))
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the horizontal scaling ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.rxRatio () AS SINGLE
   PROPERTY = m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the horizontal scaling ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.rxRatio (BYVAL rx AS SINGLE)
   m_rx = rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the vertical scaling ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.ryRatio () AS SINGLE
   PROPERTY = m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the vertical scaligng ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.ryRatio (BYVAL ry AS SINGLE)
   m_ry = ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Scales an horizontal coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.ScaleX (BYVAL cx AS SINGLE) AS SINGLE
   FUNCTION = cx * m_rx
END FUNCTION
' =====================================================================================

' =====================================================================================
' Scales a vertical coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.ScaleY (BYVAL cy AS SINGLE) AS SINGLE
   FUNCTION = cy * m_ry
END FUNCTION
' =====================================================================================

' =====================================================================================
' Unscales an horizontal coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.UnScaleX (BYVAL cx AS SINGLE) AS SINGLE
   FUNCTION = cx / m_rx
END FUNCTION
' =====================================================================================

' =====================================================================================
' Unscales a vertical coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.UnScaleY (BYVAL cy AS SINGLE) AS SINGLE
   FUNCTION = cy / m_ry
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the unscaled width of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.Width () AS LONG
   DIM rc AS RECT
   ..GetWindowRect m_hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the unscaled height of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.Height () AS LONG
   DIM rc AS RECT
   ..GetWindowRect m_hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the unscaled x-coordinate of the window relative to the screen.
' =====================================================================================
PRIVATE PROPERTY CWindow.ScreenX () AS LONG
   DIM rc AS RECT
   ' // Get the dimensions of the window
   ..GetWindowRect(m_hwnd, @rc)
   ' // Convert the coordinates to be relative to the parent
   ..MapWindowPoints(HWND_DESKTOP, ..GetParent(m_hwnd), CAST(POINT PTR, @rc), 2)
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = rc.Left / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the unscaled y-coordinate of the window relative to the screen.
' =====================================================================================
PRIVATE PROPERTY CWindow.ScreenY () AS LONG
   DIM rc AS RECT
   ' // Get the dimensions of the window
   ..GetWindowRect(m_hwnd, @rc)
   ' // Convert the coordinates to be relative to the parent
   ..MapWindowPoints(HWND_DESKTOP, ..GetParent(m_hwnd), CAST(POINT PTR, @rc), 2)
   ' // Divide by m_ry to make the result High DPI aware
   PROPERTY = rc.Top / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client width of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.ClientWidth () AS LONG
   DIM rc AS RECT
   ..GetClientRect m_hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client height of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.ClientHeight () AS LONG
   DIM rc AS RECT
   ..GetClientRect m_hwnd, @rc
   ' // Divide by m_ry to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the width of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlWidth (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   ..GetWindowRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the height of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlHeight (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   ..GetWindowRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client width of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlClientWidth (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   ..GetClientRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client height of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlClientHeight (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   ..GetClientRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Retrieves the dimensions of the bounding rectangle of the main window.
' =====================================================================================
PRIVATE SUB CWindow.GetWindowRect OVERLOAD (BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   ..GetWindowRect(m_hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetWindowRect OVERLOAD () AS RECT
   DIM rc AS RECT
   ..GetWindowRect(m_hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the coordinates of the main window's client area.
' =====================================================================================
PRIVATE SUB CWindow.GetClientRect OVERLOAD (BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   ..GetClientRect(m_hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetClientRect OVERLOAD () AS RECT
   DIM rc AS RECT
   ..GetClientRect(m_hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the dimensions of the bounding rectangle of the specified window.
' =====================================================================================
PRIVATE SUB CWindow.GetControlWindowRect (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   ..GetWindowRect(hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetControlWindowRect (BYVAL hwnd AS HWND) AS RECT
   DIM rc AS RECT
   ..GetWindowRect(hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the dimensions of a window's client area.
' =====================================================================================
PRIVATE SUB CWindow.GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   ..GetClientRect(hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND) AS RECT
   DIM rc AS RECT
   ..GetClientRect(hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the size of the work area on the primary display monitor. The work area is the
' portion of the screen not obscured by the system taskbar or by application desktop toolbars.
' =====================================================================================
PRIVATE SUB CWindow.GetWorkArea OVERLOAD (BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   SystemParametersInfoW(SPI_GETWORKAREA, 0, lpRect, 0)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetWorkArea OVERLOAD () AS RECT
   DIM rc AS RECT
   SystemParametersInfoW(SPI_GETWORKAREA, 0, @rc, 0)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves a handle to the child control specified by its identifier.
' =====================================================================================
PRIVATE FUNCTION CWindow.ControlHandle (BYVAL cID AS LONG) AS HWND
   FUNCTION = ..GetDlgItem(m_hwnd, cID)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns a pointer to the scroll window class
' =====================================================================================
'PRIVATE PROPERTY CWindow.ScrollWindowPtr () AS CScrollWindow PTR
'   PROPERTY = m_pScrollWindow
'END PROPERTY
' =====================================================================================
' =====================================================================================
' Sets a pointer to the scroll window class
' =====================================================================================
'PRIVATE PROPERTY CWindow.ScrollWindowPtr (BYVAL pScrollWindow AS CScrollWindow PTR)
'   IF m_pScrollWindow THEN Delete m_pScrollWindow
'   m_pScrollWindow = pScrollWindow
'END PROPERTY
' =====================================================================================

' =====================================================================================
' Resizes the window
' =====================================================================================
PRIVATE SUB CWindow.Resize
   ..SendMessageW m_hwnd, WM_SIZE, CAST(WPARAM, SIZE_RESTORED), MAKELONG(this.ClientWidth, this.ClientHeight)
END SUB
' =====================================================================================

' =====================================================================================
' Changes the size, position, and Z order of a child, pop-up, or top-level window.
' =====================================================================================
PRIVATE FUNCTION CWindow.SetWindowPos (BYVAL hwnd AS HWND, BYVAL hwndInsertAfter AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL uFlags AS UINT) AS BOOLEAN
   FUNCTION = ..SetWindowPos(hwnd, hwndInsertAfter, x * m_rx, y * m_ry, cx * m_rx, cy * m_ry, uFlags)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Changes the position and dimensions of the specified window.
' =====================================================================================
PRIVATE FUNCTION CWindow.MoveWindow (BYVAL hwnd AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL bRepaint AS BOOLEAN) AS BOOLEAN
   FUNCTION = ..MoveWindow(hwnd, x * m_rx, y * m_ry, nWidth * m_rx, nHeight * m_ry, bRepaint)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Creates a High DPI aware logical font.
' - wszFaceName = The typeface name.
' - lPointSize = The point size.
' - lWeight = The weight of the font in the range 0 through 1000. For example, 400 is normal
'      and 700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the font or NULL on failure.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' Usage examples:
'   hFont = CWindow.CreateFont("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
'   hFont = CWindow.CreateFont("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
'   hFont = CWindow.CreateFont("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION CWindow.CreateFont ( _
   BYREF wszFaceName AS WSTRING, _                  ' __in Typeface name of font
   BYVAL lPointSize  AS LONG, _                     ' __in Point size
   BYVAL lWeight     AS LONG = 0, _                 ' __in Font weight(bold etc.)
   BYVAL bItalic     AS UBYTE = FALSE, _            ' __in CTRUE = italic
   BYVAL bUnderline  AS UBYTE = FALSE, _            ' __in CTRUE = underline
   BYVAL bStrikeOut  AS UBYTE = FALSE, _            ' __in CTRUE = strikeout
   BYVAL bCharSet    AS UBYTE = DEFAULT_CHARSET _   ' __in character set
   ) AS HFONT                                       ' Handle of font or NULL on failure.

   DIM tlfw AS LOGFONTW
   DIM wstrFaceName AS WSTRING * LF_FACESIZE
   IF VARPTR(wszFacename) <> NULL AND LEN(wszFacename) > 0 THEN wstrFaceName = wszFaceName ELSE wstrFaceName = m_wszDefaultFontName

   ' // The result must be 1, unless the programmer has set a different DPI with the set DPI property
   DIM hDC AS HDC = ..GetDC(HWND_DESKTOP)
   lPointSize = (lPointSize * m_DPI) \ ..GetDeviceCaps(hDC, LOGPIXELSY)

   tlfw.lfHeight         = -MulDiv(lPointSize, ..GetDeviceCaps(hDC, LOGPIXELSY), 72)  ' logical font height
   tlfw.lfWidth          =  0                                                        ' average character width
   tlfw.lfEscapement     =  0                                                        ' escapement
   tlfw.lfOrientation    =  0                                                        ' orientation angles
   tlfw.lfWeight         =  lWeight                                                  ' font weight
   tlfw.lfItalic         =  bItalic                                                  ' italic(CTRUE/FALSE)
   tlfw.lfUnderline      =  bUnderline                                               ' underline(CTRUE/FALSE)
   tlfw.lfStrikeOut      =  bStrikeOut                                               ' strikeout(CTRUE/FALSE)
   tlfw.lfCharSet        =  bCharset                                                 ' character set
   tlfw.lfOutPrecision   =  OUT_TT_PRECIS                                            ' output precision
   tlfw.lfClipPrecision  =  CLIP_DEFAULT_PRECIS                                      ' clipping precision
   tlfw.lfQuality        =  DEFAULT_QUALITY                                          ' output quality
   tlfw.lfPitchAndFamily =  FF_DONTCARE                                              ' pitch and family
   tlfw.lfFaceName       =  wstrFaceName                                             ' typeface name

   ..ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = ..CreateFontIndirectW(@tlfw)

END FUNCTION
' ========================================================================================

' =====================================================================================
' Creates a High DPI aware logical font and sets it as the default font.
' Usage examples:
'   CWindow.SetFont("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
'   CWindow.SetFont("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
'   CWindow.SetFont("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' Return Value = TRUE or FALSE.
' =====================================================================================
PRIVATE FUNCTION CWindow.SetFont ( _
   BYREF wszFaceName AS WSTRING, _                  ' __in Typeface name of font
   BYVAL lPointSize  AS LONG, _                     ' __in Point size
   BYVAL lWeight     AS LONG = 0, _                 ' __in Font weight(bold etc.)
   BYVAL bItalic     AS UBYTE = FALSE, _            ' __in CTRUE = italic
   BYVAL bUnderline  AS UBYTE = FALSE, _            ' __in CTRUE = underline
   BYVAL bStrikeOut  AS UBYTE = FALSE, _            ' __in CTRUE = strikeout
   BYVAL bCharSet    AS UBYTE = DEFAULT_CHARSET _   ' __in character set
   ) AS BOOLEAN                                     ' Result

   DIM hFont AS HFONT
   DIM wstrFaceName AS WSTRING * LF_FACESIZE
   IF VARPTR(wszFacename) <> NULL AND LEN (wszFacename) > 0 THEN wstrFaceName = wszFaceName ELSE wstrFaceName = m_wszDefaultFontName
   hFont = this.CreateFont(wstrFaceName, lPointSize, lWeight, bItalic, bUnderline, bStrikeOut, bCharSet)
   IF hFont THEN
      IF m_hFont THEN ..DeleteObject m_hFont
      m_hFont = hFont
      m_wszDefaultFontName = wstrFaceName
      m_DefaultFontSize = lPointSize
      FUNCTION = TRUE
   ELSE
      FUNCTION = FALSE
   END IF

END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a control to the window
' =====================================================================================
PRIVATE FUNCTION CWindow.AddControl ( _
   BYREF wszClassName AS WSTRING, _                       ' // Class name
   BYVAL hParent AS HWND = NULL, _                        ' // Parent window handle
   BYVAL cID AS LONG_PTR = 0, _                           ' // Control identifier
   BYREF wszTitle AS WSTRING = "", _                      ' // Control caption
   BYVAL x AS LONG = 0, _                                 ' // Horizontal position
   BYVAL y AS LONG = 0, _                                 ' // Vertical position
   BYVAL nWidth AS LONG = 0, _                            ' // Control width
   BYVAL nHeight AS LONG = 0, _                           ' // Control height
   BYVAL dwStyle AS LONG = -1, _                          ' // Control style
   BYVAL dwExStyle AS LONG = -1, _                        ' // Extended style
   BYVAL lpParam AS LONG_PTR = 0, _                       ' // Pointer to custom data
   BYVAL pCtlProc AS WNDPROC = NULL _                     ' // Address of the control callback procedure
   ) AS HWND                                              ' // Control handle

   DIM hCtl AS HWND
   IF LEN(wszClassName) = 0 THEN EXIT FUNCTION
   IF hParent = NULL THEN hParent = m_hwnd
   DIM bSetFont AS LONG = CTRUE
   ' // Window styles
   DIM wsClassName AS WSTRING * 260
   wsClassName = wszClassName

   SELECT CASE UCASE(wsClassName)
      CASE "BUTTON"
         ' Adds a button to the window
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER
         IF dwStyle = BS_FLAT THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER OR BS_FLAT
         IF dwStyle = BS_DEFPUSHBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_DEFPUSHBUTTON
         IF dwStyle = BS_OWNERDRAW THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_OWNERDRAW
         #if _WIN32_WINNT = &h0602
         IF dwStyle = BS_SPLITBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_SPLITBUTTON
         IF dwStyle = BS_DEFSPLITBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_DEFSPLITBUTTON
         #endif
      CASE "CUSTOMBUTTON", "OWNERDRAWBUTTON"
         ' Adds an ownerdraw button to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_OWNERDRAW
      CASE "RADIOBUTTON", "OPTION"
         ' Adds a radio button to the window.
         ' Note: In PowerBASIC this control is called "Option".
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER
         IF dwStyle = WS_GROUP THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER OR WS_GROUP
      CASE "CHECKBOX"
         ' Adds a checkbox to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTOCHECKBOX OR BS_LEFT OR BS_VCENTER
      CASE "CHECK3STATE"
         ' Adds a 3 state checkbox to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTO3STATE OR BS_LEFT OR BS_VCENTER
      CASE "LABEL"
         ' Adds a label to the window.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SS_LEFT OR WS_GROUP OR SS_NOTIFY
      CASE "BITMAPLABEL"
         ' Adds an image label to the window.
         ' You must delete the bitmap before the application ends.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_BITMAP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "ICONLABEL"
         ' Adds an image label to the window.
         ' You must delete the icon before the application ends.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_ICON
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "BITMAPBUTTON"
         ' Adds an image button to the window.
         ' You must delete the bitmap before the application ends.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_BITMAP
      CASE "ICONBUTTON"
         ' Adds an image button to the window.
         ' You must delete the icon before the application ends.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_ICON
      CASE "CUSTOMLABEL"
         ' Adds an ownerdraw label to the window.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_OWNERDRAW
         bSetFont = FALSE
      CASE "FRAME", "FRAMEWINDOW"
         ' Adds a frame to the window.
         ' Note: This is not the same that PowerBASIC DDT's Frame control, that in fact is a Group Box.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPSIBLINGS OR WS_GROUP OR SS_BLACKFRAME
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "GROUPBOX"
         ' Adds a group box to the window.
         ' Note: This is the same that DDT's frame control.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPSIBLINGS OR WS_GROUP OR BS_GROUPBOX
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
      CASE "LINE"
         ' Adds an horizontal line to the window
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SS_ETCHEDFRAME
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "EDIT", "TEXTBOX"
         ' Adds an edit control to the window.
         wsClassName = "Edit"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR ES_AUTOHSCROLL
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "EDITMULTILINE", "MULTILINETEXTBOX"
         ' Adds an edit control to the window.
         wsClassName = "Edit"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR WS_VSCROLL OR ES_LEFT OR ES_AUTOHSCROLL OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "COMBOBOX"
         ' Adds a combo box to the window.
         IF dwStyle = -1 THEN dwStyle = WS_CHILD OR WS_VISIBLE OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWN OR CBS_HASSTRINGS OR CBS_SORT
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "COMBOBOXEX", "COMBOBOXEX32"
         ' Adds a combo box ex to the window.
         wsClassName = "ComboBoxEx32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWNLIST
      CASE "LISTBOX"
         ' Adds a list box to the window.
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR LBS_STANDARD OR LBS_HASSTRINGS OR LBS_SORT OR LBS_NOTIFY
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "PROGRESSBAR", "MSCTLS_PROGRESS32"
         ' Adds a progress bar to the window.
         wsClassName = "msctls_progress32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE
         bSetFont = FALSE
      CASE "HEADER", "SYSHEADER32"
         ' Adds an header control to the window.
         wsClassName = "SysHeader32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR CCS_TOP OR HDS_HORZ OR HDS_BUTTONS
      CASE "TREEVIEW", "SYSTREEVIEW32"
         ' Adds a tree view control to the window.
         wsClassName = "SysTreeView32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR TVS_HASBUTTONS OR TVS_HASLINES OR TVS_LINESATROOT OR TVS_SHOWSELALWAYS
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "LISTVIEW", "SYSLISTVIEW32"
         ' Adds a list view control to the window.
         wsClassName = "SysListView32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_TABSTOP OR LVS_REPORT OR LVS_SHOWSELALWAYS OR LVS_SHAREIMAGELISTS OR LVS_AUTOARRANGE OR LVS_EDITLABELS OR LVS_ALIGNTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "TOOLBAR", "TOOLBARWINDOW32"
         ' Adds a toolbar control to the window.
         wsClassName = "ToolbarWindow32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_TOP OR WS_BORDER OR TBSTYLE_FLAT OR TBSTYLE_TOOLTIPS
      CASE "REBAR", "REBARWINDOW32"
         ' Adds a rebar control to the window.
         wsClassName = "ReBarWindow32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_NODIVIDER OR RBS_VARHEIGHT OR RBS_BANDBORDERS
      CASE "DATETIMEPICKER", "SYSDATETIMEPICK32"
         ' Adds a date time picker control to the window.
         wsClassName = "SysDateTimePick32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR DTS_SHORTDATEFORMAT
      CASE "MONTHCALENDAR", "MONTHCAL", "SYSMONTHCAL32"
         ' Adds a month calendar control to the window.
         wsClassName = "SysMonthCal32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "IPADDRESS", "SYSIPADDRESS32"
         ' Adds an IPAddress control to the window.
         wsClassName = "SysIPAddress32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "HOTKEY", "MSCTLS_HOTKEY32"
         ' Adds an hotkey control to the window.
         wsClassName = "msctls_hotkey32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "ANIMATE", "ANIMATION", "SYSANIMATE32"
         ' Adds an animation control to the window.
         wsClassName = "SysAnimate32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR ACS_TRANSPARENT
      CASE "SYSLINK"
         ' Adds a SysLink control to the window.
         ' Note: The SysLink control is defined in the ComCtl32.dll version 6 and requires a manifest
         ' or directive that specifies that version 6 of the DLL should be used if it is available.
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         bSetFont = FALSE
      CASE "PAGER", "SYSPAGER"
         ' Adds a Pager control to the window.
         wsClassName = "SysPager"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR PGS_HORZ
         bSetFont = FALSE
      CASE "TAB", "TABCONTROL", "SYSTABCONTROL32"
         ' Adds a Tab control to the window.
         wsClassName = "SysTabControl32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR WS_TABSTOP OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR TCS_TABS OR TCS_SINGLELINE OR TCS_RAGGEDRIGHT
         IF dwExStyle = -1 THEN dwExStyle = 0
         dwExStyle =  dwExStyle OR WS_EX_CONTROLPARENT
      CASE "STATUSBAR", "MSCTLS_STATUSBAR32"
         ' Adds a StatusBar control to the window.
         wsClassName = "msctls_statusbar32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_BOTTOM OR SBARS_SIZEGRIP
      CASE "SIZEBAR", "SIZEBOX", "SIZEGRIP"
         ' Adds a size box to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SBS_SIZEGRIP OR SBS_SIZEBOXBOTTOMRIGHTALIGN
         bSetFont = FALSE
         nWidth = GetSystemMetrics(SM_CXVSCROLL)
         nHeight = GetSystemMetrics(SM_CYHSCROLL)
         DIM rcClient AS RECT = this.GetClientRect
         x = rcClient.Right - nWidth
         y = rcClient.Bottom - nHeight
         bSetFont = FALSE
      CASE "HSCROLLBAR"
         ' Adds an horizontal scroll bar to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR SBS_HORZ
         bSetFont = FALSE
      CASE "VSCROLLBAR"
         ' Adds a vertical scroll bar to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR SBS_VERT
         bSetFont = FALSE
      CASE "TRACKBAR", "MSCTLS_TRACKBAR32", "SLIDER"
         wsClassName = "msctls_trackbar32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR TBS_AUTOTICKS OR TBS_HORZ OR TBS_BOTTOM OR TBS_TOOLTIPS
         bSetFont = FALSE
      CASE "UPDOWN", "MSCTLS_UPDOWN32"
         wsClassName = "msctls_updown32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR UDS_WRAP OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_SETBUDDYINT
         bSetFont = FALSE
      CASE "RICHEDIT", "RICHEDIT50W"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR WS_HSCROLL OR WS_VSCROLL OR ES_AUTOHSCROLL OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_WANTRETURN OR ES_NOHIDESEL OR ES_SAVESEL
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
         wsClassName = "RichEdit50W"
         IF m_hRichEditLib = 0 THEN m_hRichEditLib = CAST(HMODULE, LoadLibraryW("MSFTEDIT.DLL"))
   END SELECT
   ' // Don't allow negative values for the styles
   IF dwStyle = -1 THEN dwStyle = 0
   IF dwExStyle = -1 THEN dwExStyle = 0
   ' // Make sure that the control has the WS_CHILD style
   dwStyle = dwStyle OR WS_CHILD
   ' // Create the control
   hCtl = ..CreateWindowExW(dwExStyle, wsClassName, wszTitle, dwStyle, x * m_rx, y * m_ry, nWidth * m_rx, nHeight * m_ry, _
      hParent, CAST(HMENU, cID), m_hInstance, CAST(LPVOID, lpParam))
   IF hCtl = NULL THEN EXIT FUNCTION
   ' // Set the font
   IF m_hFont THEN
      IF bSetFont THEN ..SendMessageW hCtl, WM_SETFONT, CAST(WPARAM, m_hFont), CTRUE
   END IF
   SELECT CASE UCASE(wszClassName)
      CASE "LISTBOX"
         ' // Adjust the height of the control so that the integral height
         ' // is based on the new font rather than the default SYSTEM_FONT
         SetWindowPos hCtl, NULL, x, y, nWidth, nHeight, SWP_NOZORDER
      CASE "DATETIMEPICKER", "SYSDATETIMEPICK32"
         ' // Sets the font to be used by the date and time picker control's child month calendar control.
         IF m_hFont THEN SendMessageW hCtl, DTM_SETMCFONT, CAST(WPARAM, m_hFont), CTRUE
      CASE "PROGRESSBAR", "MSCTLS_PROGRESS32"
         ' // Set the default range
         ..SendMessageW hCtl, PBM_SETRANGE32, 0, 100
         ' // Set the default initial value
         ..SendMessageW hCtl, PBM_SETPOS, 0, 0
      CASE "TRACKBAR", "MSCTLS_TRACKBAR32"
         ' // Set the default range values
         ..SendMessageW hCtl, TBM_SETRANGEMIN, CTRUE, 0
         ..SendMessageW hCtl, TBM_SETRANGEMAX, CTRUE, 100
         ' // Set the default page size
         ..SendMessageW hCtl, TBM_SETPAGESIZE, 0, 10
      CASE "UPDOWN", "MSCTLS_UPDOWN32"
         ' // Set the default base
         ..SendMessageW hCtl, UDM_SETBASE, 10, 0
         ' // Set the default range values
         ..SendMessageW hCtl, UDM_SETRANGE32, 100, 0
         ' // Set the default initial value
         ..SendMessageW hCtl, UDM_SETPOS32, 0, 0
         ' // Correct for Windows using a default size for the updown control
         this.SetWindowPos hCtl, NULL, x, y, nWidth, nHeight, SWP_NOZORDER
      CASE "HSCROLLBAR", "VSCROLLBAR"
         ' // Initialize the scroll bar with default values
         DIM tsi AS SCROLLINFO
         tsi.cbSize = SIZEOF(tsi)
         tsi.fMask  = SIF_PAGE OR SIF_POS OR SIF_RANGE
         tsi.nMin   = 0
         tsi.nMax   = 100
         tsi.nPage  = 0
         tsi.nPos   = 0
         ..SetScrollInfo hCtl, SB_CTL, @tsi, CTRUE
      CASE "TOOLBAR", "TOOLBARWINDOW32"
         ' // Set the button size
         DIM AS LONG nButtonWidth, nButtonHeight
         nButtonWidth = LOWORD(.SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0)) * m_rx
         nButtonHeight = HIWORD(.SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0)) * m_ry
         ..SendMessageW hCtl, TB_SETBUTTONSIZE, 0, MAKELONG(nButtonWidth, nButtonHeight)
         ' // Send this message for backward compatibility
         ..SendMessageW hCtl, TB_BUTTONSTRUCTSIZE, SIZEOF(TBBUTTON), 0
      CASE "BITMAPLABEL"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = ..LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = ..LoadImageW(m_hInstance, wszTitle, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN ..SendMessageW(hCtl, STM_SETIMAGE, IMAGE_BITMAP, CAST(LPARAM, hImage))
      CASE "ICONLABEL"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = ..LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = ..LoadImageW(m_hInstance, wszTitle, IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN ..SendMessageW(hCtl, STM_SETIMAGE, IMAGE_ICON, CAST(LPARAM, hImage))
      CASE "BITMAPBUTTON"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = ..LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = ..LoadImageW(m_hInstance, wszTitle, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN ..SendMessageW(hCtl, BM_SETIMAGE, IMAGE_BITMAP, CAST(LPARAM, hImage))
      CASE "ICONBUTTON"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = ..LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = ..LoadImageW(m_hInstance, wszTitle, IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN ..SendMessageW(hCtl, BM_SETIMAGE, IMAGE_ICON, CAST(LPARAM, hImage))
   END SELECT
   FUNCTION = hCtl
END FUNCTION
' =====================================================================================

' ########################################################################################
'                               *** ACCELERATOR TABLE ***
' ########################################################################################

' =====================================================================================
' Gets/Sets the accelerator table handle
' =====================================================================================
PRIVATE PROPERTY CWindow.AccelHandle () AS HACCEL
   PROPERTY = m_hAccel
END PROPERTY
' =====================================================================================
' =====================================================================================
PRIVATE PROPERTY CWindow.AccelHandle (BYVAL hAccel AS HACCEL)
   IF m_hAccel THEN ..DestroyAcceleratorTable(m_hAccel)
   IF UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) > -1 THEN ERASE m_rgAccelEntries
   m_hAccel = hAccel
END PROPERTY
' =====================================================================================
' =====================================================================================
' Adds an accelerator key to the table.
' =====================================================================================
PRIVATE SUB CWindow.AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYVAL wKey AS WORD, BYVAL cmd AS WORD)
   REDIM PRESERVE m_rgAccelEntries(UBOUND(m_rgAccelEntries) + 1) AS ACCEL
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).fvirt = fvirt
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).key = wKey
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).cmd = cmd
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB CWindow.AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYREF wszKey AS WSTRING, BYVAL cmd AS WORD)
   REDIM PRESERVE m_rgAccelEntries(UBOUND(m_rgAccelEntries) + 1) AS ACCEL
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).fvirt = fvirt
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).key = ASC(wszKey)
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).cmd = cmd
END SUB
' =====================================================================================
' =====================================================================================
' Creates the accelerator table.
' =====================================================================================
PRIVATE FUNCTION CWindow.CreateAcceleratorTable () AS HACCEL
   IF UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) = -1 THEN EXIT FUNCTION
   IF m_hAccel THEN ..DestroyAcceleratorTable(m_hAccel)
   m_hAccel = ..CreateAcceleratorTableW(CAST(LPACCEL, @m_rgAccelEntries(0)), UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) + 1)
   FUNCTION = m_hAccel
END FUNCTION
' =====================================================================================
' =====================================================================================
' Destroys the accelerator table.
' =====================================================================================
PRIVATE SUB CWindow.DestroyAcceleratorTable
   IF m_hAccel THEN ..DestroyAcceleratorTable(m_hAccel)
   IF UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) > -1 THEN ERASE m_rgAccelEntries
   m_hAccel = NULL
END SUB
' =====================================================================================

' ########################################################################################
'                             *** CONTROL'S USER DATA ***
' ########################################################################################

' ========================================================================================
' Gets the control user data item
' Return Value: The array index on success, or -1 on failure.
' ========================================================================================
PRIVATE FUNCTION CWindow.GetUserDataItem (BYVAL cID AS LONG) AS LONG
   FUNCTION = -1
   Dim hCtl AS HWND = GetDlgItem(m_hWnd, cID)
   FOR i AS LONG = 0 TO UBOUND(m_rgCtrlUserData)
      IF m_rgCtrlUserData(i).hCtl = hCtl THEN FUNCTION = i : EXIT FOR
   NEXT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the control's user data
' ========================================================================================
PRIVATE FUNCTION CWindow.ControlSetUser (BYVAL cID AS LONG, BYVAL index AS LONG, BYVAL userValue AS LONG_PTR) AS BOOLEAN
   IF index < 0 OR index > 9 THEN RETURN FALSE
   DIM hCtl AS HWND = GetDlgItem(m_hWnd, CID)
   IF IsWindow(hCtl) = FALSE THEN RETURN FALSE
   ' // Check if the element already exists
   DIM nItem AS LONG = this.GetUserDataItem(cID)
   ' // If it does not exist, make room for it
   IF nItem = -1 THEN
      REDIM PRESERVE m_rgCtrlUserData(UBOUND(m_rgCtrlUserData) + 1) AS AFX_CTLUSERDATA
      nItem = UBOUND(m_rgCtrlUserData)
   END IF
   m_rgCtrlUserData(nItem).hCtl = hCtl
   m_rgCtrlUserData(nItem).Values(index) = userValue
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the colors of the control
' ========================================================================================
PRIVATE FUNCTION CWindow.ControlGetUser (BYVAL cID AS LONG, BYVAL index AS LONG) AS LONG_PTR
   IF index < 0 OR index > 9 THEN RETURN FALSE
   DIM hCtl AS HWND = GetDlgItem(m_hWnd, cID)
   DIM nItem AS LONG = this.GetUserDataItem(cID)
   IF nItem > -1 THEN RETURN m_rgCtrlUserData(nItem).Values(index)
END FUNCTION
' ========================================================================================

' ########################################################################################
'                               *** LAYOUT MANAGER ***
' ########################################################################################

' ========================================================================================
' Gets the anchor item
' Return Value: The array index on success, or -1 on failure.
' ========================================================================================
PRIVATE FUNCTION CWindow.GetAnchorItem (BYVAL hwndCtl AS HWND) AS LONG
   FUNCTION = -1
   FOR i AS LONG = 0 TO UBOUND(m_rgProps)
      IF m_rgProps(i).hwndCtl = hwndCtl THEN FUNCTION = i : EXIT FOR
   NEXT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Anchors the control
' ========================================================================================
PRIVATE FUNCTION CWindow.AnchorControl (BYVAL hwndCtl AS HWND, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   IF .IsWindow(hwndCtl) = FALSE THEN RETURN FALSE
   ' // Check if the item already exists
   DIM nItem AS LONG = this.GetAnchorItem(hwndCtl)
   ' // If it does not exist, make room for it
   IF nItem = -1 THEN
      REDIM PRESERVE m_rgProps(UBOUND(m_rgProps) + 1) AS AFX_ANCHORPROPERTY
      nItem = UBOUND(m_rgProps)
   END IF
   m_rgProps(nItem).hwndCtl = hwndCtl
   ' // Get the coordinates of the control
   DIM rcCtl AS RECT, pt AS POINT
   .GetWindowRect(hwndCtl, @rcCtl)
   pt.x = rcCtl.Left
   pt.y = rcCtl.Top
   ' // Convert the screen coordinates to screen to client-area coordinates
   .ScreenToClient(GetParent(hwndCtl), @pt)
   ' // Get the client coordinates of the parent window
   DIm rcParent AS RECT
   .GetClientRect(GetParent(hwndCtl), @rcParent)
   m_rgProps(nItem).anchor    = MIN(MAX(nAnchorMode, AFX_ANCHOR_NONE), AFX_ANCHOR_CENTER)
   m_rgProps(nItem).rc.Left   = pt.x
   m_rgProps(nItem).rc.Top    = pt.y
   m_rgProps(nItem).rc.Right  = rcParent.Right - (rcCtl.Right - rcCtl.Left + pt.x)
   m_rgProps(nItem).rc.Bottom = rcParent.Bottom - (rcCtl.Bottom - rcCtl.Top + pt.y)
   m_rgProps(nItem).centerX   = pt.x - (rcParent.Right \ 2)
   m_rgProps(nItem).centerY   = pt.y - (rcParent.Bottom \ 2)
   RETURN TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWindow.AnchorControl (BYVAL cID AS LONG, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   DIM hwndCtl AS HWND = GetDlgItem(m_hwnd, cID)
   IF hwndCtl THEN RETURN this.AnchorControl(hwndCtl, nAnchorMode)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWindow.AnchorControl (BYVAL hParent AS HWND, BYVAL cID AS LONG, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   DIM hwndCtl AS HWND = GetDlgItem(hParent, cID)
   IF hwndCtl THEN RETURN this.AnchorControl(hwndCtl, nAnchorMode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback function to enumerate child windows
' ========================================================================================
PRIVATE FUNCTION CWindow.LayoutEnumChildProc (BYVAL hwndChild AS HWND, BYVAL lParam AS LPARAM) AS LONG
   DIM rcParent AS RECT, rcChild AS RECT
   DIM nWidth AS LONG, nHeight AS LONG, x AS LONG, y AS LONG, xWidth AS LONG, yHeight AS LONG
   DIM pWindow AS CWindow PTR = CAST(CWindow PTR, lParam)
   DIM nItem AS LONG = pWindow->GetAnchorItem(hwndChild)
   IF nItem > -1 THEN
      IF pWindow->m_rgProps(nItem).anchor > AFX_ANCHOR_NONE THEN
         ..GetWindowRect(hwndChild, @rcChild)
         nWidth = rcChild.Right - rcChild.Left
         nHeight = rcChild.Bottom - rcChild.Top
         ..GetClientRect(GetParent(hwndChild), @rcParent)
         x = 0: y = 0: xWidth = 0: yHeight = 0
         SELECT CASE pWindow->m_rgProps(nItem).anchor
            CASE AFX_ANCHOR_NONE
               x = pWindow->m_rgProps(nItem).rc.Left
               y = pWindow->m_rgProps(nItem).rc.Top
               xWidth = MAX(rcParent.Right - pWindow->m_rgProps(nItem).rc.Left - pWindow->m_rgProps(nItem).rc.Right, 0)
               yHeight = nHeight
            CASE AFX_ANCHOR_WIDTH
               x = pWindow->m_rgProps(nItem).rc.Left
               y = pWindow->m_rgProps(nItem).rc.Top
               xWidth = MAX(rcParent.Right - pWindow->m_rgProps(nItem).rc.Left - pWindow->m_rgProps(nItem).rc.Right, 0)
               yHeight = nHeight
            CASE AFX_ANCHOR_RIGHT
               x = rcParent.Right - nWidth - pWindow->m_rgProps(nItem).rc.Right
               y = pWindow->m_rgProps(nItem).rc.Top
               xWidth = nWidth
               yHeight = nHeight
            CASE AFX_ANCHOR_CENTER_HORZ
               x = (rcParent.Right \ 2) + pWindow->m_rgProps(nItem).centerX
               y = pWindow->m_rgProps(nItem).rc.Top
               xWidth = nWidth
               yHeight = nHeight
            CASE AFX_ANCHOR_HEIGHT
               x = pWindow->m_rgProps(nItem).rc.Left
               y = pWindow->m_rgProps(nItem).rc.Top
               xWidth = nWidth
               yHeight = MAX(rcParent.Bottom - pWindow->m_rgProps(nItem).rc.Top - pWindow->m_rgProps(nItem).rc.Bottom, 0)
            CASE AFX_ANCHOR_HEIGHT_WIDTH
               x = pWindow->m_rgProps(nItem).rc.Left
               y = pWindow->m_rgProps(nItem).rc.Top
               xWidth = MAX(rcParent.Right - pWindow->m_rgProps(nItem).rc.Left - pWindow->m_rgProps(nItem).rc.Right, 0)
               yHeight = MAX(rcParent.Bottom - pWindow->m_rgProps(nItem).rc.Top - pWindow->m_rgProps(nItem).rc.Bottom, 0)
            CASE AFX_ANCHOR_HEIGHT_RIGHT
               x = rcParent.Right - nWidth - pWindow->m_rgProps(nItem).rc.Right
               y = pWindow->m_rgProps(nItem).rc.Top
               xWidth = nWidth
               yHeight = MAX(rcParent.Bottom - pWindow->m_rgProps(nItem).rc.Top - pWindow->m_rgProps(nItem).rc.Bottom, 0)
            CASE AFX_ANCHOR_BOTTOM
               x = pWindow->m_rgProps(nItem).rc.Left
               y = rcParent.Bottom - pWindow->m_rgProps(nItem).rc.Bottom - nHeight
               xWidth = nWidth
               yHeight = nHeight
            CASE AFX_ANCHOR_BOTTOM_WIDTH
               x = pWindow->m_rgProps(nItem).rc.Left
               y = rcParent.Bottom - pWindow->m_rgProps(nItem).rc.Bottom - nHeight
               xWidth = MAX(rcParent.Right - pWindow->m_rgProps(nItem).rc.Left - pWindow->m_rgProps(nItem).rc.Right, 0)
               yHeight = nHeight
            CASE AFX_ANCHOR_BOTTOM_RIGHT
               x = rcParent.Right - nWidth - pWindow->m_rgProps(nItem).rc.Right
               y = rcParent.Bottom - pWindow->m_rgProps(nItem).rc.Bottom - nHeight
               xWidth = nWidth
               yHeight = nHeight
            CASE AFX_ANCHOR_CENTER_HORZ_BOTTOM
               x = (rcParent.Right \ 2) + pWindow->m_rgProps(nItem).centerX
               y = rcParent.Bottom - pWindow->m_rgProps(nItem).rc.Bottom - nHeight
               xWidth = nWidth
               yHeight = nHeight
            CASE AFX_ANCHOR_CENTER_VERT
               x = pWindow->m_rgProps(nItem).rc.Left
               y = (rcParent.Bottom - nHeight) \ 2
               xWidth = nWidth
               yHeight = nHeight
            CASE AFX_ANCHOR_CENTER_VERT_RIGHT
               x = rcParent.Right - nWidth - pWindow->m_rgProps(nItem).rc.Right
               y = (rcParent.Bottom - nHeight) \ 2
               xWidth = nWidth
               yHeight = nHeight
            CASE AFX_ANCHOR_CENTER
               x = (rcParent.Right \ 2) + pWindow->m_rgProps(nItem).centerX
               y = (rcParent.Bottom \ 2) + pWindow->m_rgProps(nItem).centerY
               xWidth = nWidth
               yHeight = nHeight
         END SELECT
         ..MoveWindow(hwndChild, x, y, xWidth, yHeight, CTRUE)
      END IF
   END IF

   RETURN CTRUE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjust the controls size and/or position
' ========================================================================================
PRIVATE FUNCTION CWindow.AdjustControls () AS BOOLEAN
   IF UBOUND(m_rgProps) = -1 THEN RETURN FALSE
   EnumChildWindows(m_hWnd, CAST(WNDENUMPROC, @LayoutEnumChildProc), CAST(LPARAM, @this))
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                *** HELPER FUNCTIONS ***
' ########################################################################################

' ========================================================================================
' Returns a pointer to the CWindow class given the handle of its associated window handle.
' To retrieve it from the handle of any of its child windows or controls, use AfxCWindowOwnerPtr.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL hwnd AS HWND) AS CWindow PTR
   FUNCTION = CAST(CWindow PTR, ..GetWindowLongPtrW(hwnd, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWindow class given the a pointer to the CREATESTRUCT structure.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL lParam AS LPARAM) AS CWindow PTR
   DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
   FUNCTION = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWindow class given the a pointer to the CREATESTRUCT structure.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL pCreateStruct AS CREATESTRUCT PTR) AS CWindow PTR
   FUNCTION = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWindow class given the handle of the window created with it
' or the handle of any of it's children.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowOwnerPtr OVERLOAD (BYVAL hwnd AS HWND) AS CWindow PTR
   IF hwnd = NULL THEN EXIT FUNCTION
   DIM hRootOwner AS ..HWND = ..GetAncestor(hwnd, GA_ROOTOWNER)
   IF hRootOwner = NULL THEN EXIT FUNCTION
   FUNCTION = CAST(CWindow PTR, ..GetWindowLongPtrW(hRootOwner, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CTabPage class given the handle of the tab control to which the
' tab page is associated and the zero-based tab index. If nTabIdx is ommited, the function
' will return the pointer of the selected tab, if any.
' ========================================================================================
'PRIVATE FUNCTION AfxCTabPagePtr OVERLOAD (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS CTabPage PTR
'   IF hTab = NULL THEN EXIT FUNCTION
'   IF idx = -1 THEN idx = SendMessageW(hTab, TCM_GETCURSEL, 0, 0)
'   IF idx = -1 THEN EXIT FUNCTION   ' No tab selected
'   ' // Ask to return the value of the lParam member
'   DIM tci AS TCITEMW
'   tci.mask = TCIF_PARAM
'   IF ..SendMessageW(hTab, TCM_GETITEMW, idx, CAST(LPARAM, @tci)) THEN
'      IF tci.lParam THEN
'         FUNCTION = CAST(CTabPage PTR, tci.lParam)
'      END IF
'   END IF
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CScrollWindow class given the handle of the tab control to which the
' tab page is associated and the zero-based tab index. If nTabIdx is ommited, the function
' will return the pointer associated to the selected tab, if any.
' ========================================================================================
'PRIVATE FUNCTION AfxScrollTabPagePtr (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS CScrollWindow PTR
'   FUNCTION = AfxCTabPagePtr(hTab, idx)->ScrollWindowPtr
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CScrollWindow class given the handle of the window attached to it.
' ========================================================================================
'PRIVATE FUNCTION AfxScrollWindowPtr (BYVAL hwnd AS HWND) AS CScrollWindow PTR
'   FUNCTION = cast(CWindow PTR, GetWindowLongPtr(hwnd, 0))->ScrollWindowPtr
'END FUNCTION
' ========================================================================================

END NAMESPACE
