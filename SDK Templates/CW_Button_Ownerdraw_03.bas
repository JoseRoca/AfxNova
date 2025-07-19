' ########################################################################################
' Microsoft Windows
' File: CW_Button_Ownerdraw_03.bas
' Contents: CWindow - Ownerdraw button with normal and hot icons
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/RGB_Colors.bi"
#INCLUDE ONCE "AfxNova/CWindow.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION Button_SubclassProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM, BYVAL uIdSubclass AS UINT_PTR, BYVAL dwRefData AS DWORD_PTR) AS LRESULT

CONST IDC_BUTTON = 1001

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Ownerdraw button with normal and hot icons", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // Adds an ownerdraw button
   DIM hButton AS HWND = pWindow.AddControl("Button", hWin, IDC_BUTTON, "", 200, 100, 50, 50, BS_OWNERDRAW)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDC_BUTTON, AFX_ANCHOR_BOTTOM_RIGHT)
   ' // Subclass the button
   pWindow.SetControlSubclass(hButton, @Button_SubclassProc, IDC_BUTTON, @pWindow)
   ' // Set the focus on the button
   SetFocus hButton

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
            CASE IDC_BUTTON
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  MessageBoxW hwnd, "Button clicked", "", MB_OK
                  SetFocus GetDlgItem(hwnd, IDC_BUTTON)
               END IF
         END SELECT
         RETURN 0

      CASE WM_DRAWITEM
         DIM pDis AS DRAWITEMSTRUCT PTR = CAST(DRAWITEMSTRUCT PTR, lParam)
         IF pDis->CtlId <> IDC_BUTTON THEN EXIT FUNCTION
         ' Icon identifiers in User32.dll. Hacked with ResourceHacker.
         ' IDI_APPLICATION = 100, IDI_WARNING = 101, IDI_QUESTION = 102, IDI_ERROR = 103
         ' IDI_INFORMATION = 104, IDI_WINLOGO = 105, IDI_SHIELD = 106
         ' Used here to make the example simpler.
         ' Normally, you will load an icon or bitmap from a file or resource.
         DIM hIcon AS HICON
         IF GetPropW(GetDlgItem(hwnd, IDC_BUTTON), "HOT") = 0 THEN
            hIcon = LoadImageW(GetModuleHandle("User32"), MAKEINTRESOURCEW(102), IMAGE_ICON, _
               AfxScaleX(50), AfxScaleY(50), LR_LOADTRANSPARENT)
         ELSE
            hIcon = LoadImageW(GetModuleHandle("User32"), MAKEINTRESOURCEW(103), IMAGE_ICON, _
               AfxScaleX(50), AfxScaleY(50), LR_LOADTRANSPARENT)
         END IF
         IF hIcon THEN
            DrawStateW pDis->hDC, NULL, NULL, CAST(LPARAM, hIcon), 0, 0, 0, 0, 0, DST_ICON
            DestroyIcon hIcon
         END IF
         RETURN TRUE

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Processes messages for the subclassed Button window.
' ========================================================================================
FUNCTION Button_SubclassProc ( _
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
         RETURN DLGC_WANTALLKEYS

      CASE WM_LBUTTONDOWN
         MessageBoxW(GetParent(hwnd), "Click", "FreeBasic", MB_OK)
         RETURN 0

      CASE WM_KEYDOWN
         SELECT CASE CBNMID(wParam, lParam)
            CASE VK_ESCAPE
               SendMessageW(GetParent(hwnd), WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

      CASE WM_MOUSEMOVE
         ' // Tracks the mouse movement and stores the hot state
         DIM trackMouse AS TRACKMOUSEEVENT
         IF GetPropW(hwnd, "HOT") = 0 THEN
            trackMouse.cbSize = SIZEOF(trackMouse)
            trackMouse.dwFlags = TME_LEAVE
            trackMouse.hwndTrack = hwnd
            trackMouse.dwHoverTime = 1
            TrackMouseEvent(@trackMouse)
            SetPropW hwnd, "HOT", CAST(HANDLE, CTRUE)
            InvalidateRect hwnd, NULL, 0
            UpdateWindow hwnd
            RETURN 0
         END IF

      CASE WM_MOUSELEAVE
         ' // Removes the hot state and redraws the button
         RemovePropW hwnd, "HOT"
         InvalidateRect hwnd, NULL, 0
         UpdateWindow hwnd
         RETURN 0

      CASE WM_DESTROY
         ' // REQUIRED: Remove control subclassing
         RemoveWindowSubclass hwnd, @Button_SubclassProc, uIdSubclass
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
