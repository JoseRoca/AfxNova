' ########################################################################################
' Microsoft Windows
' File: CW_Edit_DarkMode.bas
' Contents: CWindow - Edit controls
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_EDIT1 = 1001
CONST IDC_EDIT2 = 1002

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Edit controls", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center
   ' // Set the main window background color
   pWindow.SetBackColor(RGB_BLACK)

   ' // Add an Edit control
   DIM hEdit AS HWND = pWindow.AddControl("Edit", hWin, IDC_EDIT1, "", 20, 15, 360, 23)
   ' // Colors of the edit text
   pWindow.SetCtlColors(hEdit, RGB_WHITE, RGB_BLACK)
   ' // Add a multiline Edit control
   DIM hEditMulti AS HWND = pWindow.AddControl("EditMultiline", hWin, IDC_EDIT2, "", 20, 45, 360, 110)
   ' // Colors of the edit text
   pWindow.SetCtlColors(hEditMulti, RGB_WHITE, RGB_BLACK)
   ' // Anchor the controls
   pWindow.AnchorControl(hWin, IDC_EDIT1, AFX_ANCHOR_WIDTH)
   pWindow.AnchorControl(hWin, IDC_EDIT2, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Adds a button
   DIM hButton AS HWND = pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", 305, 170, 75, 30)
   ' // Set button dark mode
   SetWindowTheme(hButton, "DarkMode_Explorer", NULL)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Set the focus In the edit control
   SetFocus hEdit

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
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Theme has changed
      CASE WM_THEMECHANGED
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

      CASE WM_CLOSE
         ' // Multiline edit controls send a WM_CLOSE message when the Esc key is pressed.
         ' // This is a workaround; the proper way is to subclass the control and eat the ESC key.
         IF GetFocus = GetDlgItem(hwnd, IDC_EDIT2) THEN RETURN 0   ' // abort

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
