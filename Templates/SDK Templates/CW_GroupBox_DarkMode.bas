' ########################################################################################
' Microsoft Windows
' File: CW_GroupBox_DarkMode.bas
' Contents: Demonstrates the use of the GroupBox control.
' Comments: Drawn around several controls to indicate a visual association between them.
' Often used around related radio buttons.
' Remarks: In PowerBASIC it is mistakenly called "Frame" instead of "Group Box".
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

' // Control identifiers
ENUM
   IDC_GROUPBOX = 1001
   IDC_LABEL
   IDC_CHECK3STATE
   IDC_EDIT
   IDC_BUTTON
END ENUM

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - GroupBox control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center
   ' // Sets the main window background color
   pWindow.SetBackColor(RGB_BLACK)

   ' // Adds the controls
   DIM hGroupBox AS HWND = pWindow.AddControl("GroupBox", hWin, IDC_GROUPBOX, "GroupBox", 20, 20, 360, 120)
   DIM hLabel AS HWND = pWindow.AddControl("Label", hWin, IDC_LABEL, "Label", 60, 50, 75, 23)
   DIM hCheckState AS HWND = pWindow.AddControl("Check3State", hWin, IDC_CHECK3STATE, "Click me", 60, 100, 75, 23)
   DIM hEdit AS HWND = pWindow.AddControl("Edit", hWin, IDC_EDIT, "", 210, 50, 150, 23)
   DIM hButton AS HWND = pWindow.AddControl("Button", hWin, IDC_BUTTON, "&Ok", 210, 100, 75, 23)

   ' // Anchors the controls
   pWindow.AnchorControl(IDC_EDIT, AFX_ANCHOR_WIDTH)
   pWindow.AnchorControl(IDC_GROUPBOX, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Sets the colors of the controls
   pWindow.SetCtlColors(hLabel, RGB_WHITE, RGB_BLACK)
   pWindow.SetCtlColors(hEdit, RGB_WHITE, RGB_BLACK)

   ' // To set the color of the group box and the check state controls we need to disable the theme
   SetWindowTheme hGroupBox, "", ""
   pWindow.SetCtlColors(IDC_GROUPBOX, RGB_WHITE, RGB_BLACK)
   SetWindowTheme hCheckState, "", ""
   pWindow.SetCtlColors(hCheckState, RGB_WHITE, RGB_BLACK)
   ' // Set button dark mode
   SetWindowTheme(hButton, "DarkMode_Explorer", NULL)

   ' // Adds a button
   DIM hButtonCancel AS HWND = pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", 305, 165, 75, 30)
   ' // Set button dark mode
   SetWindowTheme(hButtonCancel, "DarkMode_Explorer", NULL)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Set the focus in the edit conrol
   SetFocus GetDlgItem(hWin, IDC_EDIT)

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
            ' // If Close button pressed, close the application sending an WM_CLOSE message
            CASE IDC_BUTTON
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  MessageBox hwnd, "Ok", "Message", MB_OK
               END IF
         END SELECT
         RETURN 0

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
