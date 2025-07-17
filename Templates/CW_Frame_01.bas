' ########################################################################################
' Microsoft Windows
' File: CW_Frame_01.bas
' Contents: Demonstrates the use of the Frame control.
' Comments: Drawn around several controls to indicate a visual association between them.
' Often used around related radio buttons. Similar to the GroupBox control but without a label.
' Remarks: The PowerBASIC "Frame" control is a GroupBox.
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
   IDC_FRAME = 1001
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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Frame control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center
   ' // Sets the main window background color
   pWindow.SetBackColor(RGB_OLDLACE)

   ' // Adds the controls
   pWindow.AddControl("Frame", hWin, IDC_FRAME, "Frame", 20, 20, 360, 120)
   pWindow.AddControl("Label", hWin, IDC_LABEL, "Label", 60, 50, 75, 23)
   pWindow.AddControl("Check3State", hWin, IDC_CHECK3STATE, "Click me", 60, 100, 75, 23)
   pWindow.AddControl("Edit", hWin, IDC_EDIT, "", 210, 50, 150, 23)
   pWindow.AddControl("Button", hWin, IDC_BUTTON, "&Ok", 210, 100, 75, 23)
   ' // Anchors the controls
   pWindow.AnchorControl(IDC_EDIT, AFX_ANCHOR_WIDTH)
   pWindow.AnchorControl(IDC_FRAME, AFX_ANCHOR_WIDTH)

   ' // Sets the colors of the controls
   pWindow.SetCtlColors(IDC_LABEL, RGB_YELLOW, RGB_VIOLET)
   pWindow.SetCtlColors(IDC_CHECK3STATE, RGB_YELLOW, RGB_LIGHTBLUE)
   pWindow.SetCtlColors(IDC_EDIT, RGB_RED, RGB_YELLOW)

   ' // Adds a button
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", 305, 165, 75, 30)
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
