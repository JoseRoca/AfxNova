' ########################################################################################
' Microsoft Windows
' File: CW_SBMouseClick.bas
' Contents: CWindow - Status Bar
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

CONST IDC_START = 1001
CONST IDC_STATUSBAR = 1002
CONST IDC_PROGRESSBAR = 1003

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Status bar - Detect mouse click", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // Addsa label
   pWindow.AddControl("LABEL", , -1, "Click one of the sections of the status bar with the mouse", 20, 40, 350, 23)

   ' // Add a status bar
   DIM hStatusbar AS HWND = pWindow.AddControl("Statusbar", hWin, IDC_STATUSBAR)
   ' // Set the parts
   DIM rgParts(1 TO 3) AS LONG = {pWindow.ScaleX(100), pWindow.ScaleX(220), -1}
   IF StatusBar_SetParts(hStatusBar, 3, @rgParts(1)) THEN
      StatusBar_Simple(hStatusBar, FALSE)
   END IF
   StatusBar_SetText(hStatusbar, 0, "Part 1")
   StatusBar_SetText(hStatusbar, 1, "Part 2")
   StatusBar_SetText(hStatusbar, 2, "Part 3")
   ' // Anchor the status bar
   pWindow.AnchorControl(IDC_STATUSBAR, AFX_ANCHOR_BOTTOM_WIDTH)

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
         END SELECT
         RETURN 0

      CASE WM_NOTIFY
         ' // Detect if the user has clicked the mouse in one of the status bar parts
         DIM nmouse AS NMMOUSE
         CBNMTYPESET(nmouse, wParam, lParam)
         IF nmouse.hdr.idFrom = IDC_STATUSBAR AND nmouse.hdr.code = NM_CLICK THEN
            ' // Display the zero-based index of the section that was clicked.
            MessageBoxW hwnd, "You have clicked section " & WSTR(nmouse.dwItemSpec + 1), "", MB_OK
         END IF

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
