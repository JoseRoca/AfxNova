' ########################################################################################
' Microsoft Windows
' File: CW_Button_Subclass_01.bas
' Contents: CWindow - Button subclass
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

' // Forward declarations
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

   ' // Create the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Button subclass", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Center the window
   pWindow.Center

   ' // Add a button
   DIM hButton AS HWND = pWindow.AddControl("Button", hWin, IDC_BUTTON, "&Click me", 270, 155, 75, 30)
   ' // Anchor the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDC_BUTTON, AFX_ANCHOR_BOTTOM_RIGHT)
   ' // Subclass the button
   pWindow.SetControlSubclass(hButton, @Button_SubclassProc, IDC_BUTTON, @pWindow)

   ' // Display the window and dispatches the Windows messages
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

      CASE WM_DESTROY
         ' // REQUIRED: Remove control subclassing
         RemoveWindowSubclass hwnd, @Button_SubclassProc, uIdSubclass
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
