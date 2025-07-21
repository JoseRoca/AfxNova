' ########################################################################################
' Microsoft Windows
' File: CW_PopupWindow.bas
' Contents: CWindow with a modal popup window
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
DECLARE FUNCTION WndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION PopupWindow (BYVAL hParent AS HWND) AS LONG
DECLARE FUNCTION PopupWndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_POPUP = 1001

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Popup Window", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // Adds a button
   pWindow.AddControl("Button", hWin, IDC_POPUP, "&Popup", 270, 155, 75, 30)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDC_POPUP, AFX_ANCHOR_BOTTOM_RIGHT)

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
            CASE IDC_POPUP
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN PopupWindow(hwnd)
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
' Popup window procedure
' ========================================================================================
FUNCTION PopupWindow (BYVAL hParent AS HWND) AS LONG

   DIM pWindow AS CWindow
   pWindow.Create(hParent, "Popup window", @PopupWndProc, , , , , _
      WS_VISIBLE OR WS_CAPTION OR WS_POPUPWINDOW OR WS_THICKFRAME, WS_EX_WINDOWEDGE)
   pWindow.Brush = GetStockObject(WHITE_BRUSH)
   pWindow.SetClientSize(300, 200)
   pWindow.Center(pWindow.hWindow, hParent)
   ' / Process Windows messages
   FUNCTION = pWindow.DoEvents

END FUNCTION
' ========================================================================================

' ========================================================================================
' Popup window procedure
' ========================================================================================
FUNCTION PopupWndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC hNewFont AS HFONT

   SELECT CASE uMsg

      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         ' // Get a pointer to the CWindow class from the CREATESTRUCT structure
         DIM pWindow AS CWindow PTR = AfxcWindowPtr(lParam)
         ' // Create a new font scaled according the DPI ratio
         IF pWindow->DPI <> 96 THEN hNewFont = pWindow->CreateFont("Times New Roman", 24)
         ' Disable parent window to make popup window modal
         EnableWindow GetParent(hwnd), FALSE
         RETURN 0

      ' // Theme has changed
      CASE WM_THEMECHANGED
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_PAINT
    		DIM ps AS PAINTSTRUCT, hOldFont AS HFONT
         DIM hDC AS HDC = BeginPaint(hWnd, @ps)
         IF hNewFont THEN hOldFont = CAST(HFONT, SelectObject(hDC, CAST(HGDIOBJ, hNewFont)))
         DIM rc AS RECT
         GetClientRect(hWnd, @rc)
         DrawTextW(hDC, "Hello, World!", -1, @rc, DT_SINGLELINE or DT_CENTER or DT_VCENTER)
         IF hNewFont THEN SelectObject(hDC, CAST(HGDIOBJ, CAST(HFONT, hOldFont)))
         EndPaint(hWnd, @ps)
         EXIT FUNCTION

      CASE WM_CLOSE
         ' // Enables parent window keeping parent's zorder
         EnableWindow GetParent(hwnd), CTRUE
         ' // Don't exit; let DefWindowProcW perform the default action

    	CASE WM_DESTROY
         ' // Destroy the new font
         IF hNewFont THEN DeleteObject(CAST(HGDIOBJ, hNewFont))
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
