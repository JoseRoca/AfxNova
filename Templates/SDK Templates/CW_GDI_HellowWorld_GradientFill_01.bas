' ########################################################################################
' Microsoft Windows
' File: CW_GDI_HelloWorld_GradientFill_01.bas
' Contents: CWindow Hello Word with gradient example (GradientFill)
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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Hello World - Gradient Fill", @WndProc)
   ' // Change the background to white
   pWindow.Brush = GetStockObject(WHITE_BRUSH)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gradient fill procedure
' ========================================================================================
SUB DrawGradient (BYVAL hwnd AS HWND, BYVAL hDC AS HDC)

   DIM rc AS RECT, vertex(1) AS TRIVERTEX

   GetClientRect hwnd, @rc

   vertex(0).x      = 0
   vertex(0).y      = 0
   vertex(0).Red    = &hFF00
   vertex(0).Green  = &hFF00
   vertex(0).Blue   = &h0000
   vertex(0).Alpha  = &h0000

   vertex(1).x      = rc.Right - rc.Left
   vertex(1).y      = rc.Bottom - rc.Top
   vertex(1).Red    = &h8000
   vertex(1).Green  = &h0000
   vertex(1).Blue   = &h0000
   vertex(1).Alpha  = &h0000

   DIM gRect AS GRADIENT_RECT
   gRect.UpperLeft  = 0
   gRect.LowerRight = 1

   GradientFill hDc, @vertex(0), 2, @gRect, 1, GRADIENT_FILL_RECT_H

END SUB
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC hNewFont AS HFONT

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         ' // Get a pointer to the CWindow class from the CREATESTRUCT structure
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
         ' // Create a new font scaled according the DPI ratio
         hNewFont = pWindow->CreateFont("Times New Roman", 16)
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

      CASE WM_PAINT
         ' // Draw the text
         DIM ps AS PAINTSTRUCT, hOldFont AS HFONT
         DIM hDC AS HDC = BeginPaint(hwnd, @ps)
         DrawGradient hwnd, hdc
         IF hNewFont THEN hOldFont = CAST(HFONT, SelectObject(hDC, CAST(HGDIOBJ, hNewFont)))
         DIM rc AS RECT
         GetClientRect hwnd, @rc
         SetBkMode hDC, TRANSPARENT
         SetTextColor hDC, &HFFFFFF
         DrawTextW hDC, "Hello, World!", -1, @rc, DT_SINGLELINE OR DT_CENTER OR DT_VCENTER
         IF hNewFont THEN SelectObject(hDC, CAST(HGDIOBJ, CAST(HFONT, hOldFont)))
         EndPaint hwnd, @ps
         RETURN TRUE

      CASE WM_ERASEBKGND
         ' // Draw the gradient
         DIM hDC AS HDC = CAST(HDC, wParam)
         DrawGradient hwnd, hDC
         RETURN TRUE

      CASE WM_DESTROY
         ' // Destroy the new font
         IF hNewFont THEN DeleteObject(CAST(HGDIOBJ, hNewFont))
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
