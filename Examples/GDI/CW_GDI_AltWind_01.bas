' ########################################################################################
' Microsoft Windows
' File: CW_GDI_AltWind_01.bas
' Contents: CWindow - Alternate and Winding Fill Modes
' This program is a translation/adaptation of the ALTWIND.C-Alternate and Winding Fill
' Modes Program © Charles Petzold, 1998, described and analysed in Chapter 5 of the book
' Programming Windows, 5th Edition.
' Displays the figure twice, once using the ALTERNATE filling mode and then using WINDING.
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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - GDI - Alternate and Winding Fill Modes", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center
   ' // Set the background color to white
   pWindow.SetBackColor(RGB_WHITE)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC cxClient AS LONG
   STATIC cyClient AS LONG
   STATIC aptFigure(9) AS POINT = {(10, 70), (50, 70), (50, 10), (90, 10), (90, 50), (30, 50), (30, 90), (70, 90), (70, 30), (10, 30)}

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

      CASE WM_PAINT
         ' // Do the painting
         DIM ps AS PAINTSTRUCT
         DIM hdc AS HDC = BeginPaint(hwnd, @ps)
         SelectObject hdc, GetStockObject(GRAY_BRUSH)
         DIM i AS LONG, apt(9) AS POINT
         FOR i = 0 TO 9
            apt(i).x = cxClient * aptFigure(i).x / 200
            apt(i).y = cyClient * aptFigure(i).y / 100
         NEXT
         SetPolyFillMode hdc, ALTERNATE
         Polygon hdc, @apt(0), 10
         FOR i = 0 TO 9
            apt(i).x = apt(i).x + cxClient / 2
         NEXT
         SetPolyFillMode hdc, WINDING
         Polygon hdc, @apt(0), 10
         EndPaint hwnd, @ps
         RETURN 0

      CASE WM_SIZE
         ' // Store the size of the client area
         cxClient = LOWORD(lParam)
         cyClient = HIWORD(lParam)
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
