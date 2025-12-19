' ########################################################################################
' Microsoft Windows
' File: CW_GDI_Bounce_01.bas
' Contents: CWindow - Bounce
' This program is a translation/adaptation of BOUNCE.C -- Bouncing Ball Program
' © Charles Petzold, 1998, described and analysed in Chapter 14 of the book Programming
' Windows, 5th Edition.
' The BOUNCE program constructs a ball that bounces around in the window's client area.
' The program uses the timer to pace the ball. The ball itself is a bitmap. The program
' first creates the ball by creating the bitmap, selecting it into a memory device
' context, and then making simple GDI function calls. The program draws the bitmapped ball
' on the display using a BitBlt from a memory device context.
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

CONST ID_TIMER = 1

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Bounce", @WndProc)
   ' // Change the background to white
   pWindow.Brush = GetStockObject(WHITE_BRUSH)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(500, 300)
   ' // Centers the window
   pWindow.Center

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC hBitmap  AS HBITMAP
   STATIC cxClient AS LONG
   STATIC cyClient AS LONG
   STATIC xCenter  AS LONG
   STATIC yCenter  AS LONG
   STATIC cxTotal  AS LONG
   STATIC cyTotal  AS LONG
   STATIC cxRadius AS LONG
   STATIC cyRadius AS LONG
   STATIC cxMove   AS LONG
   STATIC cyMove   AS LONG
   STATIC xPixel   AS LONG
   STATIC yPixel   AS LONG

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         DIM hdc AS HDC = GetDC(hwnd)
         xPixel = GetDeviceCaps(hdc, ASPECTX)
         yPixel = GetDeviceCaps(hdc, ASPECTY)
         ReleaseDC hwnd, hdc
         SetTimer hwnd, ID_TIMER, 50, NULL
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

     CASE WM_SIZE
         cxClient = LOWORD(lParam)
         cyClient = HIWORD(lParam)
         xCenter = cxClient  \ 2
         yCenter = cyClient  \ 2
         DIM iScale AS LONG = MIN(cxClient * xPixel, cyClient * yPixel) \ 16
         cxRadius = iScale \ xPixel
         cyRadius = iScale \ yPixel
         cxMove = MAX(1, cxRadius \ 2)
         cyMove = MAX(1, cyRadius \ 2)
         cxTotal = 2 * (cxRadius + cxMove)
         cyTotal = 2 * (cyRadius + cyMove)
         IF hBitmap THEN DeleteObject hBitmap
         DIM hdc AS HDC = GetDC(hwnd)
         DIM hdcMem AS HDC = CreateCompatibleDC(hdc)
         hBitmap = CreateCompatibleBitmap(hdc, cxTotal, cyTotal)
         ReleaseDC hwnd, hdc
         SelectObject hdcMem, hBitmap
         Rectangle hdcMem, -1, -1, cxTotal + 1, cyTotal + 1
         DIM hBrush AS HBRUSH = CreateHatchBrush(HS_DIAGCROSS, 0)
         SelectObject hdcMem, hBrush
         SetBkColor hdcMem, BGR(255, 0, 255)
         Ellipse hdcMem, cxMove, cyMove, cxTotal - cxMove, cyTotal - cyMove
         DeleteDC hdcMem
         DeleteObject hBrush
         RETURN 0

      CASE WM_TIMER
         IF hBitmap = NULL THEN EXIT FUNCTION
         DIM hdc AS HDC = GetDC(hwnd)
         DIM hdcMem AS HDC = CreateCompatibleDC(hdc)
         SelectObject hdcMem, hBitmap
         BitBlt hdc, xCenter - cxTotal \ 2, _
                     yCenter - cyTotal \ 2, cxTotal, cyTotal, _
                hdcMem, 0, 0, SRCCOPY
         ReleaseDC hwnd, hdc
         DeleteDC hdcMem
         xCenter = xCenter + cxMove
         yCenter = yCenter + cyMove
         IF (xCenter + cxRadius) >= cxClient OR (xCenter - cxRadius <= 0) THEN cxMove = -cxMove
         IF (yCenter + cyRadius) >= cyClient OR (yCenter - cyRadius) <= 0 THEN cyMove = -cyMove
         EXIT FUNCTION

      CASE WM_DESTROY
          ' // Dekete the botmap
         IF hBitmap THEN DeleteObject hBitmap
         ' // Kill the timer
         KillTimer hwnd, ID_TIMER
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
