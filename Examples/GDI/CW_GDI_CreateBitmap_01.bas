' ########################################################################################
' Microsoft Windows
' File: CW_GDI_CreateBitmap_01.bas
' Contents: CreateBitmap demo
' This example creates a bitmap from the desktop and displays it.
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
' Creates a bitmap from the desktop
' ========================================================================================
FUNCTION CreateBitmapFromDesktop () AS HDC

   ' // Get the desktop handle
   DIM hwnd AS HWND = GetDesktopWindow
   ' // Get the device context
   DIM hDC AS HDC = GetWindowDC(hwnd)

   ' // Fill the BITMAP structure
   DIM bmp AS BITMAP
   bmp.bmBitsPixel = GetDeviceCaps(hDC, BITSPIXEL_)
   bmp.bmPlanes    = GetDeviceCaps(hDC, PLANES_)
   bmp.bmWidth     = GetDeviceCaps(hDC, HORZRES)
   bmp.bmHeight    = GetDeviceCaps(hDC, VERTRES)

   ' // Create a memory bitmap
   DIM hMemBmp AS HBITMAP
   hMemBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, bmp.bmPlanes, bmp.bmBitsPixel, 0)
   ' // Create a compatible memory context
   DIM hMemDC AS HDC = CreateCompatibleDC(hDC)
   ' // Select the memory bitmap in the memory device context
   DIM hBmpOld AS HBITMAP = SelectObject(hMemDC, hMemBmp)
   ' // BitBlt the bitmap to the window device context
   BitBlt(hMemDC, 0, 0, bmp.bmWidth, bmp.bmHeight, hDC, 0, 0, SRCCOPY)
   ' // Release the window device context
   ReleaseDC(hwnd, hDC)

   ' // Return the memory device context
   RETURN hMemDC

END FUNCTION
' ========================================================================================

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - CreateBitmap demo", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(700, 400)
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

   STATIC hMemDC AS HDC     ' // Memory device context

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         ' // Enable dark mode for the caption
         AfxEnableDarkModeForWindow(hwnd)
         ' // Create a bitmap from the desktop
         hMemDC = CreateBitmapFromDesktop
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
         ' // Redraws the bitmap
         DIM ps AS PAINTSTRUCT
         BeginPaint(hwnd, @ps)
         BitBlt ps.hdc, ps.rcPaint.left, ps.rcPaint.top, ps.rcPaint.right - ps.rcPaint.left, _
                ps.rcPaint.bottom - ps.rcPaint.top, hMemDC, ps.rcPaint.left, ps.rcPaint.top, SRCCOPY
         EndPaint(hwnd, @ps)
         RETURN 0

      CASE WM_DESTROY
         ' // Delete the memory device context
         DeleteDC hMemDC
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
