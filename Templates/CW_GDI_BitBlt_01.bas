' ########################################################################################
' Microsoft Windows
' File: CW_CGI_B itBlt_01.bas
' Contents: BitBlt demo
' This example captures an image of the entire screen, creates a compatible device context
' and a bitmap with the appropriate dimensions, selects the bitmap into the compatible DC,
' and then copies the image using the BitBlt function.
' Later, in the WM_PAINT and WM_ERASEBKGND messages, redisplays the image calling BitBlt
' again, specifying the compatible DC as the source DC and a window DC as the target DC.
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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - BitBlt demo", @WndProc)
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

   STATIC hdcCompatible AS HDC           ' // Handle of the compatible device context
   STATIC cxBitmap      AS LONG          ' // Width of the bitmap
   STATIC cyBitmap      AS LONG          ' // Height of the bitmap
   STATIC cxClient      AS LONG          ' // Width of the client area
   STATIC cyClient      AS LONG          ' // Height of the client area
   DIM    hdcScreen     AS HDC           ' // Desktop device context
   DIM    hbmScreen     AS HBITMAP       ' // Screen bitmap handle
   DIM    hdc           AS HDC           ' // Window device context
   DIM    ps            AS PAINTSTRUCT   ' // PAINTSTRUCT structure

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         ' // Create a normal DC and a memory DC for the entire screen. The
         ' // normal DC provides a "snapshot" of the screen contents. The
         ' // memory DC keeps a copy of this "snapshot" in the associated
         ' // bitmap.
         DIM hdcScreen AS HDC = CreateDCW("DISPLAY", NULL, NULL, NULL)
         hdcCompatible = CreateCompatibleDC(hdcScreen)
         ' // Create a compatible bitmap for hdcScreen.
         cxBitmap = GetDeviceCaps(hdcScreen, HORZRES)
         cyBitmap = GetDeviceCaps(hdcScreen, VERTRES)
         hbmScreen = CreateCompatibleBitmap(hdcScreen, cxBitmap, cyBitmap)
         ' // Select the bitmaps into the compatible DC.
         SelectObject(hdcCompatible, hbmScreen)
         ' // Copy color data for the entire display into the
         ' // bitmap that is selected into a compatible DC.
         BitBlt hdcCompatible, 0, 0, cxBitmap, cyBitmap, hdcScreen, 0, 0, SRCCOPY
         ' // Delete the screen DC and bitmap
         DeleteDC hdcScreen
         DeleteObject hbmScreen
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
         EXIT FUNCTION

      CASE WM_PAINT
         hdc = BeginPaint(hwnd, @ps)
         BitBlt hdc, 0, 0, cxBitmap, cyBitmap, hdcCompatible, 0, 0, SRCCOPY
         EndPaint hwnd, @ps
         EXIT FUNCTION

      CASE WM_ERASEBKGND
         hdc = CAST(HDC, wParam)
         BitBlt hdc, 0, 0, cxBitmap, cyBitmap, hdcCompatible, 0, 0, SRCCOPY
         FUNCTION = CTRUE
         EXIT FUNCTION

      CASE WM_DESTROY
         ' // Delete the compatible device context
         IF hdcCompatible THEN DeleteDC hdcCompatible
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
