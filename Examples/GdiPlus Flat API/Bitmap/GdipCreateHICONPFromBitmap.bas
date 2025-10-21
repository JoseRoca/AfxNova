' ########################################################################################
' Microsoft Windows
' File: GdipCreateHICONFromBitmap.bas
' Contents: GDI+ Flat API - GdipCreateHICONFromBitmap example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.bi"
#INCLUDE ONCE "AfxNova/ARGB_Colors.bi"
#INCLUDE ONCE "AfxNova/CGraphCtx.inc"
#INCLUDE ONCE "file.bi"
USING AfxNova

CONST IDC_GRCTX = 1001

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG

   END wWinMain(GetModuleHandleW(NULL), NULL, wCOMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' This example converts a GDI+ Bitmap into a Windows HICON and displays it.
' ========================================================================================
SUB Example_CreateHICONFromBitmap (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a GDI+ Bitmap from file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Convert to HICON
   DIM hIcon AS HICON
   hStatus = GdipCreateHICONFromBitmap(bmp, @hIcon)

   ' // Get width and heig ht of the icon
   DIM iconInfo AS ICONINFO
   DIM iconWidth AS LONG
   DIM iconHeight AS LONG
   IF GetIconInfo(hIcon, @iconInfo) THEN
      DIM bmp AS BITMAP
      IF iconInfo.hbmColor THEN
         ' Icon has a color bitmap
         GetObject(iconInfo.hbmColor, SIZEOF(BITMAP), @bmp)
         iconWidth = bmp.bmWidth
         iconHeight = bmp.bmHeight
         DeleteObject(iconInfo.hbmColor)
      ELSEIF iconInfo.hbmMask THEN
         ' Monochrome icon: height is double, so divide by 2
         GetObject(iconInfo.hbmMask, SIZEOF(BITMAP), @bmp)
         iconWidth = bmp.bmWidth
         iconHeight = bmp.bmHeight \ 2
      END IF
      IF iconInfo.hbmMask THEN DeleteObject(iconInfo.hbmMask)
   END IF

   ' // Scale it according to the DPI settings
   DIM scaledWidth AS LONG = iconWidth * rxRatio
   DIM scaledHeight AS LONG = iconHeight * ryRatio

   ' // Draw the icon
   DrawIconEx(hdc, 0, 0, hIcon, scaledWidth, scaledHeight, 0, NULL, DI_NORMAL)

   ' // Cleanup
   IF hIcon THEN DestroyIcon(hIcon)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
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

   ' // Create the main window
   DIM pWindow AS CWindow = "MyClassName"
   pWindow.Create(NULL, "GDI+ GdipCreateHICONFromBitmap", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_WHITE
   ' // Anchor the control
   pWindow.AnchorControl(pGraphCtx.hWindow, AFX_ANCHOR_HEIGHT_WIDTH)
   
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc

   ' // Initialize GDI+
   DIM token AS ULONG_PTR = AfxGdipInit

   ' // Draw the graphics
   Example_CreateHICONFromBitmap(hdc)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Shutdown GDI+
   AfxGdipShutdown token

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

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  RETURN 0
               END IF
         END SELECT

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
