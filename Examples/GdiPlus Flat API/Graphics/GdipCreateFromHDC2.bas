' ########################################################################################
' Microsoft Windows
' File: GdipCreateFromHDC2.bas
' Contents: GDI+ Flat API - GdipCreateFromHDC2 example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.bi"
#INCLUDE ONCE "AfxNova/ARGB_Colors.bi"
#INCLUDE ONCE "AfxNova/CGraphCtx.inc"
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
' This example demonstrates how to use GdipCreateFromHDC2 with a device handle,
' and how to create a bitmap from a graphics object for off-screen rendering.
' ========================================================================================
SUB Example_CreateFromHDC2 (BYVAL hdc AS HDC)

   DIM hStatus AS LONG
   DIM graphics AS GpGraphics PTR

   DIM hPrinter AS HANDLE
   DIM printerName AS WSTRING * 128 = "Microsoft Print to PDF"
   IF OpenPrinter(printerName, @hPrinter, NULL) = FALSE THEN EXIT SUB

   ' // Create graphics object using HDC and device handle
   hStatus = GdipCreateFromHDC2(hdc, hPrinter, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a bitmap from the graphics object
   DIM bitmap AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromGraphics(300, 200, graphics, @bitmap)

   ' // Create a graphics object from the bitmap
   DIM bmpGraphics AS GpGraphics PTR
   hStatus = GdipGetImageGraphicsContext(bitmap, @bmpGraphics)

   ' // Draw something on the bitmap
   DIM brush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_GREEN, @brush)
   hStatus = GdipFillRectangle(bmpGraphics, brush, 50, 50, 200, 100)

   ' // Save the bitmap as a .png
   DIM clsid AS CLSID = AfxGdipGetEncoderClsid("image/png")  ' Helper function to get PNG encoder
   GdipSaveImageToFile(bitmap, "output.png", @clsid, NULL)

   ' // Close the printer object
   IF hPrinter THEN ClosePrinter(hPrinter)

   ' // Draw message on the window
   DIM fontFamily AS GpFontFamily PTR, font AS GpFont PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPoint, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   DIM textBrush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @textBrush)

   DIM rcf AS GpRectF
   rcf.x = 50 : rcf.y = 80 : rcf.Width = 350 : rcf.Height = 90
   DIM wszText AS WSTRING * 128 = "Image saved to output.png"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, textBrush)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF bmpGraphics THEN GdipDeleteGraphics(bmpGraphics)
   IF bitmap THEN GdipDisposeImage(bitmap)
   IF graphics THEN GdipDeleteGraphics(graphics)

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
   pWindow.Create(NULL, "GDI+ GdipCreateFromHDC2", @WndProc)
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
   Example_CreateFromHDC2(hdc)

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
      ' // creation of the window. If the application returns �1, the window is destroyed
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
