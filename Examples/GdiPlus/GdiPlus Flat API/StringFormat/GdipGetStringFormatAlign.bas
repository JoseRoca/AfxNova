' ########################################################################################
' Microsoft Windows
' File: GdipGetStringFormatAlign.bas
' Contents: GDI+ Flat API - GdipGetStringFormatAlign example
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
' The following example creates a StringFormat object, sets the string's character alignment,
' gets the alignment value, and stores it in a variable. The code then creates a second
' StringFormat object and uses the stored alignment value to set the character alignment of
' the second StringFormat object. Next, the code uses the second StringFormat object to draw
' a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetStringFormatAlign (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatAlign(format, StringAlignmentCenter)

   ' // Get the alignment setting from the StringFormat object.
   DIM strAlignment AS StringAlignment 
   hStatus = GdipGetStringFormatAlign(format, @strAlignment)

   ' // Create a second StringFormat object with the same alignment.
   DIM format2 AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, format2)
   hStatus = GdipSetStringFormatAlign(format2, strAlignment)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Draw the string
   DIM wszText AS WSTRING * 64 = "This text was formatted by a second StringFormat object."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format2, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format2 THEN GdipDeleteStringFormat(format2)
   IF format THEN GdipDeleteStringFormat(format)
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
   pWindow.Create(NULL, "GDI+ GdipGetStringFormatAlign", @WndProc)
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
   Example_GetStringFormatAlign(hdc)

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
