' ########################################################################################
' Microsoft Windows
' File: GdipMeasureCharacterRanges.bas
' Contents: GDI+ Flat API - GdipMeasureCharacterRanges example
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
' The following example defines three ranges of character positions within a string and
' sets those ranges in a StringFormat object. Next, the MeasureCharacterRanges method is
' used to get the three regions of the display that are occupied by the characters that
' are specified by the ranges. This is done for three different layout rectangles to show
' how the regions change according to the layout of the string. Also, on the third repetition
' of this, the string format flags are changed so that the regions measured will include
' trailing spaces.
' ========================================================================================
SUB Example_MeasureCharacterRanges (BYVAL hdc AS HDC)

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

   ' // Create font
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(16, TRUE), FontStyleRegular, UnitPixel, @font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

   ' // Create brushes and pen
   DIM blueBrush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @blueBrush)
   DIM redBrush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(GDIP_ARGB(100, 255, 0, 0), @redBrush)
   DIM blackPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 1.0, UnitPixel, @blackPen)
   hStatus = GdipScalePenTransform(blackPen, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create string format
   DIM stringFormat As GpStringFormat PTR
   hStatus = GdipStringFormatGetGenericDefault(@stringFormat)

   ' // Define character ranges
   DIM charRanges(2) As GpCharacterRange
   charRanges(0).First = 3  : charRanges(0).Length = 5
   charRanges(1).First = 15 : charRanges(1).Length = 2
   charRanges(2).First = 30 : charRanges(2).Length = 15
   hStatus = GdipSetStringFormatMeasurableCharacterRanges(stringFormat, 3, @charRanges(0))

   ' // Allocate regions
   DIM regions(2) AS GpRegion PTR
   FOR i AS LONG = 0 TO 2
      hStatus = GdipCreateRegion(@regions(i))
   NEXT

   ' // Text to draw
   DIM text AS WSTRING * 128 = "The quick, brown fox easily jumps over the lazy dog."

   ' // Measure and draw for layoutRectA
   DIM layoutRectA AS GpRectF = (20.0, 20.0, 130.0, 130.0)
   hStatus = GdipMeasureCharacterRanges(graphics, @text, -1, font, @layoutRectA, stringFormat, 3, @regions(0))
   hStatus = GdipDrawString(graphics, @text, -1, font, @layoutRectA, stringFormat, blueBrush)
   hStatus = GdipDrawRectangle(graphics, blackPen, layoutRectA.X, layoutRectA.Y, layoutRectA.Width, layoutRectA.Height)
   FOR i AS LONG = 0 TO 2
      hStatus = GdipFillRegion(graphics, redBrush, regions(i))
   NEXT

   ' // Measure and draw for layoutRectB
   DIM layoutRectB AS GpRectF = (160.0, 20.0, 165.0, 130.0)
   hStatus = GdipMeasureCharacterRanges(graphics, @text, -1, font, @layoutRectB, stringFormat, 3, @regions(0))
   hStatus = GdipDrawString(graphics, @text, -1, font, @layoutRectB, stringFormat, blueBrush)
   hStatus = GdipDrawRectangle(graphics, blackPen, layoutRectB.X, layoutRectB.Y, layoutRectB.Width, layoutRectB.Height)
   FOR i AS LONG = 0 TO 2
      hStatus = GdipFillRegion(graphics, redBrush, regions(i))
   NEXT

   ' // Set trailing space flag and draw for layoutRectC
   DIM layoutRectC As GpRectF = (335.0, 20.0, 165.0, 130.0)
   hStatus = GdipSetStringFormatFlags(stringFormat, StringFormatFlagsMeasureTrailingSpaces)
   hStatus = GdipMeasureCharacterRanges(graphics, @text, -1, font, @layoutRectC, stringFormat, 3, @regions(0))
   hStatus = GdipDrawString(graphics, @text, -1, font, @layoutRectC, stringFormat, blueBrush)
   hStatus = GdipDrawRectangle(graphics, blackPen, layoutRectC.X, layoutRectC.Y, layoutRectC.Width, layoutRectC.Height)
   FOR i AS LONG = 0 TO 2
      hStatus = GdipFillRegion(graphics, redBrush, regions(i))
   NEXT

   ' // Cleanup
   FOR i AS LONG = 0 TO 2
      IF regions(i) THEN GdipDeleteRegion(regions(i))
   NEXT
   IF font THEN GdipDeleteFont(font)
   IF stringFormat THEN GdipDeleteStringFormat(stringFormat)
   IF blueBrush THEN GdipDeleteBrush(blueBrush)
   IF redBrush THEN GdipDeleteBrush(redBrush)
   IF blackPen THEN GdipDeletePen(blackPen)
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
   pWindow.Create(NULL, "GDI+ GdipMeasureCharacterRanges", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(520, 250)
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
   Example_MeasureCharacterRanges(hdc)

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
      ' // creation of the window. If the application returns 1, the window is destroyed
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

