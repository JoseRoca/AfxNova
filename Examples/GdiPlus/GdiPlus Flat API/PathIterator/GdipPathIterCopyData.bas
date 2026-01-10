' ########################################################################################
' Microsoft Windows
' File: GdipPathIterCopyData.bas
' Contents: GDI+ Flat API - GdipPathIterCopyData example
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
' Creates a path with two figures.
' Uses GdipPathIterCopyData to extract the first three points and their types.
' Draws the extracted segment and displays point info in the window.
' ========================================================================================
SUB Example_PathIterCopyData (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with multiple segments
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' Add a triangle
   hStatus = GdipAddPathLine(path, 20, 20, 120, 20)
   hStatus = GdipAddPathLine(path, 120, 20, 70, 100)
   hStatus = GdipClosePathFigure(path)

   ' Add a zigzag line
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 150, 30, 200, 80)
   hStatus = GdipAddPathLine(path, 200, 80, 150, 130)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Copy a subset of the path (e.g., indices 0 to 2)
   DIM resultCount AS LONG
   DIM points(0 TO 2) AS GpPointF
   DIM types(0 TO 2) AS BYTE
   hStatus = GdipPathIterCopyData(iterator, @resultCount, @points(0), @types(0), 0, 2)

   ' Draw the copied segment using a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2.0, UnitPixel, @pen)
   hStatus = GdipDrawLines(graphics, pen, @points(0), resultCount)

   ' Display info about copied points
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM yOffset AS SINGLE = 140
   FOR i AS LONG = 0 TO resultCount - 1
      DIM info AS STRING
      info = "Point " & i & ": (" & points(i).x & ", " & points(i).y & ") Type=" & types(i)
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)
      yOffset += 20.0
   NEXT

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF pen THEN GdipDeletePen(pen)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
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
   pWindow.Create(NULL, "GDI+ GdipPathIterCopyData", @WndProc)
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
   Example_PathIterCopyData(hdc)

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
