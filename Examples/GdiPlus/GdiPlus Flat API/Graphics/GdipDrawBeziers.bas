' ########################################################################################
' Microsoft Windows
' File: GdipDrawBeziers.bas
' Contents: GDI+ Flat API - GdipDrawBeziers example
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
' The following example draws a pair of Bézier curves.
' ========================================================================================
SUB Example_DrawBeziers (BYVAL hdc AS HDC)

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

   ' // Create a green Pen
   DIM greenPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 1, UnitWorld, @greenPen)

   DIM startPoint AS GpPointF = (100, 100)
   DIM ctrlPoint1 AS GpPointF = (200,  50)
   DIM ctrlPoint2 AS GpPointF = (400,  10)
   DIM endPoint1  AS GpPointF = (500, 100)
   DIM ctrlPoint3 AS GpPointF = (600, 200)
   DIM ctrlPoint4 AS GpPointF = (700, 400)
   DIM endPoint2  AS GpPointF = (500, 500)

   DIM curvePoints(6) AS GpPointF
   curvePoints(0) = startPoint
   curvePoints(1) = ctrlPoint1
   curvePoints(2) = ctrlPoint2
   curvePoints(3) = endPoint1
   curvePoints(4) = ctrlPoint3
   curvePoints(5) = ctrlPoint4
   curvePoints(6) = endPoint2

   ' // Draw the Bézier curves
   hStatus = GdipDrawBeziers(graphics, greenPen, @curvePoints(0), 7)

   ' // Create the brushes
   DIM AS GpBrush PTR redBrush, blueBrush
   hStatus = GdipCreateSolidFill(ARGB_RED, @redBrush)
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @blueBrush)

   ' //Draw the end points and control points.
   hStatus = GdipFillEllipse(graphics, redBrush, 100 - 5, 100 - 5, 10, 10)
   hStatus = GdipFillEllipse(graphics, redBrush, 500 - 5, 100 - 5, 10, 10)
   hStatus = GdipFillEllipse(graphics, redBrush, 500 - 5, 500 - 5, 10, 10)
   hStatus = GdipFillEllipse(graphics, blueBrush, 200 - 5, 50 - 5, 10, 10)
   hStatus = GdipFillEllipse(graphics, blueBrush, 400 - 5, 10 - 5, 10, 10)
   hStatus = GdipFillEllipse(graphics, blueBrush, 600 - 5, 200 - 5, 10, 10)
   hStatus = GdipFillEllipse(graphics, blueBrush, 700 - 5, 400 - 5, 10, 10)

   ' // Cleanup
   IF blueBrush THEN GdipDeleteBrush(blueBrush)
   IF redBrush THEN GdipDeleteBrush(redBrush)
   IF greenPen THEN GdipDeletePen(greenPen)
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
   pWindow.Create(NULL, "GDI+ GdipDrawBeziers", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(800, 530)
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
   Example_DrawBeziers(hdc)

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
