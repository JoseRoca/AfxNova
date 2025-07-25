' ########################################################################################
' Microsoft Windows
' File: DrawClosedCurve.bas
' Contents: GDI+ - DrawClosedCurve example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CGdiPlus.inc"
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
' The following example draws a closed cardinal spline.
' ========================================================================================
SUB Example_DrawClosedCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Define a Pen object and an array of PointF objects.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM point1 AS GpPointF : point1.x = 100.0 : point1.y = 100.0
   DIM point2 AS GpPointF : point2.x = 200.0 : point2.y = 50.0
   DIM point3 AS GpPointF : point3.x = 400.0 : point3.y = 10.0
   DIM point4 AS GpPointF : point4.x = 500.0 : point4.y = 100.0
   DIM point5 AS GpPointF : point5.x = 600.0 : point5.y = 200.0
   DIM point6 AS GpPointF : point6.x = 700.0 : point6.y = 400.0
   DIM point7 AS GpPointF : point7.x = 500.0 : point7.y = 500.0

   DIM curvePoints(6) AS GpPointF
   curvePoints(0) = point1
   curvePoints(1) = point2
   curvePoints(2) = point3
   curvePoints(3) = point4
   curvePoints(4) = point5
   curvePoints(5) = point6
   curvePoints(6) = point7

   ' // Draw the closed curve.
   graphics.DrawClosedCurve(@greenPen, @curvePoints(0), 7, 1.0)

   ' // Draw the points in the curve.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 95, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 495, 10, 10)
   graphics.FillEllipse(@redBrush, 195, 45, 10, 10)
   graphics.FillEllipse(@redBrush, 395, 5, 10, 10)
   graphics.FillEllipse(@redBrush, 595, 195, 10, 10)
   graphics.FillEllipse(@redBrush, 695, 395, 10, 10)

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
   pWindow.CreateOverlapped(NULL, "GDI+ DrawClosedCurve", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(800, 550)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_WHITE
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_DrawClosedCurve(hdc)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

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
