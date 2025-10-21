' ########################################################################################
' Microsoft Windows
' File: GdipVectorTransformMatrixPoints.bas
' Contents: GDI+ Flat API - GdipVectorTransformMatrixPoints example
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
' The following example creates a vector and a point. The tip of the vector and the point
' are at the same location: (100, 50). The code creates a Matrix object and initializes its
' elements so that it represents a clockwise rotation followed by a translation 100 units
' to the right. The code calls the GdipTransformMatrixPoints functioon of the matrix
' to transform the point and calls the GdipVectorTransformMatrixPoints function to transform
' the vector. The entire transformation (rotation followed by translation) is performed on
' the point, but only the rotation part of the transformation is performed on the vector.
' The elements of the matrix that represent translation are ignored by the
' GdipVectorTransformMatrixPoints function
' ========================================================================================
SUB Example_VectorTransformMatrixPoints (BYVAL hdc AS HDC)

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
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2 * rxRatio, UnitPixel, @pen)
   hStatus = GdipSetPenEndCap(pen, LineCapArrowAnchor)
   
   ' // Create a solid brush
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @brush)
   
   ' // A point and a vector, same representation but different behavior
   DIM pt AS GpPointF = (100, 50)
   DIM vector AS GpPointF = (100, 50)

   ' // Draw the original point and vector in blue
   hStatus = GdipFillEllipse(graphics, brush, pt.x - 5.0, pt.y - 5.0, 10.0, 10.0)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, vector.x, vector.y)

   ' // Transform
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(0.8, 0.6, -0.6, 0.8, 100.0, 0.0,@matrix)
   hStatus = GdipTransformMatrixPoints(matrix, @pt, 1)
   hStatus = GdipVectorTransformMatrixPoints(matrix, @vector, 1)

   ' // Draw the transformed point and vector in red.
   hStatus = GdipSetPenColor(pen, ARGB_RED)
   hStatus = GdipSetSolidFillColor(brush, ARGB_RED)
   hStatus = GdipFillEllipse(graphics, brush, pt.X - 5.0, pt.Y - 5.0, 10.0, 10.0)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, vector.x, vector.y)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF brush THEN GdipDeleteBrush(brush)
   IF matrix THEN GdipDeleteMatrix(matrix)
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
   pWindow.Create(NULL, "GDI+ GdipVectorTransformMatrixPoints", @WndProc)
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
   Example_VectorTransformMatrixPoints(hdc)

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
