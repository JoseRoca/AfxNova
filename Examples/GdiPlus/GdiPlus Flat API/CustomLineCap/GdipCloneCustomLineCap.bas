' ########################################################################################
' Microsoft Windows
' File: GdipCloneCustomLineCap.bas
' Contents: GDI+ Flat API - GdipCloneCustomLineCap example
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
' The following example creates a CustomLineCap object with a stroke path, creates a second
' CustomLineCap object by cloning the first, and then assigns the cloned cap to a Pen object.
' It then draws a line by using the Pen object.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the window device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE, dpiY AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM ryRatio AS SINGLE = dpiY / 96

   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create the first CustomLineCap object
   DIM firstCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapFlat, 0, @firstCap)

   ' // Clone the first CustomLineCap into secondCap
   DIM secondCap AS GpCustomLineCap PTR
   hStatus = GdipCloneCustomLineCap(firstCap, @secondCap)

   ' // Create a violet Pen with width 1
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_VIOLET, 1, UnitWorld, @pen)

   ' // Assign the cloned cap as the end cap of the pen
   hStatus = GdipSetPenCustomEndCap(pen, secondCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 100, 100)

   ' // Clean up resources
   IF pen THEN GdipDeletePen(pen)
   IF secondCap THEN GdipDeleteCustomLineCap(secondCap)
   IF firstCap THEN GdipDeleteCustomLineCap(firstCap)
   IF capPath THEN GdipDeletePath(capPath)
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
   pWindow.Create(NULL, "GDI+ GdipCloneCustomLineCap", @WndProc)
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
   Example_Clone(hdc)

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
