' ########################################################################################
' Microsoft Windows
' File: GdipGetCustomLineCapType.bas
' Contents: GDI+ Flat API - GdipGetCustomLineCapType example
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
' The following example creates two AdjustableArrowCap objects, arrowCapStart and
' arrowCapEnd, and sets the fill mode to TRUE. The code then creates a Pen object and
' assigns arrowCapStart as the starting line cap for this Pen object and arrowCapEnd as
' the ending line cap. Next, draws a line.
' ========================================================================================
SUB Example_LineCapType (BYVAL hdc AS HDC)

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
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create an AdjustableArrowCap that is filled.
   DIM arrowCapStart AS GpAdjustableArrowCap PTR
   hStatus = GdipCreateAdjustableArrowCap(10, 10, TRUE, @arrowCapStart)
   ' // Adjust to DPI by setting the scale width
   hStatus = GdipSetCustomLineCapWidthScale(arrowCapStart, rxRatio)

   ' // Get the type of CustomLineCap
   ' // It will return 1 (CustomLineCapTypeAdjustableArrow)
   DIM nType AS CustomLineCapType
   hStatus = GdipGetCustomLineCapType(arrowCapStart, @nType)
   AfxMsg "arrowCapStart type = " & WSTR(nType)

   ' // Create an AdjustableArrowCap that is not filled.
   DIM arrowCapEnd AS GpAdjustableArrowCap PTR
   hStatus = GdipCreateAdjustableArrowCap(15, 15, FALSE, @arrowCapEnd)
   ' // Adjust to DPI by setting the scale width
   hStatus = GdipSetCustomLineCapWidthScale(arrowCapEnd, rxRatio)

   ' // Get the type of CustomLineCap
   ' // It will return 1 (CustomLineCapTypeAdjustableArrow)
   hStatus = GdipGetCustomLineCapType(arrowCapEnd, @nType)
   AfxMsg "arrowCapEnd type = " & WSTR(nType)

   ' // Create a Pen
   DIM arrowPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_Violet, 1, UnitWorld, @arrowPen)

   ' Assign arrowCapStart as the start cap.
   hStatus = GdipSetPenCustomStartCap(arrowPen, arrowCapStart)
   ' Assign arrowCapEnd as the end cap.
   hStatus = GdipSetPenCustomEndCap(arrowPen, arrowCapEnd)

   ' // Draw a line using arrowPen.
   hStatus = GdipDrawLine(graphics, arrowPen, 0, 0, 100, 100)

   ' // Cleanup
   IF arrowCapStart THEN GdipDeleteCustomLineCap(arrowCapStart)
   IF arrowCapEnd THEN GdipDeleteCustomLineCap(arrowCapEnd)
   IF arrowPen THEN GdipDeletePen(arrowPen)
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
   pWindow.Create(NULL, "GDI+ GdipGetCustomLineCapType", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_FLORALWHITE
   ' // Anchor the control
   pWindow.AnchorControl(pGraphCtx.hWindow, AFX_ANCHOR_HEIGHT_WIDTH)
   
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc

   ' // Initialize GDI+
   DIM token AS ULONG_PTR = AfxGdipInit

   ' // Draw the graphics
   Example_LineCapType(hdc)

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
