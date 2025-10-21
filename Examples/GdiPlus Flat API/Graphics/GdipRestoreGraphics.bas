' ########################################################################################
' Microsoft Windows
' File: Example_RestoreGraphics.bas
' Contents: GDI+ Flat API - Example_RestoreGraphics example
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
' Restoring Only the First of Two Nested Saved States
' The following example sets the world transformation of a Graphics object to a rotation and
' then saves the state of the Graphics object. Next, the code calls GdipTranslateWorldTransform,
' and saves the state again. Then the code calls ScaleTransform. At that point, the world
' transformation of the Graphics object is a composite transformation: first rotate, then
' translate, then scale. The code uses a red pen to draw an ellipse that is transformed by
' that composite transformation.
' The code passes state1, which was returned by the first call to Save, to the GdipRestoreGraphics
' function, and draws the ellipse again using a green pen. The green ellipse is rotated but
' not translated or scaled.
' Next the code attempts to restore the state identified by state2. The attempt fails because
' the call GdipRestoreGraphics(state1) removed the information blocks identified by both
' state1 and state2 from the stack.
' ========================================================================================
SUB Example_RestoreGraphics (BYVAL hdc AS HDC)

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

   ' // Apply rotation
   hStatus = GdipRotateWorldTransform(graphics, 30.0, MatrixOrderPrepend)
   DIM state1 AS GraphicsState
   hStatus = GdipSaveGraphics(graphics, @state1)

   ' // Apply translation
   hStatus = GdipTranslateWorldTransform(graphics, 100.0, 0.0, MatrixOrderAppend)
   DIM state2 AS GraphicsState
   hStatus = GdipSaveGraphics(graphics, @state2)

   ' // Apply scaling
   hStatus = GdipScaleWorldTransform(graphics, 1.0, 3.0, MatrixOrderAppend)

   ' // Draw red ellipse (rotate + translate + scale)
   DIM redPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1.0, UnitWorld, @redPen)
   hStatus = GdipDrawEllipse(graphics, redPen, 0, 0, 100, 20)

   ' // Restore to state1 and draw the ellipse again. 
   ' // Only the rotation transformation applies.
     hStatus = GdipRestoreGraphics(graphics, state1)
   DIM greenPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_GREEN, 1.0, UnitWorld, @greenPen)
   hStatus = GdipDrawEllipse(graphics, greenPen, 0, 0, 100, 20)

   ' // The information block identified by state2 has been lost.
   ' // The following call to Restore has no effect because
   ' // GdipRestoreGraphics(state1) removed from the stack the
   ' // information blocks identified by state1 and state2.
     hStatus = GdipRestoreGraphics(graphics, state2)
   DIM bluePen AS GpPen PTR

   ' // The Graphics object is still in the state identified by state1.
   ' // The following code draws a blue ellipse on top of the previously
   ' // drawn green ellipse.
   hStatus = GdipCreatePen1(ARGB_BLUE, 1.0, UnitWorld, @bluePen)
   hStatus = GdipDrawEllipse(graphics, bluePen, 0, 0, 100, 20)

   ' // Cleanup
   IF bluePen THEN GdipDeletePen(bluePen)
   IF greenPen THEN GdipDeletePen(greenPen)
   IF redPen THEN GdipDeletePen(redPen)
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
   pWindow.Create(NULL, "GDI+ GdipSaveGraphics2", @WndProc)
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
   Example_RestoreGraphics(hdc)

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
