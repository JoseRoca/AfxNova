' ########################################################################################
' Microsoft Windows
' File: GraphicsPathWarp.bas
' Contents: GDI+ - GraphicsPathWarp example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
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
' The following example creates a GraphicsPath object and adds a closed figure to the path.
' The code defines a warp transformation by specifying a source rectangle and an array of
' four destination points. The source rectangle and destination points are passed to the
' Warp method. The code draws the path twice: once before it has been warped and once after
' it has been warped.
' ========================================================================================
SUB Example_Warp (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scale transform
   graphics.ScaleTransformForDpi

   ' // Create a path.
   DIM points(0 TO 7) AS GpPointF = {(20.0, 60.0), (30.0, 90.0), _
      (15.0, 110.0), (15.0, 145.0), (55.0, 145.0), (55.0, 110.0), (40.0, 90.0), (50.0, 60.0)}
   DIM path AS CGpGraphicsPath
   path.AddLines(@points(0), 8)
   path.CloseFigure

   ' // Draw the path before applying a warp transformation.
   DIM bluePen AS CGpPen = ARGB_RED
   graphics.DrawPath(@bluePen, @path)

   ' // Define a warp transformation, and warp the path.
   DIM srcRect AS GpRectF = (10.0, 50.0, 50.0, 100.0)
   DIM destPts(0 TO 3) AS GpPointF = {(220.0, 10.0), (280.0, 10.0), (100.0, 150.0), (400.0, 150.0)}
   path.Warp(@destPts(0), 4, @srcRect)

   ' // Draw the warped path.
   graphics.DrawPath(@bluePen, @path)

   ' // Draw the source rectangle and the destination polygon.
   DIM blackPen AS CGpPen = ARGB_BLACK
   graphics.DrawRectangle(@blackPen, @srcRect)
   graphics.DrawLine(@blackPen, @destPts(0), @destPts(1))
   graphics.DrawLine(@blackPen, @destPts(0), @destPts(2))
   graphics.DrawLine(@blackPen, @destPts(1), @destPts(3))
   graphics.DrawLine(@blackPen, @destPts(2), @destPts(3))

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
   pWindow.Create(NULL, "GDI+ GraphicsPathWarp", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(420, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_WHITE
   ' // Anchor the control
   pWindow.AnchorControl(pGraphCtx.hWindow, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_Warp(hdc)

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
