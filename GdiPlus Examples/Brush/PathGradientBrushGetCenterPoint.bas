' ########################################################################################
' Microsoft Windows
' File: PathGradientBrushGetCenterPoint.bas
' Contents: GDI+ - PathGradientBrushGetCenterPoint example
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
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.GetCenterPoint and PathGradientBrush.SetCenterColor. The code creates
' a PathGradientBrush object and then sets the brush's center color and boundary color.
' The code calls the PathGradientBrush.GetCenterPoint method to determine the center point
' of the path gradient and then draws a line from the origin to that center point.
' ========================================================================================
SUB Example_GetCenterPoint (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM centerPoint AS GpPointF
   pthGrBrush.GetCenterPoint(@centerPoint)

   ' // Draw a line from the origin to the center of the ellipse.
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM pt AS GpPointF = GDIP_POINTF(0, 0)
   graphics.DrawLine(@pen, @pt, @centerPoint)

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
   pWindow.CreateOverlapped(NULL, "GDI+ PathGradientBrushGetCenterPoint", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_WHITE
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_GetCenterPoint(hdc)

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
