' ########################################################################################
' Microsoft Windows
' File: PathGradientBrushGetInterpolationColors.bas
' Contents: GDI+ - PathGradientBrushGetInterpolationColors example
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
' The following example creates a PathGradientBrush object from a triangular path. The code
' sets the preset colors to red, blue, and aqua and sets the blend positions to 0, 0.6, and 1.
' The code calls the PathGradientBrush.GetInterpolationColorCount method of the PathGradientBrush
' object to obtain the number of preset colors currently set for the brush. Next, the code
' allocates two buffers: one to hold the array of preset colors, and one to hold the array
' of blend positions. The call to the PathGradientBrush.GetInterpolationColors method of
' the PathGradientBrush object fills the buffers with the preset colors and the blend positions.
' Finally the code fills a small square with each of the preset colors.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 255)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.6, 1.0}

   pthGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the path gradient brush.
   DIM colorCount AS LONG = pthGrBrush.GetInterpolationColorCount
   DIM rgColors(colorCount - 1) AS ARGB
   DIM rgPositions(colorCount - 1) AS SINGLE
   pthGrBrush.GetInterpolationColors(@rgColors(0), @rgPositions(0), colorCount)

   ' // Fill a small square with each of the interpolation colors.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 255, 255)

   FOR j AS LONG = 0 TO colorCount - 1
      solidBrush.SetColor(rgColors(j))
      graphics.FillRectangle(@solidBrush, 15 * j, 0, 10, 10)
   NEXT

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
   pWindow.Create(NULL, "GDI+ PathGradientBrushGetInterpolationColors", @WndProc)
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
   ' // Draw the graphics
   Example_GetInterpolationColors(hdc)

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
