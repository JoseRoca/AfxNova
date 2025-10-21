' ########################################################################################
' Microsoft Windows
' File: GdipDrawSpheres.bas
' Contents: GDI+ Flat API - GdipDrawSpheres example
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
' Drawing Spheres with PathGradientBrush
' Creates two ellipses filled with radial gradients to simulate spheres. The center color
' mimics light reflection, while the translucent surround color adds depth and shading.
' Antialiasing and DPI scaling ensure smooth rendering across devices.
' ========================================================================================
SUB Example_DrawSpheres (BYVAL hdc AS HDC)

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

   ' // Create a GraphicsPath object.
   DIM path1 AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path1)

   ' // Add an ellipse to the path.
   DIM x AS LONG = 90
   DIM y AS LONG = 15
   DIM nSize AS LONG = 220
   hStatus = GdipAddPathEllipseI(path1, x, y, nSize, nSize)

   ' // Create a path gradient based on the ellipse.
   DIM brush1 AS GpPathGradient PTR
   hStatus = GdipCreatePathGradientFromPath(path1, @brush1)

  ' // Set the middle color of the path.
   DIM MiddleColorToOpaque AS ARGB = GDIP_ARGB_SetAlpha(ARGB_MediumAquamarine, 0)
   hStatus = GdipSetPathGradientCenterColor(brush1, MiddleColorToOpaque)

   ' // Set the entire path boundary to Alpha Black using 50% translucency.
   DIM count AS LONG = 1
   DIM BlackFullTranslucent AS ARGB = GDIP_ARGB_SetAlpha(ARGB_Black, 128)
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush1, @BlackFullTranslucent, @count)

   ' // Draw the ellipse, keeping the exact coordinates defined for the path,
   ' // and using antialising mode (+ 2 and - 4 are used to better achieve antialiasing)
   hStatus = GdipSetSmoothingMode(graphics, SmoothingModeAntiAlias)
   hStatus = GdipFillEllipseI(graphics, brush1, x + 2, y + 2, nSize - 4, nSize - 4)

   ' // Create a second GraphicsPath object.
   DIM path2 AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, path2)

   ' // Add an ellipse to the path
   x = 200 : y = 200 : nSize = 150
   hStatus = GdipAddPathEllipse(path2, x, y, nSize, nSize)
   
   ' // Create a path gradient based on the ellipse.
   DIM brush2 AS GpPathGradient PTR
   hStatus = GdipCreatePathGradientFromPath(path2, @brush2)

   ' // Set the middle color of the path.
   MiddleColorToOpaque = GDIP_ARGB_SetAlpha(ARGB_Yellow, 64)
   hStatus = GdipSetPathGradientCenterColor(brush2, MiddleColorToOpaque)
   
   ' // Set the entire path boundary to Alpha Black using 50% translucency
   count = 1
   BlackFullTranslucent = GDIP_ARGB_SetAlpha(ARGB_Red, 128)
   GdipSetPathGradientSurroundColorsWithCount(brush2, @BlackFullTranslucent, @count)

   ' // Draw the ellipse, keeping the exact coords we defined for the path
   hStatus = GdipFillEllipseI(graphics, brush2, x + 2, y + 2, nSize - 4, nSize - 4)

   ' // Cleanup
   IF path1 THEN GdipDeletePath(path1)
   IF path2 THEN GdipDeletePath(path2)
   IF brush1 THEN GdipDeleteBrush(brush1)
   IF brush2 THEN GdipDeleteBrush(brush2)
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
   pWindow.Create(NULL, "GDI+ GdipDrawSpheres", @WndProc)
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
   Example_DrawSpheres(hdc)

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
