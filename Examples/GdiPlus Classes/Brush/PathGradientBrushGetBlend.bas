' ########################################################################################
' Microsoft Windows
' File: PathGradientBrushGetBlend.bas
' Contents: GDI+ - PathGradientBrushGetBlend example
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
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.SetBlend, PathGradientBrush.GetBlendCount, and PathGradientBrush.GetBlend.
' The code creates a PathGradientBrush object and calls the PathGradientBrush.SetBlend method
' to establish a set of blend factors and blend positions for the brush. Then the code calls
' the PathGradientBrush.GetBlendCount method to retrieve the number of blend factors. After
' the number of blend factors is retrieved, the code allocates two buffers: one to receive
' the array of blend factors and one to receive the array of blend positions. Then the code
' calls the PathGradientBrush.GetBlend method to retrieve the blend factors and the blend
' positions.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(ARGB_BLUE)

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {ARGB_RED}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}
   pthGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 100, 70, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM blendCount AS LONG = pthGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   pthGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugStringW WSTR(rgFactors(j)) & wSTR(rgPositions(j))
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
   pWindow.Create(NULL, "GDI+ PathGradientBrushGetBlend", @WndProc)
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
   Example_GetBlend(hdc)

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
