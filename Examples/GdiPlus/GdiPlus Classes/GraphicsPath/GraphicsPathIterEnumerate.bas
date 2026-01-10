' ########################################################################################
' Microsoft Windows
' File: GraphicsPathIterEnumerate.bas
' Contents: GDI+ - GraphicsPathIterEnumerate example
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
' Using GdipPathIterEnumerate to Extract All Path Data.
' Builds a path with two figures.
' Uses GdipPathIterEnumerate to extract all points and types.
' Draws the path and overlays each point’s coordinates and type.
' ========================================================================================
SUB Example_PathIterEnumerate (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scale transform
   graphics.ScaleTransformForDpi

   ' // Create a GraphicsPath with two figures
   DIM path AS CGpGraphicsPath
   path.AddLine(20, 20, 120, 20)
   path.AddLine(120, 20, 70, 100)
   path.CloseFigure

   path.StartFigure
   path.AddLine(150, 30, 200, 80)
   path.AddLine(200, 80, 150, 130)

   ' // Create a GraphicsPathIterator
   DIM iterator AS CGpGraphicsPathIterator = CGpGraphicsPathIterator(@path)

   ' // Get total number of points in the path
   DIM totalCount AS LONG = iterator.GetCount

   ' // Allocate arrays for points and types
   DIM points(0 TO totalCount - 1) AS GpPointF
   DIM types(0 TO totalCount - 1) AS BYTE

   ' // Enumerate all path data
   DIM resultCount AS LONG = iterator.Enumerate(@points(0), @types(0), totalCount)

   ' Draw the path using a blue pen
   DIM pen AS CGpPen = CGpPen(ARGB_BLUE, 2.0)
   graphics.DrawPath(@pen, @path)

   ' Create font and brush for text output
   DIM fontFamily AS CGpFontFamily = CGpFontFamily("Arial")
   DIM font AS CGpFont = CGpFont(@fontFamily, AfxGdipPointsToPixels(9, TRUE), FontStyleRegular)
   DIM brush AS CGpSolidBrush = ARGB_BLACK

   ' Display point coordinates and type below the drawing
   DIM yOffset AS SINGLE = 150
   FOR i AS LONG = 0 TO resultCount - 1
      DIM info AS STRING = "Point " & i & ": (" & points(i).x & ", " & points(i).y & ") Type=" & types(i)
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      graphics.DrawString(info, -1, @font, @layout, @brush)
      yOffset += 20.0
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
   pWindow.Create(NULL, "GDI+ GraphicsPathIterEnumerate", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 300)
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
   Example_PathIterEnumerate(hdc)

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
