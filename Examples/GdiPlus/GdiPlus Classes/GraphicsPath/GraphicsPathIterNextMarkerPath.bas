' ########################################################################################
' Microsoft Windows
' File: GraphicsPathIterNextMarkerPath.bas
' Contents: GDI+ - GraphicsPathIterNextMarkerPath example
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
' Using the NextMarker method to extract and draw marker sections.
' Lets you isolate and render each marker-defined section independently.
' Useful for animations, selective editing, or exporting segments.
' ========================================================================================
SUB Example_NextMarkerPath (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scale transform
   graphics.ScaleTransformForDpi

   ' // Create a GraphicsPath with two marker sections
   DIM path AS CGpGraphicsPath
   path.AddLine(30, 30, 130, 30)
   path.AddLine(130, 30, 80, 100)
   path.CloseFigure
   path.SetMarker

   path.StartFigure
   path.AddLine(160, 40, 210, 90)
   path.AddLine(210, 90, 160, 140)
   path.SetMarker

   ' // Create a GraphicsPathIterator
   DIM iterator AS CGpGraphicsPathIterator = CGpGraphicsPathIterator(@path)

   ' // Create pen for drawing
   DIM pen AS CGpPen = CGpPen(ARGB_BLUE, 2.0)

   ' // Create font and brush for text
   DIM fontFamily AS CGpFontFamily = CGpFontFamily("Arial")
   DIM font AS CGpFont = CGpFont(@fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular)
   DIM brush AS CGpSolidBrush = ARGB_BLACK

   ' // Extract and draw each marker-defined section
   DIM resultCount AS LONG
   DIM yOffset AS SINGLE = 170
   DIM markerIndex AS LONG = 1

   DO
      DIM markerPath AS CGpGraphicsPath
      resultCount = iterator.NextMarker(@markerPath)
      IF resultCount = 0 THEN EXIT DO
      ' // Draw the extracted marker path
      graphics.DrawPath(@pen, @markerPath)
      ' // Display marker info
      DIM info AS STRING = "Marker " & markerIndex & ": Points=" & resultCount
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      graphics.DrawString(info, -1, @font, @layout, @brush)
      yOffset += 20.0
      markerIndex += 1
   LOOP

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
   pWindow.Create(NULL, "GDI+ GraphicsPathIterNextMarkerPath", @WndProc)
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
   Example_NextMarkerPath(hdc)

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
