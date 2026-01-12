' ########################################################################################
' Microsoft Windows
' File: Gdip_PathIterNextMarkerPath.bas
' Contents: GDI+ Flat API - GdipPathIterNextMarkerPath example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2026 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
'#define _GDIP_DEBUG_ 1
#INCLUDE ONCE "AfxNova/AfxGdipObjects.inc"
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
' Example: Using GdipPathIterNextMarkerPath to Extract and Draw Marker Sections.
' Lets you isolate and render each marker-defined section independently.
' Useful for animations, selective editing, or exporting segments.
' ========================================================================================
SUB Example_PathIterNextMarkerPath (BYVAL hdc AS HDC)

   DIM status AS GpStatus

   ' // Create a graphics object from the device context
   DIM graphics AS GdiPlusGraphics = hdc
   ' // Set the scale transform
   DIM dpiRatio AS SINGLE = graphics.DpiRatio
   status = graphics.ScaleTransform(dpiRatio)

   ' Create a GraphicsPath with two marker sections
   DIM path AS GdiPlusGraphicsPath = FillModeAlternate

   ' First section: triangle
   status = GdipAddPathLine(*path, 30, 30, 130, 30)
   status = GdipAddPathLine(*path, 130, 30, 80, 100)
   status = GdipClosePathFigure(*path)
   status = GdipSetPathMarker(*path)

   ' Second section: zigzag
   status = GdipStartPathFigure(*path)
   status = GdipAddPathLine(*path, 160, 40, 210, 90)
   status = GdipAddPathLine(*path, 210, 90, 160, 140)
   status = GdipSetPathMarker(*path)

   ' Create PathIterator
   DIM iterator AS GdiPlusPathIterator = *path

   ' Create pen for drawing
   DIM pen AS GdiPlusPen = GdiPlusPen(ARGB_BLUE, 2.0, UnitPixel)

   ' Create font and brush for text
   DIM fontFamily AS GdiPlusFontFamily ="Arial"
   DIM font AS GdiPlusFont = GdiPlusFont(*fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel)
   DIM brush AS GdiPlusSolidBrush = ARGB_BLACK

   ' Extract and draw each marker section
   DIM resultCount AS LONG
   DIM markerPath AS GpPath PTR
   DIM yOffset AS SINGLE = 170
   DIM markerIndex AS LONG = 1

   DO
      status = GdipCreatePath(FillModeAlternate, @markerPath)
      status = GdipPathIterNextMarkerPath(*iterator, @resultCount, markerPath)
      IF resultCount = 0 THEN EXIT DO
      ' Draw the marker path
      status = GdipDrawPath(*graphics, *pen, markerPath)
      ' Display marker info
      DIM info AS STRING = "Marker " & markerIndex & ": Points=" & resultCount
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      status = GdipDrawString(*graphics, info, -1, *font, @layout, NULL, *brush)
      yOffset += 20.0
      markerIndex += 1
      IF markerPath THEN GdipDeletePath(markerPath)
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
   pWindow.Create(NULL, "GDI+ GdipPathIterNextMarkerPath", @WndProc)
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
   Example_PathIterNextMarkerPath(hdc)

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
