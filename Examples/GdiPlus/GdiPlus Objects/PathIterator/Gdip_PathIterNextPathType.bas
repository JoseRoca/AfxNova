' ########################################################################################
' Microsoft Windows
' File: Gdip_PathIterNextPathType.bas
' Contents: GDI+ Flat API - GdipPathIterNextPathType example
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
' Example: Using GdipPathIterNextPathType to Group Points by Type.
' Helps you understand how a path is constructed—lines vs. curves.
' Enables selective rendering, editing, or flattening based on segment type.
' Returns StatusOk, but resultCount 0.
' ========================================================================================
SUB Example_PathIterNextPathType (BYVAL hdc AS HDC)

   DIM status AS GpStatus

   ' // Create a graphics object from the device context
   DIM graphics AS GdiPlusGraphics = hdc
   ' // Set the scale transform
   DIM dpiRatio AS SINGLE = graphics.DpiRatio
   status = graphics.ScaleTransform(dpiRatio)

   ' Create a GraphicsPath with grouped types
   DIM path AS GdiPlusGraphicsPath = FillModeAlternate

   ' Add 3 lines (same type group)
   status = GdipAddPathLine(*path, 20, 20, 120, 20)
   status = GdipAddPathLine(*path, 120, 20, 70, 100)
   status = GdipAddPathLine(*path, 70, 100, 20, 20)

   ' Add Bézier curve (new type group)
   status = GdipAddPathBezier(*path, 150, 30, 180, 10, 210, 50, 240, 30)

   ' Create PathIterator
   DIM iterator AS GdiPlusPathIterator = *path

   ' Create font and brush
   DIM fontFamily AS GdiPlusFontFamily ="Arial"
   DIM font AS GdiPlusFont = GdiPlusFont(*fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel)
   DIM brush AS GdiPlusSolidBrush = ARGB_BLACK

   ' Iterate through path types
   DIM resultCount AS LONG, pathType AS BYTE
   DIM startIdx AS LONG, endIdx AS LONG
   DIM yOffset AS SINGLE = 10.0
   DIM segmentIndex AS LONG = 1

   DO
      status = GdipPathIterNextPathType(*iterator, @resultCount, @pathType, @startIdx, @endIdx)
      IF resultCount = 0 THEN EXIT DO
      DIM typeName AS STRING
      SELECT CASE pathType AND &H07
         CASE PathPointTypeStart
            typeName = "Start"
         CASE PathPointTypeLine
            typeName = "Line"
         CASE PathPointTypeBezier
            typeName = "Bezier"
         CASE ELSE
            typeName = "Unknown"
      END SELECT
      DIM info AS STRING
      info = "Segment " & segmentIndex & ": Type=" & typeName & ", Start=" & startIdx & ", End=" & endIdx
      DIM layout AS GpRectF = (10.0, yOffset, 400.0, 20.0)
      status = GdipDrawString(*graphics, info, -1, *font, @layout, NULL, *brush)
      yOffset += 20.0
      segmentIndex += 1
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
   pWindow.Create(NULL, "GDI+ GdipPathIterNextPathType", @WndProc)
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
   Example_PathIterNextPathType(hdc)

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
