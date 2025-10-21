' ########################################################################################
' Microsoft Windows
' File: GdipGetMatrixElements.bas
' Contents: GDI+ Flat API - GdipGetMatrixElements example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
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
' Retrieves the six elements of a transformation matrix. This is useful for inspecting or
' debugging the current transformation state.
' ========================================================================================
SUB Example_GetMatrixElements (BYVAL hdc AS HDC)

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
   DIM ryRatio AS SINGLE = dpiY / 96
   GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)

   ' // First translate
   hStatus = GdipTranslateMatrix(matrix, 40 * rxRatio, 30 * ryRatio, MatrixOrderAppend)
   ' // Then rotate
   hStatus = GdipRotateMatrix(matrix, 45, MatrixOrderAppend)

   ' // Retrieve matrix elements
   DIM elements(0 TO 5) AS SINGLE
   hStatus = GdipGetMatrixElements(matrix, @elements(0))

   ' // Create font and brush
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' // Display matrix values
   DIM text AS STRING
   text = "Matrix Elements:" & !"\n"
   text &= "m11 = " & STR(elements(0)) & !"\n"
   text &= "m12 = " & STR(elements(1)) & !"\n"
   text &= "m21 = " & STR(elements(2)) & !"\n"
   text &= "m22 = " & STR(elements(3)) & !"\n"
   text &= "dx  = " & STR(elements(4)) & !"\n"
   text &= "dy  = " & STR(elements(5))
   DIM layout AS GpRectF = (10.0, 10.0, 300.0, 200.0)
   GdipDrawString(graphics, text, LEN(text), font, @layout, NULL, brush)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
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
   pWindow.Create(NULL, "GDI+ GdipGetMatrixElements", @WndProc)
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
   Example_GetMatrixElements(hdc)

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
