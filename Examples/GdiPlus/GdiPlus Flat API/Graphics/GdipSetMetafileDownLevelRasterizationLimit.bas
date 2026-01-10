' ########################################################################################
' Microsoft Windows
' File: GdipSetMetafileDownLevelRasterizationLimit.bas
' Contents: GDI+ Flat API - GdipSetMetafileDownLevelRasterizationLimit example
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
' Compare Low vs High DPI Rasterization.
' Low DPI version will look pixelated or blurry, especially with gradients or textures.
' High DPI version will appear crisp and smooth, but may result in larger file size.
' ========================================================================================
SUB Example_SetMetafileDownLevelRasterizationLimit (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   DIM rcfFrame AS GpRectF = (0.0, 0.0, 300.0, 200.0)
   DIM description AS WSTRING * 64 = "Rasterization test"

   ' --- Low DPI Metafile ---
   DIM metafileLow AS GpMetafile PTR
   hStatus = GdipRecordMetafile(hdc, EmfTypeEmfPlusDual, @rcfFrame, MetafileFrameUnitPixel, @description, @metafileLow)
   GdipSetMetafileDownLevelRasterizationLimit(metafileLow, 72)  ' Low quality

   DIM graphicsLow AS GpGraphics PTR
   GdipGetImageGraphicsContext(metafileLow, @graphicsLow)
   DIM brush AS GpBrush PTR
   GdipCreateSolidFill(ARGB_GREEN, @brush)
   GdipFillRectangle(graphicsLow, brush, 50.0, 50.0, 200.0, 100.0)
   GdipFlush(graphicsLow, FlushIntentionFlush)
   GdipDeleteGraphics(graphicsLow)

   ' --- High DPI Metafile ---
   DIM metafileHigh AS GpMetafile PTR
   hStatus = GdipRecordMetafile(hdc, EmfTypeEmfPlusDual, @rcfFrame, MetafileFrameUnitPixel, @description, @metafileHigh)
   GdipSetMetafileDownLevelRasterizationLimit(metafileHigh, 300)  ' High quality

   DIM graphicsHigh AS GpGraphics PTR
   GdipGetImageGraphicsContext(metafileHigh, @graphicsHigh)
   GdipFillRectangle(graphicsHigh, brush, 50.0, 50.0, 200.0, 100.0)
   GdipFlush(graphicsHigh, FlushIntentionFlush)
   GdipDeleteGraphics(graphicsHigh)

   ' --- Playback ---
   DIM playback AS GpGraphics PTR
   GdipCreateFromHDC(hdc, @playback)

   ' Label and draw low DPI version
   DIM rcf AS GpRectF = (10, 10, 100, 20)
   GdipDrawString(playback, "Low DPI", -1, NULL, @rcf, NULL, brush)
   GdipDrawImage(playback, metafileLow, 10.0, 30.0)

   ' Label and draw high DPI version
   DIM rcf2 AS GpRectF = (320, 10, 100, 20)
   GdipDrawString(playback, "High DPI", -1, NULL, @rcf2, NULL, brush)
   GdipDrawImage(playback, metafileHigh, 320.0, 30.0)

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF playback THEN GdipDeleteGraphics(playback)
   IF metafileLow THEN GdipDisposeImage(metafileLow)
   IF metafileHigh THEN GdipDisposeImage(metafileHigh)

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
   pWindow.Create(NULL, "GDI+ GdipSetMetafileDownLevelRasterizationLimit", @WndProc)
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
   Example_SetMetafileDownLevelRasterizationLimit(hdc)

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
