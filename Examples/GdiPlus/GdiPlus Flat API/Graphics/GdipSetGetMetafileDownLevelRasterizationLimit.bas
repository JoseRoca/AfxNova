' ########################################################################################
' Microsoft Windows
' File: GdipSetGetMetafileDownLevelRasterizationLimit.bas
' Contents: GDI+ Flat API - GdipSetGetMetafileDownLevelRasterizationLimit example
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
' Record Metafile and Set/Get Rasterization Limit
' Metafile is recorded, not loaded—required for setting the rasterization limit.
' Limit is set before drawing, which ensures it affects rasterization.
' Saved using GdipGetHemfFromMetafile, not GdipSaveImageToFile.
' ========================================================================================
SUB Example_MetafileDownLevelRasterizationLimit (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Define the frame rectangle
   DIM rcfFrame AS GpRectF = (0.0, 0.0, 300.0, 200.0)
   DIM description AS WSTRING * 64 = "Rasterization limit test"

   ' // Record a new EMF+ metafile
   DIM metafile AS GpMetafile PTR
   hStatus = GdipRecordMetafile(hdc, EmfTypeEmfPlusDual, @rcfFrame, MetafileFrameUnitPixel, @description, @metafile)
   IF hStatus <> 0 OR metafile = NULL THEN
      AfxMsg "Failed to record metafile: " & STR(hStatus)
      EXIT SUB
   END IF

   ' // Set rasterization limit (must be done before drawing)
   DIM dpiLimit AS UINT = 150
   hStatus = GdipSetMetafileDownLevelRasterizationLimit(metafile, dpiLimit)
   IF hStatus <> 0 THEN
      AfxMsg "Failed to set rasterization limit: " & STR(hStatus)
   END IF

   ' // Draw into the metafile
   DIM graphics AS GpGraphics PTR
   hStatus = GdipGetImageGraphicsContext(metafile, @graphics)
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2.0, UnitPixel, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, 50.0, 50.0, 200.0, 100.0)

   ' // Flush graphics
   hStatus = GdipFlush(graphics, FlushIntentionFlush)

   ' // Retrieve rasterization limit
   DIM currentLimit AS UINT
   hStatus = GdipGetMetafileDownLevelRasterizationLimit(metafile, @currentLimit)
   IF hStatus = 0 THEN
      AfxMsg "Current rasterization limit: " & STR(currentLimit) & " DPI"
   ELSE
      AfxMsg "Failed to get rasterization limit: " & STR(hStatus)
   END IF

   ' // Save metafile to disk
   DIM hEmf AS HENHMETAFILE
   hStatus = GdipGetHemfFromMetafile(metafile, @hEmf)
   IF hEmf THEN
      DIM saved AS HENHMETAFILE = CopyEnhMetaFile(hEmf, "RasterizedOutput.emf")
      IF saved THEN DeleteEnhMetaFile(saved)
      DeleteEnhMetaFile(hEmf)
   END IF

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)
   IF metafile THEN GdipDisposeImage(metafile)

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
   pWindow.Create(NULL, "GDI+ GdipSetGetMetafileDownLevelRasterizationLimit", @WndProc)
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
   Example_MetafileDownLevelRasterizationLimit(hdc)

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
