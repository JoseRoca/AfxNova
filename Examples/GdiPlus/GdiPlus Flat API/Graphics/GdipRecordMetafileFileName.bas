' ########################################################################################
' Microsoft Windows
' File: GdipRecordMetafileFileName.bas
' Contents: GDI+ Flat API - GdipRecordMetafileFileName example
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
' The following example creates a Metafile. The code uses the Metafile to record a sequence
' of graphics commands and then saves the recorded commands in a file named MetafileTest.emf.
' Note that the function that creates the metafile (GdipRecordMetafileFileName) receives a
' device context handle, and we retrieve a Graphics handle calling the GdipGetImageGraphicsContext
' function passing the metafile handle. The recording stops (and the recorded commands are
' saved to the file) when the Graphics handle is deleted with GdipDeleteGraphics. To display
' the metafile, a new Graphics handle is retrieved calling the GdipCreateFromHDC function
' and passing the Graphics and Metafile handles to the GdipDrawImage function. Note that
' the code uses the same Metafile handle to record and to display (play back) the metafile.
' A metafile contains its own graphics state, which is defined by the Graphics object used
' to record the metafile. Any properties of the Graphics object (clip region, world
' transformation, smoothing mode, and the like) that you set while recording the metafile
' will be stored in the metafile. When you display the metafile, the drawing will be done
' according to those stored properties.
' ========================================================================================
SUB Example_RecordMetafileFileName (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a Metafile object for recording
   DIM metafile AS GpMetaFile PTR
   DIM wszFileName AS WSTRING * 64 = "MetafileTest.emf"
   hStatus = GdipRecordMetafileFileName(wszFileName, hdc, EmfTypeEmfPlusDual, NULL, _
             MetafileFrameUnitGdi, "", @metafile)
   IF hStatus <> StatusOk THEN EXIT SUB

   ' // Create a Graphics object that is associated with the Metafile object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipGetImageGraphicsContext(metafile, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96

'   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a pen and a solid brush
   DIM greenPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 1 * rxRatio, UnitPixel, @greenPen)
   ' // Add a rectangle and an ellipse to the metafile
   hStatus = GdipDrawRectangle(graphics, greenPen, 50, 10, 25, 75)
   hStatus = GdipDrawEllipse(graphics, greenPen, 100, 10, 25, 75)
   ' // Add an ellipse (drawn with antialiasing) to the metafile.
   hStatus = GdipSetSmoothingMode(graphics, SmoothingModeHighQuality)
   hStatus = GdipDrawEllipse(graphics, greenPen, 150, 10, 25, 75)
   ' // Add some text (drawn with antialiasing) to the metafile.
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   IF fontFamily THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF
   hStatus = GdipSetTextRenderingHint(graphics, TextRenderingHintAntiAlias)
   hStatus = GdipRotateWorldTransform(graphics, 30.0, MatrixOrderPrepend)
   DIM wszText AS WSTRING * 64 = "Smooth Text"
   DIM rcf AS GpRectF
   rcf.x = 50.0 : rcf.y = 50.0
   DIM solidBrush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)
   ' // Stop recording by deleting the Graphics object
   GdipDeleteGraphics(graphics)
   ' // Free resources
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF greenPen THEN GdipDeletePen(greenPen)

   ' // Create a Graphics object that is associated with the device context
   DIM playbackGraphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @playbackGraphics)
   IF playbackGraphics THEN
      ' // Display the metafile
       hStatus = GdipDrawImage(playbackGraphics, metafile, 60 * rxRatio, 10 * ryRatio)
      ' // Delete the Graphics object
      GdipDeleteGraphics(playbackGraphics)
   END IF

   ' // Dispose the Metafile object
   IF @metafile THEN GdipDisposeImage(metafile)

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
   pWindow.Create(NULL, "GDI+ GdipRecordMetafileFileName", @WndProc)
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
   Example_RecordMetafileFileName(hdc)

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
