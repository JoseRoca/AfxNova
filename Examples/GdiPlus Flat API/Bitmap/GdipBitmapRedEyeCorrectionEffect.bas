' ########################################################################################
' Microsoft Windows
' File: Gdip_RedEyeCorrectionEffect.bas
' Contents: GDI+ Flat API - Gdip_RedEyeCorrectionEffect example
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
' The RedEyeCorrection effect enables you to correct the red eyes that sometimes occur in
' flash photographs. To specify areas of the bitmap that have red eyes, pass a 
' RedEyeCorrectionParams structure to the GdipSetEffectParameters function.
' ========================================================================================
SUB Example_Gdip_RedEyeCorrectionEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("RedEyes.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a levels effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(RedEyeCorrectionEffectGuid, @effect)

   ' // We need to specify one or more rectangles that enclose the red-eye areas.
   ' // These are passed as an array of RECT structures.

   ' // Define two rectangles around the eyes
   ' // Change the values according the coordinates of your image
   DIM eyeRects(0 TO 1) AS RECT
   eyeRects(0).left   = 1 : eyeRects(0).top    = 1
   eyeRects(0).right  = 400 : eyeRects(0).bottom = 250
   eyeRects(1).left   = 1 : eyeRects(1).top    = 1
   eyeRects(1).right  = 400 : eyeRects(1).bottom = 250

   ' // Fill a RedEyeCorrectionParams structure
   DIM redeyeParams AS RedEyeCorrectionParams
   redeyeParams.numberOfAreas = 2
   redeyeParams.areas = @eyeRects(0)
   ' // Calculate the size of the paameters
   DIM paramsSize AS UINT = SIZEOF(RedEyeCorrectionParams) + redeyeParams.numberOfAreas * SIZEOF(RECT)
   ' // Set the parameters
   hStatus = GdipSetEffectParameters(effect, @redeyeParams, paramsSize)

   ' // Apply effect
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

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
   pWindow.Create(NULL, "GDI+ Gdip_RedEyeCorrectionEffect", @WndProc)
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
   Example_Gdip_RedEyeCorrectionEffect(hdc)

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
