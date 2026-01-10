' ########################################################################################
' Microsoft Windows
' File: GdipSetImageAttributesOutputChannel.bas
' Contents: GDI+ Flat API - GdipSetImageAttributesOutputChannel example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.bi"
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
' The following example creates an Image object and calls the GdipDrawImageRect function
' to draw the image. Then the code creates an ImageAttributes object and sets its bitmap
' output channel to cyan (ColorChannelFlagsC). The code calls GdipDrawImageRect a second
' time, passing the address of the Image object and the address of the ImageAttributes object.
' The cyan channel of each pixel is calculated, and the rendered image shows the intensities
' of the cyan channel as shades of gray. The code calls GdipDrawImageRect three more times
' to show the intensities of the magenta, yellow, and black channels.
' ========================================================================================
SUB Example_SetOutputChannel (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("Mosaic2.bmp", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Draw the image unaltered.
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Draw the image, showing the intensity of the cyan channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsC)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             110, 10, nWidth, nHeight, _   ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Draw the image, showing the intensity of the magenta channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsM)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             210, 10, nWidth, nHeight, _   ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Draw the image, showing the intensity of the yellow channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsY)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             10, 110, nWidth, nHeight, _   ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Draw the image, showing the intensity of the black channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsK)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             110, 110, nWidth, nHeight, _  ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
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
   pWindow.Create(NULL, "GDI+ GdipSetImageAttributesOutputChannel", @WndProc)
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
   Example_SetOutputChannel(hdc)

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
