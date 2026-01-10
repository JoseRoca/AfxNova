' ########################################################################################
' Microsoft Windows
' File: ImageAttributesSetOutputChannel.bas
' Contents: GDI+ - ImageAttributesSetOutputChannel example
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
' The following example creates an Image object and calls the DrawImage method to draw the
' image. Then the code creates an ImageAttributes object and sets its bitmap output channel
' to cyan (ColorChannelFlagsC). The code calls DrawImage a second time, passing the address
' of the Image object and the address of the ImageAttributes object. The cyan channel of
' each pixel is calculated, and the rendered image shows the intensities of the cyan channel
' as shades of gray. The code calls DrawImage three more times to show the intensities of
' the magenta, yellow, and black channels.
' ========================================================================================
SUB Example_SetOutputChannel (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create an Image object based on a .bmp file.
   ' // The image has one stripe with RGB components (160, 0, 0)
   ' // and one stripe with RGB components (0, 140, 0).
   DIM image AS CGpBitmap = "Mosaic2.bmp"
   image.SetResolutionForDpi

   ' // Draw the image unaltered.
   DIM nWidth AS UINT = image.GetWidth
   DIM nHeight AS UINT = image.GetHeight
   graphics.DrawImage(@image, 10, 10, nWidth, nHeight)

   ' // Create an ImageAttributes object, and set its bitmap threshold to 0.6.
   DIM imgAttr AS CGpImageAttributes

   ' // Draw the image, showing the intensity of the cyan channel.
   imgAttr.SetOutputChannel(ColorChannelFlagsC)
   DIM rc AS GpRect = (110, 10, nWidth, nHeight)
   graphics.DrawImage(@image, @rc, 0, 0, nWidth, nHeight, UnitPixel, @imgAttr)

   ' // Draw the image, showing the intensity of the magenta channel.
   imgAttr.SetOutputChannel(ColorChannelFlagsM)
   rc = Type<GpRect>(210, 10, nWidth, nHeight)
   graphics.DrawImage(@image, @rc, 0, 0, nWidth, nHeight, UnitPixel, @imgAttr)
 
    ' // Draw the image, showing the intensity of the yellow channel.
  imgAttr.SetOutputChannel(ColorChannelFlagsY)
   rc = Type<GpRect>(10, 110, nWidth, nHeight)
   graphics.DrawImage(@image, @rc, 0, 0, nWidth, nHeight, UnitPixel, @imgAttr)

    ' // Draw the image, showing the intensity of the black channel.
  imgAttr.SetOutputChannel(ColorChannelFlagsK)
   rc = Type<GpRect>(110, 110, nWidth, nHeight)
   graphics.DrawImage(@image, @rc, 0, 0, nWidth, nHeight, UnitPixel, @imgAttr)

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
   pWindow.Create(NULL, "GDI+ ImageAttributesSetOutputChannel", @WndProc)
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
   Example_SetOutputChannel(hdc)

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
