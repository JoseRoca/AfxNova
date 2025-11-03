' ########################################################################################
' Microsoft Windows
' File: ImageAttributesSetColorMatrices.bas
' Contents: GDI+ - ImageAttributesSetColorMatrices example
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
' This example sets the color-adjustment matrix and the grayscale-adjustment matrix for a specified category.
' SetColorMatrices lets you apply both a primary color matrix and a grayscale matrix in one call.
' If you only want one effect, you can pass NULL for the second matrix.
' This is ideal for layered effects like sepia + desaturation, or tint + gamma correction.
' ========================================================================================
SUB Example_SetColorMatrices (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create an Image object based on a .bmp file.
   ' // The image has one stripe with RGB components (160, 0, 0)
   ' // and one stripe with RGB components (0, 140, 0).
   DIM image AS CGpBitmap = "climber.jpg"
   image.SetResolutionForDpi

   ' // Get the width and height of the image
   DIM nWidth AS UINT = image.GetWidth
   DIM nHeight AS UINT = image.GetHeight

   ' // Create an ImageAttributes object, and set its bitmap threshold to 0.6.
   DIM imgAttr AS CGpImageAttributes

   ' Define sepia color matrix
   DIM sepiaMatrix AS ColorMatrix
   sepiaMatrix.m(0,0) = 0.393 : sepiaMatrix.m(0,1) = 0.349 : sepiaMatrix.m(0,2) = 0.272
   sepiaMatrix.m(1,0) = 0.769 : sepiaMatrix.m(1,1) = 0.686 : sepiaMatrix.m(1,2) = 0.534
   sepiaMatrix.m(2,0) = 0.189 : sepiaMatrix.m(2,1) = 0.168 : sepiaMatrix.m(2,2) = 0.131
   sepiaMatrix.m(3,3) = 1
   sepiaMatrix.m(4,4) = 1

   ' Define grayscale matrix (optional, for layering)
   DIM grayMatrix AS ColorMatrix
   grayMatrix.m(0,0) = 0.3 : grayMatrix.m(0,1) = 0.3 : grayMatrix.m(0,2) = 0.3
   grayMatrix.m(1,0) = 0.59: grayMatrix.m(1,1) = 0.59: grayMatrix.m(1,2) = 0.59
   grayMatrix.m(2,0) = 0.11: grayMatrix.m(2,1) = 0.11: grayMatrix.m(2,2) = 0.11
   grayMatrix.m(3,3) = 1
   grayMatrix.m(4,4) = 1

   ' // Apply the color matrix
   imgAttr.SetColorMatrices(@sepiaMatrix, @grayMatrix)

   ' // Draw original image
   graphics.DrawImage(@image, 10, 10, nWidth, nHeight)

   ' // Draw image with transparency applied
   DIM rc AS GpRect = (200, 10, nWidth, nHeight)
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
   pWindow.Create(NULL, "GDI+ ImageAttributesSetColorMatrices", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(390, 250)
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
   Example_SetColorMatrices(hdc)

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
