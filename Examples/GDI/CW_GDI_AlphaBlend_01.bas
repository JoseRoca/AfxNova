' ########################################################################################
' Microsoft Windows
' File: CW_GDI_AlphaBlend_01.bas
' Contents: CWindow - Alpha Blending a Bitmap
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' The following code sample divides a window into three horizontal areas. Then it draws an
' alpha-blended bitmap in each of the window areas as follows:
'    * In the top area, constant alpha = 50% but there is no source alpha.
'    * In the middle area, constant alpha = 100% (disabled) and source alpha is 0
'      (transparent) in the middle of the bitmap and 0xff (opaque) elsewhere.
'    * In the bottom area, constant alpha = 75% and source alpha changes.
' ========================================================================================
SUB DrawAlphaBlend (BYVAL hwnd AS HWND, BYVAL hDCwnd AS HDC)

   ' // Get window dimensions
   DIM rc AS RECT
   GetClientRect(hwnd, @rc)

   ' // Calculate window width/height
   DIM ulWindowWidth AS LONG = rc.right - rc.left
   DIM ulWindowHeight AS LONG = rc.bottom - rc.top

   ' // Make sure we have at least some window size
   IF ulWindowWidth = 0 OR ulWindowHeight = 0 THEN EXIT SUB

   ' // Divide the window into 3 horizontal areas
   ulWindowHeight \= 3

   ' Create a DC for our bitmap -- the source DC for AlphaBlend
   DIM hdc AS HDC = CreateCompatibleDC(hDCwnd)

   ' // Setup bitmap info
   ' // Set the bitmap width and height to 60% of the width and height of each of
   ' // the three horizontal areas. Later on, the blending will occur in the center
   ' // of each of the three areas.
   DIM ulBitmapWidth AS LONG = ulWindowWidth  - (ulWindowWidth  \ 5) * 2
   DIM ulBitmapHeight AS LONG = ulWindowHeight - (ulWindowHeight \ 5) * 2

   ' // Setup BITMAPINFO
   DIM bmi AS BITMAPINFO
   bmi.bmiHeader.biSize = SIZEOF(BITMAPINFOHEADER)
   bmi.bmiHeader.biWidth = ulBitmapWidth
   bmi.bmiHeader.biHeight = -ulBitmapHeight   ' top-down
   bmi.bmiHeader.biPlanes = 1
   bmi.bmiHeader.biBitCount = 32
   bmi.bmiHeader.biCompression = BI_RGB
   bmi.bmiHeader.biSizeImage = 0

   ' // Create a DIB section
   DIM hbitmap AS HBITMAP
   DIM pvBits AS ANY PTR   ' pointer to DIB section
   hbitmap = CreateDIBSection(hdc, @bmi, DIB_RGB_COLORS, @pvBits, 0, 0)
   IF hbitmap = 0 OR pvBits = 0 THEN
      DeleteDC(hdc)
      EXIT SUB
   END IF
   ' // Select the bitmap into the device context
   SelectObject(hdc, hbitmap)

   ' // Get real stride and pixel pointer
   DIM bmp AS BITMAP
   GetObject(hbitmap, SIZEOF(bmp), @bmp)
   DIM pPixels AS ULONG PTR
   DIM stride AS LONG
   pPixels = cast(ULONG PTR, bmp.bmBits)
   stride = bmp.bmWidthBytes \ 4   ' stride in pixels

   ' ---------------------------------------------------------
   ' Top area
   ' in top window area, constant alpha = 50%, but no source alpha
   ' the color format for each pixel is &Haarrggbb
   ' set all pixels to blue and set source alpha to zero
   ' ---------------------------------------------------------
   FOR y AS LONG = 0 TO ulBitmapHeight - 1
      FOR x AS LONG = 0 To ulBitmapWidth - 1
         pPixels[x + y * stride] = &H000000FF
      NEXT
   NEXT

   DIM bf AS BLENDFUNCTION         ' structure for alpha blending
   bf.BlendOp = AC_SRC_OVER
   bf.BlendFlags = 0
   bf.SourceConstantAlpha = &H7F   ' half of &HFF = 50% transparency
   bf.AlphaFormat = 0              ' ignore source alpha channel

    AlphaBlend(hDCwnd, ulWindowWidth\5, ulWindowHeight\5, _
               ulBitmapWidth, ulBitmapHeight, hdc, 0, 0, _
               ulBitmapWidth, ulBitmapHeight, bf)

   ' ---------------------------------------------------------
   ' Middle area
   ' In middle window area, constant alpha = 100% (disabled), source
   ' alpha is 0 in middle of bitmap and opaque in rest of bitmap
    ' ---------------------------------------------------------
   FOR y AS LONG = 0 TO ulBitmapHeight - 1
      FOR x AS LONG = 0 To ulBitmapWidth - 1
         IF (x > ulBitmapWidth \ 5) AND (x < ulBitmapWidth - ulBitmapWidth \ 5) AND _
            (y > ulBitmapHeight \ 5) AND (y < ulBitmapHeight - ulBitmapHeight \ 5) THEN
            pPixels[x + y * stride] = &H00FF0000
         ELSE
            pPixels[x + y * stride] = &HFF0000FF
         END IF
      NEXT
   NEXT

   bf.AlphaFormat = AC_SRC_ALPHA
   bf.SourceConstantAlpha = &HFF

   AlphaBlend(hDCwnd, ulWindowWidth\5, ulWindowHeight\5 + ulWindowHeight, _
              ulBitmapWidth, ulBitmapHeight, hdc, 0, 0, _
              ulBitmapWidth, ulBitmapHeight, bf)

   ' ---------------------------------------------------------
   ' Bottom area
   ' Bottom window area, use constant alpha = 75% and a changing
   ' source alpha. Create a gradient effect using source alpha, and
   ' then fade it even more with constant alpha
   ' ---------------------------------------------------------
   DIM ubAlpha AS UBYTE          ' used for doing transparent gradient
   DIM ubRed AS UBYTE
   DIM ubGreen AS UBYTE
   DIM ubBlue AS UBYTE
   DIM fAlphaFactor AS SINGLE    ' used to do premultiply
   DIM f AS SINGLE

   FOR y AS LONG = 0 TO ulBitmapHeight - 1
      FOR x AS LONG = 0 TO ulBitmapWidth - 1
         ' For a simple gradient, base the alpha value on the x value of the pixel
         f = (x / ulBitmapWidth) * 255
         ubAlpha = CBYTE(f)
         ' // Calculate the factor by which we multiply each component
         fAlphaFactor = ubAlpha / &HFF
         ubBlue = CBYTE(&HFF * fAlphaFactor)
         pPixels[x + y * stride] = _
                (ubAlpha SHL 24) OR _
                (ubRed   SHL 16) OR _
                (ubGreen SHL 8)  OR _
                ubBlue

      NEXT
   NEXT

   bf.AlphaFormat = 0
   bf.SourceConstantAlpha = &HBF   ' use constant alpha, with 75% opaqueness

   AlphaBlend(hDCwnd, ulWindowWidth\5, ulWindowHeight\5 + 2 * ulWindowHeight, _
              ulBitmapWidth, ulBitmapHeight, hdc, 0, 0, _
              ulBitmapWidth, ulBitmapHeight, bf)

    DeleteObject(hbitmap)
    DeleteDC(hdc)

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

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - GDI - AlphaBlend Demo", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(500, 320)
   ' // Centers the window
   pWindow.Center

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
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

      CASE WM_PAINT
         DIM ps AS PAINTSTRUCT
         DIM hdc AS HDC = BeginPaint(hwnd, @ps)
         DrawAlphaBlend hwnd, hdc
         EndPaint hwnd, @ps
         RETURN 0

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
