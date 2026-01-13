' ########################################################################################
' Microsoft Windows
' File: Gdip_GetStringFormatFlags.bas
' Contents: GDI+ Flat API - GdipGetStringFormatFlags example
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
' The following example creates a StringFormat object, sets the string's format flags, and
' then gets the 32-bit value that contains the format flags and stores it in a variable.
' The code then creates another StringFormat object and uses the stored format flags value
' to set the format flags of the second StringFormat object. Next, the code uses the second
' StringFormat object to draw a formatted string. The code also draws the string's layout
' rectangle.
' ========================================================================================
SUB Example_GetStringFormatFlags (BYVAL hdc AS HDC)

   DIM status AS GpStatus

   ' // Create a graphics object from the device context
   DIM graphics AS GdiPlusGraphics = hdc
   ' // Set the scale transform
   DIM dpiRatio AS SINGLE = graphics.DpiRatio
   status = graphics.ScaleTransform(dpiRatio)

   ' // Create the font
   DIM fontFamily AS GdiPlusFontFamily = "Times New Roman"
   DIm font AS GdiPlusFont = GdiPlusFont(*fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel)

   ' // Create a StringFormat object, and set its format flags.
   DIM format AS GdiPlusStringFormat = GdiPlusStringFormat(0, LANG_NEUTRAL)
   status = GdipSetStringFormatFlags(*format, StringFormatFlagsDirectionVertical OR StringFormatFlagsNoFitBlackBox)

   ' // Get the format flags from the StringFormat object.
   DIM flags AS LONG
   status = GdipGetStringFormatFlags(*format, @flags)

   ' // Create a second StringFormat object with the same flags.
   DIM format2 AS GdiPlusStringFormat = GdiPlusStringFormat(0, LANG_NEUTRAL)
   status = GdipSetStringFormatFlags(*format2, flags)

   ' // Create a solid brush
   DIM solidBrush AS GdiPlusSolidBrush = ARGB_RED

   ' // Draw the string using the second StringFormat object
   DIM wszText AS WSTRING * 64 = "This text is vertical because of a format flag."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   status = GdipDrawString(*graphics, wszText, LEN(wszText), *font, @rcf, *format2, *solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GdiPlusPen = GdiPlusPen(ARGB_RED, 1, UnitPixel)
   status = GdipScalePenTransform(*pen, graphics.dpiRatioX, graphics.dpiRatioY, MatrixOrderPrepend)
   status = GdipDrawRectangle(*graphics, *pen, rcf.x, rcf.y, rcf.width, rcf.height)

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
   pWindow.Create(NULL, "GDI+ GdipGetStringFormatFlags", @WndProc)
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
   Example_GetStringFormatFlags(hdc)

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
