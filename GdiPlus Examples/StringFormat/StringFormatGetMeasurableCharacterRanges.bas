' ########################################################################################
' Microsoft Windows
' File: StringFormatGetMeasurableCharacterRanges.bas
' Contents: GDI+ - StringFormatGetMeasurableCharacterRanges example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
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
' The following example defines three ranges of character positions within a string and sets
' those ranges in a StringFormat object. Next, the StringFormat::GetMeasurableCharacterRangeCount
' method is used to get the number of character ranges that are currently set in the StringFormat
' object. This number is then used to allocate a buffer large enough to store the regions
' that correspond with the ranges. Then, the MeasureCharacterRanges method is used to get
' the three regions of the display that are occupied by the characters that are specified
' by the ranges.
' Remarks: It doesn't work with the 64-bit headers because they lack a declare for the
' GdipMeasureCharacterRanges function.
' ========================================================================================
SUB Example_GetMeasurableCharacterRanges (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Brushes and pens used for drawing and painting
   DIM blueBrush AS CGpSOlidBrush = GDIP_ARGB(255, 0, 0, 255)
   DIM redBrush AS CGpSOlidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM blackPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)

   ' // Layout rectangles used for drawing strings
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(20.0, 20.0, 130.0, 130.0)

   ' // Three ranges of character positions within the string
   DIM charRanges(2) AS CharacterRange
   charRanges(0).First = 3  : charRanges(0).Length = 5
   charRanges(1).First = 15 : charRanges(1).Length = 2
   charRanges(2).First = 30 : charRanges(2).Length = 15

   ' // Font and string format used to apply to string when drawing
   DIM myFont AS CGpFont = CGpFont("Times New Roman", AfxPointsToPixelsX(16) / rxRatio, FontStyleRegular, UnitPixel)
   DIM strFormat AS CGpStringFormat

   DIM wszText AS WSTRING * 260
   wszText = "The quick, brown fox easily jumps over the lazy dog."

   ' // Set three ranges of character positions.
   strFormat.SetMeasurableCharacterRanges(3, @charRanges(0))

   ' // Get the number of ranges that have been set, and allocate memory to
   ' // store the regions that correspond to the ranges.
   DIM nCount AS LONG = strFormat.GetMeasurableCharacterRangeCount
   DIM rgCharRangeRegions(nCount - 1) AS CGpRegion

   ' // Get the regions that correspond to the ranges within the string
   graphics.MeasureCharacterRanges(@wszText, -1, @myFont, @layoutRect, @strFormat, nCount, @rgCharRangeRegions(0))
   graphics.DrawString(@wszText, -1, @myFont, @layoutRect, @strFormat, @blueBrush)
   graphics.DrawRectangle(@blackPen, @layoutRect)
   FOR i AS LONG = 0 TO nCount - 1
      graphics.FillRegion(@redBrush, @rgCharRangeRegions(i))
   NEXT

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
   pWindow.CreateOverlapped(NULL, "GDI+ GetMeasurableCharacterRanges", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(430, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_WHITE
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_GetMeasurableCharacterRanges(hdc)

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
