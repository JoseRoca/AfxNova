' ########################################################################################
' Microsoft Windows
' File: CW_GDI_DrawingMarkers_01.bas
' You can use the line functions to draw markers. A marker is a symbol centered over a
' point. Drawing applications use markers to designate starting points, ending points,
' and control points. Spreadsheet applications use markers to designate points of interest
' on a chart or graph.
' In the following code sample, the application-defined Marker function creates a marker
' by using the MoveToEx and LineTo functions. These functions draw two intersecting lines,
' 20 pixels in length, centered over the cursor coordinates.
' The system stores the coordinates of the cursor in the lParam parameter of the
' WM_LBUTTONDOWN message when the user presses the left mouse button. The following code
' demonstrates how an application gets these coordinates, determines whether they lie
' within its client area, and passes them to the Marker function to draw the marker.
' https://learn.microsoft.com/en-us/windows/win32/gdi/drawing-markers?redirectedfrom=MSDN
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/AfxExt.bi"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - GDI - Drawing markers", @WndProc)
   pWindow.Brush = GetStockObject(WHITE_BRUSH)
   pWindow.WindowStyle = WS_POPUPWINDOW OR WS_CAPTION   ' // make the window not resizable
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(500, 300)
   ' // Centers the window
   pWindow.Center

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Draw the marker
' ========================================================================================
SUB Marker(BYVAL x AS LONG, BYVAL y AS LONG, BYVAL hwnd AS HWND)

   DIM dpiRatio AS UINT = GetDpiForWindow(hwnd) / 96
   DIM hdc AS HDC = GetDC(hwnd)
   MoveToEx hdc, x - (10 * dpiRatio), y, NULL
   LineTo hdc, x + (10 * dpiRatio), y
   MoveToEx hdc, x, y - (10 * dpiRatio), NULL
   LineTo hdc, x, y + (10 * dpiRatio)
   ReleaseDC hwnd, hdc

END SUB
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC ptMouseDown(32) AS POINT
   STATIC index AS LONG
   DIM  rc AS RECT
   DIM  hrgn AS HRGN

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

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

      CASE WM_LBUTTONDOWN
         IF index < 32 THEN
            ' // Create the region from the client area.
            GetClientRect hwnd, @rc
            hrgn = CreateRectRgn(rc.Left, rc.Top, rc.Right, rc.Bottom)
            ptMouseDown(index).x = LOWORD(lParam)
            ptMouseDown(index).y = HIWORD(lParam)
            ' // Test for a hit in the client rectangle.
            IF PtInRegion(hrgn, ptMouseDown(index).x, ptMouseDown(index).y) THEN
               ' // If a hit occurs, record the mouse coords.
               Marker ptMouseDown(index).x, ptMouseDown(index).y, hwnd
               index += 1
            END IF
         END IF
         EXIT FUNCTION

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
