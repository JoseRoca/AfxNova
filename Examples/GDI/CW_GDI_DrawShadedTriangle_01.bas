' ########################################################################################
' Microsoft Windows
' File: CW_GDI_DrawShadedTriangle_01.bas
' Contents: Drawing a Shaded Triangle
' To draw a shaded triangle, define a TRIVERTEX structure with three elements and a single
' GRADIENT_TRIANGLE structure. The following code example shows how to draw a shaded
' triangle using the GradientFill function with the GRADIENT_FILL_TRIANGLE mode defined.
' Compiler: FreeBasic 32 & 64 bit
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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - GDI - Drawing a shaded triangle", @WndProc)
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
' Draws a shaded triangle using the GradientFill function.
' ========================================================================================
SUB DrawShadedTriangle (BYVAL hdc AS HDC, BYVAL rxRatio AS SINGLE, BYVAL ryRatio AS SINGLE)

   ' // Create an array of TRIVERTEX structures that describe
   ' // positional and color values for each vertex. For a rectangle,
   ' // only two vertices need to be defined: upper-left and lower-right.

   DIM vertex(2) AS TRIVERTEX

   vertex(0).x     = 250 * rxRatio
   vertex(0).y     = 50 * ryRatio
   vertex(0).Red   = &Hff00
   vertex(0).Green = &H8000
   vertex(0).Blue  = &H0000
   vertex(0).Alpha = &H0000

   vertex(1).x     = 50 * rxRatio
   vertex(1).y     = 250 * ryRatio
   vertex(1).Red   = &H9000
   vertex(1).Green = &H0000
   vertex(1).Blue  = &H9000
   vertex(1).Alpha = &H0000

   vertex(2).x     = 450 * rxRatio
   vertex(2).y     = 250 * ryRatio
   vertex(2).Red   = &h9000
   vertex(2).Green = &h0000
   vertex(2).Blue  = &h9000
   vertex(2).Alpha = &h0000

   ' // Create a GRADIENT_TRIANGLE structure that
   ' // references the TRIVERTEX vertices.

   DIM gTriangle AS GRADIENT_TRIANGLE
   gTriangle.Vertex1 = 0
   gTriangle.Vertex2 = 1
   gTriangle.Vertex3 = 2

   ' // Draw a shaded triangle.
   GradientFill hDc, @vertex(0), 3, @gTriangle, 1, GRADIENT_FILL_TRIANGLE

END SUB
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC rxRatio AS SINGLE     ' // Horizontal scaling ratio
   STATIC ryRatio AS SINGLE     ' // Vertical scaling ratio

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         ' // Get a pointer to the CWindow class from the CREATESTRUCT structure
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
         ' // Get the scaling ratios
         rxRatio = pWindow->rxRatio
         ryRatio = pWindow->ryRatio
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

      CASE WM_PAINT
         DIM ps AS PAINTSTRUCT
         DIM hdc AS HDC = BeginPaint(hwnd, @ps)
         DrawShadedTriangle hdc, rxRatio, ryRatio
         EndPaint hwnd, @ps
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
