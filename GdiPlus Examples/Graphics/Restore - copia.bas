' ########################################################################################
' Microsoft Windows
' File: Restore.bas
' Contents: GDI+ - Restore example
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
' Restoring Nested Saved States
' The following example sets the world transformation of a Graphics object to a rotation
' and then saves the state of the Graphics object. Next, the code calls TranslateTransform,
' and saves the state again. Then the code calls ScaleTransform. At that point, the world
' transformation of the Graphics object is a composite transformation: first rotate, then
' translate, then scale. The code uses a red pen to draw an ellipse that is transformed by
' that composite transformation.
' The code passes state2, which was returned by the second call to Save, to the Graphics.Restore
' method, and draws the ellipse again using a green pen. The green ellipse is rotated and
' translated but not scaled. Finally the code passes state1, which was returned by the
' first call to Save, to the Graphics.Restore method, and draws the ellipse again using a
' blue pen. The blue ellipse is rotated but not translated or scaled.
' ========================================================================================
SUB Example_Restore (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Three transformations apply: rotate, then translate, then scale.
   DIM AS GraphicsState state1, state2
   graphics.RotateTransform(30)
   state1 = graphics.Save
   graphics.TranslateTransform(100 * rxRatio, 0, MatrixOrderAppend)
   state2 = graphics.Save
   graphics.ScaleTransform(1, 3, MatrixOrderAppend)

   ' // Draw an ellipse.
   DIM redPen AS CGpPen = ARGB_RED
   graphics.DrawEllipse(@redPen, 0, 0, 100, 20)

   ' // Restore to state2 and draw the ellipse again.
   ' // Two transformations apply: rotate then translate.
   graphics.Restore(state2)
   DIM greenPen AS CGpPen = ARGB_GREEN
   graphics.DrawEllipse(@greenPen, 0, 0, 100, 20)

   ' // Restore to state1 and draw the ellipse again.
   ' // Only the rotation transformation applies.
   graphics.Restore(state1)
   DIM bluePen AS CGpPen = ARGB_BLUE
   graphics.DrawEllipse(@bluePen, 0, 0, 100, 20)

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
   pWindow.CreateOverlapped(NULL, "GDI+ Restore", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_WHITE
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_Restore(hdc)

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
